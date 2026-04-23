"""
广告联盟Offer管理工具
支持从PartnerBoost获取offer信息，并与Google Ads和飞书表格联动管理
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import requests
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import threading
import os
import re
import time
from google.ads.googleads.client import GoogleAdsClient
from google.oauth2 import service_account


# 硬编码API基础URL
PB_API_BASE_URL = "https://app.partnerboost.com"
FEISHU_API_BASE_URL = "https://open.feishu.cn/open-apis"

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

# 日志文件路径
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
LOG_FILE = os.path.join(LOG_DIR, "app.log")
DEBUG_LOG_FILE = os.path.join(LOG_DIR, "debug.log")  # 调试日志文件
LOG_MAX_SIZE_MB = 10  # 日志文件最大大小（MB）
LOG_KEEP_DAYS = 7     # 保留最近几天的日志

# 旧广告系列花费记录文件
OLD_CAMPAIGN_CONSUME_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "old_campaign_consume.xlsx")

# 广告系列花费快照文件（用于追踪每个广告系列的花费变化）
CAMPAIGN_COST_SNAPSHOT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "campaign_cost_snapshot.xlsx")


class OfferToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("广告联盟Offer管理工具")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # 初始化日志系统
        self.init_log_system()
        
        # 设置样式
        style = ttk.Style()
        style.configure('TLabel', font=('Microsoft YaHei', 10))
        style.configure('TButton', font=('Microsoft YaHei', 10))
        style.configure('Header.TLabel', font=('Microsoft YaHei', 12, 'bold'))
        
        # 加载配置
        self.config = self.load_config()
        
        # 创建Notebook（选项卡）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建各个选项卡
        self.create_offer_get_tab()
        self.create_offer_manage_tab()
        self.create_config_tab()
    
    def init_log_system(self):
        """初始化日志系统"""
        # 创建日志目录
        if not os.path.exists(LOG_DIR):
            os.makedirs(LOG_DIR)
        
        # 清理旧日志
        self.clean_old_logs()
        
        # 清空调试日志（每次启动时重新开始）
        try:
            with open(DEBUG_LOG_FILE, 'w', encoding='utf-8') as f:
                f.write(f"=== 调试日志 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
        except Exception:
            pass
        
        # 写入启动日志
        self.write_log_to_file(f"\n{'='*60}")
        self.write_log_to_file(f"程序启动 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.write_log_to_file(f"{'='*60}")
    
    def write_log_to_file(self, message):
        """写入日志到文件"""
        try:
            # 检查文件大小，如果超过限制则轮转
            if os.path.exists(LOG_FILE):
                size_mb = os.path.getsize(LOG_FILE) / (1024 * 1024)
                if size_mb > LOG_MAX_SIZE_MB:
                    self.rotate_log_file()
            
            # 写入日志
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {message}\n")
        except Exception as e:
            print(f"写入日志失败: {e}")
    
    def rotate_log_file(self):
        """轮转日志文件"""
        try:
            if os.path.exists(LOG_FILE):
                # 重命名为带时间戳的备份文件
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_file = os.path.join(LOG_DIR, f"app_{timestamp}.log")
                os.rename(LOG_FILE, backup_file)
        except Exception as e:
            print(f"轮转日志失败: {e}")
    
    def clean_old_logs(self):
        """清理旧日志文件"""
        try:
            if not os.path.exists(LOG_DIR):
                return
            
            cutoff_date = datetime.now() - timedelta(days=LOG_KEEP_DAYS)
            
            for filename in os.listdir(LOG_DIR):
                if filename.endswith('.log'):
                    filepath = os.path.join(LOG_DIR, filename)
                    # 获取文件修改时间
                    file_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                    if file_time < cutoff_date:
                        os.remove(filepath)
                        print(f"已清理旧日志: {filename}")
        except Exception as e:
            print(f"清理日志失败: {e}")
        
    def load_config(self):
        """加载配置文件"""
        default_config = {
            # PartnerBoost配置
            "pb_token": "5wXyGERfQ3rEQTdI",
            # 飞书配置
            "feishu_app_id": "cli_a8517363cf3bd013",
            "feishu_app_secret": "O4Sm3UNHjpykF9OZq3LroblsrCVYyQEp",
            "feishu_spreadsheet_token": "KnJ1wphpBiVMrGkWl5ncUkMGnfe",
            "feishu_sheet_id": "kPlW5z",
            # Google Ads配置
            "google_developer_token": "1YsRjWGxV6XUdxtX8MiT3Q",
            "google_mcc_id": "6885177935",
            "google_service_account_file": "credentials/google_ads_service_account.json",
            # 保存路径
            "save_path": os.path.join(os.path.expanduser("~"), "Desktop", "offers.xlsx")
        }
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    default_config.update(saved_config)
        except Exception:
            pass
        return default_config
    
    def save_config(self):
        """保存配置文件"""
        try:
            config_to_save = {
                "pb_token": self.pb_token_var.get().strip(),
                "feishu_app_id": self.feishu_app_id_var.get().strip(),
                "feishu_app_secret": self.feishu_app_secret_var.get().strip(),
                "feishu_spreadsheet_token": self.feishu_spreadsheet_var.get().strip(),
                "feishu_sheet_id": self.feishu_sheet_id_var.get().strip(),
                "google_developer_token": self.google_dev_token_var.get().strip(),
                "google_mcc_id": self.google_mcc_id_var.get().strip(),
                "google_service_account_file": self.google_sa_file_var.get().strip(),
                "save_path": self.save_path_var.get()
            }
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")

    # ==================== Offer获取选项卡 ====================
    def create_offer_get_tab(self):
        """创建Offer获取选项卡"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="  Offer获取  ")
        
        main_frame = ttk.Frame(tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # API参数配置区域
        api_frame = ttk.LabelFrame(main_frame, text="API参数配置", padding="10")
        api_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 使用Grid布局，2列
        ttk.Label(api_frame, text="国家代码:", width=15, anchor=tk.W).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.country_code_var = tk.StringVar(value="")
        ttk.Entry(api_frame, textvariable=self.country_code_var, width=20).grid(row=0, column=1, sticky=tk.W, pady=5, padx=(0, 30))
        
        ttk.Label(api_frame, text="品牌ID:", width=15, anchor=tk.W).grid(row=0, column=2, sticky=tk.W, pady=5)
        self.brand_id_var = tk.StringVar(value="")
        ttk.Entry(api_frame, textvariable=self.brand_id_var, width=20).grid(row=0, column=3, sticky=tk.W, pady=5)
        
        ttk.Label(api_frame, text="ASIN(逗号分隔):", width=15, anchor=tk.W).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.asins_var = tk.StringVar(value="")
        ttk.Entry(api_frame, textvariable=self.asins_var, width=20).grid(row=1, column=1, sticky=tk.W, pady=5, padx=(0, 30))
        
        ttk.Label(api_frame, text="排序方式:", width=15, anchor=tk.W).grid(row=1, column=2, sticky=tk.W, pady=5)
        self.sort_var = tk.StringVar(value="")
        ttk.Entry(api_frame, textvariable=self.sort_var, width=20).grid(row=1, column=3, sticky=tk.W, pady=5)
        
        ttk.Label(api_frame, text="Relationship:", width=15, anchor=tk.W).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.relationship_var = tk.StringVar(value="1")
        ttk.Combobox(api_frame, textvariable=self.relationship_var, values=["0", "1"], width=17).grid(row=2, column=1, sticky=tk.W, pady=5, padx=(0, 30))
        
        # 复选框
        self.default_filter_var = tk.IntVar(value=0)
        ttk.Checkbutton(api_frame, text="默认筛选", variable=self.default_filter_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.is_original_currency_var = tk.IntVar(value=0)
        ttk.Checkbutton(api_frame, text="原始货币", variable=self.is_original_currency_var).grid(row=3, column=2, columnspan=2, sticky=tk.W, pady=5)
        
        self.has_promo_code_var = tk.IntVar(value=0)
        ttk.Checkbutton(api_frame, text="有促销码", variable=self.has_promo_code_var).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.has_acc_var = tk.IntVar(value=0)
        ttk.Checkbutton(api_frame, text="有ACC佣金", variable=self.has_acc_var).grid(row=4, column=2, columnspan=2, sticky=tk.W, pady=5)
        
        self.filter_sexual_wellness_var = tk.IntVar(value=0)
        ttk.Checkbutton(api_frame, text="过滤成人用品", variable=self.filter_sexual_wellness_var).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # UID追踪参数
        ttk.Label(api_frame, text="UID模式:", width=15, anchor=tk.W).grid(row=6, column=0, sticky=tk.W, pady=5)
        self.uid_mode_var = tk.StringVar(value="随机")
        uid_mode_combo = ttk.Combobox(api_frame, textvariable=self.uid_mode_var, values=["无", "随机", "固定"], width=17, state='readonly')
        uid_mode_combo.grid(row=6, column=1, sticky=tk.W, pady=5, padx=(0, 30))
        uid_mode_combo.bind('<<ComboboxSelected>>', self._on_uid_mode_changed)
        
        ttk.Label(api_frame, text="固定UID:", width=15, anchor=tk.W).grid(row=6, column=2, sticky=tk.W, pady=5)
        self.fixed_uid_var = tk.StringVar(value="")
        self.fixed_uid_entry = ttk.Entry(api_frame, textvariable=self.fixed_uid_var, width=20, state='disabled')
        self.fixed_uid_entry.grid(row=6, column=3, sticky=tk.W, pady=5)
        
        # 保存配置区域
        save_frame = ttk.LabelFrame(main_frame, text="保存配置", padding="10")
        save_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(save_frame, text="保存路径:").pack(side=tk.LEFT)
        self.save_path_var = tk.StringVar(value=self.config.get("save_path", ""))
        save_entry = ttk.Entry(save_frame, textvariable=self.save_path_var, width=50)
        save_entry.pack(side=tk.LEFT, padx=(10, 10))
        ttk.Button(save_frame, text="浏览...", command=self.browse_save_path).pack(side=tk.LEFT)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.extract_btn = ttk.Button(button_frame, text="提取Offer", command=self.extract_offers)
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.copy_offer_btn = ttk.Button(button_frame, text="复制Offer", command=self.copy_offers)
        self.copy_offer_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.brand_search_btn = ttk.Button(button_frame, text="获取品牌搜索量报告", command=self.get_brand_search_volume)
        self.brand_search_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.progress_label_get = ttk.Label(button_frame, text="")
        self.progress_label_get.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.progress_get = ttk.Progressbar(button_frame, mode='indeterminate', length=150)
        self.progress_get.pack(side=tk.RIGHT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text_get = scrolledtext.ScrolledText(log_frame, height=12, font=('Consolas', 9))
        self.log_text_get.pack(fill=tk.BOTH, expand=True)

    # ==================== Offer管理选项卡 ====================
    def create_offer_manage_tab(self):
        """创建Offer管理选项卡"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="  Offer管理  ")
        
        main_frame = ttk.Frame(tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 说明区域
        info_frame = ttk.LabelFrame(main_frame, text="功能说明", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        info_text = """• "开始统计"：获取Google Ads/PB数据，更新offer状态、广告系列数量、花费、佣金
• "更新已有offer"：按品牌批量刷新已有offer信息
• "offer顺序整理"：对Offer表和广告系列表排序（按状态分组+总佣金降序）"""
        ttk.Label(info_frame, text=info_text, justify=tk.LEFT).pack(anchor=tk.W)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.start_stats_btn = ttk.Button(button_frame, text="开始统计", command=self.start_statistics)
        self.start_stats_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = ttk.Button(button_frame, text="停止", command=self.stop_statistics, state='disabled')
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.update_offers_btn = ttk.Button(button_frame, text="更新已有offer", command=self.start_update_offers)
        self.update_offers_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.sort_tables_btn = ttk.Button(button_frame, text="offer顺序整理", command=self.start_sort_tables)
        self.sort_tables_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.progress_label_manage = ttk.Label(button_frame, text="")
        self.progress_label_manage.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.progress_manage = ttk.Progressbar(button_frame, mode='indeterminate', length=200)
        self.progress_manage.pack(side=tk.RIGHT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text_manage = scrolledtext.ScrolledText(log_frame, height=20, font=('Consolas', 9))
        self.log_text_manage.pack(fill=tk.BOTH, expand=True)
        
        # 停止标志
        self.stop_flag = False

    # ==================== 配置选项卡 ====================
    def create_config_tab(self):
        """创建配置选项卡"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="  API配置  ")
        
        main_frame = ttk.Frame(tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # PartnerBoost配置
        pb_frame = ttk.LabelFrame(main_frame, text="PartnerBoost配置", padding="10")
        pb_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(pb_frame, text="API Token:", width=20, anchor=tk.W).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.pb_token_var = tk.StringVar(value=self.config.get("pb_token", ""))
        ttk.Entry(pb_frame, textvariable=self.pb_token_var, width=50).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 飞书配置
        feishu_frame = ttk.LabelFrame(main_frame, text="飞书配置", padding="10")
        feishu_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(feishu_frame, text="App ID:", width=20, anchor=tk.W).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.feishu_app_id_var = tk.StringVar(value=self.config.get("feishu_app_id", ""))
        ttk.Entry(feishu_frame, textvariable=self.feishu_app_id_var, width=50).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(feishu_frame, text="App Secret:", width=20, anchor=tk.W).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.feishu_app_secret_var = tk.StringVar(value=self.config.get("feishu_app_secret", ""))
        ttk.Entry(feishu_frame, textvariable=self.feishu_app_secret_var, width=50, show="*").grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(feishu_frame, text="电子表格Token:", width=20, anchor=tk.W).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.feishu_spreadsheet_var = tk.StringVar(value=self.config.get("feishu_spreadsheet_token", ""))
        ttk.Entry(feishu_frame, textvariable=self.feishu_spreadsheet_var, width=50).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(feishu_frame, text="工作表ID:", width=20, anchor=tk.W).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.feishu_sheet_id_var = tk.StringVar(value=self.config.get("feishu_sheet_id", ""))
        ttk.Entry(feishu_frame, textvariable=self.feishu_sheet_id_var, width=50).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Google Ads配置
        google_frame = ttk.LabelFrame(main_frame, text="Google Ads配置", padding="10")
        google_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(google_frame, text="Developer Token:", width=20, anchor=tk.W).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.google_dev_token_var = tk.StringVar(value=self.config.get("google_developer_token", ""))
        ttk.Entry(google_frame, textvariable=self.google_dev_token_var, width=50).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(google_frame, text="MCC ID:", width=20, anchor=tk.W).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.google_mcc_id_var = tk.StringVar(value=self.config.get("google_mcc_id", ""))
        ttk.Entry(google_frame, textvariable=self.google_mcc_id_var, width=50).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(google_frame, text="服务账号文件:", width=20, anchor=tk.W).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.google_sa_file_var = tk.StringVar(value=self.config.get("google_service_account_file", ""))
        ttk.Entry(google_frame, textvariable=self.google_sa_file_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Button(google_frame, text="浏览", command=self.browse_sa_file).grid(row=2, column=2, sticky=tk.W, pady=5, padx=5)
        
        # 保存按钮
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="保存配置", command=self.save_and_notify).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="测试连接", command=self.test_connections).pack(side=tk.LEFT, padx=10)

    # ==================== 通用方法 ====================
    def browse_save_path(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            title="选择保存位置"
        )
        if file_path:
            self.save_path_var.set(file_path)
    
    def browse_sa_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            title="选择服务账号文件"
        )
        if file_path:
            self.google_sa_file_var.set(file_path)
    
    def save_and_notify(self):
        self.save_config()
        messagebox.showinfo("成功", "配置已保存！")
    
    def log_get(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text_get.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text_get.see(tk.END)
        self.root.update_idletasks()
        # 同时写入日志文件
        self.write_log_to_file(f"[Offer获取] {message}")
    
    def log_manage(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text_manage.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text_manage.see(tk.END)
        self.root.update_idletasks()
        # 同时写入日志文件
        self.write_log_to_file(message)
    
    def log_debug(self, message):
        """写入调试日志（仅写入文件，不显示在GUI中）"""
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(DEBUG_LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(f"[{timestamp}] {message}\n")
        except Exception:
            pass  # 调试日志写入失败不影响主程序
    
    def update_progress_get(self, text):
        self.progress_label_get.config(text=text)
        self.root.update_idletasks()
    
    def update_progress_manage(self, text):
        self.progress_label_manage.config(text=text)
        self.root.update_idletasks()

    # ==================== Offer获取功能 ====================
    def build_request_body(self, page):
        body = {
            "token": self.pb_token_var.get().strip(),
            "page_size": 100,
            "page": page,
            "default_filter": self.default_filter_var.get(),
            "country_code": self.country_code_var.get().strip(),
            "brand_id": self.brand_id_var.get().strip() if self.brand_id_var.get().strip() else None,
            "sort": self.sort_var.get().strip(),
            "asins": self.asins_var.get().strip(),
            "relationship": int(self.relationship_var.get()),
            "is_original_currency": self.is_original_currency_var.get(),
            "has_promo_code": self.has_promo_code_var.get(),
            "has_acc": self.has_acc_var.get(),
            "filter_sexual_wellness": self.filter_sexual_wellness_var.get()
        }
        return body
    
    def generate_random_uid(self):
        """生成7位随机UID（避免重复）"""
        import secrets
        import string
        
        # 加载已使用的UID
        used_uids_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "used_uids.txt")
        used_uids = set()
        if os.path.exists(used_uids_file):
            with open(used_uids_file, 'r', encoding='utf-8') as f:
                for line in f:
                    uid = line.strip()
                    if uid:
                        used_uids.add(uid)
        
        # 生成不重复的UID
        alphabet = string.ascii_letters + string.digits
        max_attempts = 1000
        for _ in range(max_attempts):
            uid = ''.join(secrets.choice(alphabet) for _ in range(7))
            if uid not in used_uids:
                # 保存新UID
                used_uids.add(uid)
                with open(used_uids_file, 'a', encoding='utf-8') as f:
                    f.write(uid + '\n')
                return uid
        
        # 如果多次尝试都重复，使用时间戳
        import time
        uid = f"T{int(time.time()) % 10000000}"
        with open(used_uids_file, 'a', encoding='utf-8') as f:
            f.write(uid + '\n')
        return uid
    
    def _on_uid_mode_changed(self, event=None):
        """UID模式改变时的回调函数"""
        mode = self.uid_mode_var.get()
        if mode == "固定":
            self.fixed_uid_entry.config(state='normal')
        else:
            self.fixed_uid_entry.config(state='disabled')
            self.fixed_uid_var.set("")
    
    def get_uid_for_offer(self):
        """根据UID模式获取用于投放链接的UID"""
        mode = self.uid_mode_var.get()
        if mode == "无":
            return ""
        elif mode == "随机":
            return self.generate_random_uid()
        elif mode == "固定":
            uid = self.fixed_uid_var.get().strip()
            if not uid:
                self.log_get("警告: 固定UID模式但未输入UID，将使用随机UID")
                return self.generate_random_uid()
            return uid
        return ""
    
    def resolve_redirect_url(self, tracking_link):
        """访问投放链接，从返回的HTML中提取JS重定向的最终URL
        
        pboost.me使用JavaScript location.replace()做客户端跳转，
        而非HTTP 301/302重定向，因此需要解析HTML提取目标URL。
        """
        if not tracking_link:
            return ""
        try:
            resp = requests.get(tracking_link, timeout=15)
            if resp.status_code == 200:
                html = resp.text
                # 从 JavaScript location.replace(u) 中提取URL
                # 格式: var u = "https://www.amazon.com/dp/...";
                match = re.search(r'var\s+u\s*=\s*"([^"]+)";\s*\n\s*location\.replace', html)
                if match:
                    return match.group(1)
                # 备选：从 noscript meta refresh 中提取
                match = re.search(r'<meta\s+http-equiv="refresh"\s+content="[^;]*;\s*url=([^"]+)"', html)
                if match:
                    return match.group(1)
            return ""
        except Exception as e:
            self.log_get(f"    [警告] 解析重定向失败: {str(e)[:80]}")
            return ""
    
    def get_partnerboost_link(self, asin, country_code, uid=""):
        try:
            url = f"{PB_API_BASE_URL}/api/datafeed/get_amazon_link_by_asin"
            body = {
                "token": self.pb_token_var.get().strip(),
                "asins": asin,
                "country_code": country_code,
                "uid": uid,  # 使用传入的uid
                "return_partnerboost_link": 1
            }
            response = requests.post(url, json=body, timeout=30)
            if response.status_code == 200:
                data = response.json()
                status_code = data.get("status", {}).get("code")
                if status_code == 0:
                    link_data = data.get("data", [])
                    if link_data:
                        link = link_data[0].get("partnerboost_link", "")
                        if link:
                            return link
                        else:
                            # 数据中没有partnerboost_link字段
                            self.log_get(f"    [警告] ASIN={asin} 返回数据中无投放链接")
                    else:
                        # data为空列表
                        self.log_get(f"    [警告] ASIN={asin} API返回空数据")
                else:
                    # API返回错误码
                    error_msg = data.get("status", {}).get("message", "未知错误")
                    self.log_get(f"    [错误] ASIN={asin} API错误: {error_msg}")
            else:
                self.log_get(f"    [错误] ASIN={asin} HTTP状态码: {response.status_code}")
            return ""
        except Exception as e:
            self.log_get(f"    [异常] ASIN={asin} 获取链接失败: {str(e)}")
            return ""
    
    def extract_offers(self):
        if not self.pb_token_var.get().strip():
            messagebox.showerror("错误", "请在API配置中输入PartnerBoost Offer Token!")
            return
        
        self.extract_btn.config(state='disabled')
        self.progress_get.start()
        self.log_text_get.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._do_extract)
        thread.daemon = True
        thread.start()
    
    def _do_extract(self):
        all_offers = []
        page = 1
        has_more = True
        
        try:
            endpoint = "/api/datafeed/get_fba_products"
            full_url = PB_API_BASE_URL + endpoint
            self.log_get(f"开始获取Offer数据...")
            
            headers = {"Content-Type": "application/json", "Accept": "application/json"}
            
            while has_more:
                request_body = self.build_request_body(page)
                self.root.after(0, lambda p=page, t=len(all_offers): self.update_progress_get(f"第{p}页 | 已获取{t}条"))
                
                response = requests.post(full_url, json=request_body, headers=headers, timeout=60)
                
                if response.status_code == 200:
                    data = response.json()
                    if data.get("status", {}).get("code") == 0:
                        offers = data.get("data", {}).get("list", [])
                        has_more = data.get("data", {}).get("has_more", False)
                        all_offers.extend(offers)
                        self.log_get(f"第{page}页: 获取 {len(offers)} 个offer，累计 {len(all_offers)} 个")
                        page += 1
                    else:
                        self.log_get(f"API返回错误: {data.get('status', {}).get('msg')}")
                        break
                else:
                    self.log_get(f"HTTP请求失败: {response.status_code}")
                    break
            
            self.log_get(f"Offer数据获取完成，共 {len(all_offers)} 个")
            
            if all_offers:
                self.log_get(f"开始获取投放链接...")
                
                # 获取UID模式
                uid_mode = self.uid_mode_var.get()
                self.log_get(f"UID模式: {uid_mode}")
                
                total = len(all_offers)
                for idx, offer in enumerate(all_offers):
                    asin = offer.get("asin", "")
                    country_code = offer.get("country_code", "")
                    if asin and country_code:
                        self.root.after(0, lambda i=idx+1, t=total: self.update_progress_get(f"获取链接 {i}/{t}"))
                        # 根据UID模式获取UID
                        uid = self.get_uid_for_offer()
                        link = self.get_partnerboost_link(asin, country_code, uid)
                        offer["partnerboost_link"] = link
                        
                        # 访问投放链接获取重定向后的最终URL，替换产品链接
                        if link:
                            redirect_url = self.resolve_redirect_url(link)
                            if redirect_url:
                                offer["url"] = redirect_url
                        
                        if (idx + 1) % 10 == 0:
                            self.log_get(f"已获取 {idx + 1}/{total} 个投放链接")
                        time.sleep(0.1)
                    else:
                        offer["partnerboost_link"] = ""
                
                self.save_to_xlsx(all_offers)
            else:
                self.log_get("没有获取到任何offer数据")
                
        except Exception as e:
            self.log_get(f"错误: {str(e)}")
        finally:
            self.root.after(0, self._restore_get_ui)
    
    def _restore_get_ui(self):
        self.extract_btn.config(state='normal')
        self.copy_offer_btn.config(state='normal')
        self.brand_search_btn.config(state='normal')
        self.progress_get.stop()
        self.update_progress_get("")
    
    def copy_offers(self):
        """复制offer功能：识别只有国家代码和ASIN的行，复制完整信息并生成新投放链接"""
        if not self.feishu_app_id_var.get().strip() or not self.feishu_app_secret_var.get().strip():
            messagebox.showerror("错误", "请在API配置中输入飞书App ID和App Secret!")
            return
        
        if not self.pb_token_var.get().strip():
            messagebox.showerror("错误", "请在API配置中输入PartnerBoost Offer Token!")
            return
        
        self.extract_btn.config(state='disabled')
        self.copy_offer_btn.config(state='disabled')
        self.progress_get.start()
        self.log_text_get.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._do_copy_offers)
        thread.daemon = True
        thread.start()
    
    def _do_copy_offers(self):
        """执行复制offer操作"""
        try:
            self.log_get("=" * 50)
            self.log_get("开始复制offer...")
            self.log_get("=" * 50)
            
            # 获取飞书访问令牌
            token = self.get_feishu_token()
            if not token:
                self.log_get("获取飞书访问令牌失败")
                return
            
            spreadsheet_token = self.feishu_spreadsheet_var.get().strip()
            sheet_id = self.feishu_sheet_id_var.get().strip()
            
            if not spreadsheet_token or not sheet_id:
                self.log_get("飞书电子表格配置缺失")
                return
            
            # 读取飞书数据
            self.log_get("\n【步骤1】读取飞书offer表格数据...")
            self.update_progress_get("读取飞书数据...")
            
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!A1:Z2000"
            headers = {"Authorization": f"Bearer {token}"}
            resp = requests.get(url, headers=headers)
            data = resp.json()
            
            if data.get('code') != 0:
                self.log_get(f"读取飞书数据失败: {data.get('msg')}")
                return
            
            values = data.get('data', {}).get('valueRange', {}).get('values', [])
            if len(values) < 2:
                self.log_get("飞书表格数据不足")
                return
            
            headers_row = values[0]
            
            # 找到所有列的索引
            col_indices = {}
            for i, h in enumerate(headers_row):
                h_str = str(h).strip() if h else ''
                if h_str:
                    col_indices[h_str] = i
            
            self.log_get(f"  检测到的列: {list(col_indices.keys())}")
            
            # 必需的列
            if 'ASIN' not in col_indices or '国家代码' not in col_indices:
                self.log_get("未找到ASIN或国家代码列")
                return
            
            asin_col = col_indices['ASIN']
            country_col = col_indices['国家代码']
            tracking_link_col = col_indices.get('投放链接')
            product_link_col = col_indices.get('产品链接')
            
            # 定义需要复制的字段（除投放链接和公式列外）
            # 注意："每单佣金"是公式列，不需要复制，会自动计算
            copy_fields = ['品牌名称', '品牌ID', '折扣价', '佣金', '产品名称', '产品链接', 
                          '库存状态', '原价', '货币', '更新时间', '图片URL', '分类', '子分类',
                          '评分', '评论数', '折扣', '折扣码', '优惠券']
            
            # 第一遍扫描：识别复制源和复制对象
            self.log_get("\n【步骤2】识别复制源和复制对象...")
            
            copy_sources = []  # [(row_idx, asin, country), ...]
            complete_offers = {}  # {(asin, country): (row_idx, row_data), ...} 第一个完整offer
            
            for row_idx, row in enumerate(values[1:], start=2):
                asin = str(row[asin_col]).strip() if asin_col < len(row) and row[asin_col] else ''
                country = str(row[country_col]).strip().upper() if country_col < len(row) and row[country_col] else ''
                
                if not asin or not country:
                    continue
                
                # 检查是否只有ASIN和国家代码有值（复制源）
                has_other_data = False
                for field in copy_fields:
                    if field in col_indices:
                        field_col = col_indices[field]
                        if field_col < len(row) and row[field_col]:
                            cell_value = str(row[field_col]).strip()
                            if cell_value and cell_value.lower() not in ['none', '']:
                                has_other_data = True
                                break
                
                # 检查投放链接
                has_tracking_link = False
                if tracking_link_col is not None and tracking_link_col < len(row) and row[tracking_link_col]:
                    link_value = row[tracking_link_col]
                    if isinstance(link_value, list) and len(link_value) > 0 and isinstance(link_value[0], dict):
                        link_value = link_value[0].get('link', '') or link_value[0].get('text', '')
                    if str(link_value).strip():
                        has_tracking_link = True
                
                key = (asin, country)
                
                if not has_other_data and not has_tracking_link:
                    # 这是复制源（只有ASIN和国家代码）
                    copy_sources.append((row_idx, asin, country))
                elif has_other_data:
                    # 这是有完整信息的offer，记录第一个作为复制对象
                    if key not in complete_offers:
                        complete_offers[key] = (row_idx, row)
            
            self.log_get(f"  找到 {len(copy_sources)} 个复制源（只有ASIN和国家代码）")
            self.log_get(f"  找到 {len(complete_offers)} 个有完整信息的offer")
            
            if not copy_sources:
                self.log_get("没有找到需要复制的offer（复制源）")
                return
            
            # 第二遍处理：执行复制
            self.log_get("\n【步骤3】执行复制操作...")
            self.update_progress_get("复制offer...")
            
            updates = []  # [(row_idx, col_idx, value), ...]
            style_updates = []  # [(row_idx, col_idx, style_type), ...] 样式更新
            success_count = 0
            skip_count = 0
            
            for row_idx, asin, country in copy_sources:
                key = (asin, country)
                
                if key not in complete_offers:
                    self.log_get(f"  [跳过] {asin}_{country}: 未找到同ASIN+国家的完整offer")
                    skip_count += 1
                    continue
                
                source_row_idx, source_row = complete_offers[key]
                self.log_get(f"  [复制] {asin}_{country}: 从第{source_row_idx}行复制到第{row_idx}行")
                
                # 复制字段（除投放链接外）
                for field in copy_fields:
                    if field in col_indices:
                        field_col = col_indices[field]
                        if field_col < len(source_row) and source_row[field_col]:
                            source_value = source_row[field_col]
                            # 处理URL类型的单元格
                            if isinstance(source_value, list) and len(source_value) > 0 and isinstance(source_value[0], dict):
                                source_value = source_value[0].get('link', '') or source_value[0].get('text', '')
                            if source_value and str(source_value).strip():
                                updates.append((row_idx, field_col, source_value))
                
                # 为"每单佣金"列写入公式（公式格式: =L{row}*M{row}）
                if '每单佣金' in col_indices:
                    commission_per_order_col = col_indices['每单佣金']
                    formula = f"=L{row_idx}*M{row_idx}"
                    updates.append((row_idx, commission_per_order_col, formula))
                    self.log_get(f"    写入每单佣金公式: {formula}")
                
                # 设置"状态"列为"新复制"（蓝色加粗）
                if '状态' in col_indices:
                    status_col = col_indices['状态']
                    updates.append((row_idx, status_col, '新复制'))
                    style_updates.append((row_idx, status_col, 'blue_bold'))
                    self.log_get(f"    设置状态: 新复制 (蓝色加粗)")
                
                # 生成新的投放链接
                self.log_get(f"    获取新投放链接...")
                uid = self.generate_random_uid()  # 为复制的offer生成随机uid
                new_link = self.get_partnerboost_link(asin, country, uid)
                
                if new_link and tracking_link_col is not None:
                    updates.append((row_idx, tracking_link_col, new_link))
                    self.log_get(f"    新投放链接: {new_link[-20:]}... (uid={uid})")
                    
                    # 解析投放链接重定向，更新产品链接
                    if product_link_col is not None:
                        redirect_url = self.resolve_redirect_url(new_link)
                        if redirect_url:
                            updates.append((row_idx, product_link_col, redirect_url))
                            self.log_get(f"    产品链接已更新: {redirect_url[:50]}...")
                    
                    success_count += 1
                else:
                    self.log_get(f"    [警告] 获取投放链接失败")
                    success_count += 1  # 即使没有链接，其他字段也复制了
                
                time.sleep(0.3)  # API调用间隔
            
            # 应用更新到飞书
            if updates:
                self.log_get(f"\n【步骤4】写入飞书表格... ({len(updates)} 个单元格)")
                self.update_progress_get("写入飞书...")
                self._apply_copy_updates(token, spreadsheet_token, sheet_id, updates, style_updates, headers_row)
            
            # 输出统计
            self.log_get("\n" + "=" * 50)
            self.log_get("复制完成！统计信息：")
            self.log_get(f"  成功复制: {success_count} 个offer")
            self.log_get(f"  跳过（无复制对象）: {skip_count} 个")
            self.log_get("=" * 50)
            
        except Exception as e:
            self.log_get(f"复制offer时发生错误: {str(e)}")
            import traceback
            self.log_get(traceback.format_exc())
        finally:
            self.root.after(0, self._restore_get_ui)
    
    def _apply_copy_updates(self, token, spreadsheet_token, sheet_id, updates, style_updates, headers_row):
        """应用复制offer的更新到飞书表格"""
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # 构建批量更新请求
        value_ranges = []
        for row_idx, col_idx, value in updates:
            col_letter = self.index_to_column_letter(col_idx)
            value_ranges.append({
                'range': f"{sheet_id}!{col_letter}{row_idx}:{col_letter}{row_idx}",
                'values': [[value]]
            })
        
        # 分批处理，每批最多100个range
        batch_size = 100
        for i in range(0, len(value_ranges), batch_size):
            batch = value_ranges[i:i+batch_size]
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
            body = {"valueRanges": batch}
            
            try:
                response = requests.post(url, headers=headers, json=body, timeout=30)
                data = response.json()
                if data.get('code') == 0:
                    self.log_get(f"  批次 {i//batch_size + 1}: 更新成功 {len(batch)} 个单元格")
                else:
                    self.log_get(f"  批次 {i//batch_size + 1}: 更新失败 - {data.get('msg', 'Unknown error')}")
            except Exception as e:
                self.log_get(f"  批次 {i//batch_size + 1}: 更新异常 - {str(e)}")
        
        # 应用样式更新（蓝色加粗）
        if style_updates:
            self._apply_copy_style_updates(token, spreadsheet_token, sheet_id, style_updates)
    
    def _apply_copy_style_updates(self, token, spreadsheet_token, sheet_id, style_updates):
        """应用复制offer的样式更新（蓝色加粗）"""
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        for row_idx, col_idx, style_type in style_updates:
            col_letter = self.index_to_column_letter(col_idx)
            range_str = f"{sheet_id}!{col_letter}{row_idx}:{col_letter}{row_idx}"
            
            if style_type == 'blue_bold':
                style = {
                    "font": {
                        "bold": True
                    },
                    "foreColor": "#0000FF"  # 蓝色
                }
            else:
                style = {
                    "font": {
                        "bold": False
                    },
                    "foreColor": "#000000"  # 黑色
                }
            
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/style"
            body = {
                "appendStyle": {
                    "range": range_str,
                    "style": style
                }
            }
            
            try:
                resp = requests.put(url, headers=headers, json=body, timeout=10)
                result = resp.json()
                if result.get('code') != 0:
                    self.log_get(f"  样式更新失败: {result.get('msg', 'Unknown error')}")
            except Exception as e:
                self.log_get(f"  样式更新异常: {str(e)}")
    
    def generate_random_uid(self):
        """生成随机7位UID"""
        import random
        import string
        return ''.join(random.choices(string.ascii_lowercase + string.digits, k=7))
    
    # ==================== 品牌搜索量报告 ====================
    def get_brand_search_volume(self):
        """获取品牌搜索量报告"""
        if not self.pb_token_var.get().strip():
            messagebox.showerror("错误", "请在API配置中输入PartnerBoost Offer Token!")
            return
        
        self.extract_btn.config(state='disabled')
        self.copy_offer_btn.config(state='disabled')
        self.brand_search_btn.config(state='disabled')
        self.progress_get.start()
        self.log_text_get.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._do_brand_search_volume)
        thread.daemon = True
        thread.start()
    
    def _do_brand_search_volume(self):
        """执行品牌搜索量报告生成"""
        try:
            self.log_get("=" * 50)
            self.log_get("开始获取品牌搜索量报告...")
            self.log_get("=" * 50)
            
            # 步骤1：获取PB partnered品牌列表
            self.log_get("\n【步骤1】获取PartnerBoost已合作品牌列表...")
            self.update_progress_get("获取品牌列表...")
            
            brands = self._get_partnered_brands()
            if not brands:
                self.log_get("未获取到任何品牌数据")
                return
            
            self.log_get(f"  共获取到 {len(brands)} 个品牌")
            
            # 步骤2：解析品牌名和国家代码
            self.log_get("\n【步骤2】解析品牌名和国家代码...")
            self.update_progress_get("解析品牌信息...")
            
            parsed_brands = []
            for brand in brands:
                bid = brand.get('bid', '')
                raw_name = brand.get('brand_name', '')
                brand_name, country_code = self._parse_brand_name_country(raw_name)
                parsed_brands.append({
                    'bid': bid,
                    'raw_name': raw_name,
                    'brand_name': brand_name,
                    'country_code': country_code
                })
            
            # 统计国家分布
            country_stats = {}
            for b in parsed_brands:
                cc = b['country_code']
                country_stats[cc] = country_stats.get(cc, 0) + 1
            self.log_get(f"  国家分布: {country_stats}")
            self.log_get(f"  前5个品牌示例:")
            for b in parsed_brands[:5]:
                self.log_get(f"    • {b['raw_name']} -> 品牌名={b['brand_name']}, 国家={b['country_code']}, BID={b['bid']}")
            
            # 步骤3：调用Google Ads API获取搜索量
            self.log_get("\n【步骤3】调用Google Ads API获取关键词搜索量...")
            self.update_progress_get("获取搜索量...")
            
            client = self.get_google_ads_client()
            if not client:
                self.log_get("  ✗ 无法创建Google Ads客户端，请检查配置")
                return
            
            mcc_id = self.google_mcc_id_var.get().strip()
            
            # 按国家代码分组，批量查询
            country_brand_groups = {}  # {country_code: [brand_info, ...]}
            for b in parsed_brands:
                cc = b['country_code']
                if cc not in country_brand_groups:
                    country_brand_groups[cc] = []
                country_brand_groups[cc].append(b)
            
            results = []
            total_brands = len(parsed_brands)
            processed = 0
            
            for country_code, brand_list in country_brand_groups.items():
                self.log_get(f"\n  查询国家 {country_code} 的品牌搜索量（{len(brand_list)} 个品牌）...")
                
                # 分批查询，每批最多20个关键词
                batch_size = 20
                max_retries = 5
                for i in range(0, len(brand_list), batch_size):
                    batch = brand_list[i:i+batch_size]
                    keywords = [b['brand_name'] for b in batch]
                    
                    success = False
                    for attempt in range(max_retries):
                        try:
                            metrics = self._get_keyword_historical_metrics(client, mcc_id, keywords, country_code)
                            
                            for b in batch:
                                keyword = b['brand_name']
                                metric = metrics.get(keyword.lower(), {})
                                b['avg_monthly_searches'] = metric.get('avg_monthly_searches', 0)
                                b['competition'] = metric.get('competition', 'N/A')
                                b['competition_index'] = metric.get('competition_index', 'N/A')
                                b['low_top_of_page_bid'] = metric.get('low_top_of_page_bid', 0)
                                b['high_top_of_page_bid'] = metric.get('high_top_of_page_bid', 0)
                                results.append(b)
                            
                            processed += len(batch)
                            self.update_progress_get(f"搜索量 {processed}/{total_brands}...")
                            success = True
                            break
                            
                        except Exception as e:
                            error_str = str(e)
                            # 解析429限流错误中的等待时间
                            is_rate_limit = '429' in error_str or 'Resource has been exhausted' in error_str
                            if is_rate_limit and attempt < max_retries - 1:
                                # 尝试从错误信息中提取等待秒数
                                wait_seconds = 5  # 默认等待时间
                                retry_match = re.search(r'Retry in (\d+) seconds', error_str)
                                if retry_match:
                                    wait_seconds = int(retry_match.group(1)) + 1
                                else:
                                    wait_seconds = min(5 * (2 ** attempt), 60)  # 指数退避，最长60秒
                                self.log_get(f"    [限流] 第{attempt+1}次重试，等待{wait_seconds}秒...")
                                self.update_progress_get(f"限流等待{wait_seconds}秒... ({processed}/{total_brands})")
                                time.sleep(wait_seconds)
                            else:
                                self.log_get(f"    [错误] 批次查询失败 (已重试{attempt}次): {error_str[:200]}")
                                break
                    
                    if not success:
                        # 所有重试都失败，记录错误结果
                        for b in batch:
                            b['avg_monthly_searches'] = 'ERROR'
                            b['competition'] = 'ERROR'
                            b['competition_index'] = 'ERROR'
                            b['low_top_of_page_bid'] = 'ERROR'
                            b['high_top_of_page_bid'] = 'ERROR'
                            results.append(b)
                        processed += len(batch)
                    
                    time.sleep(1)  # API调用间隔
            
            self.log_get(f"\n  搜索量查询完成，共 {len(results)} 个品牌")
            
            # 步骤4：查询品牌Storefront链接
            self.log_get("\n【步骤4】查询品牌Storefront链接...")
            self.update_progress_get("查询Storefront链接...")
            
            all_bids = [r['bid'] for r in results if r.get('bid')]
            storefront_map = self._get_brand_storefront_links(all_bids)
            
            for r in results:
                r['storefront_link'] = storefront_map.get(str(r.get('bid', '')), '')
            
            found_count = sum(1 for v in storefront_map.values() if v)
            self.log_get(f"  Storefront链接查询完成: {found_count}/{len(all_bids)} 个品牌有链接")
            
            # 步骤5：查询各品牌产品数量
            self.log_get("\n【步骤5】查询各品牌产品数量...")
            self.update_progress_get("查询产品数量...")
            
            product_count_map = self._get_brand_product_counts(all_bids)
            
            for r in results:
                r['product_count'] = product_count_map.get(str(r.get('bid', '')), 0)
            
            has_products = sum(1 for v in product_count_map.values() if v > 0)
            total_products = sum(product_count_map.values())
            self.log_get(f"  产品数量查询完成: {has_products}/{len(all_bids)} 个品牌有产品，共 {total_products} 个产品")
            
            # 步骤6：生成报告
            self.log_get("\n【步骤6】生成Excel报告...")
            self.update_progress_get("生成报告...")
            
            self._save_brand_search_report(results)
            
            self.log_get("\n" + "=" * 50)
            self.log_get("品牌搜索量报告生成完成！")
            self.log_get("=" * 50)
            
        except Exception as e:
            self.log_get(f"生成品牌搜索量报告时发生错误: {str(e)}")
            import traceback
            self.log_get(traceback.format_exc())
        finally:
            self.root.after(0, self._restore_get_ui)
    
    def _get_partnered_brands(self):
        """获取PB已合作品牌列表（自动翻页）"""
        all_brands = []
        page = 1
        page_size = 50
        
        token = self.pb_token_var.get().strip()
        url = f"{PB_API_BASE_URL}/api/datafeed/get_amazon_joined_brands"
        
        while True:
            body = {
                "token": token,
                "bids": "",
                "page_size": page_size,
                "page": page
            }
            
            try:
                response = requests.post(url, json=body, timeout=30)
                data = response.json()
                
                if data.get('status', {}).get('code') != 0:
                    self.log_get(f"  获取品牌列表失败 (页{page}): {data.get('status', {}).get('msg', 'Unknown error')}")
                    break
                
                brand_list = data.get('data', {}).get('list', [])
                if not brand_list:
                    break
                
                all_brands.extend(brand_list)
                has_more = data.get('data', {}).get('hasMore', False)
                
                self.log_get(f"  第{page}页: 获取 {len(brand_list)} 个品牌 (累计 {len(all_brands)})")
                
                if not has_more:
                    break
                
                page += 1
                time.sleep(0.3)
                
            except Exception as e:
                self.log_get(f"  获取品牌列表异常 (页{page}): {str(e)}")
                break
        
        return all_brands
    
    def _get_brand_storefront_links(self, bids):
        """批量查询品牌storefront链接，返回 {bid: storefront_link} 映射"""
        storefront_map = {}
        token = self.pb_token_var.get().strip()
        url = f"{PB_API_BASE_URL}/api/datafeed/get_fba_brand_link"
        
        batch_size = 50
        for i in range(0, len(bids), batch_size):
            batch = bids[i:i+batch_size]
            bids_str = ','.join(str(b) for b in batch)
            
            try:
                response = requests.post(url, json={"token": token, "bids": bids_str}, timeout=30)
                data = response.json()
                
                if data.get('status', {}).get('code') == 0 and data.get('data'):
                    for item in data['data']:
                        bid = item.get('bid', '')
                        link = item.get('storefront_link', '')
                        if link:
                            storefront_map[str(bid)] = link
                
                self.log_get(f"    Storefront查询批次 {i//batch_size+1}: 成功 {len([b for b in batch if str(b) in storefront_map])}/{len(batch)}")
                
            except Exception as e:
                self.log_get(f"    Storefront查询批次 {i//batch_size+1} 失败: {str(e)[:100]}")
            
            if i + batch_size < len(bids):
                time.sleep(0.5)
        
        return storefront_map
    
    def _get_brand_product_counts(self, bids):
        """查询各品牌的产品数量，返回 {bid: count} 映射"""
        count_map = {}
        token = self.pb_token_var.get().strip()
        url = f"{PB_API_BASE_URL}/api/datafeed/get_fba_products"
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        
        total = len(bids)
        for idx, bid in enumerate(bids):
            bid_str = str(bid).strip()
            if not bid_str:
                continue
            
            product_count = 0
            page = 1
            has_more = True
            
            while has_more:
                try:
                    body = {
                        "token": token,
                        "page_size": 100,
                        "page": page,
                        "default_filter": 0,
                        "country_code": "",
                        "brand_id": bid_str,
                        "sort": "",
                        "asins": "",
                        "relationship": 1,
                        "is_original_currency": 0,
                        "has_promo_code": 0,
                        "has_acc": 0,
                        "filter_sexual_wellness": 0
                    }
                    resp = requests.post(url, json=body, headers=headers, timeout=60)
                    data = resp.json()
                    
                    if data.get("status", {}).get("code") == 0:
                        products = data.get("data", {}).get("list", [])
                        product_count += len(products)
                        has_more = data.get("data", {}).get("has_more", False)
                        page += 1
                    else:
                        break
                except Exception:
                    break
                
                time.sleep(0.1)
            
            count_map[bid_str] = product_count
            
            if (idx + 1) % 10 == 0 or idx == total - 1:
                self.log_get(f"    产品数量查询进度: {idx+1}/{total}")
                self.update_progress_get(f"产品数量 {idx+1}/{total}...")
        
        return count_map
    
    def _parse_brand_name_country(self, raw_name):
        """从品牌名中智能解析品牌名和国家代码
        
        示例：
            "Odoland_UN-US" -> ("Odoland", "US")
            "PARWIN PRO BEAUTY--DE" -> ("PARWIN PRO BEAUTY", "DE")
            "DE-Anker" -> ("Anker", "DE")
            "Brand_FR" -> ("Brand", "FR")
            "SimpleBrand" -> ("SimpleBrand", "US")  # 默认US
        """
        if not raw_name:
            return (raw_name, 'US')
        
        country_codes = ['US', 'UK', 'DE', 'FR', 'IT', 'ES', 'CA', 'JP', 'AU', 'NL', 'BE', 'MX', 'BR', 'IN', 'SG', 'AE', 'SA', 'PL', 'SE', 'TR', 'EG', 'GB']
        
        name = raw_name.strip()
        
        # 模式1: 结尾有国家代码，用各种分隔符 如 "Brand_UN-US", "Brand--DE", "Brand_FR", "Brand-US"
        for code in country_codes:
            # 结尾匹配：各种分隔符+国家代码
            patterns = [
                (f'[-_]+UN[-_]+{code}$', code),        # _UN-US, -UN_US
                (f'[-_]{{2,}}{code}$', code),           # --DE, __DE
                (f'[-_]{code}$', code),                 # -US, _FR
            ]
            for pattern, matched_code in patterns:
                match = re.search(pattern, name, re.IGNORECASE)
                if match:
                    brand_name = name[:match.start()].strip()
                    if brand_name:
                        return (brand_name, matched_code)
        
        # 模式2: 开头有国家代码 如 "DE-Anker"
        for code in country_codes:
            match = re.match(rf'^({code})[-_]+(.+)$', name, re.IGNORECASE)
            if match:
                return (match.group(2).strip(), code)
        
        # 没有找到国家代码，默认US
        return (name, 'US')
    
    # 国家代码到Google Ads Geo Target Constant ID的映射
    GEO_TARGET_MAP = {
        'US': '2840', 'UK': '2826', 'GB': '2826', 'DE': '2276', 'FR': '2250',
        'IT': '2380', 'ES': '2724', 'CA': '2124', 'JP': '2392', 'AU': '2036',
        'NL': '2528', 'BE': '2056', 'MX': '2484', 'BR': '2076', 'IN': '2356',
        'SG': '2702', 'AE': '2784', 'SA': '2682', 'PL': '2616', 'SE': '2752',
        'TR': '2792', 'EG': '2818',
    }
    
    # 国家代码到Google Ads Language Constant ID的映射
    LANGUAGE_MAP = {
        'US': '1000', 'UK': '1000', 'GB': '1000', 'CA': '1000', 'AU': '1000',  # English
        'DE': '1001',  # German
        'FR': '1002',  # French
        'IT': '1004',  # Italian
        'ES': '1003',  # Spanish
        'JP': '1005',  # Japanese
        'NL': '1010',  # Dutch
        'BE': '1010',  # Dutch (Belgium)
        'MX': '1003',  # Spanish
        'BR': '1014',  # Portuguese
        'IN': '1000',  # English (India)
        'SG': '1000',  # English (Singapore)
        'AE': '1019',  # Arabic
        'SA': '1019',  # Arabic
        'PL': '1030',  # Polish
        'SE': '1015',  # Swedish
        'TR': '1037',  # Turkish
        'EG': '1019',  # Arabic
    }
    
    def _get_keyword_historical_metrics(self, client, customer_id, keywords, country_code):
        """调用Google Ads API获取关键词历史搜索量指标
        
        Args:
            client: GoogleAdsClient实例
            customer_id: MCC客户ID
            keywords: 关键词列表
            country_code: 国家代码
            
        Returns:
            {keyword_lower: {avg_monthly_searches, competition, ...}}
        """
        googleads_service = client.get_service("GoogleAdsService")
        keyword_plan_idea_service = client.get_service("KeywordPlanIdeaService")
        
        request = client.get_type("GenerateKeywordHistoricalMetricsRequest")
        request.customer_id = customer_id
        request.keywords.extend(keywords)
        
        # 设置地理位置
        geo_id = self.GEO_TARGET_MAP.get(country_code.upper(), '2840')  # 默认US
        request.geo_target_constants.append(
            googleads_service.geo_target_constant_path(geo_id)
        )
        
        # 设置搜索网络
        request.keyword_plan_network = client.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
        
        # 设置语言
        lang_id = self.LANGUAGE_MAP.get(country_code.upper(), '1000')  # 默认English
        request.language = googleads_service.language_constant_path(lang_id)
        
        response = keyword_plan_idea_service.generate_keyword_historical_metrics(
            request=request
        )
        
        metrics_map = {}
        for result in response.results:
            keyword = result.text.lower() if result.text else ''
            m = result.keyword_metrics
            metrics_map[keyword] = {
                'avg_monthly_searches': m.avg_monthly_searches if m.avg_monthly_searches else 0,
                'competition': m.competition.name if m.competition else 'N/A',
                'competition_index': m.competition_index if m.competition_index else 0,
                'low_top_of_page_bid': round(m.low_top_of_page_bid_micros / 1_000_000, 2) if m.low_top_of_page_bid_micros else 0,
                'high_top_of_page_bid': round(m.high_top_of_page_bid_micros / 1_000_000, 2) if m.high_top_of_page_bid_micros else 0,
            }
        
        return metrics_map
    
    def _save_brand_search_report(self, results):
        """保存品牌搜索量报告为xlsx"""
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
        from datetime import datetime
        
        wb = Workbook()
        ws = wb.active
        ws.title = "品牌搜索量报告"
        
        # 表头
        headers = ['Brand ID', 'Storefront Link', '产品数量', '原始品牌名', '品牌名', '国家代码', '月均搜索量', 
                   '竞争程度', '竞争指数', '页首低价出价($)', '页首高价出价($)']
        ws.append(headers)
        
        # 数据行
        for r in results:
            ws.append([
                r.get('bid', ''),
                r.get('storefront_link', ''),
                r.get('product_count', 0),
                r.get('raw_name', ''),
                r.get('brand_name', ''),
                r.get('country_code', ''),
                r.get('avg_monthly_searches', 0),
                r.get('competition', 'N/A'),
                r.get('competition_index', 0),
                r.get('low_top_of_page_bid', 0),
                r.get('high_top_of_page_bid', 0),
            ])
        
        # 调整列宽
        for col_idx in range(1, len(headers) + 1):
            col_letter = get_column_letter(col_idx)
            max_length = len(str(headers[col_idx - 1]))
            for row_idx in range(2, min(len(results) + 2, 100)):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value:
                    max_length = max(max_length, min(len(str(cell_value)), 40))
            ws.column_dimensions[col_letter].width = max_length + 2
        
        # 按月均搜索量降序排序（表头行之后）
        data_rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))
        data_rows.sort(key=lambda x: (x[6] if isinstance(x[6], (int, float)) else 0), reverse=True)
        
        # 清除数据行，重新写入排序后的数据
        for row_idx in range(2, ws.max_row + 1):
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col_idx).value = None
        
        for row_idx, row_data in enumerate(data_rows, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx).value = value
        
        # 保存文件
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_dir = os.path.dirname(self.save_path_var.get())
        if not save_dir:
            save_dir = os.path.join(os.path.expanduser("~"), "Desktop")
        save_path = os.path.join(save_dir, f"品牌搜索量报告_{timestamp}.xlsx")
        
        wb.save(save_path)
        self.log_get(f"  报告已保存到: {save_path}")
        self.root.after(0, lambda: self._ask_open_file(save_path))
    
    def save_to_xlsx(self, offers):
        try:
            save_path = self.save_path_var.get()
            wb = Workbook()
            ws = wb.active
            ws.title = "Offers"
            
            columns_config = [
                ("投放链接", "partnerboost_link", False),
                ("国家代码", "country_code", False),
                ("品牌名称", "brand_name", False),
                ("折扣价", "discount_price", False),
                ("佣金", "commission", False),
                ("产品ID", "product_id", True),
                ("产品名称", "product_name", False),
                ("图片URL", "image", True),
                ("ASIN", "asin", False),
                ("折扣", "discount", True),
                ("折扣码", "discount_code", True),
                ("优惠券", "coupon", True),
                ("分类", "category", True),
                ("子分类", "subcategory", True),
                ("父ASIN", "parent_asin", True),
                ("变体ASIN", "variant_asin", True),
                ("库存状态", "availability", False),
                ("评分", "rating", True),
                ("评论数", "reviews", True),
                ("产品链接", "url", False),
                ("品牌ID", "brand_id", False),
                ("更新时间", "update_time", False),
                ("Relationship", "relationship", True),
                ("原价", "original_price", False),
                ("货币", "currency", False),
                ("ACC佣金", "acc_commission", True),
                ("ACC开始日期", "acc_start_date", True),
                ("ACC结束日期", "acc_end_date", True),
                ("促销码列表", "promo_code_list", True),
            ]
            
            headers = [col[0] for col in columns_config]
            ws.append(headers)
            
            for offer in offers:
                row = []
                for header, field, hidden in columns_config:
                    if field == "promo_code_list":
                        promo_codes = offer.get("promo_code_list", [])
                        if promo_codes:
                            promo_code_str = "; ".join([f"{p.get('promo_code', '')}({p.get('promotional_price', '')})" for p in promo_codes])
                        else:
                            promo_code_str = ""
                        row.append(promo_code_str)
                    else:
                        row.append(offer.get(field, ""))
                ws.append(row)
            
            for col_idx, (header, field, hidden) in enumerate(columns_config, start=1):
                col_letter = get_column_letter(col_idx)
                if hidden:
                    ws.column_dimensions[col_letter].hidden = True
                else:
                    max_length = len(header)
                    for row_idx in range(1, min(len(offers) + 2, 100)):
                        cell_value = ws.cell(row=row_idx, column=col_idx).value
                        if cell_value:
                            max_length = max(max_length, min(len(str(cell_value)), 50))
                    ws.column_dimensions[col_letter].width = max_length + 2
            
            wb.save(save_path)
            self.log_get(f"数据已保存到: {save_path}")
            self.root.after(0, lambda: self._ask_open_file(save_path))
            
        except Exception as e:
            self.log_get(f"保存文件失败: {str(e)}")
    
    def _ask_open_file(self, path):
        if messagebox.askyesno("保存成功", f"文件已保存到:\n{path}\n\n是否打开文件?"):
            os.startfile(path)

    # ==================== Offer管理功能 ====================
    def test_connections(self):
        """测试所有API连接"""
        self.log_text_manage.delete(1.0, tk.END)
        thread = threading.Thread(target=self._do_test_connections)
        thread.daemon = True
        thread.start()
    
    def _do_test_connections(self):
        self.log_manage("开始测试API连接...")
        
        # 测试飞书
        self.log_manage("\n[飞书API]")
        try:
            token = self.get_feishu_token()
            if token:
                self.log_manage("  ✓ 飞书Token获取成功")
            else:
                self.log_manage("  ✗ 飞书Token获取失败")
        except Exception as e:
            self.log_manage(f"  ✗ 飞书连接失败: {e}")
        
        # 测试Google Ads
        self.log_manage("\n[Google Ads API]")
        try:
            client = self.get_google_ads_client()
            if client:
                customer_service = client.get_service('CustomerService')
                accessible_customers = customer_service.list_accessible_customers()
                self.log_manage(f"  ✓ Google Ads连接成功，可访问 {len(accessible_customers.resource_names)} 个账户")
            else:
                self.log_manage("  ✗ Google Ads客户端创建失败")
        except Exception as e:
            self.log_manage(f"  ✗ Google Ads连接失败: {e}")
        
        # 测试PartnerBoost
        self.log_manage("\n[PartnerBoost API]")
        try:
            url = f"{PB_API_BASE_URL}/api.php?mod=medium&op=transaction"
            body = {
                "token": self.pb_token_var.get().strip(),
                "begin_date": (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d'),
                "end_date": datetime.now().strftime('%Y-%m-%d'),
                "type": "json",
                "status": "All",
                "limit": 1,
                "page": 1
            }
            response = requests.post(url, json=body, timeout=30)
            data = response.json()
            if data.get("status", {}).get("code") == 0:
                self.log_manage("  ✓ PartnerBoost连接成功")
            else:
                self.log_manage(f"  ✗ PartnerBoost返回错误: {data.get('status', {}).get('msg')}")
        except Exception as e:
            self.log_manage(f"  ✗ PartnerBoost连接失败: {e}")
        
        self.log_manage("\n测试完成！")
    
    def get_feishu_token(self):
        """获取飞书tenant_access_token"""
        url = f"{FEISHU_API_BASE_URL}/auth/v3/tenant_access_token/internal"
        body = {
            "app_id": self.feishu_app_id_var.get().strip(),
            "app_secret": self.feishu_app_secret_var.get().strip()
        }
        response = requests.post(url, json=body, timeout=30)
        data = response.json()
        if data.get("code") == 0:
            return data.get("tenant_access_token")
        return None
    
    def get_google_ads_client(self):
        """创建Google Ads客户端"""
        sa_file = self.google_sa_file_var.get().strip()
        if not os.path.exists(sa_file):
            return None
        
        credentials = service_account.Credentials.from_service_account_file(
            sa_file,
            scopes=['https://www.googleapis.com/auth/adwords']
        )
        
        client = GoogleAdsClient(
            credentials=credentials,
            developer_token=self.google_dev_token_var.get().strip(),
            login_customer_id=self.google_mcc_id_var.get().strip(),
            use_proto_plus=True
        )
        return client
    
    def start_statistics(self):
        """开始统计"""
        self.stop_flag = False
        self.start_stats_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.sort_tables_btn.config(state='disabled')
        self.progress_manage.start()
        self.log_text_manage.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._do_statistics)
        thread.daemon = True
        thread.start()
    
    def stop_statistics(self):
        """停止统计"""
        self.stop_flag = True
        self.log_manage("正在停止...")
    
    def _restore_manage_ui(self):
        self.start_stats_btn.config(state='normal')
        self.stop_btn.config(state='disabled')
        self.update_offers_btn.config(state='normal')
        self.sort_tables_btn.config(state='normal')
        self.progress_manage.stop()
        self.update_progress_manage("")
    
    def start_sort_tables(self):
        """开始offer顺序整理"""
        self.stop_flag = False
        self.start_stats_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.update_offers_btn.config(state='disabled')
        self.sort_tables_btn.config(state='disabled')
        self.progress_manage.start()
        self.log_text_manage.delete(1.0, tk.END)
        
        thread = threading.Thread(target=self._do_sort_tables)
        thread.daemon = True
        thread.start()
    
    def _do_sort_tables(self):
        """执行offer顺序整理操作"""
        try:
            self.log_manage("=" * 50)
            self.log_manage("开始Offer顺序整理...")
            self.log_manage("=" * 50)
            
            # 获取飞书Token
            feishu_token = self.get_feishu_token()
            if not feishu_token:
                self.log_manage("✗ 获取飞书Token失败")
                return
            
            # 排序Offer表格
            self.log_manage("\n【1/2】对Offer表格行排序")
            self.update_progress_manage("排序Offer表...")
            try:
                self.sort_offer_table(feishu_token)
            except Exception as e:
                self.log_manage(f"  Offer表格排序出错: {str(e)}")
                import traceback
                self.log_manage(traceback.format_exc())
            
            if self.stop_flag:
                self.log_manage("\n已停止")
                return
            
            # 排序广告系列表格
            self.log_manage("\n【2/2】对广告系列表格行排序")
            self.update_progress_manage("排序广告系列表...")
            try:
                self.sort_campaigns_table(feishu_token)
            except Exception as e:
                self.log_manage(f"  广告系列表格排序出错: {str(e)}")
                import traceback
                self.log_manage(traceback.format_exc())
            
            self.log_manage("\n" + "=" * 50)
            self.log_manage("✅ Offer顺序整理完成！")
            self.log_manage("=" * 50)
            
        except Exception as e:
            self.log_manage(f"\n错误: {str(e)}")
            import traceback
            self.log_manage(traceback.format_exc())
        finally:
            self.root.after(0, self._restore_manage_ui)
    
    def start_update_offers(self):
        """开始更新已有offer"""
        self.stop_flag = False
        self.start_stats_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.update_offers_btn.config(state='disabled')
        self.sort_tables_btn.config(state='disabled')
        self.progress_manage.start()
        
        thread = threading.Thread(target=self._do_update_offers)
        thread.daemon = True
        thread.start()
    
    def _do_update_offers(self):
        """执行更新已有offer操作（按品牌批量获取）"""
        try:
            self.log_manage("=" * 50)
            self.log_manage("开始更新已有offer（按品牌批量获取）...")
            self.log_manage("=" * 50)
            
            # 获取飞书访问令牌
            token = self.get_feishu_token()
            if not token:
                self.log_manage("获取飞书访问令牌失败")
                return
            
            spreadsheet_token = self.config.get('feishu_spreadsheet_token', '')
            sheet_id = self.config.get('feishu_sheet_id', '')
            
            if not spreadsheet_token or not sheet_id:
                self.log_manage("飞书电子表格配置缺失")
                return
            
            # 读取飞书数据
            self.log_manage("\n【步骤1】读取飞书现有offer数据...")
            self.update_progress_manage("读取飞书数据...")
            
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!A1:Z2000"
            headers = {"Authorization": f"Bearer {token}"}
            resp = requests.get(url, headers=headers)
            data = resp.json()
            
            if data.get('code') != 0:
                self.log_manage(f"读取飞书数据失败: {data.get('msg')}")
                return
            
            values = data.get('data', {}).get('valueRange', {}).get('values', [])
            if len(values) < 2:
                self.log_manage("飞书表格数据不足")
                return
            
            headers_row = values[0]
            
            # 找到关键列的索引
            col_indices = {}
            target_cols = ['ASIN', '国家代码', '品牌名称', '品牌ID', '佣金', '库存状态', '更新时间', 
                          '投放链接', '折扣价', '产品名称', '产品链接', '原价', '货币']
            for i, h in enumerate(headers_row):
                h_str = str(h).strip() if h else ''
                if h_str in target_cols:
                    col_indices[h_str] = i
            
            # 输出检测到的列索引（调试用）
            self.log_manage(f"  检测到的列: {col_indices}")
            
            if 'ASIN' not in col_indices:
                self.log_manage("未找到ASIN列")
                return
            
            # 获取表格实际的列数（用于验证）
            max_col_idx = len(headers_row) - 1
            
            # 收集飞书现有offer和品牌ID
            # 修改：支持同一个(asin, country)有多行（复制的offer）
            existing_offers = {}  # (asin, country) -> [(row_idx, offer_data), ...]  改为列表
            existing_asins = {}   # asin -> (row_idx, country)  用于ASIN回退匹配
            brand_ids = set()
            
            for row_idx, row in enumerate(values[1:], start=2):
                asin_idx = col_indices.get('ASIN')
                country_idx = col_indices.get('国家代码')
                brand_id_idx = col_indices.get('品牌ID')
                
                asin = str(row[asin_idx]).strip() if asin_idx is not None and asin_idx < len(row) else ''
                country = str(row[country_idx]).strip().upper() if country_idx is not None and country_idx < len(row) else ''
                brand_id = str(row[brand_id_idx]).strip() if brand_id_idx is not None and brand_id_idx < len(row) else ''
                
                # 跳过"总计"行或空ASIN行
                if not asin or asin.lower() in ['none', '总计', 'total']:
                    continue
                
                key = (asin, country)
                if key not in existing_offers:
                    existing_offers[key] = []
                existing_offers[key].append((row_idx, row))
                existing_asins[asin] = (row_idx, country)
                # 只添加有效的brand_id（过滤掉空值和None）
                if brand_id and brand_id.lower() not in ['none', '']:
                    brand_ids.add(brand_id)
            
            # 统计唯一的(asin, country)组合数和总行数
            total_rows = sum(len(rows) for rows in existing_offers.values())
            self.log_manage(f"  飞书现有offer: {total_rows} 行（{len(existing_offers)} 个唯一组合）")
            self.log_manage(f"  涉及品牌ID: {len(brand_ids)} 个")
            
            if self.stop_flag:
                self.log_manage("已停止")
                return
            
            # 按品牌ID获取PB的offer数据
            self.log_manage("\n【步骤2】按品牌ID获取PB offer数据...")
            self.update_progress_manage("获取PB数据...")
            
            all_pb_offers = {}  # (asin, country) -> offer_data
            
            for i, brand_id in enumerate(brand_ids):
                if self.stop_flag:
                    self.log_manage("已停止")
                    return
                
                self.log_manage(f"  获取品牌 {brand_id} 的offer... ({i+1}/{len(brand_ids)})")
                offers = self.get_offers_by_brand(brand_id)
                self.log_manage(f"    品牌 {brand_id}: PB返回 {len(offers)} 个offer")
                
                for offer in offers:
                    asin = str(offer.get('asin', '')).strip()
                    pb_country = str(offer.get('country_code', '')).strip().upper()
                    brand_name = str(offer.get('brand_name', '')).strip()
                    
                    if not asin:
                        continue
                    
                    # 原始国家代码（用于调试）
                    original_pb_country = pb_country
                    
                    # 如果PB的country_code为空，尝试从品牌名称中提取国家代码
                    if not pb_country and brand_name:
                        extracted_country = self.extract_country_from_merchant_name(brand_name)
                        if extracted_country:
                            pb_country = extracted_country
                    
                    # 尝试匹配：只做精确匹配，不做ASIN回退（避免国家混淆）
                    matched_country = pb_country  # 默认用PB返回的（或从品牌名提取的）国家代码
                    
                    # 如果PB有明确的国家代码，检查是否存在精确匹配
                    if pb_country:
                        if (asin, pb_country) in existing_offers:
                            matched_country = pb_country
                        elif asin in existing_asins:
                            # PB有国家但飞书里是另一个国家，不强制回退
                            feishu_country = existing_asins[asin][1]
                            if feishu_country != pb_country:
                                # 跳过，避免用错误国家的数据覆盖
                                continue
                    else:
                        # PB没有国家代码，尝试用飞书里该ASIN的国家
                        if asin in existing_asins:
                            matched_country = existing_asins[asin][1]
                        else:
                            matched_country = 'US'  # 默认US
                    
                    offer['_matched_country'] = matched_country
                    offer['_original_pb_country'] = original_pb_country
                    all_pb_offers[(asin, matched_country)] = offer
                
                time.sleep(0.2)  # 品牌间间隔
            
            self.log_manage(f"  PB总共返回 {len(all_pb_offers)} 个有效offer")
            
            if self.stop_flag:
                self.log_manage("已停止")
                return
            
            # 对比并更新
            self.log_manage("\n【步骤3】对比并更新offer...")
            self.update_progress_manage("对比更新...")
            
            updates = []  # (row_idx, col_idx, new_value)
            style_updates = []  # (row_idx, col_idx, style_type)  style_type: 'red_bold' or 'black_normal'
            updated_count = 0
            unchanged_count = 0
            new_offers = []
            
            # 检查现有offer的更新
            # 修改：existing_offers现在是 {(asin, country): [(row_idx, row), ...]}
            for (asin, country), rows_list in existing_offers.items():
                pb_offer = all_pb_offers.get((asin, country))
                
                # 对该(asin, country)的所有行进行相同的更新
                for row_idx, row in rows_list:
                    # 调试前几行
                    if row_idx <= 5:
                        self.log_manage(f"    调试: 第{row_idx}行 ASIN={asin} 国家={country} PB匹配={'是' if pb_offer else '否'} 行数据长度={len(row)}")
                    
                    if not pb_offer:
                        # PB未返回此offer，可能已下架，标记为OUT_OF_STOCK
                        current_stock = str(row[col_indices.get('库存状态', -1)]).strip() if col_indices.get('库存状态', -1) < len(row) else ''
                        if current_stock != 'OUT_OF_STOCK':
                            stock_col = col_indices.get('库存状态')
                            if stock_col is not None:
                                updates.append((row_idx, stock_col, 'OUT_OF_STOCK'))
                                style_updates.append((row_idx, stock_col, 'red_bold'))
                                self.log_manage(f"  [下架] {asin}_{country} (行{row_idx}): PB未返回，标记为OUT_OF_STOCK")
                                updated_count += 1
                        continue
                    
                    # 对比字段
                    changes = []
                    
                    # 佣金对比
                    commission_col = col_indices.get('佣金')
                    if commission_col is not None:
                        # 安全获取旧佣金值
                        if commission_col < len(row):
                            old_commission = str(row[commission_col]).strip()
                        else:
                            old_commission = ''
                            # 调试：行数据不够长
                            if row_idx <= 5:  # 只打印前几行的调试信息
                                self.log_manage(f"    调试: 第{row_idx}行数据长度{len(row)}，佣金列索引{commission_col}超出范围")
                        new_commission = str(pb_offer.get('commission', '')).strip()
                        
                        # 标准化佣金比较
                        old_val = self.normalize_commission(old_commission)
                        new_val = self.normalize_commission(new_commission)
                        
                        if abs(old_val - new_val) > 0.0001:
                            updates.append((row_idx, commission_col, new_commission))
                            changes.append(f"佣金: {old_commission} → {new_commission}")
                            if new_val < old_val:
                                style_updates.append((row_idx, commission_col, 'red_bold'))
                    
                    # 库存状态对比
                    stock_col = col_indices.get('库存状态')
                    if stock_col is not None:
                        old_stock = str(row[stock_col]).strip() if stock_col < len(row) else ''
                        new_stock = str(pb_offer.get('availability', '')).strip()
                        
                        if old_stock != new_stock:
                            updates.append((row_idx, stock_col, new_stock))
                            changes.append(f"库存: {old_stock} → {new_stock}")
                            if new_stock == 'OUT_OF_STOCK':
                                style_updates.append((row_idx, stock_col, 'red_bold'))
                            elif old_stock == 'OUT_OF_STOCK' and new_stock == 'IN_STOCK':
                                style_updates.append((row_idx, stock_col, 'black_normal'))
                    
                    # 更新时间对比
                    update_time_col = col_indices.get('更新时间')
                    if update_time_col is not None:
                        old_time = str(row[update_time_col]).strip() if update_time_col < len(row) else ''
                        new_time = str(pb_offer.get('update_time', '')).strip()  # PB返回的是update_time
                        
                        if old_time != new_time and new_time:
                            updates.append((row_idx, update_time_col, new_time))
                    
                    if changes:
                        brand_name = pb_offer.get('brand_name', '')
                        pb_country_info = pb_offer.get('_original_pb_country', '')
                        rows_count = len(rows_list)
                        if rows_count > 1:
                            self.log_manage(f"  [更新] {asin}_{country} (行{row_idx}, 共{rows_count}行): {', '.join(changes)} [PB品牌:{brand_name}, PB国家:{pb_country_info}]")
                        else:
                            self.log_manage(f"  [更新] {asin}_{country}: {', '.join(changes)} [PB品牌:{brand_name}, PB国家:{pb_country_info}]")
                        updated_count += 1
                    else:
                        unchanged_count += 1
            
            # 检查新offer
            for (asin, country), pb_offer in all_pb_offers.items():
                # 检查是否是真正的新offer（ASIN在飞书中完全不存在）
                if asin not in existing_asins:
                    # 过滤掉佣金过低的offer
                    commission_val = self.normalize_commission(pb_offer.get('commission', '0'))
                    if commission_val <= 0.0001:
                        continue
                    new_offers.append(pb_offer)
            
            if self.stop_flag:
                self.log_manage("已停止")
                return
            
            # 应用更新
            if updates:
                self.log_manage(f"\n【步骤4】应用更新到飞书... ({len(updates)} 个单元格)")
                self.update_progress_manage("写入飞书...")
                self._apply_offer_updates(token, updates, style_updates)
            
            # 添加新offer
            if new_offers:
                self.log_manage(f"\n【步骤5】添加新offer到飞书... ({len(new_offers)} 个)")
                self.update_progress_manage("添加新offer...")
                self._append_new_offers(token, new_offers, col_indices)
            
            # 输出统计
            self.log_manage("\n" + "=" * 50)
            self.log_manage("更新完成！统计信息：")
            self.log_manage(f"  已更新offer: {updated_count} 个")
            self.log_manage(f"  无变化offer: {unchanged_count} 个")
            self.log_manage(f"  新增offer: {len(new_offers)} 个")
            self.log_manage("=" * 50)
            
        except Exception as e:
            self.log_manage(f"更新offer时发生错误: {str(e)}")
            import traceback
            self.log_manage(traceback.format_exc())
        finally:
            self._restore_manage_ui()
    
    def get_offers_by_brand(self, brand_id):
        """按品牌ID获取offer（使用与获取offer功能相同的参数格式）"""
        all_offers = []
        page = 1
        has_more = True
        
        # 跳过无效的brand_id
        if not brand_id or str(brand_id).strip() in ['', 'None']:
            return all_offers
        
        url = f"{PB_API_BASE_URL}/api/datafeed/get_fba_products"
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        
        while has_more:
            try:
                # 使用与获取offer功能相同的参数格式
                request_body = {
                    "token": self.pb_token_var.get().strip(),
                    "page_size": 100,
                    "page": page,
                    "default_filter": 0,
                    "country_code": "",
                    "brand_id": str(brand_id).strip(),  # 字符串格式
                    "sort": "",
                    "asins": "",
                    "relationship": 1,  # 重要：必须是1
                    "is_original_currency": 0,
                    "has_promo_code": 0,
                    "has_acc": 0,
                    "filter_sexual_wellness": 0
                }
                
                resp = requests.post(url, json=request_body, headers=headers, timeout=60)
                
                if resp.status_code == 200:
                    data = resp.json()
                    
                    if data.get("status", {}).get("code") == 0:
                        offers = data.get("data", {}).get("list", [])
                        has_more = data.get("data", {}).get("has_more", False)
                        all_offers.extend(offers)
                        page += 1
                    else:
                        self.log_manage(f"    API错误: {data.get('status', {}).get('msg')}")
                        break
                else:
                    self.log_manage(f"    HTTP错误: {resp.status_code}")
                    break
                
                time.sleep(0.1)
                
            except Exception as e:
                self.log_manage(f"    异常: {str(e)}")
                break
        
        return all_offers
    
    def _apply_offer_updates(self, token, updates, style_updates):
        """应用offer更新到飞书"""
        spreadsheet_token = self.config.get('feishu_spreadsheet_token', '')
        sheet_id = self.config.get('feishu_sheet_id', '')
        
        # 列索引转换为列字母的辅助函数（支持超过26列）
        def col_idx_to_letter(idx):
            result = ""
            while idx >= 0:
                result = chr(ord('A') + (idx % 26)) + result
                idx = idx // 26 - 1
            return result
        
        # 过滤掉超出合理范围的更新（假设表格最多26列A-Z）
        max_col = 25  # 最大列索引（Z列）
        valid_updates = [(r, c, v) for r, c, v in updates if c <= max_col]
        
        if len(valid_updates) < len(updates):
            self.log_manage(f"    警告: 跳过 {len(updates) - len(valid_updates)} 个超出列范围的更新")
        
        # 批量更新值
        batch_size = 50
        for i in range(0, len(valid_updates), batch_size):
            batch = valid_updates[i:i+batch_size]
            value_ranges = []
            
            for row_idx, col_idx, value in batch:
                col_letter = col_idx_to_letter(col_idx)
                # 飞书API要求range格式为 sheetId!开始位置:结束位置
                range_str = f"{sheet_id}!{col_letter}{row_idx}:{col_letter}{row_idx}"
                value_ranges.append({
                    "range": range_str,
                    "values": [[value]]
                })
            
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            body = {"valueRanges": value_ranges}
            
            resp = requests.post(url, headers=headers, json=body)
            result = resp.json()
            
            if result.get('code') != 0:
                self.log_manage(f"    批量更新失败: {result.get('msg')}")
            
            time.sleep(0.2)
        
        # 应用样式更新（也过滤）
        if style_updates:
            valid_style_updates = [(r, c, s) for r, c, s in style_updates if c <= max_col]
            self._apply_style_updates(token, valid_style_updates)
    
    def _apply_style_updates(self, token, style_updates):
        """应用样式更新（红色加粗或黑色普通）"""
        spreadsheet_token = self.config.get('feishu_spreadsheet_token', '')
        sheet_id = self.config.get('feishu_sheet_id', '')
        
        # 列索引转换为列字母
        def col_idx_to_letter(idx):
            result = ""
            while idx >= 0:
                result = chr(ord('A') + (idx % 26)) + result
                idx = idx // 26 - 1
            return result
        
        for row_idx, col_idx, style_type in style_updates:
            col_letter = col_idx_to_letter(col_idx)
            range_str = f"{sheet_id}!{col_letter}{row_idx}:{col_letter}{row_idx}"
            
            if style_type == 'red_bold':
                style = {
                    "font": {
                        "bold": True
                    },
                    "foreColor": "#FF0000"  # 红色字体
                }
            else:  # black_normal
                style = {
                    "font": {
                        "bold": False
                    },
                    "foreColor": "#000000"  # 黑色字体
                }
            
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/style"
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            body = {
                "appendStyle": {
                    "range": range_str,
                    "style": style
                }
            }
            
            resp = requests.put(url, headers=headers, json=body)
            result = resp.json()
            if result.get('code') != 0:
                self.log_manage(f"    样式更新失败 {col_letter}{row_idx}: {result.get('msg')}")
            time.sleep(0.1)
    
    def _append_new_offers(self, token, new_offers, col_indices):
        """添加新offer到飞书表格末尾（使用动态列索引）"""
        spreadsheet_token = self.config.get('feishu_spreadsheet_token', '')
        sheet_id = self.config.get('feishu_sheet_id', '')
        
        # 过滤掉佣金过低的offer
        valid_offers = []
        for offer in new_offers:
            commission_val = self.normalize_commission(offer.get('commission', '0'))
            if commission_val > 0.0001:
                valid_offers.append(offer)
        
        if not valid_offers:
            self.log_manage("  没有有效的新offer需要添加")
            return
        
        # 找到第一个空白行（使用ASIN列的索引）
        asin_col_idx = col_indices.get('ASIN', 0)
        first_empty_row = self._find_first_empty_row(token, spreadsheet_token, sheet_id, asin_col_idx)
        if first_empty_row is None:
            self.log_manage("  无法确定空白行位置")
            return
        
        self.log_manage(f"  从第 {first_empty_row} 行开始添加新offer")
        
        # 确保有足够的行
        needed_rows = first_empty_row + len(valid_offers)
        self._ensure_sheet_rows(token, spreadsheet_token, sheet_id, needed_rows)
        
        # 列索引转换为列字母的辅助函数
        def col_idx_to_letter(idx):
            result = ""
            while idx >= 0:
                result = chr(ord('A') + (idx % 26)) + result
                idx = idx // 26 - 1
            return result
        
        # 要写入的字段映射：列名 -> PB字段名（根据PB API返回的实际字段名）
        # 注意：产品链接不从PB直接获取(PB返回的是简单URL)，而是通过投放链接重定向获取完整版
        field_mapping = {
            '投放链接': '_tracking_link',  # 特殊处理
            '国家代码': '_country',  # 特殊处理
            '品牌名称': 'brand_name',
            '折扣价': 'discount_price',  # PB返回的是discount_price
            '佣金': 'commission',
            '产品名称': 'product_name',
            'ASIN': 'asin',
            '库存状态': 'availability',
            '产品链接': '_product_link',  # 特殊处理：通过投放链接重定向获取
            '品牌ID': 'brand_id',
            '更新时间': 'update_time',  # PB返回的是update_time
            '原价': 'original_price',  # PB返回的是original_price
            '货币': 'currency'
        }
        
        # 批量写入
        batch_size = 50
        current_row = first_empty_row
        
        for i in range(0, len(valid_offers), batch_size):
            batch = valid_offers[i:i+batch_size]
            value_ranges = []
            
            for offer in batch:
                asin = offer.get('asin', '')
                
                # 确定国家代码
                country = offer.get('country_code', '').strip().upper()
                if not country:
                    brand_name = offer.get('brand_name', '')
                    country = self.extract_country_from_merchant_name(brand_name)
                if not country:
                    country = 'US'
                
                # 生成投放链接
                uid = self.generate_random_uid()
                tracking_link = self.get_partnerboost_link(asin, country, uid)
                
                if not tracking_link:
                    self.log_manage(f"    警告: {asin}_{country} 无法获取投放链接")
                
                # 解析投放链接重定向，获取完整产品链接
                product_link = ''
                if tracking_link:
                    product_link = self.resolve_redirect_url(tracking_link)
                
                # 为每个需要写入的列创建单独的range
                for col_name, pb_field in field_mapping.items():
                    col_idx = col_indices.get(col_name)
                    if col_idx is None:
                        continue  # 该列不存在于表格中
                    
                    # 获取值
                    if col_name == '投放链接':
                        value = tracking_link or ''
                    elif col_name == '国家代码':
                        value = country
                    elif col_name == 'ASIN':
                        value = asin
                    elif col_name == '产品链接':
                        value = product_link or ''
                    else:
                        value = offer.get(pb_field, '')
                    
                    col_letter = col_idx_to_letter(col_idx)
                    range_str = f"{sheet_id}!{col_letter}{current_row}:{col_letter}{current_row}"
                    value_ranges.append({
                        "range": range_str,
                        "values": [[value]]
                    })
                
                self.log_manage(f"    新增: {asin}_{country} - {offer.get('brand_name', '')}")
                current_row += 1
            
            # 批量写入
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            body = {"valueRanges": value_ranges}
            
            resp = requests.post(url, headers=headers, json=body)
            result = resp.json()
            
            if result.get('code') != 0:
                self.log_manage(f"    批量写入失败: {result.get('msg')}")
            else:
                self.log_manage(f"    成功写入 {len(batch)} 个新offer")
            
            time.sleep(0.3)
    
    def _find_first_empty_row(self, token, spreadsheet_token, sheet_id, asin_col_idx):
        """找到第一个空白行（基于ASIN列，只要文本为空就算空白）"""
        # 列索引转换为列字母
        def col_idx_to_letter(idx):
            result = ""
            while idx >= 0:
                result = chr(ord('A') + (idx % 26)) + result
                idx = idx // 26 - 1
            return result
        
        col_letter = col_idx_to_letter(asin_col_idx)
        self.log_manage(f"    查找空白行: 检查ASIN列 {col_letter} (索引{asin_col_idx})")
        
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!{col_letter}1:{col_letter}3000"
        headers = {"Authorization": f"Bearer {token}"}
        
        resp = requests.get(url, headers=headers)
        data = resp.json()
        
        if data.get('code') != 0:
            self.log_manage(f"    读取失败: {data.get('msg')}")
            return None
        
        values = data.get('data', {}).get('valueRange', {}).get('values', [])
        self.log_manage(f"    ASIN列读取到 {len(values)} 行")
        
        # 显示前几行数据用于调试
        if values and len(values) > 0:
            self.log_manage(f"    前5行ASIN数据: {[str(v[0] if v else '') for v in values[:5]]}")
        
        # 找最后一个有数据的行
        last_data_row = 0
        for i, row in enumerate(values):
            cell_value = str(row[0]).strip() if row and len(row) > 0 else ''
            # 跳过表头、空值、None
            if cell_value and cell_value.lower() not in ['none', '', 'asin']:
                last_data_row = i + 1  # 1-based
        
        # 第一个空白行 = 最后有数据的行 + 1
        first_empty = last_data_row + 1
        self.log_manage(f"    最后有数据行: {last_data_row}, 第一个空白行: {first_empty}")
        return first_empty
    
    def _ensure_sheet_rows(self, token, spreadsheet_token, sheet_id, min_rows):
        """确保工作表有足够的行数"""
        # 获取当前工作表信息
        url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}"
        headers = {"Authorization": f"Bearer {token}"}
        
        resp = requests.get(url, headers=headers)
        data = resp.json()
        
        current_rows = data.get('data', {}).get('sheet', {}).get('grid_properties', {}).get('row_count', 1000)
        
        if current_rows < min_rows:
            # 需要扩展行数
            rows_to_add = min_rows - current_rows + 100  # 多加100行buffer
            
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/dimension_range"
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            body = {
                "dimension": {
                    "sheetId": sheet_id,
                    "majorDimension": "ROWS",
                    "startIndex": current_rows,
                    "endIndex": current_rows + rows_to_add
                }
            }
            
            resp = requests.post(url, headers=headers, json=body)
            result = resp.json()
            
            if result.get('code') == 0:
                self.log_manage(f"    已扩展工作表行数: {current_rows} → {current_rows + rows_to_add}")
            else:
                self.log_manage(f"    扩展行数失败: {result.get('msg')}")
    
    def normalize_commission(self, val):
        """标准化佣金值为小数形式"""
        if not val:
            return 0.0
        val_str = str(val).strip()
        if '%' in val_str:
            try:
                return float(val_str.replace('%', '')) / 100
            except:
                return 0.0
        else:
            try:
                f = float(val_str)
                if f > 1:
                    return f / 100
                return f
            except:
                return 0.0
    
    def _do_statistics(self):
        """执行统计操作"""
        try:
            self.log_manage("=" * 50)
            self.log_manage("开始Offer统计...")
            self.log_manage("=" * 50)
            
            # 步骤1：获取Google Ads广告系列信息
            self.log_manage("\n【步骤1】获取Google Ads广告系列信息")
            self.update_progress_manage("获取广告系列...")
            
            campaign_data, succeeded_account_ids = self.get_all_campaigns_with_asin()
            if self.stop_flag:
                self.log_manage("已停止")
                return
            
            self.log_manage(f"  获取到 {len(campaign_data)} 个广告系列（来自 {len(succeeded_account_ids)} 个账户）")
            
            # 调试：显示前5个广告系列的信息
            for i, c in enumerate(campaign_data[:5]):
                name = c.get('campaign_name', '')[:40]
                self.log_manage(f"    [{i+1}] {name}...")
                urls = c.get('final_urls', [])
                self.log_manage(f"        Final URL: {urls[0][:80] if urls else 'None'}...")
                self.log_manage(f"        提取: ASIN={c.get('asin', 'None')}, 国家={c.get('country', 'None')}, UIDs={c.get('uids', [])}, LinkIDs={len(c.get('link_ids', []))}")
            
            # 构建(ASIN, 国家) -> 广告系列映射
            asin_country_campaigns = {}  # {(asin, country): [campaign_info, ...]}
            campaigns_without_asin = 0
            campaigns_without_country = 0
            for campaign in campaign_data:
                asin = campaign.get('asin')
                country = campaign.get('country')
                if asin and country:
                    key = (asin, country)
                    if key not in asin_country_campaigns:
                        asin_country_campaigns[key] = []
                    asin_country_campaigns[key].append(campaign)
                elif not asin:
                    campaigns_without_asin += 1
                else:
                    campaigns_without_country += 1
            
            self.log_manage(f"  关联到 {len(asin_country_campaigns)} 个(ASIN+国家)组合")
            self.log_manage(f"  无法提取ASIN的广告系列: {campaigns_without_asin} 个")
            self.log_manage(f"  无法提取国家的广告系列: {campaigns_without_country} 个")
            if asin_country_campaigns:
                sample_keys = list(asin_country_campaigns.keys())[:5]
                self.log_manage(f"  样例: {sample_keys}...")
            
            # 详细显示每个广告系列的ASIN和国家提取情况（写入调试日志）
            self.log_debug(f"=== 广告系列详情（共{len(campaign_data)}个）===")
            for i, campaign in enumerate(campaign_data):
                c_name = campaign.get('campaign_name', '')
                c_asin = campaign.get('asin', 'None')
                c_country = campaign.get('country', 'None')
                c_status = campaign.get('status', 'Unknown')
                c_cost = campaign.get('cost_usd', 0)
                self.log_debug(f"  [{i+1}] {c_name} | ASIN={c_asin} | 国家={c_country} | 状态={c_status} | 花费=${c_cost:.2f}")
            
            # 步骤2：获取飞书表格数据
            self.log_manage("\n【步骤2】获取飞书表格数据")
            self.update_progress_manage("获取飞书数据...")
            
            feishu_token = self.get_feishu_token()
            if not feishu_token:
                self.log_manage("  ✗ 获取飞书Token失败")
                return
            
            feishu_data = self.get_feishu_sheet_data(feishu_token)
            if self.stop_flag:
                return
            
            self.log_manage(f"  获取到 {len(feishu_data)} 行数据")
            
            # 调试：显示表头和前3行数据
            if feishu_data:
                first_row = feishu_data[0]
                self.log_manage(f"  表头字段: {list(first_row.keys())}")
                for i, row in enumerate(feishu_data[:3]):
                    asin_val = row.get('ASIN', 'N/A')
                    status_val = row.get('状态', 'N/A')
                    self.log_manage(f"    [{i+1}] ASIN={asin_val}, 状态={status_val}")
            
            # 读取当前"总计"行的数值（用于后续计算新增量）
            previous_total_cost = 0
            previous_total_commission = 0
            for row in feishu_data:
                if row.get('row_index') == 2:  # 总计行
                    # 读取当前总花费
                    cost_val = row.get('广告系列总花费', '')
                    if cost_val:
                        try:
                            # 处理格式：$1838.68 或 1838.68
                            cost_str = str(cost_val).replace('$', '').replace(',', '').strip()
                            # 如果有括号（如"$1143.82（$35.63）"），只取括号前的部分
                            if '（' in cost_str:
                                cost_str = cost_str.split('（')[0]
                            if cost_str:
                                previous_total_cost = float(cost_str)
                        except (ValueError, TypeError):
                            pass
                    
                    # 读取当前总佣金
                    commission_val = row.get('总佣金', '')
                    if commission_val:
                        try:
                            commission_str = str(commission_val).replace('$', '').replace(',', '').strip()
                            # 支持新格式: "4481.79 (+133.91-278.33)" -> 取第一个数字
                            # 也支持旧格式: "1143.82（35.63+80.93）" 或 "1143.82"
                            if ' (' in commission_str:
                                commission_str = commission_str.split(' (')[0]
                            elif '（' in commission_str:
                                commission_str = commission_str.split('（')[0]
                            if commission_str:
                                previous_total_commission = float(commission_str)
                        except (ValueError, TypeError):
                            pass
                    break
            
            self.log_manage(f"  当前总计行: 总花费=${previous_total_cost:.2f}, 总佣金=${previous_total_commission:.2f}")
            
            # 步骤3：获取PartnerBoost佣金数据
            self.log_manage("\n【步骤3】获取PartnerBoost佣金数据")
            self.update_progress_manage("获取佣金数据...")
            
            commission_data = self.get_all_commissions()
            if self.stop_flag:
                return
            
            # 构建(ASIN, 国家) -> 佣金映射（排除Rejected）
            asin_country_commission = {}  # {(asin, country): total_commission}
            
            # 调试：显示前几条交易数据的完整内容
            if commission_data:
                self.log_manage(f"  交易数据字段: {list(commission_data[0].keys()) if commission_data else '无'}")
                self.log_manage(f"  === 前3条交易完整数据 ===")
                for i, trans in enumerate(commission_data[:3]):
                    self.log_manage(f"    [{i+1}] 完整数据:")
                    for k, v in trans.items():
                        if v:  # 只显示非空字段
                            self.log_manage(f"        {k}: {v}")
            
            skipped_no_asin = 0
            skipped_rejected = 0
            rejected_commission_total = 0  # Rejected状态的佣金总额
            country_from_customer = 0      # 从 customer_country 字段获取
            country_from_merchant = 0      # 从 merchant_name 提取
            country_from_mcid = 0          # 从 mcid 提取
            country_default_us = 0         # 默认为 US
            
            # 构建四种佣金映射（严格区分uid和非uid，避免重复计算）：
            # 1. (ASIN, 国家) -> 佣金（所有交易的总和，仅用于总计计算）
            # 2. (ASIN, 国家, uid) -> 佣金（仅有uid的交易，用于uid精确匹配）
            # 3. (ASIN, 国家) -> 佣金（仅无uid的交易，用于ASIN+国家匹配）
            # 4. ASIN -> 佣金（仅无uid且国家默认US的交易，用于ASIN-only备选匹配）
            asin_country_uid_commission = {}  # {(asin, country, uid): total_commission} 有uid的交易
            asin_country_no_uid_commission = {}  # {(asin, country): total_commission} 无uid的交易
            asin_only_commission = {}  # {asin: total_commission} 无uid且国家默认US
            
            for trans in commission_data:
                if trans.get('status') == 'Rejected':
                    skipped_rejected += 1
                    # 累加Rejected佣金
                    rejected_comm = float(trans.get('sale_comm', 0) or 0)
                    rejected_commission_total += rejected_comm
                    continue
                    
                asin = trans.get('prod_id', '')
                merchant_name = trans.get('merchant_name', '')
                mcid = trans.get('mcid', '')
                uid = trans.get('uid', '')  # 获取uid
                
                # 优先从 customer_country 字段获取国家代码
                country = trans.get('customer_country', '') or trans.get('geo', '') or trans.get('country', '')
                country_source = 'customer_country'
                
                if not country and merchant_name:
                    # 尝试从 merchant_name 提取国家代码（如 "DE-Anker" -> "DE"）
                    country = self.extract_country_from_merchant_name(merchant_name)
                    if country:
                        country_source = 'merchant_name'
                
                if not country and mcid:
                    # 尝试从 mcid 提取国家代码（如 "amzdeanker" -> "DE"）
                    country = self.extract_country_from_mcid(mcid)
                    if country:
                        country_source = 'mcid'
                
                if not country:
                    # 品牌名中无国家代码，默认为 US
                    country = 'US'
                    country_source = 'default_us'
                
                if not asin:
                    skipped_no_asin += 1
                    continue
                
                comm = float(trans.get('sale_comm', 0) or 0)
                
                # 统计国家代码来源
                if country_source == 'customer_country':
                    country_from_customer += 1
                elif country_source == 'merchant_name':
                    country_from_merchant += 1
                elif country_source == 'mcid':
                    country_from_mcid += 1
                else:
                    country_default_us += 1
                
                country = country.upper()
                key = (asin, country)
                
                # 总映射（所有交易，用于总计计算）
                asin_country_commission[key] = asin_country_commission.get(key, 0) + comm
                
                # 严格按uid分池
                if uid:
                    # 有uid的交易 → uid池（按uid精确匹配到offer行）
                    uid_key = (asin, country, uid)
                    asin_country_uid_commission[uid_key] = asin_country_uid_commission.get(uid_key, 0) + comm
                else:
                    # 无uid的交易 → 非uid池（按ASIN+国家匹配，每个组合只分配一次）
                    asin_country_no_uid_commission[key] = asin_country_no_uid_commission.get(key, 0) + comm
                    # 对于国家默认US的情况，额外加入ASIN-only备选映射
                    if country_source == 'default_us':
                        asin_only_commission[asin] = asin_only_commission.get(asin, 0) + comm
            
            self.log_manage(f"  获取到 {len(commission_data)} 条交易")
            self.log_manage(f"  佣金池分布:")
            self.log_manage(f"    • UID池(按uid匹配): {len(asin_country_uid_commission)} 个uid组合, 金额=${sum(asin_country_uid_commission.values()):.2f}")
            self.log_manage(f"    • 非UID池(按ASIN+国家匹配): {len(asin_country_no_uid_commission)} 个组合, 金额=${sum(asin_country_no_uid_commission.values()):.2f}")
            self.log_manage(f"    • ASIN-only备选池: {len(asin_only_commission)} 个ASIN")
            self.log_manage(f"    • 合计(非Rejected): {len(asin_country_commission)} 个(ASIN+国家)组合, 金额=${sum(asin_country_commission.values()):.2f}")
            self.log_manage(f"  国家代码来源统计:")
            self.log_manage(f"    • 从customer_country字段: {country_from_customer} 条")
            self.log_manage(f"    • 从merchant_name提取: {country_from_merchant} 条")
            self.log_manage(f"    • 从mcid提取: {country_from_mcid} 条")
            self.log_manage(f"    • 默认为US: {country_default_us} 条")
            if skipped_rejected > 0 or skipped_no_asin > 0:
                self.log_manage(f"  跳过: Rejected={skipped_rejected}条(佣金${rejected_commission_total:.2f}), 无ASIN={skipped_no_asin}")
            if asin_only_commission:
                sample_asins = list(asin_only_commission.keys())[:5]
                self.log_manage(f"  佣金ASIN样例: {sample_asins}")
            # 输出带uid佣金的详细信息
            if asin_country_uid_commission:
                self.log_manage(f"  带UID佣金详情（共{len(asin_country_uid_commission)}条）:")
                for uid_key, comm_value in asin_country_uid_commission.items():
                    asin, country, uid = uid_key
                    self.log_manage(f"    • UID={uid}, ASIN={asin}, 国家={country}, 佣金=${comm_value:.2f}")
            
            # 步骤4：更新飞书表格
            self.log_manage("\n【步骤4】更新飞书表格")
            self.update_progress_manage("更新飞书...")
            
            
            
            # 步骤4.5：通过link_id精确分配广告系列到具体的offer行
            self.log_manage("\n  🔗 通过link_id精确匹配广告系列到offer行...")
            self.update_progress_manage("精确匹配广告系列...")
            
            row_campaigns = self._resolve_campaign_to_row_mapping(
                feishu_data, asin_country_campaigns, campaign_data
            )
            if row_campaigns:
                self.log_manage(f"  已精确分配 {len(row_campaigns)} 行的广告系列")
            
            updates = self.calculate_updates(feishu_data, asin_country_campaigns, asin_country_commission, asin_only_commission, asin_country_uid_commission, asin_country_no_uid_commission, row_campaigns)
            
            self.log_manage(f"  计算得到 {len(updates)} 个需要更新的offer")
            
            if updates:
                # 调试：显示前5个更新
                for i, u in enumerate(updates[:5]):
                    self.log_manage(f"    [{i+1}] ASIN={u.get('asin')}, 国家={u.get('country')}, 状态={u.get('status')}")
                self.apply_feishu_updates(feishu_token, updates)
            else:
                self.log_manage("  没有需要更新的内容")
                self.log_manage("  可能原因：")
                self.log_manage("    1. 广告系列中提取不到ASIN")
                self.log_manage("    2. 飞书表格中的ASIN与广告系列不匹配")
                self.log_manage("    3. 所有offer状态都是'未测试'且没有广告系列")
            
            # 生成统计报告
            self.log_manage("\n" + "=" * 50)
            self.log_manage("【统计报告】")
            self.log_manage("=" * 50)
            
            # 构建飞书表格中所有的 (ASIN, 国家) 组合
            feishu_asin_country_set = set()
            feishu_asin_set = set()
            for row in feishu_data:
                asin = row.get('ASIN', '')
                country = row.get('国家代码', '')
                if asin:
                    feishu_asin_set.add(asin)
                    if country:
                        feishu_asin_country_set.add((asin, country.upper().strip()))
            
            # 1. 未匹配到offer的广告系列统计
            unmatched_campaigns = []
            for key, campaigns_list in asin_country_campaigns.items():
                if key not in feishu_asin_country_set:
                    for c in campaigns_list:
                        unmatched_campaigns.append({
                            'asin': key[0],
                            'country': key[1],
                            'campaign_name': c.get('campaign_name', ''),
                            'cost_usd': c.get('cost_usd', 0)
                        })
            
            # 2. 未匹配到offer的PB佣金订单统计
            unmatched_commission_asin_country = {}  # {(asin, country): commission}
            unmatched_commission_asin_only = {}  # {asin: commission} - ASIN-only匹配也失败的
            
            for key, comm in asin_country_commission.items():
                if key not in feishu_asin_country_set:
                    unmatched_commission_asin_country[key] = comm
            
            for asin, comm in asin_only_commission.items():
                if asin not in feishu_asin_set:
                    unmatched_commission_asin_only[asin] = comm
            
            # 3. 更新统计（按状态分类）
            status_counts = {}
            commission_with_asterisk = 0  # 带"*"标记的佣金数量
            total_cost_updated = 0
            total_commission_updated = 0
            
            for u in updates:
                status = u.get('status', '仅更新佣金')
                status_counts[status] = status_counts.get(status, 0) + 1
                
                if u.get('total_cost'):
                    total_cost_updated += u.get('total_cost', 0)
                
                commission = u.get('commission')
                if commission:
                    if isinstance(commission, str) and '*' in str(commission):
                        commission_with_asterisk += 1
                        # 提取数值部分
                        try:
                            total_commission_updated += float(str(commission).replace('$', '').replace('*', ''))
                        except:
                            pass
                    else:
                        total_commission_updated += float(commission) if commission else 0
            
            # 输出报告
            self.log_manage("\n📊 数据概览:")
            self.log_manage(f"  • Google Ads 广告系列总数: {len(campaign_data)}")
            self.log_manage(f"  • 飞书表格 Offer 总数: {len(feishu_data)}")
            self.log_manage(f"  • PartnerBoost 交易总数: {len(commission_data)}")
            self.log_manage(f"  • 本次更新的 Offer 数: {len(updates)}")
            
            self.log_manage("\n⚠️ 未匹配数据:")
            self.log_manage(f"  • 未匹配到飞书Offer的广告系列: {len(unmatched_campaigns)} 个")
            if unmatched_campaigns:
                # 按(ASIN, 国家)分组显示
                unmatched_by_key = {}
                for c in unmatched_campaigns:
                    key = (c['asin'], c['country'])
                    if key not in unmatched_by_key:
                        unmatched_by_key[key] = []
                    unmatched_by_key[key].append(c)
                
                for (asin, country), campaigns_list in list(unmatched_by_key.items())[:5]:
                    cost_sum = sum(c['cost_usd'] for c in campaigns_list)
                    self.log_manage(f"    - ASIN={asin}, 国家={country}, 广告系列数={len(campaigns_list)}, 花费=${cost_sum:.2f}")
                if len(unmatched_by_key) > 5:
                    self.log_manage(f"    ... 还有 {len(unmatched_by_key) - 5} 个(ASIN+国家)组合未显示")
            
            self.log_manage(f"  • 未匹配到飞书Offer的PB佣金(ASIN+国家): {len(unmatched_commission_asin_country)} 个")
            # 计算ASIN+国家未匹配的佣金总额
            total_unmatched_asin_country_comm = sum(unmatched_commission_asin_country.values()) if unmatched_commission_asin_country else 0
            if unmatched_commission_asin_country:
                self.log_manage(f"    未匹配佣金总额: ${total_unmatched_asin_country_comm:.2f}")
                for (asin, country), comm in list(unmatched_commission_asin_country.items())[:3]:
                    self.log_manage(f"    - ASIN={asin}, 国家={country}, 佣金=${comm:.2f}")
                if len(unmatched_commission_asin_country) > 3:
                    self.log_manage(f"    ... 还有 {len(unmatched_commission_asin_country) - 3} 条未显示")
            
            self.log_manage(f"  • 未匹配到飞书Offer的PB佣金(仅ASIN): {len(unmatched_commission_asin_only)} 个")
            # 计算仅ASIN未匹配的佣金总额
            total_unmatched_asin_only_comm = sum(unmatched_commission_asin_only.values()) if unmatched_commission_asin_only else 0
            if unmatched_commission_asin_only:
                self.log_manage(f"    未匹配佣金总额: ${total_unmatched_asin_only_comm:.2f}")
            
            # 计算未匹配佣金总额：直接用 PB总额(非Rejected) - 已匹配 来避免重复计算
            # （旧方法 asin_country + asin_only 相加会对"ASIN完全不在飞书中"的情况重复计算）
            total_pb_non_rejected_for_unmatched = sum(asin_country_commission.values())
            total_matched_commission = sum(
                comm for key, comm in asin_country_commission.items()
                if key in feishu_asin_country_set
            )
            total_unmatched_commission = total_pb_non_rejected_for_unmatched - total_matched_commission
            self.log_manage(f"  • 未匹配佣金合计: ${total_unmatched_commission:.2f}（= PB总额${total_pb_non_rejected_for_unmatched:.2f} - 已匹配${total_matched_commission:.2f}）")
            self.log_manage(f"  • Rejected佣金合计: ${rejected_commission_total:.2f}")
            
            self.log_manage("\n📈 更新统计:")
            if status_counts:
                for status, count in status_counts.items():
                    self.log_manage(f"  • {status}: {count} 个")
            else:
                self.log_manage(f"  • 无状态更新")
            
            self.log_manage(f"\n💰 本次更新涉及的金额:")
            self.log_manage(f"  • 广告花费: ${total_cost_updated:.2f}")
            self.log_manage(f"  • 佣金: ${total_commission_updated:.2f}")
            if commission_with_asterisk > 0:
                self.log_manage(f"  • 带*标记的佣金数 (可能不准确): {commission_with_asterisk} 个")
                self.log_manage(f"    (原因: PB数据无国家信息，且飞书有多个相同ASIN的Offer)")
            
            # 计算广告数据中无法提取信息的统计
            self.log_manage("\n🔍 数据质量:")
            self.log_manage(f"  • 无法提取ASIN的广告系列: {campaigns_without_asin} 个")
            self.log_manage(f"  • 无法提取国家的广告系列: {campaigns_without_country} 个")
            self.log_manage(f"  • PB交易中Rejected状态: {skipped_rejected} 条")
            self.log_manage(f"  • PB交易中无ASIN: {skipped_no_asin} 条")
            self.log_manage(f"  • PB佣金国家来源: customer_country={country_from_customer}, merchant_name={country_from_merchant}, mcid={country_from_mcid}, 默认US={country_default_us}")
            
            # 步骤5：更新"总计"行
            self.log_manage("\n【步骤5】更新飞书表格'总计'行")
            self.update_progress_manage("更新总计行...")
            
            try:
                # 直接从PB源数据计算已匹配佣金（不再从单元格求和，避免旧值导致重复计算）
                # 已匹配佣金 = PB中(ASIN+国家)在飞书表格中存在的佣金之和
                total_commission_sum = 0
                for key, comm in asin_country_commission.items():
                    if key in feishu_asin_country_set:
                        total_commission_sum += comm
                
                # 校验：PB总佣金（非Rejected） = 已匹配 + 未匹配
                total_pb_non_rejected = sum(asin_country_commission.values())
                self.log_manage(f"  PB佣金校验: PB总额(非Rejected)=${total_pb_non_rejected:.2f} = 已匹配${total_commission_sum:.2f} + 未匹配${total_unmatched_commission:.2f}")
                
                # 总花费：从MCC获取2026-01-01至今的全部花费（包含已删除的广告系列）
                self.log_manage(f"  从MCC获取全部花费（含已删除广告系列）...")
                mcc_total_cost = self.get_mcc_total_cost()
                if mcc_total_cost is not None:
                    total_cost_sum = mcc_total_cost
                else:
                    # 回退：使用campaign_data中的花费（不含已删除）
                    total_cost_sum = sum(c.get('cost_usd', 0) for c in campaign_data)
                    self.log_manage(f"  ⚠ 无法获取MCC全部花费，使用当前广告系列花费")
                
                self.log_manage(f"  飞书表格所有Offer汇总:")
                self.log_manage(f"    • 已匹配佣金: ${total_commission_sum:.2f}")
                self.log_manage(f"    • 未匹配佣金: ${total_unmatched_commission:.2f}")
                self.log_manage(f"    • Rejected佣金: ${rejected_commission_total:.2f}")
                self.log_manage(f"    • MCC总花费(USD): ${total_cost_sum:.2f}")
                
                # 计算新增量（本次统计值 - 上次记录值）
                new_cost_increment = total_cost_sum - previous_total_cost
                new_commission_increment = total_commission_sum - previous_total_commission
                
                self.log_manage(f"  📈 新增数据（相比上次运行）:")
                self.log_manage(f"    • 新增广告花费: ${new_cost_increment:.2f}")
                self.log_manage(f"    • 新增佣金: ${new_commission_increment:.2f}")
                
                # 更新第二行（总计行），传入未匹配佣金和Rejected佣金
                self.update_feishu_summary_row(feishu_token, total_commission_sum, total_cost_sum, total_unmatched_commission, rejected_commission_total)
                
            except Exception as e:
                self.log_manage(f"  更新总计行时出错: {str(e)}")
            
            # 步骤6：更新"广告系列"表格
            self.log_manage("\n【步骤6】更新'广告系列'表格")
            self.update_progress_manage("更新广告系列表格...")
            
            try:
                # 构建 campaign_id → tracking_link 的精确映射（基于row_campaigns的匹配结果）
                campaign_id_to_tracking_link = {}
                if row_campaigns:
                    # 先建一个 row_index → tracking_link 的映射
                    row_index_to_link = {}
                    for row in feishu_data:
                        ri = row.get('row_index')
                        if ri in row_campaigns:
                            tl = row.get('投放链接', '')
                            if isinstance(tl, list) and len(tl) > 0 and isinstance(tl[0], dict):
                                tl = tl[0].get('link', '') or tl[0].get('text', '')
                            if isinstance(tl, str) and tl:
                                row_index_to_link[ri] = tl
                    # 然后 campaign_id → tracking_link
                    for ri, cams in row_campaigns.items():
                        link = row_index_to_link.get(ri, '')
                        if link:
                            for c in cams:
                                cid = c.get('campaign_id')
                                if cid:
                                    campaign_id_to_tracking_link[cid] = link
                    if campaign_id_to_tracking_link:
                        self.log_manage(f"  构建了 {len(campaign_id_to_tracking_link)} 个广告系列→投放链接的精确映射")
                
                self.update_campaigns_sheet(
                    feishu_token=feishu_token,
                    campaign_data=campaign_data,
                    commission_data=commission_data,
                    feishu_data=feishu_data,
                    campaign_id_to_tracking_link=campaign_id_to_tracking_link
                )
            except Exception as e:
                self.log_manage(f"  更新广告系列表格时出错: {str(e)}")
                import traceback
                self.log_manage(traceback.format_exc())
            
            self.log_manage("\n" + "=" * 50)
            self.log_manage("✅ 统计完成！")
            self.log_manage("=" * 50)
            
        except Exception as e:
            self.log_manage(f"\n错误: {str(e)}")
            import traceback
            self.log_manage(traceback.format_exc())
        finally:
            self.root.after(0, self._restore_manage_ui)
    
    def update_campaigns_sheet(self, feishu_token, campaign_data, commission_data, feishu_data, campaign_id_to_tracking_link=None):
        """更新'广告系列'表格
        
        Args:
            feishu_token: 飞书访问令牌
            campaign_data: Google Ads广告系列数据列表
            commission_data: PB佣金数据列表
            feishu_data: offer表格数据（用于获取广告系列名称-投放链接映射）
        """
        # 新表格的配置
        campaigns_spreadsheet_token = "KnJ1wphpBiVMrGkWl5ncUkMGnfe"
        campaigns_sheet_id = "XrkOF7"
        
        # 步骤1：从offer表格构建 (ASIN, 国家) -> 投放链接 和 每单佣金 的映射
        self.log_manage("  构建(ASIN+国家)到投放链接和每单佣金的映射...")
        if campaign_id_to_tracking_link is None:
            campaign_id_to_tracking_link = {}
        
        asin_country_to_tracking_links = {}  # {(asin, country): [tracking_link, ...]} 改为列表，支持多个链接
        asin_country_to_commission_per_order = {}  # {(asin, country): commission_per_order}
        
        rows_with_tracking_link = 0
        for row in feishu_data:
            asin = row.get('ASIN', '')
            country = row.get('国家代码', '')
            tracking_link = row.get('投放链接', '')
            # 直接读取折扣价和佣金列来计算每单佣金，而不是读取公式列
            discount_price = row.get('折扣价', '')
            commission_rate = row.get('佣金', '')
            
            if asin and country:
                key = (asin, country.upper())
                if tracking_link:
                    rows_with_tracking_link += 1
                    if key not in asin_country_to_tracking_links:
                        asin_country_to_tracking_links[key] = []
                    asin_country_to_tracking_links[key].append(tracking_link)
                # 计算每单佣金 = 折扣价 * 佣金（而不是直接读取公式列）
                if discount_price and commission_rate:
                    try:
                        # 处理折扣价（可能带$符号）
                        price_str = str(discount_price).replace('$', '').replace(',', '').strip()
                        price_val = float(price_str) if price_str else 0
                        
                        # 处理佣金（可能是百分比如"8%"或小数如"0.08"）
                        comm_str = str(commission_rate).strip()
                        if '%' in comm_str:
                            comm_val = float(comm_str.replace('%', '')) / 100
                        else:
                            comm_val = float(comm_str)
                            # 如果大于1，认为是百分比形式
                            if comm_val > 1:
                                comm_val = comm_val / 100
                        
                        # 计算每单佣金
                        commission_per_order = price_val * comm_val
                        if commission_per_order > 0:
                            asin_country_to_commission_per_order[key] = commission_per_order
                    except (ValueError, TypeError):
                        pass  # 无法计算，跳过
        
        self.log_manage(f"    找到 {len(asin_country_to_tracking_links)} 个(ASIN+国家)到投放链接的映射")
        self.log_manage(f"    找到 {len(asin_country_to_commission_per_order)} 个(ASIN+国家)到每单佣金的映射")
        
        # 调试日志
        self.log_debug(f"=== (ASIN+国家)到投放链接映射 ===")
        for key, links in list(asin_country_to_tracking_links.items())[:20]:
            self.log_debug(f"  {key[0]}_{key[1]} -> {len(links)}个链接")
        
        # 步骤2：构建 uid -> 佣金数据 的映射（只处理带uid的佣金）
        self.log_manage("  构建UID到佣金数据的映射...")
        uid_commission = {}  # {uid: {'total_commission': float, 'asins': set()}}
        
        for trans in commission_data:
            if trans.get('status') == 'Rejected':
                continue
            
            uid = trans.get('uid', '')
            if uid:  # 只处理带uid的佣金
                comm = float(trans.get('sale_comm', 0) or 0)
                asin = trans.get('prod_id', '')
                
                # 获取国家代码
                country = trans.get('customer_country', '') or trans.get('geo', '') or trans.get('country', '')
                if not country:
                    merchant_name = trans.get('merchant_name', '')
                    if merchant_name:
                        country = self.extract_country_from_merchant_name(merchant_name)
                if not country:
                    country = 'US'
                country = country.upper()
                
                if uid not in uid_commission:
                    uid_commission[uid] = {'total_commission': 0, 'asins': set()}
                uid_commission[uid]['total_commission'] += comm
                if asin:
                    uid_commission[uid]['asins'].add(f"{asin}_{country}")
        
        self.log_manage(f"    找到 {len(uid_commission)} 个带UID的佣金记录")
        
        # 步骤3：构建MCC中所有广告系列的信息
        self.log_manage("  处理MCC广告系列数据...")
        mcc_campaigns = {}  # {campaign_name: campaign_info}
        
        for campaign in campaign_data:
            campaign_name = campaign.get('campaign_name', '')
            if not campaign_name:
                continue
            
            if campaign_name not in mcc_campaigns:
                mcc_campaigns[campaign_name] = {
                    'account_id': campaign.get('account_id', ''),
                    'cost_usd': 0,
                    'status': campaign.get('status', 'UNKNOWN'),
                    'campaign_id': campaign.get('campaign_id', ''),
                    'asin': campaign.get('asin', ''),
                    'country': campaign.get('country', ''),
                    'final_urls': campaign.get('final_urls', []),
                    'final_url_suffix': campaign.get('final_url_suffix', '')
                }
            
            # 累加花费
            mcc_campaigns[campaign_name]['cost_usd'] += campaign.get('cost_usd', 0)
            
            # 如果之前没有ASIN/国家，尝试从当前campaign获取
            if not mcc_campaigns[campaign_name]['asin'] and campaign.get('asin'):
                mcc_campaigns[campaign_name]['asin'] = campaign.get('asin')
            if not mcc_campaigns[campaign_name]['country'] and campaign.get('country'):
                mcc_campaigns[campaign_name]['country'] = campaign.get('country')
        
        self.log_manage(f"    MCC中共有 {len(mcc_campaigns)} 个广告系列")
        
        # 步骤4：读取现有的"广告系列"表格数据
        self.log_manage("  读取现有广告系列表格...")
        existing_campaigns_data = self.read_campaigns_sheet(feishu_token, campaigns_spreadsheet_token, campaigns_sheet_id)
        
        if existing_campaigns_data is None:
            self.log_manage("  ✗ 无法读取广告系列表格")
            return
        
        existing_rows, column_map, first_empty_row = existing_campaigns_data
        self.log_manage(f"    读取到 {len(existing_rows)} 行现有数据，第一个空行: {first_empty_row}")
        
        # 构建现有广告系列名称到行索引、投放链接、上次花费/佣金的映射
        existing_campaign_names = {}  # {campaign_name: row_index}
        existing_campaign_links = {}  # {campaign_name: tracking_link} 已有的投放链接
        existing_campaign_cost = {}   # {campaign_name: float} 上次运行时的花费
        existing_campaign_commission = {}  # {campaign_name: float} 上次运行时的佣金
        existing_campaign_status = {}  # {campaign_name: str} 上次运行时的状态
        for row_data in existing_rows:
            name = row_data.get('广告系列名称', '')
            row_idx = row_data.get('row_index')
            if name and row_idx:
                existing_campaign_names[name] = row_idx
                existing_link = row_data.get('投放链接', '')
                if isinstance(existing_link, list) and len(existing_link) > 0 and isinstance(existing_link[0], dict):
                    existing_link = existing_link[0].get('link', '') or existing_link[0].get('text', '')
                if existing_link:
                    existing_campaign_links[name] = existing_link
                # 记录上次的花费和佣金（用于计算新增量）
                existing_campaign_cost[name] = self.parse_commission_value(row_data.get('广告系列总花费', ''))
                existing_campaign_commission[name] = self.parse_commission_value(row_data.get('总佣金', ''))
                # 记录上次的状态（用于检测状态变更）
                existing_campaign_status[name] = str(row_data.get('状态', '') or '').strip()
        
        # 调试日志
        self.log_debug(f"广告系列表格列映射: {column_map}")
        self.log_debug(f"现有表格中有 {len(existing_campaign_names)} 个广告系列名称")
        for name in list(mcc_campaigns.keys()):
            asin = mcc_campaigns[name].get('asin', 'None')
            country = mcc_campaigns[name].get('country', 'None')
            in_existing = "已存在" if name in existing_campaign_names else "新增"
            self.log_debug(f"  MCC广告系列: {name} | ASIN={asin} | 国家={country} | [{in_existing}]")
        
        # 步骤5：准备更新数据
        self.log_manage("  准备更新数据...")
        updates = []  # 更新现有行
        new_rows = []  # 新增行
        
        # 用于追踪每个(ASIN+国家)已分配的投放链接索引，确保不同campaign获得不同链接
        key_link_assigned_idx = {}  # {(asin, country): next_index}
        
        # 处理MCC中的所有广告系列
        for campaign_name, campaign_info in mcc_campaigns.items():
            # 通过(ASIN+国家)获取投放链接和每单佣金
            asin = campaign_info.get('asin', '')
            country = campaign_info.get('country', '')
            tracking_link = ''
            commission_per_order = ''
            campaign_id = campaign_info.get('campaign_id', '')
            
            if asin and country:
                key = (asin, country.upper())
                
                # 优先使用精确映射（基于offer表的link_id匹配结果）
                if campaign_id and campaign_id in campaign_id_to_tracking_link:
                    tracking_link = campaign_id_to_tracking_link[campaign_id]
                else:
                    # 无精确映射时的策略：
                    links = asin_country_to_tracking_links.get(key, [])
                    if len(links) == 1:
                        # 只有一个链接，可以安全分配
                        tracking_link = links[0]
                    elif campaign_name in existing_campaign_links:
                        # 多个链接但无精确映射：保留已有的投放链接，避免错误覆盖
                        tracking_link = existing_campaign_links[campaign_name]
                    elif links:
                        # 新增campaign且有多个链接：按序分配
                        idx = key_link_assigned_idx.get(key, 0)
                        tracking_link = links[idx] if idx < len(links) else links[0]
                        key_link_assigned_idx[key] = idx + 1
                
                commission_per_order_value = asin_country_to_commission_per_order.get(key)
                if commission_per_order_value is not None:
                    if isinstance(commission_per_order_value, (int, float)):
                        commission_per_order = f"${commission_per_order_value:.2f}"
                    else:
                        commission_per_order = str(commission_per_order_value)
            
            # 通过投放链接最后7位匹配UID获取佣金
            total_commission = 0
            commission_asins = set()
            
            if tracking_link:
                # 提取投放链接最后7位作为UID
                link_uid = tracking_link[-7:] if len(tracking_link) >= 7 else tracking_link
                if link_uid in uid_commission:
                    total_commission = uid_commission[link_uid]['total_commission']
                    commission_asins = uid_commission[link_uid]['asins']
                    self.log_debug(f"  广告系列UID匹配成功: {campaign_name} | UID={link_uid} | 佣金=${total_commission:.2f} | ASINs={commission_asins}")
                else:
                    self.log_debug(f"  广告系列UID匹配失败: {campaign_name} | UID={link_uid} (佣金数据中无此UID)")
            else:
                self.log_debug(f"  广告系列无投放链接: {campaign_name} | ASIN={asin} | 国家={country}")
            
            # 计算ROI
            cost = campaign_info['cost_usd']
            roi = round(total_commission / cost, 1) if cost > 0 else 0
            
            # 确定状态
            status_raw = campaign_info['status']
            prev_status = existing_campaign_status.get(campaign_name, '')
            if status_raw == 'ENABLED':
                status = '投放中'
                status_color = 'green'
            elif status_raw == 'PAUSED':
                # 检查是否是从其他状态变为暂停
                if prev_status.startswith('广告系列暂停中'):
                    # 已经是暂停状态，保留原来的日期
                    status = prev_status
                else:
                    # 从其他状态（投放中、新增等）变为暂停，附加变更日期（如 2026-2-15）
                    now = datetime.now()
                    today_str = f"{now.year}-{now.month}-{now.day}"
                    status = f'广告系列暂停中{today_str}'
                status_color = 'orange'
            else:
                status = '投放中'  # 默认
                status_color = 'green'
            
            # 格式化佣金ASIN列表
            asins_str = ', '.join(sorted(commission_asins)) if commission_asins else ''
            
            # 构建产品链接 = Amazon最终到达网址 + ? + 最终到达网址后缀
            product_link_url = ''
            final_urls = campaign_info.get('final_urls', [])
            suffix = campaign_info.get('final_url_suffix', '')
            if final_urls and suffix:
                base_url = final_urls[0]
                separator = '&' if '?' in base_url else '?'
                product_link_url = f"{base_url}{separator}{suffix}"
            
            # 计算新增量（本次值 - 上次值）
            prev_cost = existing_campaign_cost.get(campaign_name, 0)
            prev_commission = existing_campaign_commission.get(campaign_name, 0)
            cost_increment = cost - prev_cost
            commission_increment = total_commission - prev_commission
            
            # 准备更新数据
            update_data = {
                'campaign_name': campaign_name,
                '状态': status,
                '广告系列名称': campaign_name,
                '投放中的ads': campaign_info['account_id'].replace('-', ''),
                '投放链接': tracking_link,
                '广告系列总花费': f"${cost:.2f}",
                '总佣金': f"${total_commission:.2f}",
                'ROI': f"{roi}",
                '佣金ASIN': asins_str,
                '每单佣金': commission_per_order,
                '产品链接': product_link_url,
                '新增广告系列花费': f"${cost_increment:.2f}",
                '新增佣金': f"${commission_increment:.2f}",
                'status_color': status_color
            }
            
            if campaign_name in existing_campaign_names:
                # 更新现有行
                update_data['row_index'] = existing_campaign_names[campaign_name]
                updates.append(update_data)
            else:
                # 新增行（上次不存在，新增量就是本次的全部值）
                new_rows.append(update_data)
        
        # 处理已删除的广告系列（在表格中存在但MCC中不存在）
        for campaign_name, row_idx in existing_campaign_names.items():
            if campaign_name not in mcc_campaigns:
                updates.append({
                    'campaign_name': campaign_name,
                    'row_index': row_idx,
                    '状态': '投放已结束',
                    'status_color': 'black'
                })
        
        self.log_manage(f"    需要更新 {len(updates)} 行，新增 {len(new_rows)} 行")
        
        # 步骤6：执行更新
        if updates or new_rows:
            self.apply_campaigns_sheet_updates(
                feishu_token,
                campaigns_spreadsheet_token,
                campaigns_sheet_id,
                updates,
                new_rows,
                column_map,
                first_empty_row
            )
        else:
            self.log_manage("  没有需要更新的内容")
    
    def read_campaigns_sheet(self, token, spreadsheet_token, sheet_id):
        """读取广告系列表格数据
        
        Returns:
            (rows_data, column_map, first_empty_row) 或 None
        """
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # 读取表格数据（扩大范围到AZ列，支持更多列）
        url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!A1:AZ1000"
        
        try:
            response = requests.get(url, headers=headers, timeout=30)
            data = response.json()
            
            if data.get('code') != 0:
                self.log_manage(f"    读取失败: {data.get('msg', 'Unknown error')}")
                return None
            
            values = data.get('data', {}).get('valueRange', {}).get('values', [])
            
            if not values:
                self.log_manage("    表格为空")
                return [], {}, 1
            
            # 解析表头（第一行）
            header_row = values[0] if values else []
            column_map = {}  # {列名: 列字母}
            
            for i, header in enumerate(header_row):
                if header:
                    col_letter = self.index_to_column_letter(i)
                    column_map[str(header).strip()] = col_letter
            
            # 解析数据行，同时找到第一个空行
            rows_data = []
            first_empty_row = 2  # 默认从第2行开始（第1行是表头）
            campaign_name_col_idx = None
            
            # 找到"广告系列名称"列的索引
            for i, header in enumerate(header_row):
                if header and str(header).strip() == '广告系列名称':
                    campaign_name_col_idx = i
                    break
            
            for row_idx, row in enumerate(values[1:], start=2):  # 从第2行开始（跳过表头）
                row_data = {'row_index': row_idx}
                has_campaign_name = False
                
                for i, cell in enumerate(row):
                    if i < len(header_row) and header_row[i]:
                        row_data[str(header_row[i]).strip()] = cell if cell else ''
                    
                    # 检查是否有广告系列名称
                    if campaign_name_col_idx is not None and i == campaign_name_col_idx and cell:
                        has_campaign_name = True
                
                rows_data.append(row_data)
                
                # 更新第一个空行位置
                if has_campaign_name:
                    first_empty_row = row_idx + 1
            
            # 返回第一个空行位置，而不是总行数
            return rows_data, column_map, first_empty_row
            
        except Exception as e:
            self.log_manage(f"    读取异常: {str(e)}")
            return None
    
    def apply_campaigns_sheet_updates(self, token, spreadsheet_token, sheet_id, updates, new_rows, column_map, first_empty_row):
        """应用广告系列表格更新"""
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # 定义字段到列的映射（完全基于表头文字匹配，不使用硬编码列位置）
        all_fields = ['状态', '广告系列名称', '投放中的ads', '投放链接', '广告系列总花费', 
                      '总佣金', 'ROI', '佣金ASIN', '每单佣金', '产品链接',
                      '新增广告系列花费', '新增佣金']
        field_to_column = {}
        missing_fields = []
        for f in all_fields:
            if f in column_map:
                field_to_column[f] = column_map[f]
            else:
                missing_fields.append(f)
        if missing_fields:
            self.log_manage(f"    ⚠ 广告系列表中未找到以下列: {missing_fields}")
            self.log_manage(f"    表格现有列: {list(column_map.keys())}")
        self.log_debug(f"  广告系列表列映射: {field_to_column}")
        
        value_ranges = []
        style_updates = []
        
        # 处理更新
        for update in updates:
            row_idx = update.get('row_index')
            if not row_idx:
                continue
            
            for field, col in field_to_column.items():
                if field in update and update[field] is not None:
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_idx}:{col}{row_idx}",
                        'values': [[update[field]]]
                    })
            
            # 收集样式更新
            if 'status_color' in update:
                style_updates.append({
                    'row_index': row_idx,
                    'color': update['status_color'],
                    'column': field_to_column['状态']
                })
        
        # 处理新增行
        current_row = first_empty_row
        for new_row in new_rows:
            for field, col in field_to_column.items():
                if field in new_row and new_row[field] is not None:
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{current_row}:{col}{current_row}",
                        'values': [[new_row[field]]]
                    })
            
            # 收集样式更新
            if 'status_color' in new_row:
                style_updates.append({
                    'row_index': current_row,
                    'color': new_row['status_color'],
                    'column': field_to_column['状态']
                })
            
            current_row += 1
        
        # 批量更新值
        if value_ranges:
            # 分批处理，每批最多100个range
            batch_size = 100
            for i in range(0, len(value_ranges), batch_size):
                batch = value_ranges[i:i+batch_size]
                url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
                body = {"valueRanges": batch}
                
                try:
                    response = requests.post(url, headers=headers, json=body, timeout=30)
                    data = response.json()
                    if data.get('code') == 0:
                        self.log_manage(f"    值更新成功: {len(batch)} 个单元格")
                    else:
                        self.log_manage(f"    值更新失败: {data.get('msg', 'Unknown error')}")
                except Exception as e:
                    self.log_manage(f"    值更新异常: {str(e)}")
        
        # 批量更新样式
        if style_updates:
            self.apply_campaigns_style_updates(token, spreadsheet_token, sheet_id, style_updates)
        
        self.log_manage(f"  ✅ 广告系列表格更新完成")
    
    def apply_campaigns_style_updates(self, token, spreadsheet_token, sheet_id, style_updates):
        """应用广告系列表格样式更新"""
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        color_map = {
            'green': {'red': 0, 'green': 128, 'blue': 0},
            'orange': {'red': 255, 'green': 165, 'blue': 0},
            'black': {'red': 0, 'green': 0, 'blue': 0},
            'red': {'red': 255, 'green': 0, 'blue': 0}
        }
        
        # 准备样式更新请求
        style_ranges = []
        for update in style_updates:
            row_idx = update['row_index']
            color = update['color']
            col = update['column']
            
            rgb = color_map.get(color, color_map['black'])
            
            style_ranges.append({
                'ranges': f"{sheet_id}!{col}{row_idx}:{col}{row_idx}",
                'style': {
                    'font': {
                        'bold': True
                    },
                    'foreColor': f"#{rgb['red']:02X}{rgb['green']:02X}{rgb['blue']:02X}"
                }
            })
        
        if style_ranges:
            url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/styles_batch_update"
            body = {"data": style_ranges}
            
            try:
                response = requests.put(url, headers=headers, json=body, timeout=30)
                data = response.json()
                if data.get('code') == 0:
                    self.log_manage(f"    样式更新成功: {len(style_ranges)} 个单元格")
                else:
                    self.log_manage(f"    样式更新失败: {data.get('msg', 'Unknown error')}")
            except Exception as e:
                self.log_manage(f"    样式更新异常: {str(e)}")
    
    def get_all_campaigns_with_asin(self):
        """获取所有广告系列及其ASIN（从广告层级的最终到达网址提取）
        
        返回:
            (campaigns, succeeded_account_ids): 广告系列列表和成功查询的账户ID集合
        """
        campaigns = []
        succeeded_account_ids = set()
        
        try:
            client = self.get_google_ads_client()
            if not client:
                self.log_manage("  ✗ 无法创建Google Ads客户端")
                return campaigns, succeeded_account_ids
            
            ga_service = client.get_service('GoogleAdsService')
            mcc_id = self.google_mcc_id_var.get().strip()
            
            # 获取所有子账户（包含货币信息）
            query = """
                SELECT customer_client.id, customer_client.descriptive_name, customer_client.manager, customer_client.currency_code
                FROM customer_client
                WHERE customer_client.level <= 1
            """
            response = ga_service.search(customer_id=mcc_id, query=query)
            
            sub_accounts = []
            for row in response:
                if not row.customer_client.manager:
                    sub_accounts.append({
                        'id': str(row.customer_client.id),
                        'name': row.customer_client.descriptive_name,
                        'currency': row.customer_client.currency_code
                    })
            
            self.log_manage(f"  找到 {len(sub_accounts)} 个子账户")
            
            # CNY到USD的汇率（可根据实际情况调整）
            CNY_TO_USD_RATE = 0.14  # 1 CNY ≈ 0.14 USD
            
            # 遍历每个子账户获取广告系列和广告
            for account in sub_accounts:
                if self.stop_flag:
                    break
                
                try:
                    # 获取账户货币类型
                    account_currency = account.get('currency', 'USD')
                    
                    # 查询广告层级，获取最终到达网址和跟踪后缀
                    ad_query = """
                        SELECT
                            campaign.id,
                            campaign.name,
                            campaign.status,
                            campaign.final_url_suffix,
                            ad_group.id,
                            ad_group_ad.ad.final_urls,
                            metrics.cost_micros
                        FROM ad_group_ad
                        WHERE campaign.status != 'REMOVED'
                            AND ad_group_ad.status != 'REMOVED'
                    """
                    
                    ad_response = ga_service.search(customer_id=account['id'], query=ad_query)
                    
                    # 用于去重：每个广告系列只记录一次，但收集所有ASIN
                    campaign_info = {}  # {campaign_id: campaign_data}
                    
                    for row in ad_response:
                        c = row.campaign
                        ad = row.ad_group_ad.ad
                        m = row.metrics
                        
                        campaign_id = str(c.id)
                        
                        # 从广告的最终到达网址提取ASIN
                        asin = None
                        ad_link_id = None
                        ad_uid = None
                        if ad.final_urls:
                            for url in ad.final_urls:
                                if not asin:
                                    asin = self.extract_asin_from_url(url)
                                # 也尝试从final_url中提取uid和link_id（适用于pboost.me短链接）
                                if not ad_uid:
                                    uid_match = re.search(r'[?&]uid=([^&]+)', url)
                                    if uid_match:
                                        ad_uid = uid_match.group(1)
                                if not ad_link_id:
                                    lid_match = re.search(r'[?&]aa_adgroupid=([^&]+)', url)
                                    if lid_match:
                                        ad_link_id = lid_match.group(1)
                        
                        # 从campaign的final_url_suffix（跟踪后缀）中提取link_id和uid
                        suffix = c.final_url_suffix if c.final_url_suffix else ''
                        if suffix:
                            if not ad_link_id:
                                lid_match = re.search(r'(?:^|&)aa_adgroupid=([^&]+)', suffix)
                                if lid_match:
                                    ad_link_id = lid_match.group(1)
                            if not ad_uid:
                                uid_match = re.search(r'(?:^|&)uid=([^&]+)', suffix)
                                if uid_match:
                                    ad_uid = uid_match.group(1)
                        
                        if campaign_id not in campaign_info:
                            # 从广告系列名称提取国家代码
                            country = self.extract_country_from_campaign_name(c.name)
                            campaign_info[campaign_id] = {
                                'account_id': account['id'],
                                'account_name': account['name'],
                                'campaign_id': campaign_id,
                                'campaign_name': c.name,
                                'status': c.status.name,
                                'cost_usd': 0,
                                'currency': account_currency,
                                'asin': asin,
                                'country': country,
                                'final_urls': list(ad.final_urls) if ad.final_urls else [],
                                'final_url_suffix': suffix,
                                'link_ids': set(),  # 该广告系列关联的所有link_id
                                'uids': set()       # 该广告系列关联的所有uid
                            }
                        
                        # 累加花费（根据货币类型转换为USD）
                        cost_in_original = m.cost_micros / 1000000 if m.cost_micros else 0
                        if account_currency == 'CNY':
                            cost_in_usd = cost_in_original * CNY_TO_USD_RATE
                        else:
                            cost_in_usd = cost_in_original  # 假设其他货币都是USD
                        campaign_info[campaign_id]['cost_usd'] += cost_in_usd
                        
                        # 如果之前没有ASIN，尝试从这个广告获取
                        if not campaign_info[campaign_id]['asin'] and asin:
                            campaign_info[campaign_id]['asin'] = asin
                        
                        # 收集link_id和uid（用于精确匹配offer行）
                        if ad_link_id:
                            campaign_info[campaign_id]['link_ids'].add(ad_link_id)
                        if ad_uid:
                            campaign_info[campaign_id]['uids'].add(ad_uid)
                    
                    # 将set转为list方便后续处理
                    for cinfo in campaign_info.values():
                        cinfo['link_ids'] = list(cinfo.get('link_ids', set()))
                        cinfo['uids'] = list(cinfo.get('uids', set()))
                    campaigns.extend(campaign_info.values())
                    succeeded_account_ids.add(account['id'])
                        
                except Exception as e:
                    self.log_manage(f"  ⚠ 账户 {account['name']}({account['id']}) 获取失败: {str(e)[:100]}")
            
            if len(succeeded_account_ids) < len(sub_accounts):
                failed_count = len(sub_accounts) - len(succeeded_account_ids)
                self.log_manage(f"  ⚠ {failed_count} 个账户查询失败，这些账户的广告系列不会参与删除检测")
            
        except Exception as e:
            self.log_manage(f"  获取广告系列失败: {e}")
        
        return campaigns, succeeded_account_ids
    
    def _resolve_campaign_to_row_mapping(self, feishu_data, asin_country_campaigns, campaign_data):
        """将广告系列精确匹配到具体的offer行
        
        匹配策略（按优先级）：
        0. 产品链接匹配（最高优先级）：从offer行产品链接中提取aa_adgroupid，与广告系列的aa_adgroupid比较（一对一匹配）
        1. 直接uid匹配：从Google Ads final URL提取的uid与飞书投放链接的uid比较
        2. link_id匹配：通过PB API获取link_id，与Google Ads final URL中的aa_adgroupid比较
        
        返回:
            {row_index: [campaign_info, ...]} 精确匹配的映射
        """
        row_campaigns = {}
        
        # 1. 构建 uid → campaigns 和 link_id → campaigns 映射
        uid_to_campaigns = {}     # {uid: [campaign]}
        link_id_to_campaigns = {} # {link_id: [campaign]}
        has_uids = False
        has_link_ids = False
        
        for campaign in campaign_data:
            for uid in campaign.get('uids', []):
                if uid:
                    has_uids = True
                    if uid not in uid_to_campaigns:
                        uid_to_campaigns[uid] = []
                    uid_to_campaigns[uid].append(campaign)
            for lid in campaign.get('link_ids', []):
                if lid:
                    has_link_ids = True
                    if lid not in link_id_to_campaigns:
                        link_id_to_campaigns[lid] = []
                    link_id_to_campaigns[lid].append(campaign)
        
        if not has_uids and not has_link_ids:
            self.log_manage("    （广告系列中未找到uid或link_id，跳过精确匹配）")
            return row_campaigns
        
        self.log_manage(f"    广告系列标识: {len(uid_to_campaigns)}个uid, {len(link_id_to_campaigns)}个link_id")
        
        # 2. 策略零（最高优先级）：通过产品链接中的aa_adgroupid直接匹配
        # 每个投放链接重定向后的产品链接包含唯一的aa_adgroupid参数，与广告系列后缀一致
        if has_link_ids:
            product_link_matched = 0
            for row in feishu_data:
                row_index = row.get('row_index')
                if row_index in row_campaigns:
                    continue
                
                product_link = row.get('产品链接', '')
                if isinstance(product_link, list) and len(product_link) > 0 and isinstance(product_link[0], dict):
                    product_link = product_link[0].get('link', '') or product_link[0].get('text', '')
                if not isinstance(product_link, str) or not product_link:
                    continue
                
                # 从产品链接URL中提取aa_adgroupid参数
                adgroupid_match = re.search(r'[?&]aa_adgroupid=([^&]+)', product_link)
                if not adgroupid_match:
                    continue
                
                row_adgroupid = adgroupid_match.group(1)
                if row_adgroupid in link_id_to_campaigns:
                    row_campaigns[row_index] = link_id_to_campaigns[row_adgroupid]
                    product_link_matched += 1
            
            if product_link_matched:
                self.log_manage(f"    ✓ 通过产品链接aa_adgroupid匹配: {product_link_matched} 行")
        
        # 3. 找出所有需要精确匹配的(ASIN, 国家)组合
        # 不仅处理多行offer的组合，还处理所有有广告系列的组合
        # 这样才能为广告系列表建立完整的 campaign_id -> tracking_link 映射
        asin_country_rows = {}
        for row in feishu_data:
            asin = row.get('ASIN', '') or ''
            country = (row.get('国家代码', '') or '').upper().strip()
            if not asin or not country:
                continue
            key = (asin, country)
            if key not in asin_country_rows:
                asin_country_rows[key] = []
            asin_country_rows[key].append(row)
        
        # 处理所有有广告系列的key，而不仅仅是多行的
        keys_to_match = {
            key for key, rows in asin_country_rows.items()
            if key in asin_country_campaigns
        }
        
        if not keys_to_match:
            return row_campaigns
        
        multi_row_count = sum(1 for k in keys_to_match if len(asin_country_rows[k]) > 1)
        self.log_manage(f"    需要精确匹配: {len(keys_to_match)} 个(ASIN+国家)组合 (其中{multi_row_count}个有多行offer)")
        
        # 3. 策略一：直接用uid匹配（最快，无需API调用）
        uid_matched = 0
        for key in keys_to_match:
            rows = asin_country_rows[key]
            for row in rows:
                row_index = row.get('row_index')
                tracking_link = row.get('投放链接', '')
                
                # 提取投放链接中的uid
                if isinstance(tracking_link, list) and len(tracking_link) > 0 and isinstance(tracking_link[0], dict):
                    tracking_link = tracking_link[0].get('link', '') or tracking_link[0].get('text', '')
                if not isinstance(tracking_link, str) or not tracking_link:
                    continue
                
                row_uid = ''
                uid_match = re.search(r'[?&]uid=([^&]+)', tracking_link)
                if uid_match:
                    row_uid = uid_match.group(1)
                
                if not row_uid:
                    continue
                
                # 直接在uid_to_campaigns中查找
                if row_uid in uid_to_campaigns:
                    row_campaigns[row_index] = uid_to_campaigns[row_uid]
                    uid_matched += 1
        
        if uid_matched:
            self.log_manage(f"    ✓ 通过uid直接匹配: {uid_matched} 行")
        
        # 4. 策略二：对未匹配的行，通过PB API获取link_id匹配
        if has_link_ids:
            unmatched_keys = {
                key for key in keys_to_match
                if any(row.get('row_index') not in row_campaigns for row in asin_country_rows[key])
            }
            
            if unmatched_keys:
                token = self.pb_token_var.get().strip()
                api_calls = 0
                link_matched = 0
                
                for key in unmatched_keys:
                    if self.stop_flag:
                        break
                    rows = asin_country_rows[key]
                    asin, country = key
                    
                    for row in rows:
                        row_index = row.get('row_index')
                        if row_index in row_campaigns:
                            continue  # 已经通过uid匹配了
                        
                        tracking_link = row.get('投放链接', '')
                        if isinstance(tracking_link, list) and len(tracking_link) > 0 and isinstance(tracking_link[0], dict):
                            tracking_link = tracking_link[0].get('link', '') or tracking_link[0].get('text', '')
                        if not isinstance(tracking_link, str) or not tracking_link:
                            continue
                        
                        row_uid = ''
                        uid_m = re.search(r'[?&]uid=([^&]+)', tracking_link)
                        if uid_m:
                            row_uid = uid_m.group(1)
                        elif len(tracking_link) >= 7:
                            row_uid = tracking_link[-7:]
                        if not row_uid:
                            continue
                        
                        try:
                            api_url = f"{PB_API_BASE_URL}/api/datafeed/get_amazon_link_by_asin"
                            body = {
                                "token": token, "asins": asin,
                                "country_code": country, "uid": row_uid,
                                "return_partnerboost_link": 1
                            }
                            resp = requests.post(api_url, json=body, timeout=15)
                            api_calls += 1
                            data = resp.json()
                            
                            if data.get('status', {}).get('code') == 0:
                                link_data = data.get('data', [])
                                if link_data:
                                    link_id = link_data[0].get('link_id', '')
                                    if link_id and link_id in link_id_to_campaigns:
                                        row_campaigns[row_index] = link_id_to_campaigns[link_id]
                                        link_matched += 1
                            time.sleep(0.1)
                        except Exception as e:
                            self.log_debug(f"    解析link_id失败 ASIN={asin} uid={row_uid}: {str(e)[:80]}")
                
                if api_calls:
                    self.log_manage(f"    ✓ 通过link_id匹配: {link_matched} 行 ({api_calls}次API调用)")
        
        # 去重：确保每个广告系列只分配到一个offer行（一一对应原则）
        # 如果同一个广告系列被分配到多行（例如PB API对同ASIN+国家返回相同link_id），
        # 只保留第一个匹配的行，其余行移除该广告系列
        assigned_campaign_ids = {}  # {campaign_id: row_index} 记录已分配的广告系列
        rows_to_remove = []
        
        for row_index in sorted(row_campaigns.keys()):
            campaigns = row_campaigns[row_index]
            unassigned = []
            for c in campaigns:
                cid = c.get('campaign_id')
                if cid not in assigned_campaign_ids:
                    assigned_campaign_ids[cid] = row_index
                    unassigned.append(c)
            if unassigned:
                row_campaigns[row_index] = unassigned
            else:
                rows_to_remove.append(row_index)
        
        for ri in rows_to_remove:
            del row_campaigns[ri]
        
        if rows_to_remove:
            self.log_manage(f"    去重: 移除了 {len(rows_to_remove)} 个重复分配（确保一一对应）")
        
        total = len(row_campaigns)
        self.log_manage(f"    精确匹配总计: {total} 行")
        
        return row_campaigns
    
    def extract_asin_from_url(self, url):
        """从最终到达网址中提取ASIN"""
        if not url:
            return None
        
        # ASIN格式：https://www.amazon.com/dp/B0F9NK4M7Q 或类似格式
        # 尝试多种匹配模式
        patterns = [
            r'/dp/([A-Z0-9]{10})',           # 标准格式 /dp/ASIN
            r'/gp/product/([A-Z0-9]{10})',   # 旧格式 /gp/product/ASIN
            r'/product/([A-Z0-9]{10})',      # 简化格式
            r'[?&]asin=([A-Z0-9]{10})',      # URL参数 ?asin=ASIN
            r'[?&]ASIN=([A-Z0-9]{10})',      # URL参数大写
        ]
        
        for pattern in patterns:
            match = re.search(pattern, url, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        
        return None
    
    def extract_country_from_campaign_name(self, campaign_name):
        """从广告系列名称中提取国家代码
        
        广告系列名称格式示例：2037-2589-anker-US-Search-20260122205857
        提取其中的国家代码（如US, DE, UK等）
        """
        if not campaign_name:
            return None
        
        # 常见国家代码列表
        country_codes = ['US', 'UK', 'DE', 'FR', 'IT', 'ES', 'CA', 'JP', 'AU', 'NL', 'BE', 'MX', 'BR', 'IN', 'SG', 'AE', 'SA', 'PL', 'SE', 'TR', 'EG']
        
        # 尝试从广告系列名称中匹配国家代码
        # 格式：xxxx-xxxx-brand-COUNTRY-type-date
        parts = campaign_name.split('-')
        for part in parts:
            part_upper = part.upper()
            if part_upper in country_codes:
                return part_upper
        
        # 也尝试用正则匹配，以防格式略有不同
        for code in country_codes:
            # 匹配被-包围的国家代码，不区分大小写
            pattern = rf'[-_]({code})[-_]'
            match = re.search(pattern, campaign_name, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        
        return None
    
    def extract_country_from_merchant_name(self, merchant_name):
        """从品牌名称中提取国家代码
        
        品牌名称格式示例：DE-Anker, UK-Anker, Anker
        提取其中的国家代码（如 DE, UK 等）
        
        返回:
            提取到的国家代码（大写），如果没有则返回 None
        """
        if not merchant_name:
            return None
        
        # 常见国家代码列表
        country_codes = ['US', 'UK', 'DE', 'FR', 'IT', 'ES', 'CA', 'JP', 'AU', 'NL', 'BE', 'MX', 'BR', 'IN', 'SG', 'AE', 'SA', 'PL', 'SE', 'TR', 'EG']
        
        # 尝试匹配开头的国家代码（如 "DE-Anker" -> "DE"）
        # 格式：COUNTRY-BrandName
        parts = merchant_name.split('-')
        if len(parts) >= 2:
            first_part = parts[0].upper()
            if first_part in country_codes:
                return first_part
        
        # 也尝试匹配其他位置的国家代码
        for code in country_codes:
            # 匹配被-或_包围的国家代码，或在开头/结尾
            patterns = [
                rf'^({code})[-_]',      # 开头：DE-Anker
                rf'[-_]({code})$',      # 结尾：Anker-DE
                rf'[-_]({code})[-_]',   # 中间：Brand-DE-Name
            ]
            for pattern in patterns:
                match = re.search(pattern, merchant_name, re.IGNORECASE)
                if match:
                    return match.group(1).upper()
        
        return None
    
    def extract_country_from_mcid(self, mcid):
        """从 mcid 中提取国家代码
        
        mcid 格式示例：amzdeanker, amzusanker, amzukanker
        其中 amz 是 Amazon 前缀，de/us/uk 是国家代码，anker 是品牌名
        
        返回:
            提取到的国家代码（大写），如果没有则返回 None
        """
        if not mcid:
            return None
        
        mcid_lower = mcid.lower()
        
        # 常见国家代码及其在 mcid 中的表示形式
        # mcid 格式：amz + 国家代码 + 品牌名
        country_patterns = {
            'de': 'DE',   # amzdeanker -> DE
            'us': 'US',   # amzusanker -> US
            'uk': 'UK',   # amzukanker -> UK
            'fr': 'FR',   # amzfranker -> FR
            'it': 'IT',   # amzitanker -> IT
            'es': 'ES',   # amzesanker -> ES
            'ca': 'CA',   # amzcanker -> CA (注意可能有歧义)
            'jp': 'JP',   # amzjpanker -> JP
            'au': 'AU',   # amzauanker -> AU
            'nl': 'NL',   # amznlanker -> NL
            'be': 'BE',   # amzbeanker -> BE
            'mx': 'MX',   # amzmxanker -> MX
            'br': 'BR',   # amzbranker -> BR
            'in': 'IN',   # amzinanker -> IN
            'sg': 'SG',   # amzsganker -> SG
            'ae': 'AE',   # amzaeanker -> AE
            'sa': 'SA',   # amzsaanker -> SA
            'pl': 'PL',   # amzplanker -> PL
            'se': 'SE',   # amzseanker -> SE
            'tr': 'TR',   # amztranker -> TR
            'eg': 'EG',   # amzeganker -> EG
        }
        
        # 检查 mcid 是否以 amz 开头
        if mcid_lower.startswith('amz'):
            # 尝试提取国家代码（amz 后面的 2 个字符）
            if len(mcid_lower) >= 5:
                potential_country = mcid_lower[3:5]
                if potential_country in country_patterns:
                    return country_patterns[potential_country]
        
        return None
    
    def get_feishu_sheet_data(self, token):
        """获取飞书表格数据"""
        spreadsheet_token = self.feishu_spreadsheet_var.get().strip()
        sheet_id = self.feishu_sheet_id_var.get().strip()
        
        # 获取工作表元数据
        meta_url = f"{FEISHU_API_BASE_URL}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}"
        headers = {"Authorization": f"Bearer {token}"}
        
        response = requests.get(meta_url, headers=headers, timeout=30)
        meta = response.json()
        row_count = meta.get("data", {}).get("sheet", {}).get("grid_properties", {}).get("row_count", 100)
        
        # 读取数据
        range_str = f"{sheet_id}!A1:Z{row_count}"  # 扩大范围以支持更多列
        data_url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range_str}"
        
        response = requests.get(data_url, headers=headers, timeout=30)
        data = response.json()
        
        values = data.get("data", {}).get("valueRange", {}).get("values", [])
        
        # 解析数据，跳过表头
        rows = []
        column_map = {}  # 表头名称 -> 列字母的映射
        
        if len(values) > 1:
            headers_row = values[0]
            
            # 构建列映射：表头名称 -> 列字母
            for j, header in enumerate(headers_row):
                if header:
                    col_letter = self.index_to_column_letter(j)
                    column_map[header] = col_letter
            
            for i, row in enumerate(values[1:], start=2):  # 从第2行开始
                row_data = {'row_index': i}
                for j, cell in enumerate(row):
                    if j < len(headers_row):
                        # 处理URL类型的单元格
                        if isinstance(cell, list) and len(cell) > 0 and isinstance(cell[0], dict):
                            cell = cell[0].get('link', '') or cell[0].get('text', '')
                        row_data[headers_row[j]] = cell
                rows.append(row_data)
        
        # 保存列映射供写入时使用
        self.feishu_column_map = column_map
        
        return rows
    
    def index_to_column_letter(self, index):
        """将列索引（0开始）转换为Excel列字母（A, B, ..., Z, AA, AB, ...）"""
        result = ""
        while index >= 0:
            result = chr(index % 26 + ord('A')) + result
            index = index // 26 - 1
        return result
    
    def load_old_campaign_consume(self):
        """读取旧广告系列花费记录
        
        返回:
            dict: {(asin, country): {'old_cost': float, 'last_updated': str, 'brand_name': str}}
        """
        old_consume_data = {}
        
        if not os.path.exists(OLD_CAMPAIGN_CONSUME_FILE):
            return old_consume_data
        
        try:
            from openpyxl import load_workbook
            wb = load_workbook(OLD_CAMPAIGN_CONSUME_FILE)
            ws = wb.active
            
            # 跳过表头，从第2行开始读取
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:  # ASIN 和 国家代码 不为空
                    asin = str(row[0]).strip()
                    country = str(row[1]).strip().upper()
                    old_cost = float(row[2]) if row[2] else 0
                    brand_name = str(row[3]) if row[3] else ''
                    last_updated = str(row[4]) if row[4] else ''
                    
                    key = (asin, country)
                    old_consume_data[key] = {
                        'old_cost': old_cost,
                        'brand_name': brand_name,
                        'last_updated': last_updated
                    }
            
            wb.close()
        except Exception as e:
            self.log_manage(f"  读取旧广告系列花费文件失败: {e}")
        
        return old_consume_data
    
    def save_old_campaign_consume(self, old_consume_data):
        """保存旧广告系列花费记录
        
        参数:
            old_consume_data: dict, {(asin, country): {'old_cost': float, 'last_updated': str, 'brand_name': str}}
        """
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "旧广告系列花费"
            
            # 写入表头
            headers = ['ASIN', '国家代码', '旧广告系列累计花费', '品牌名称', '最后更新时间']
            ws.append(headers)
            
            # 写入数据
            for (asin, country), data in old_consume_data.items():
                ws.append([
                    asin,
                    country,
                    data.get('old_cost', 0),
                    data.get('brand_name', ''),
                    data.get('last_updated', '')
                ])
            
            # 调整列宽
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 20
            ws.column_dimensions['E'].width = 20
            
            wb.save(OLD_CAMPAIGN_CONSUME_FILE)
            self.log_manage(f"  旧广告系列花费已保存到: {OLD_CAMPAIGN_CONSUME_FILE}")
        except Exception as e:
            self.log_manage(f"  保存旧广告系列花费文件失败: {e}")
    
    def add_old_campaign_cost(self, old_consume_data, asin, country, cost_to_add, brand_name=''):
        """添加或累加旧广告系列花费
        
        参数:
            old_consume_data: dict, 旧花费数据字典
            asin: str, ASIN码
            country: str, 国家代码
            cost_to_add: float, 要累加的花费
            brand_name: str, 品牌名称
        """
        key = (asin, country.upper())
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        if key in old_consume_data:
            # 累加花费
            old_consume_data[key]['old_cost'] += cost_to_add
            old_consume_data[key]['last_updated'] = current_time
            if brand_name:
                old_consume_data[key]['brand_name'] = brand_name
        else:
            # 新增记录
            old_consume_data[key] = {
                'old_cost': cost_to_add,
                'brand_name': brand_name,
                'last_updated': current_time
            }
    
    def load_campaign_cost_snapshot(self):
        """读取广告系列花费快照
        
        返回:
            dict: {campaign_id: {'asin': str, 'country': str, 'cost': float, 'campaign_name': str}}
        """
        snapshot = {}
        
        if not os.path.exists(CAMPAIGN_COST_SNAPSHOT_FILE):
            return snapshot
        
        try:
            from openpyxl import load_workbook
            wb = load_workbook(CAMPAIGN_COST_SNAPSHOT_FILE)
            ws = wb.active
            
            # 跳过表头，从第2行开始读取
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:  # campaign_id 不为空
                    campaign_id = str(row[0]).strip()
                    snapshot[campaign_id] = {
                        'asin': str(row[1]).strip() if row[1] else '',
                        'country': str(row[2]).strip().upper() if row[2] else '',
                        'cost': float(row[3]) if row[3] else 0,
                        'campaign_name': str(row[4]) if row[4] else '',
                        'account_id': str(row[5]) if row[5] else ''
                    }
            
            wb.close()
        except Exception as e:
            self.log_manage(f"  读取广告系列快照文件失败: {e}")
        
        return snapshot
    
    def save_campaign_cost_snapshot(self, snapshot):
        """保存广告系列花费快照
        
        参数:
            snapshot: dict, {campaign_id: {'asin': str, 'country': str, 'cost': float, 'campaign_name': str}}
        """
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "广告系列花费快照"
            
            # 写入表头
            headers = ['广告系列ID', 'ASIN', '国家代码', '花费(USD)', '广告系列名称', '账户ID']
            ws.append(headers)
            
            # 写入数据
            for campaign_id, data in snapshot.items():
                ws.append([
                    campaign_id,
                    data.get('asin', ''),
                    data.get('country', ''),
                    data.get('cost', 0),
                    data.get('campaign_name', ''),
                    data.get('account_id', '')
                ])
            
            # 调整列宽
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 10
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 40
            ws.column_dimensions['F'].width = 15
            
            wb.save(CAMPAIGN_COST_SNAPSHOT_FILE)
        except Exception as e:
            self.log_manage(f"  保存广告系列快照文件失败: {e}")
    
    def get_all_commissions(self):
        """获取所有历史佣金数据"""
        all_transactions = []
        url = f"{PB_API_BASE_URL}/api.php?mod=medium&op=transaction"
        token = self.pb_token_var.get().strip()
        
        start_from = datetime(2024, 1, 1)
        current_end = datetime.now()
        
        self.log_manage(f"  查询佣金数据，日期范围: {start_from.strftime('%Y-%m-%d')} 到 {current_end.strftime('%Y-%m-%d')}")
        
        api_calls = 0
        api_errors = 0
        
        while current_end > start_from:
            if self.stop_flag:
                break
            
            current_begin = current_end - timedelta(days=60)
            if current_begin < start_from:
                current_begin = start_from
            
            begin_str = current_begin.strftime('%Y-%m-%d')
            end_str = current_end.strftime('%Y-%m-%d')
            
            page = 1
            while True:
                body = {
                    'token': token,
                    'begin_date': begin_str,
                    'end_date': end_str,
                    'type': 'json',
                    'status': 'All',
                    'limit': 2000,
                    'page': page
                }
                
                try:
                    response = requests.post(url, json=body, timeout=60)
                    data = response.json()
                    api_calls += 1
                    
                    status_code = data.get('status', {}).get('code')
                    if status_code != 0:
                        error_msg = data.get('status', {}).get('message', 'Unknown error')
                        self.log_manage(f"  API返回错误 ({begin_str}~{end_str}): code={status_code}, msg={error_msg}")
                        api_errors += 1
                        break
                    
                    transactions = data.get('data', {}).get('list', [])
                    total_page = int(data.get('data', {}).get('total_page', 0))
                    total_count = int(data.get('data', {}).get('total_count', 0))
                    
                    if page == 1 and transactions:
                        self.log_manage(f"  {begin_str}~{end_str}: 共{total_count}条，{total_page}页")
                    
                    all_transactions.extend(transactions)
                    
                    if page >= total_page or not transactions:
                        break
                    page += 1
                    
                except Exception as e:
                    self.log_manage(f"  API请求异常 ({begin_str}~{end_str}): {str(e)[:100]}")
                    api_errors += 1
                    break
            
            current_end = current_begin - timedelta(days=1)
        
        self.log_manage(f"  API调用统计: {api_calls}次请求, {api_errors}次错误")
        
        return all_transactions
    
    def calculate_updates(self, feishu_data, asin_country_campaigns, asin_country_commission, asin_only_commission=None, asin_country_uid_commission=None, asin_country_no_uid_commission=None, row_campaigns=None):
        """计算需要更新的内容
        
        参数:
            feishu_data: 飞书表格数据
            asin_country_campaigns: {(asin, country): [campaign_info, ...]} 广告系列映射
            asin_country_commission: {(asin, country): total_commission} 所有佣金总映射（用于总计）
            asin_only_commission: {asin: total_commission} 无uid且国家默认US的佣金（ASIN-only备选匹配）
            asin_country_uid_commission: {(asin, country, uid): total_commission} 有uid的佣金（UID池）
            asin_country_no_uid_commission: {(asin, country): total_commission} 无uid的佣金（非UID池）
            row_campaigns: {row_index: [campaign_info, ...]} 通过uid/link_id精确匹配的行级广告系列映射
        
        返回:
            updates: 更新列表
        """
        if asin_only_commission is None:
            asin_only_commission = {}
        if asin_country_uid_commission is None:
            asin_country_uid_commission = {}
        if asin_country_no_uid_commission is None:
            asin_country_no_uid_commission = {}
        if row_campaigns is None:
            row_campaigns = {}
        updates = []
        
        # 构建已被link_id精确匹配"认领"的广告系列ID集合
        # 用于在回退到(ASIN+国家)匹配时，排除已认领的广告系列
        claimed_campaign_ids = set()
        for row_idx, row_cams in row_campaigns.items():
            for c in row_cams:
                claimed_campaign_ids.add(c.get('campaign_id'))
        
        # 获取所有正在投放的(ASIN, 国家)组合
        active_keys = set(asin_country_campaigns.keys())
        
        # 第一遍遍历：构建每个(ASIN, 国家)组合的状态分布
        # 用于判断佣金应该更新到哪一行
        # active_rows: 正在投放的行（投放中、广告系列暂停中等）
        # ended_rows: 投放已结束的行
        # untested_rows: 未测试的行
        asin_country_rows = {}  # {(asin, country): {'active_rows': [...], 'ended_rows': [...], 'untested_rows': [...]}}
        
        # 统计每个ASIN在飞书表格中出现的次数（用于判断是否有多个相同ASIN的offer）
        asin_row_count = {}  # {asin: count}
        # 记录每个ASIN有哪些国家的offer
        asin_countries = {}  # {asin: set(country1, country2, ...)}
        # 记录哪些ASIN有US offer
        asin_has_us_offer = set()
        # 记录每行的uid（从投放链接最后7位提取）
        row_uid_map = {}  # {row_index: uid}
        # 记录每个(asin, country)有多少个不同的投放链接（用于判断是否有复制的offer）
        asin_country_uids = {}  # {(asin, country): [uid1, uid2, ...]}
        
        for row in feishu_data:
            asin = row.get('ASIN', '')
            country = row.get('国家代码', '')
            current_status = row.get('状态', '')
            row_index = row.get('row_index')
            tracking_link = row.get('投放链接', '')
            
            if not asin:
                continue
            
            # 提取投放链接的uid（最后7位）
            row_uid = ''
            if tracking_link:
                # 处理URL类型的单元格
                if isinstance(tracking_link, list) and len(tracking_link) > 0 and isinstance(tracking_link[0], dict):
                    tracking_link = tracking_link[0].get('link', '') or tracking_link[0].get('text', '')
                if isinstance(tracking_link, str) and len(tracking_link) >= 7:
                    row_uid = tracking_link[-7:]
            row_uid_map[row_index] = row_uid
            
            # 统计每个ASIN的出现次数
            if asin not in asin_row_count:
                asin_row_count[asin] = 0
            asin_row_count[asin] += 1
            
            # 标准化国家代码
            if country:
                country = country.upper().strip()
            
            # 记录每个ASIN有哪些国家
            if country:
                if asin not in asin_countries:
                    asin_countries[asin] = set()
                asin_countries[asin].add(country)
                
                # 记录有US offer的ASIN
                if country == 'US':
                    asin_has_us_offer.add(asin)
            
            key = (asin, country) if country else None
            if not key:
                continue
            
            # 记录每个(asin, country)的uid列表
            if key not in asin_country_uids:
                asin_country_uids[key] = []
            if row_uid:
                asin_country_uids[key].append(row_uid)
            
            if key not in asin_country_rows:
                asin_country_rows[key] = {'active_rows': [], 'ended_rows': [], 'untested_rows': []}
            
            if current_status == '投放已结束':
                asin_country_rows[key]['ended_rows'].append(row_index)
            elif current_status == '未测试' or current_status == '新复制' or not current_status:
                asin_country_rows[key]['untested_rows'].append(row_index)
            else:
                # 投放中、广告系列暂停中等活跃状态
                asin_country_rows[key]['active_rows'].append(row_index)
        
        # 用于跟踪非uid池的佣金是否已分配给某个(asin, country)的offer
        asin_country_no_uid_assigned = set()
        # 用于跟踪ASIN-only备选佣金是否已分配
        asin_only_commission_assigned = set()
        
        def get_commission_for_row(asin, country, row_index, is_first_offer):
            """获取某行的佣金值
            
            匹配规则（两个池子互不重叠）：
            1. UID池：PB交易有uid → 匹配到投放链接末尾uid相同的offer行
            2. 非UID池：PB交易无uid → 按(ASIN+国家)匹配，每个(ASIN,国家)只分配一次
            
            一行offer可以同时获得UID池和非UID池的佣金（因为它们是不同的PB交易）
            
            Args:
                asin: ASIN
                country: 国家代码
                row_index: 行索引
                is_first_offer: 是否是该(asin, country)的第一个offer（已弃用，保留参数兼容）
                
            Returns:
                (commission_value, used_asin_only_match, should_clear)
            """
            key = (asin, country)
            row_uid = row_uid_map.get(row_index, '')
            
            total_for_row = 0
            used_asin_only_match = False
            has_any_match = False
            
            # === 第1步：UID池匹配 ===
            # 如果offer行有uid（投放链接最后7位），尝试从UID池匹配
            if row_uid:
                uid_key = (asin, country, row_uid)
                if uid_key in asin_country_uid_commission:
                    uid_comm = asin_country_uid_commission[uid_key]
                    total_for_row += uid_comm
                    has_any_match = True
                    self.log_debug(f"  UID池匹配: row={row_index}, ASIN={asin}, 国家={country}, UID={row_uid}, 佣金=${uid_comm:.2f}")
                else:
                    self.log_debug(f"  UID池未命中: row={row_index}, ASIN={asin}, 国家={country}, UID={row_uid}")
            
            # === 第2步：非UID池匹配 ===
            # PB交易无uid的佣金，按(ASIN+国家)匹配，每个组合只分配给一行offer
            if key not in asin_country_no_uid_assigned:
                if key in asin_country_no_uid_commission:
                    no_uid_comm = asin_country_no_uid_commission[key]
                    total_for_row += no_uid_comm
                    asin_country_no_uid_assigned.add(key)
                    has_any_match = True
                    self.log_debug(f"  非UID池匹配: row={row_index}, ASIN={asin}, 国家={country}, 佣金=${no_uid_comm:.2f}")
                elif asin in asin_only_commission and asin not in asin_has_us_offer and asin not in asin_only_commission_assigned:
                    # ASIN-only备选匹配：PB交易国家默认为US但ASIN在飞书中没有US的offer
                    asin_only_comm = asin_only_commission[asin]
                    total_for_row += asin_only_comm
                    asin_only_commission_assigned.add(asin)
                    used_asin_only_match = True
                    has_any_match = True
                    self.log_debug(f"  ASIN-only备选匹配: row={row_index}, ASIN={asin}, 国家={country}, 佣金=${asin_only_comm:.2f}")
            
            if has_any_match:
                self.log_debug(f"  → row={row_index} 最终佣金=${total_for_row:.2f}")
                return (total_for_row, used_asin_only_match, False)
            
            return (None, False, False)
        
        # 记录每个(asin, country)已处理的行数，用于判断是否是第一个offer
        asin_country_processed = {}  # {(asin, country): processed_count}
        
        # 调试日志：检查广告系列和offer表格的匹配情况
        self.log_debug("=== 广告系列与offer表格匹配检查 ===")
        for campaign_key in asin_country_campaigns.keys():
            if campaign_key in asin_country_rows:
                rows_info = asin_country_rows[campaign_key]
                self.log_debug(f"  ✓ {campaign_key[0]}_{campaign_key[1]}: 找到匹配 (active={len(rows_info['active_rows'])}, ended={len(rows_info['ended_rows'])}, untested={len(rows_info['untested_rows'])})")
            else:
                self.log_debug(f"  ✗ {campaign_key[0]}_{campaign_key[1]}: 未在offer表格中找到")
        
        # 第二遍遍历：处理更新
        for row in feishu_data:
            row_index = row.get('row_index')
            asin = row.get('ASIN', '')
            country = row.get('国家代码', '')
            current_status = row.get('状态', '')
            
            if not asin:
                continue
            
            # 标准化国家代码
            if country:
                country = country.upper().strip()
            
            key = (asin, country) if country else None
            
            # 记录当前行是否是该(asin, country)的第一个offer
            if key:
                if key not in asin_country_processed:
                    asin_country_processed[key] = 0
                asin_country_processed[key] += 1
                is_first_offer = (asin_country_processed[key] == 1)
            else:
                is_first_offer = True
            
            # 处理"投放已结束"状态的offer
            if current_status == '投放已结束':
                # 首先检查是否有新的广告系列，如果有，则不跳过，继续后面的处理让状态变成"投放中"
                if key and key in asin_country_campaigns:
                    # 有广告系列，不跳过，继续后面的正常处理流程
                    pass
                else:
                    # 没有广告系列，只更新佣金
                    # 使用辅助函数获取佣金
                    commission_value, used_asin_only_match, _ = get_commission_for_row(asin, country, row_index, is_first_offer)
                    
                    if commission_value is not None:
                        commission_display = round(commission_value, 2)
                        if used_asin_only_match and len(asin_countries.get(asin, set())) > 1:
                            commission_display = f"{commission_display}*"
                        
                        update = {
                            'row_index': row_index,
                            'asin': asin,
                            'country': country,
                            'commission': commission_display
                        }
                        updates.append(update)
                    # 没有广告系列，跳过其他处理
                    continue
            
            # 处理"未测试"或"新复制"状态的offer（没有广告系列的情况）
            if (current_status == '未测试' or current_status == '新复制' or not current_status):
                # 检查是否有广告系列
                if key and key in asin_country_campaigns:
                    # 有广告系列，继续后面的正常处理流程
                    pass
                else:
                    # 没有广告系列，检查是否需要更新佣金
                    # 使用辅助函数获取佣金
                    commission_value, used_asin_only_match, _ = get_commission_for_row(asin, country, row_index, is_first_offer)
                    
                    if commission_value is not None:
                        commission_display = round(commission_value, 2)
                        if used_asin_only_match and len(asin_countries.get(asin, set())) > 1:
                            commission_display = f"{commission_display}*"
                        
                        update = {
                            'row_index': row_index,
                            'asin': asin,
                            'country': country,
                            'commission': commission_display
                        }
                        updates.append(update)
                    # 跳过其他处理（未测试且没有广告系列，不更新状态等）
                    continue
            
            update = {'row_index': row_index, 'asin': asin, 'country': country}
            
            # 优先使用link_id精确匹配的广告系列，否则回退到(ASIN+国家)匹配
            has_precise_match = row_index in row_campaigns
            
            if has_precise_match:
                campaigns = row_campaigns[row_index]
            elif key and key in asin_country_campaigns:
                # 回退到(ASIN+国家)匹配，但排除已被其他行精确认领的广告系列
                all_campaigns = asin_country_campaigns[key]
                if claimed_campaign_ids:
                    campaigns = [c for c in all_campaigns if c.get('campaign_id') not in claimed_campaign_ids]
                else:
                    campaigns = all_campaigns
            else:
                campaigns = []
            
            has_campaigns = len(campaigns) > 0
            
            if has_campaigns:
                total_campaigns = len(campaigns)
                enabled_campaigns = [c for c in campaigns if c['status'] == 'ENABLED']
                paused_campaigns = [c for c in campaigns if c['status'] == 'PAUSED']
                
                # 计算当前广告系列的花费（从Google Ads获取，即全部花费）
                total_cost = sum(c['cost_usd'] for c in campaigns)
                
                # 获取所有账户ID（去除横杠）
                account_ids = list(set(c['account_id'].replace('-', '') for c in campaigns))
                
                # 获取所有广告系列名称
                campaign_names = [c['campaign_name'] for c in campaigns]
                
                # 确定状态
                if len(enabled_campaigns) == total_campaigns:
                    # 全部启用
                    update['status'] = '投放中'
                    update['status_color'] = 'green'
                elif len(enabled_campaigns) > 0:
                    # 部分启用
                    update['status'] = f'投放中({len(enabled_campaigns)}/{total_campaigns})'
                    update['status_color'] = 'green'
                else:
                    # 全部暂停
                    update['status'] = '广告系列暂停中'
                    update['status_color'] = 'orange'
                
                update['ads_ids'] = ','.join(account_ids)
                update['campaign_count'] = total_campaigns
                update['total_cost'] = round(total_cost, 2)
                # 多个广告系列名称用逗号分隔
                update['campaign_names'] = ', '.join(campaign_names)
                
            else:
                # 这个ASIN+国家组合没有广告系列
                if current_status and current_status not in ['未测试', '新复制']:
                    # 之前有状态，现在没有广告系列了
                    update['status'] = '投放已结束'
                    update['status_color'] = 'black'
                    update['ads_ids'] = ''
                    update['campaign_count'] = ''
                    update['campaign_names'] = ''  # 同时清空广告系列名称
                    update['total_cost'] = None  # 保留原有花费数据
                else:
                    # 未测试或新复制状态，不更新
                    continue
            
            # 添加佣金数据 - 非"投放已结束"状态的行
            # 使用辅助函数获取佣金
            commission_value, used_asin_only_match, _ = get_commission_for_row(asin, country, row_index, is_first_offer)
            
            if commission_value is not None:
                commission_display = round(commission_value, 2)
                if used_asin_only_match and len(asin_countries.get(asin, set())) > 1:
                    commission_display = f"{commission_display}*"
                update['commission'] = commission_display
            
            updates.append(update)
        
        return updates
    
    def apply_feishu_updates(self, token, updates):
        """应用飞书表格更新"""
        spreadsheet_token = self.feishu_spreadsheet_var.get().strip()
        sheet_id = self.feishu_sheet_id_var.get().strip()
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # 获取列映射（表头名称 -> 列字母）
        column_map = getattr(self, 'feishu_column_map', {})
        
        # 定义数据字段与表头名称的对应关系
        field_to_header = {
            'status': '状态',
            'ads_ids': '投放中的ads',
            'campaign_count': '广告系列数量',
            'campaign_names': '广告系列名称',
            'total_cost': '广告系列总花费',
            'commission': '总佣金'
        }
        
        # 检查必要的列是否存在
        missing_columns = []
        for field, header in field_to_header.items():
            if header not in column_map:
                missing_columns.append(header)
        
        if missing_columns:
            self.log_manage(f"  警告：以下列在表头中未找到: {missing_columns}")
            self.log_manage(f"  当前表头列: {list(column_map.keys())}")
        
        self.log_manage(f"  需要更新 {len(updates)} 行数据")
        self.log_manage(f"  列位置映射: 状态={column_map.get('状态', '?')}, 投放中的ads={column_map.get('投放中的ads', '?')}, 广告系列数量={column_map.get('广告系列数量', '?')}, 广告系列名称={column_map.get('广告系列名称', '?')}, 广告系列总花费={column_map.get('广告系列总花费', '?')}, 总佣金={column_map.get('总佣金', '?')}")
        
        # 分批处理，每批最多10条更新
        batch_size = 10
        total_updates = len(updates)
        
        for batch_start in range(0, total_updates, batch_size):
            if self.stop_flag:
                break
            
            batch_end = min(batch_start + batch_size, total_updates)
            batch_updates = updates[batch_start:batch_end]
            
            # 准备批量更新的数据
            value_ranges = []
            style_updates = []
            
            for update in batch_updates:
                row_index = update['row_index']
                
                # 状态列
                if 'status' in update and '状态' in column_map:
                    col = column_map['状态']
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_index}:{col}{row_index}",
                        'values': [[update['status']]]
                    })
                
                # 投放中的ads列
                if 'ads_ids' in update and '投放中的ads' in column_map:
                    col = column_map['投放中的ads']
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_index}:{col}{row_index}",
                        'values': [[update['ads_ids']]]
                    })
                
                # 广告系列数量列
                if 'campaign_count' in update and '广告系列数量' in column_map:
                    col = column_map['广告系列数量']
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_index}:{col}{row_index}",
                        'values': [[update['campaign_count']]]
                    })
                
                # 广告系列名称列
                if 'campaign_names' in update and '广告系列名称' in column_map:
                    col = column_map['广告系列名称']
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_index}:{col}{row_index}",
                        'values': [[update['campaign_names']]]
                    })
                
                # 广告系列总花费列
                if 'total_cost' in update and update['total_cost'] is not None and '广告系列总花费' in column_map:
                    col = column_map['广告系列总花费']
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_index}:{col}{row_index}",
                        'values': [[f"${update['total_cost']}"]]
                    })
                
                # 总佣金列
                if 'commission' in update and '总佣金' in column_map:
                    col = column_map['总佣金']
                    value_ranges.append({
                        'range': f"{sheet_id}!{col}{row_index}:{col}{row_index}",
                        'values': [[f"${update['commission']}"]]
                    })
                
                # 收集样式更新（状态列的样式）
                if 'status_color' in update and '状态' in column_map:
                    style_updates.append({
                        'row_index': row_index,
                        'color': update['status_color'],
                        'column': column_map['状态']
                    })
            
            # 批量更新值 - 使用 values_batch_update API
            if value_ranges:
                try:
                    url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
                    body = {
                        "valueRanges": value_ranges
                    }
                    self.log_manage(f"    调用API: POST {url}")
                    self.log_manage(f"    更新范围数: {len(value_ranges)}")
                    response = requests.post(url, headers=headers, json=body, timeout=60)
                    result = response.json()
                    if result.get('code') != 0:
                        self.log_manage(f"  批量更新失败 (code={result.get('code')}): {result.get('msg', 'Unknown error')}")
                    else:
                        # 成功时记录更新的单元格数
                        responses = result.get('data', {}).get('responses', [])
                        total_cells = sum(r.get('updatedCells', 0) for r in responses)
                        self.log_manage(f"    成功更新 {total_cells} 个单元格")
                except Exception as e:
                    self.log_manage(f"  批量更新异常: {e}")
            
            # 批量更新样式
            if style_updates:
                self.batch_update_cell_styles(token, spreadsheet_token, sheet_id, style_updates)
            
            self.log_manage(f"  已更新 {batch_end}/{total_updates} 行")
            self.update_progress_manage(f"更新中 {batch_end}/{total_updates}")
            
            time.sleep(0.3)  # 避免请求过快
        
        self.log_manage(f"  更新完成")
    
    def batch_update_cell_styles(self, token, spreadsheet_token, sheet_id, style_updates):
        """批量更新单元格样式（字体颜色+加粗）"""
        # 字体颜色映射 - 使用十六进制颜色代码
        color_map = {
            'green': '#00AA00',    # 绿色
            'orange': '#FF8C00',   # 橙色
            'black': '#333333'     # 深灰/黑色
        }
        
        # 按颜色分组
        color_groups = {}
        for style in style_updates:
            color = style['color']
            col = style.get('column', 'A')  # 使用动态列，默认为A
            if color not in color_groups:
                color_groups[color] = []
            color_groups[color].append(f"{sheet_id}!{col}{style['row_index']}:{col}{style['row_index']}")
        
        # 构建批量样式更新数据
        data = []
        for color, ranges in color_groups.items():
            font_color = color_map.get(color, '#000000')
            data.append({
                "ranges": ranges,
                "style": {
                    "font": {
                        "bold": True
                    },
                    "foreColor": font_color
                }
            })
        
        if not data:
            return
        
        try:
            url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/styles_batch_update"
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            body = {
                "data": data
            }
            response = requests.put(url, headers=headers, json=body, timeout=60)
            result = response.json()
            if result.get('code') != 0:
                self.log_manage(f"    样式更新失败 (code={result.get('code')}): {result.get('msg', 'Unknown error')}")
            else:
                total_cells = result.get('data', {}).get('totalUpdatedCells', 0)
                self.log_manage(f"    样式更新成功: {total_cells} 个单元格")
        except Exception as e:
            self.log_manage(f"    样式更新异常: {e}")
    
    def update_cell_style(self, token, spreadsheet_token, sheet_id, row_index, color, column='A'):
        """更新单个单元格样式（字体颜色+加粗）- 兼容旧接口"""
        self.batch_update_cell_styles(token, spreadsheet_token, sheet_id, [{'row_index': row_index, 'color': color, 'column': column}])
    
    def get_mcc_total_cost(self):
        """从MCC获取所有子账户从2026-01-01至今的总花费（含已删除广告系列），转为USD"""
        try:
            client = self.get_google_ads_client()
            if not client:
                self.log_manage("  ✗ 无法创建Google Ads客户端")
                return None
            
            ga_service = client.get_service('GoogleAdsService')
            mcc_id = self.google_mcc_id_var.get().strip()
            
            # 获取所有子账户
            query = """
                SELECT customer_client.id, customer_client.manager, customer_client.currency_code
                FROM customer_client
                WHERE customer_client.level <= 1
            """
            response = ga_service.search(customer_id=mcc_id, query=query)
            
            sub_accounts = []
            for row in response:
                if not row.customer_client.manager:
                    sub_accounts.append({
                        'id': str(row.customer_client.id),
                        'currency': row.customer_client.currency_code
                    })
            
            CNY_TO_USD_RATE = 0.14
            today = datetime.now().strftime('%Y-%m-%d')
            total_cost_usd = 0
            
            for account in sub_accounts:
                if self.stop_flag:
                    break
                try:
                    account_currency = account.get('currency', 'USD')
                    
                    # 查询该账户从2026-01-01至今的总花费（包含所有广告系列，含已删除的）
                    cost_query = f"""
                        SELECT metrics.cost_micros
                        FROM customer
                        WHERE segments.date >= '2026-01-01'
                            AND segments.date <= '{today}'
                    """
                    cost_response = ga_service.search(customer_id=account['id'], query=cost_query)
                    
                    for row in cost_response:
                        cost_original = row.metrics.cost_micros / 1000000 if row.metrics.cost_micros else 0
                        if account_currency == 'CNY':
                            total_cost_usd += cost_original * CNY_TO_USD_RATE
                        else:
                            total_cost_usd += cost_original
                            
                except Exception as e:
                    self.log_manage(f"  ⚠ 获取账户 {account['id']} 花费失败: {str(e)[:80]}")
            
            return total_cost_usd
            
        except Exception as e:
            self.log_manage(f"  获取MCC总花费失败: {e}")
            return None
    
    def update_feishu_summary_row(self, token, total_commission, total_cost, unmatched_commission=0, rejected_commission=0):
        """更新飞书表格第二行'总计'行的汇总数据
        
        参数:
            token: 飞书访问令牌
            total_commission: 已匹配的总佣金
            total_cost: 总花费
            unmatched_commission: 未匹配的佣金（找不到对应offer的）
            rejected_commission: Rejected状态的佣金
        """
        spreadsheet_token = self.feishu_spreadsheet_var.get().strip()
        sheet_id = self.feishu_sheet_id_var.get().strip()
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # 获取列映射
        column_map = getattr(self, 'feishu_column_map', {})
        
        # 准备更新数据
        value_ranges = []
        
        # 总佣金列 - 格式：$4481.79 (+$133.91-$278.33)
        # 含义：已匹配佣金 (+未匹配佣金-Rejected佣金)
        if '总佣金' in column_map:
            col = column_map['总佣金']
            commission_display = f"${total_commission:.2f} (+${unmatched_commission:.2f}-${rejected_commission:.2f})"
            value_ranges.append({
                'range': f"{sheet_id}!{col}2:{col}2",
                'values': [[commission_display]]
            })
        else:
            self.log_manage("  警告：未找到'总佣金'列")
        
        # 广告系列总花费列
        if '广告系列总花费' in column_map:
            col = column_map['广告系列总花费']
            value_ranges.append({
                'range': f"{sheet_id}!{col}2:{col}2",
                'values': [[f"${total_cost:.2f}"]]
            })
        else:
            self.log_manage("  警告：未找到'广告系列总花费'列")
        
        if not value_ranges:
            self.log_manage("  无法更新总计行：缺少必要的列")
            return
        
        try:
            url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{spreadsheet_token}/values_batch_update"
            body = {
                "valueRanges": value_ranges
            }
            response = requests.post(url, headers=headers, json=body, timeout=60)
            result = response.json()
            if result.get('code') != 0:
                self.log_manage(f"  更新总计行失败 (code={result.get('code')}): {result.get('msg', 'Unknown error')}")
            else:
                self.log_manage(f"  ✅ 总计行更新成功: 总佣金=${total_commission:.2f} (+${unmatched_commission:.2f}-${rejected_commission:.2f}), 总花费=${total_cost:.2f}")
        except Exception as e:
            self.log_manage(f"  更新总计行异常: {e}")

    # ==================== 表格行排序功能 ====================
    
    def parse_commission_value(self, val):
        """从佣金字符串中提取数值
        
        支持格式: "$123.45", "*$123.45", "$4481.79 (+$133.91-$278.33)", "123.45", 数字等
        """
        if not val:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        val_str = str(val).strip()
        # 去除$、*、逗号
        val_str = val_str.replace('$', '').replace('*', '').replace(',', '')
        # 如果有空格+括号补充信息（如 "4481.79 (+133.91-278.33)"），只取第一部分
        if ' ' in val_str:
            val_str = val_str.split(' ')[0]
        # 如果有中文括号（如 "1143.82（35.63）"）
        if '（' in val_str:
            val_str = val_str.split('（')[0]
        try:
            return float(val_str) if val_str else 0.0
        except (ValueError, TypeError):
            return 0.0
    
    def feishu_move_dimension(self, token, spreadsheet_token, sheet_id, start_index, end_index, destination_index):
        """调用飞书API移动行
        
        参数:
            token: 飞书访问令牌
            spreadsheet_token: 电子表格token
            sheet_id: 工作表ID
            start_index: 要移动的起始行（0-based，包含）
            end_index: 要移动的结束行（0-based，包含）
            destination_index: 目标位置（0-based，移动后的位置，基于移除源行后的索引）
        
        返回:
            bool: 是否成功
        """
        url = f"{FEISHU_API_BASE_URL}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}/move_dimension"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8"
        }
        body = {
            "source": {
                "major_dimension": "ROWS",
                "start_index": start_index,
                "end_index": end_index
            },
            "destination_index": destination_index
        }
        
        try:
            response = requests.post(url, headers=headers, json=body, timeout=30)
            result = response.json()
            if result.get('code') != 0:
                self.log_manage(f"    移动行失败 (src={start_index}, dst={destination_index}): code={result.get('code')}, {result.get('msg', '')}")
                return False
            return True
        except Exception as e:
            self.log_manage(f"    移动行异常: {e}")
            return False
    
    def sort_sheet_rows(self, token, spreadsheet_token, sheet_id, data_rows, data_start_0based, table_name=""):
        """对飞书表格的数据行进行排序（通用方法）
        
        参数:
            token: 飞书访问令牌
            spreadsheet_token: 电子表格token
            sheet_id: 工作表ID
            data_rows: 数据行列表，每项为 {'row_index': int, 'status': str, 'total_commission': float}，
                       必须按当前物理顺序排列
            data_start_0based: 第一个数据行的0-based索引
            table_name: 表格名称（用于日志）
        
        返回:
            int: 执行的移动次数
        """
        if len(data_rows) <= 1:
            return 0
        
        # 定义状态优先级：投放中(0) > 暂停中(1) > 投放已结束(2) > 未测试(3)
        status_priority = {
            '投放中': 0,
            '广告系列暂停中': 1,
            '暂停中': 1,
            '暂停': 1,
            '投放已结束': 2,
            '未测试': 3,
            '新复制': 3,
        }
        
        def get_status_priority(status_str):
            """获取状态优先级，支持带日期后缀的状态（如'广告系列暂停中2026-2-15'）"""
            if status_str in status_priority:
                return status_priority[status_str]
            # 前缀匹配（处理带日期的状态）
            for key, pri in status_priority.items():
                if status_str.startswith(key):
                    return pri
            return 99
        
        # 排序：先按状态优先级升序，再按总佣金降序
        sorted_rows = sorted(data_rows, key=lambda r: (
            get_status_priority(r['status']),
            -r['total_commission']
        ))
        
        # 统计各状态数量
        status_count = {}
        for r in sorted_rows:
            s = r['status'] if r['status'] else '(空)'
            status_count[s] = status_count.get(s, 0) + 1
        status_summary = ', '.join(f"{s}:{c}行" for s, c in status_count.items())
        self.log_manage(f"    {table_name}状态分布: {status_summary}")
        
        # 对比当前顺序和目标顺序
        current_order = [r['row_index'] for r in data_rows]
        desired_order = [r['row_index'] for r in sorted_rows]
        
        if current_order == desired_order:
            self.log_manage(f"    {table_name}已经是正确的排序顺序，无需移动")
            return 0
        
        # 预先计算实际需要移动的次数（模拟排序过程）
        temp_order = list(current_order)
        expected_moves = 0
        for i in range(len(temp_order)):
            target = desired_order[i]
            j = temp_order.index(target)
            if j != i:
                expected_moves += 1
                item = temp_order.pop(j)
                temp_order.insert(i, item)
        
        # 100次/分钟 => 每次调用间隔至少0.6秒
        MIN_INTERVAL = 0.6
        estimated_seconds = int(expected_moves * MIN_INTERVAL)
        self.log_manage(f"    需要执行 {expected_moves} 次移动，预计耗时约 {estimated_seconds // 60} 分 {estimated_seconds % 60} 秒")
        
        # 使用选择排序算法，从上往下依次将正确的行移到目标位置
        working_order = list(current_order)
        move_count = 0
        start_time = time.time()
        last_call_time = 0  # 上一次API调用的时间戳
        
        for i in range(len(working_order)):
            if self.stop_flag:
                self.log_manage(f"    排序被用户中断")
                break
            
            target = desired_order[i]
            j = working_order.index(target)
            if j == i:
                continue  # 已在正确位置
            
            # 精确控制API调用频率：确保两次调用之间间隔至少MIN_INTERVAL秒
            now = time.time()
            elapsed_since_last = now - last_call_time
            if elapsed_since_last < MIN_INTERVAL:
                time.sleep(MIN_INTERVAL - elapsed_since_last)
            
            # 计算飞书API所需的0-based行索引
            src_0based = data_start_0based + j
            dst_0based = data_start_0based + i
            
            last_call_time = time.time()
            success = self.feishu_move_dimension(
                token, spreadsheet_token, sheet_id,
                src_0based, src_0based, dst_0based
            )
            
            if success:
                # 更新工作顺序：模拟移动操作（pop后insert）
                item = working_order.pop(j)
                working_order.insert(i, item)
                move_count += 1
                
                # 每20次移动输出一次进度
                if move_count % 20 == 0:
                    elapsed = time.time() - start_time
                    remaining = (expected_moves - move_count) * (elapsed / move_count) if move_count > 0 else 0
                    self.log_manage(f"    进度: {move_count}/{expected_moves} ({move_count*100//expected_moves}%)，已用 {elapsed:.0f}s，剩余约 {remaining:.0f}s")
                    self.update_progress_manage(f"排序{table_name} {move_count}/{expected_moves}")
            else:
                self.log_manage(f"    排序中断：移动行失败")
                break
        
        elapsed_total = time.time() - start_time
        self.log_manage(f"    排序完成: {move_count}/{expected_moves} 次移动，耗时 {elapsed_total:.0f} 秒")
        return move_count
    
    def sort_offer_table(self, token):
        """对Offer表格进行排序
        
        排序规则：
        1. 表头行(第1行)、总计行(第2行)固定不动
        2. 数据行按状态排序：投放中 > 暂停中 > 投放已结束 > 未测试
        3. 同状态内按总佣金降序排列
        """
        self.log_manage("  对Offer表格行排序...")
        
        spreadsheet_token = self.feishu_spreadsheet_var.get().strip()
        sheet_id = self.feishu_sheet_id_var.get().strip()
        
        # 重新读取最新数据
        feishu_data = self.get_feishu_sheet_data(token)
        
        if not feishu_data:
            self.log_manage("    无法读取Offer表格数据")
            return
        
        # 找到最后一个有内容的数据行（排除总计行 row_index=2）
        last_data_ri = 0
        row_by_index = {}
        for row in feishu_data:
            ri = row.get('row_index')
            if ri is None or ri == 2:  # 跳过总计行
                continue
            if ri >= 3:
                row_by_index[ri] = row
                # 检查行是否有实际内容
                has_content = (
                    row.get('ASIN') or 
                    row.get('状态') or 
                    row.get('品牌名称') or
                    row.get('总佣金')
                )
                if has_content:
                    last_data_ri = max(last_data_ri, ri)
        
        if last_data_ri < 3:
            self.log_manage("    没有需要排序的数据行")
            return
        
        # 构建连续的数据行列表（从第3行到最后有内容的行）
        data_rows = []
        for ri in range(3, last_data_ri + 1):
            row = row_by_index.get(ri, {})
            status = str(row.get('状态', '') or '').strip()
            commission = self.parse_commission_value(row.get('总佣金', ''))
            data_rows.append({
                'row_index': ri,
                'status': status,
                'total_commission': commission
            })
        
        self.log_manage(f"    共 {len(data_rows)} 行数据需要排序")
        
        # Offer表数据行从第3行开始（1-based），即0-based索引为2
        move_count = self.sort_sheet_rows(
            token, spreadsheet_token, sheet_id,
            data_rows, data_start_0based=2, table_name="Offer表"
        )
        
        self.log_manage(f"  ✅ Offer表格排序完成，共执行 {move_count} 次移动")
    
    def sort_campaigns_table(self, token):
        """对广告系列表格进行排序
        
        排序规则：
        1. 表头行(第1行)固定不动
        2. 数据行按状态排序：投放中 > 广告系列暂停中 > 投放已结束 > 未测试
        3. 同状态内按总佣金降序排列
        """
        self.log_manage("  对广告系列表格行排序...")
        
        campaigns_spreadsheet_token = "KnJ1wphpBiVMrGkWl5ncUkMGnfe"
        campaigns_sheet_id = "XrkOF7"
        
        # 读取最新数据
        result = self.read_campaigns_sheet(token, campaigns_spreadsheet_token, campaigns_sheet_id)
        if result is None:
            self.log_manage("    无法读取广告系列表格数据")
            return
        
        existing_rows, column_map, first_empty_row = result
        
        if not existing_rows:
            self.log_manage("    广告系列表格无数据")
            return
        
        # 找到最后一个有内容的数据行
        last_data_ri = 0
        row_by_index = {}
        for row in existing_rows:
            ri = row.get('row_index')
            if ri is None:
                continue
            row_by_index[ri] = row
            campaign_name = row.get('广告系列名称', '')
            if campaign_name:
                last_data_ri = max(last_data_ri, ri)
        
        if last_data_ri < 2:
            self.log_manage("    没有需要排序的数据行")
            return
        
        # 构建连续的数据行列表（从第2行到最后有内容的行）
        data_rows = []
        for ri in range(2, last_data_ri + 1):
            row = row_by_index.get(ri, {})
            status = str(row.get('状态', '') or '').strip()
            commission = self.parse_commission_value(row.get('总佣金', ''))
            data_rows.append({
                'row_index': ri,
                'status': status,
                'total_commission': commission
            })
        
        self.log_manage(f"    共 {len(data_rows)} 行数据需要排序")
        
        # 广告系列表数据行从第2行开始（1-based），即0-based索引为1
        move_count = self.sort_sheet_rows(
            token, campaigns_spreadsheet_token, campaigns_sheet_id,
            data_rows, data_start_0based=1, table_name="广告系列表"
        )
        
        self.log_manage(f"  ✅ 广告系列表格排序完成，共执行 {move_count} 次移动")


def main():
    root = tk.Tk()
    app = OfferToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
