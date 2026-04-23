"""
一次性脚本：为飞书表格中所有offer刷新投放链接（添加UID追踪参数）
"""
import requests
import json
import os
import time
import secrets
import string
import sys

# 强制实时输出
sys.stdout.reconfigure(line_buffering=True)

# 配置
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
USED_UIDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "used_uids.txt")

PB_API_BASE_URL = "https://app.partnerboost.com"
FEISHU_API_BASE_URL = "https://open.feishu.cn/open-apis"

# 加载配置
with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
    config = json.load(f)

PB_TOKEN = config.get("pb_token", "")
FEISHU_APP_ID = config.get("feishu_app_id", "")
FEISHU_APP_SECRET = config.get("feishu_app_secret", "")
FEISHU_SPREADSHEET_TOKEN = config.get("feishu_spreadsheet_token", "")
FEISHU_SHEET_ID = config.get("feishu_sheet_id", "")

# UID管理
used_uids = set()

def load_used_uids():
    """加载已使用的UID"""
    global used_uids
    if os.path.exists(USED_UIDS_FILE):
        with open(USED_UIDS_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                uid = line.strip()
                if uid:
                    used_uids.add(uid)
    print(f"已加载 {len(used_uids)} 个历史UID")

def generate_random_uid():
    """生成7位真随机UID，确保不重复"""
    global used_uids
    alphabet = string.ascii_letters + string.digits
    
    for _ in range(1000):
        uid = ''.join(secrets.choice(alphabet) for _ in range(7))
        if uid not in used_uids:
            used_uids.add(uid)
            # 即时保存
            with open(USED_UIDS_FILE, 'a', encoding='utf-8') as f:
                f.write(uid + '\n')
            return uid
    
    return f"{int(time.time()) % 10000000:07d}"

def get_feishu_token():
    """获取飞书访问令牌"""
    url = f"{FEISHU_API_BASE_URL}/auth/v3/tenant_access_token/internal"
    body = {
        "app_id": FEISHU_APP_ID,
        "app_secret": FEISHU_APP_SECRET
    }
    response = requests.post(url, json=body, timeout=30)
    data = response.json()
    if data.get("code") == 0:
        return data.get("tenant_access_token")
    return None

def get_feishu_sheet_data(token):
    """获取飞书表格数据"""
    # 获取行数
    meta_url = f"{FEISHU_API_BASE_URL}/sheets/v3/spreadsheets/{FEISHU_SPREADSHEET_TOKEN}/sheets/{FEISHU_SHEET_ID}"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(meta_url, headers=headers, timeout=30)
    meta = response.json()
    row_count = meta.get("data", {}).get("sheet", {}).get("grid_properties", {}).get("row_count", 500)
    
    # 读取数据
    range_str = f"{FEISHU_SHEET_ID}!A1:Z{row_count}"
    data_url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{FEISHU_SPREADSHEET_TOKEN}/values/{range_str}"
    response = requests.get(data_url, headers=headers, timeout=30)
    data = response.json()
    values = data.get("data", {}).get("valueRange", {}).get("values", [])
    
    rows = []
    column_map = {}
    
    if len(values) > 0:
        headers_row = values[0]
        for j, header in enumerate(headers_row):
            if header:
                # 列索引转字母
                col_letter = ""
                idx = j
                while idx >= 0:
                    col_letter = chr(idx % 26 + ord('A')) + col_letter
                    idx = idx // 26 - 1
                column_map[header] = col_letter
        
        for i, row in enumerate(values[1:], start=2):
            row_data = {'row_index': i}
            for j, cell in enumerate(row):
                if j < len(headers_row):
                    if isinstance(cell, list) and len(cell) > 0 and isinstance(cell[0], dict):
                        cell = cell[0].get('link', '') or cell[0].get('text', '')
                    row_data[headers_row[j]] = cell
            rows.append(row_data)
    
    return rows, column_map

def get_partnerboost_link(asin, country_code, uid):
    """获取带UID的投放链接"""
    url = f"{PB_API_BASE_URL}/api/datafeed/get_amazon_link_by_asin"
    body = {
        "token": PB_TOKEN,
        "asins": asin,
        "country_code": country_code,
        "uid": uid,
        "return_partnerboost_link": 1
    }
    try:
        response = requests.post(url, json=body, timeout=30)
        if response.status_code == 200:
            data = response.json()
            if data.get("status", {}).get("code") == 0:
                link_data = data.get("data", [])
                if link_data:
                    return link_data[0].get("partnerboost_link", "")
    except Exception as e:
        print(f"    获取链接异常: {e}")
    return ""

def batch_update_feishu(token, updates, link_col):
    """批量更新飞书表格"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    batch_size = 20
    total = len(updates)
    
    for batch_start in range(0, total, batch_size):
        batch_end = min(batch_start + batch_size, total)
        batch = updates[batch_start:batch_end]
        
        value_ranges = []
        for u in batch:
            value_ranges.append({
                'range': f"{FEISHU_SHEET_ID}!{link_col}{u['row_index']}:{link_col}{u['row_index']}",
                'values': [[u['new_link']]]
            })
        
        url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{FEISHU_SPREADSHEET_TOKEN}/values_batch_update"
        body = {"valueRanges": value_ranges}
        response = requests.post(url, headers=headers, json=body, timeout=60)
        result = response.json()
        
        if result.get('code') == 0:
            print(f"  已更新 {batch_end}/{total} 行")
        else:
            print(f"  更新失败: {result.get('msg', 'Unknown error')}")
        
        time.sleep(0.3)

def main():
    print("=" * 60)
    print("开始刷新投放链接（添加UID追踪参数）")
    print("=" * 60)
    
    # 1. 加载已使用的UID
    print("\n【步骤1】加载UID历史记录")
    load_used_uids()
    
    # 2. 获取飞书Token
    print("\n【步骤2】获取飞书访问令牌")
    token = get_feishu_token()
    if not token:
        print("  ✗ 获取Token失败")
        return
    print("  ✓ 获取Token成功")
    
    # 3. 获取飞书数据
    print("\n【步骤3】获取飞书表格数据")
    rows, column_map = get_feishu_sheet_data(token)
    print(f"  获取到 {len(rows)} 行数据")
    
    if '投放链接' not in column_map:
        print("  ✗ 未找到'投放链接'列")
        return
    link_col = column_map['投放链接']
    print(f"  投放链接列: {link_col}")
    
    # 4. 遍历生成新链接
    print("\n【步骤4】为每个offer生成带UID的投放链接")
    updates = []
    success = 0
    skipped = 0
    failed = 0
    
    for idx, row in enumerate(rows):
        row_index = row.get('row_index')
        asin = row.get('ASIN', '')
        country = row.get('国家代码', '')
        
        # 跳过无效行
        if not asin or not country or row_index == 2:
            skipped += 1
            continue
        
        # 生成UID和新链接
        uid = generate_random_uid()
        new_link = get_partnerboost_link(asin, country, uid)
        
        if new_link:
            updates.append({
                'row_index': row_index,
                'asin': asin,
                'country': country,
                'uid': uid,
                'new_link': new_link
            })
            success += 1
        else:
            failed += 1
            print(f"  [失败] {asin}_{country}")
        
        # 进度
        if (idx + 1) % 50 == 0:
            print(f"  已处理 {idx + 1}/{len(rows)} 个offer...")
        
        time.sleep(0.2)  # API限速
    
    print(f"\n  生成完成: 成功={success}, 跳过={skipped}, 失败={failed}")
    
    # 5. 更新飞书
    if updates:
        print(f"\n【步骤5】更新飞书表格（共 {len(updates)} 条）")
        batch_update_feishu(token, updates, link_col)
        
        # 显示部分结果
        print("\n  更新详情（前10条）:")
        for u in updates[:10]:
            print(f"    • {u['asin']}_{u['country']}: uid={u['uid']}")
        if len(updates) > 10:
            print(f"    ... 还有 {len(updates) - 10} 条")
    
    print("\n" + "=" * 60)
    print("✅ 完成！")
    print(f"  • 成功更新: {success} 个")
    print(f"  • 跳过: {skipped} 个")
    print(f"  • 失败: {failed} 个")
    print("=" * 60)

if __name__ == "__main__":
    main()

