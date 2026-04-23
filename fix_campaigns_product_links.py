"""修复脚本：
1. 清除J列（多余的表头+数据）
2. 自动检测正确的'产品链接'和'投放链接'列
3. 对每行的投放链接解析重定向，写入正确的产品链接列"""
import requests
import re
import time

FEISHU_API_BASE_URL = "https://open.feishu.cn/open-apis"
SPREADSHEET_TOKEN = "KnJ1wphpBiVMrGkWl5ncUkMGnfe"
SHEET_ID = "XrkOF7"

def get_token():
    url = f"{FEISHU_API_BASE_URL}/auth/v3/tenant_access_token/internal"
    body = {"app_id": "cli_a8517363cf3bd013", "app_secret": "O4Sm3UNHjpykF9OZq3LroblsrCVYyQEp"}
    resp = requests.post(url, json=body, timeout=30)
    data = resp.json()
    return data.get("tenant_access_token") if data.get("code") == 0 else None

def resolve_redirect_url(tracking_link):
    if not tracking_link:
        return ""
    try:
        resp = requests.get(tracking_link, timeout=15)
        if resp.status_code == 200:
            html = resp.text
            match = re.search(r'var\s+u\s*=\s*"([^"]+)";\s*\n\s*location\.replace', html)
            if match:
                return match.group(1)
            match = re.search(r'<meta\s+http-equiv="refresh"\s+content="[^;]*;\s*url=([^"]+)"', html)
            if match:
                return match.group(1)
        return ""
    except:
        return ""

def idx_to_col(index):
    result = ""
    while True:
        result = chr(ord('A') + index % 26) + result
        index = index // 26 - 1
        if index < 0:
            break
    return result

def main():
    token = get_token()
    if not token:
        return
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # 1. 先清除J列表头（之前误写的）
    print("清除J列多余表头...")
    url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_batch_update"
    resp = requests.post(url, headers=headers, json={"valueRanges": [
        {"range": f"{SHEET_ID}!J1:J1", "values": [[""]]}
    ]}, timeout=30)
    print(f"  {'✓' if resp.json().get('code')==0 else '✗'}")
    
    # 2. 读取表格（宽范围）
    print("读取广告系列表格...")
    url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:J1000"
    resp = requests.get(url, headers=headers, timeout=30)
    data = resp.json()
    if data.get('code') != 0:
        print(f"失败: {data.get('msg')}")
        return
    values = data.get('data', {}).get('valueRange', {}).get('values', [])
    if len(values) < 2:
        print("数据不足")
        return
    
    headers_row = values[0]
    print(f"表头: {headers_row}")
    
    # 3. 自动检测列（文字匹配，取第一个匹配的）
    tracking_col = None
    product_col = None
    for i, h in enumerate(headers_row):
        h_str = str(h).strip() if h else ''
        if h_str == '投放链接' and tracking_col is None:
            tracking_col = i
        if h_str == '产品链接' and product_col is None:
            product_col = i
    
    if tracking_col is None:
        print("未找到'投放链接'列！")
        return
    if product_col is None:
        print("未找到'产品链接'列！")
        return
    
    product_col_letter = idx_to_col(product_col)
    print(f"投放链接列: {idx_to_col(tracking_col)} (索引{tracking_col})")
    print(f"产品链接列: {product_col_letter} (索引{product_col})")
    
    # 4. 遍历每行
    updates = []
    total = len(values) - 1
    for row_idx, row in enumerate(values[1:], start=2):
        tracking_link = ''
        if tracking_col < len(row):
            raw = row[tracking_col]
            if raw is None:
                continue
            if isinstance(raw, list) and len(raw) > 0 and isinstance(raw[0], dict):
                tracking_link = raw[0].get('link', '') or raw[0].get('text', '')
            elif isinstance(raw, str):
                tracking_link = raw.strip()
        
        if not tracking_link or not tracking_link.startswith('http'):
            continue
        
        print(f"  行{row_idx}/{total+1}: {tracking_link[:45]}... ", end="", flush=True)
        redirect_url = resolve_redirect_url(tracking_link)
        if redirect_url:
            updates.append((row_idx, redirect_url))
            print("✓")
        else:
            print("✗")
        time.sleep(0.15)
    
    print(f"\n共 {len(updates)} 行需要更新")
    if not updates:
        return
    
    # 5. 写入正确的产品链接列
    print(f"写入{product_col_letter}列...")
    value_ranges = [{"range": f"{SHEET_ID}!{product_col_letter}{r}:{product_col_letter}{r}", "values": [[v]]} for r, v in updates]
    
    batch_size = 100
    for i in range(0, len(value_ranges), batch_size):
        batch = value_ranges[i:i+batch_size]
        url = f"{FEISHU_API_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values_batch_update"
        resp = requests.post(url, headers=headers, json={"valueRanges": batch}, timeout=30)
        result = resp.json()
        print(f"  批次{i//batch_size+1}: {'✓ '+str(len(batch))+'个' if result.get('code')==0 else '✗ '+result.get('msg','')}")
    
    print("\n✅ 全部完成！")

if __name__ == '__main__':
    main()
