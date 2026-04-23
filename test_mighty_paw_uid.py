"""查看 Mighty Paw 品牌的佣金数据是否带有 uid"""

import requests
from datetime import datetime, timedelta
import json

# 从 config.json 读取配置
config_file = "config.json"
with open(config_file, 'r', encoding='utf-8') as f:
    config = json.load(f)

PB_TOKEN = config.get("pb_token", "")
BASE_URL = "https://app.partnerboost.com"

print("=" * 60)
print("查看 Mighty Paw 品牌的佣金数据")
print("=" * 60)

# 获取最近60天的交易数据
end_date = datetime.now()
begin_date = end_date - timedelta(days=60)

url = f"{BASE_URL}/api/transaction/get_transactions"
body = {
    "token": PB_TOKEN,
    "begin_date": begin_date.strftime("%Y-%m-%d"),
    "end_date": end_date.strftime("%Y-%m-%d"),
    "type": "json",
    "status": "All",
    "limit": 2000,
    "page": 1
}

print(f"\n查询日期范围: {begin_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")

response = requests.post(url, json=body, timeout=60)
data = response.json()

if data.get("status", {}).get("code") != 0:
    print(f"API 错误: {data.get('status', {}).get('message')}")
else:
    transactions = data.get("data", {}).get("list", [])
    print(f"获取到 {len(transactions)} 条交易记录")
    
    # 筛选 Mighty Paw 品牌的交易
    mighty_paw_trans = [t for t in transactions if "Mighty" in str(t.get("merchant_name", "")) or "mighty" in str(t.get("merchant_name", "")).lower()]
    
    print(f"\nMighty Paw 品牌的交易记录: {len(mighty_paw_trans)} 条")
    print("-" * 60)
    
    if mighty_paw_trans:
        for i, trans in enumerate(mighty_paw_trans, 1):
            print(f"\n[{i}] 交易详情:")
            print(f"    merchant_name: {trans.get('merchant_name')}")
            print(f"    prod_id (ASIN): {trans.get('prod_id')}")
            print(f"    order_id: {trans.get('order_id')}")
            print(f"    sale_amount: ${trans.get('sale_amount')}")
            print(f"    sale_comm: ${trans.get('sale_comm')}")
            print(f"    status: {trans.get('status')}")
            print(f"    click_ref: {trans.get('click_ref')}")
            print(f"    ----- UID 字段 -----")
            print(f"    uid: '{trans.get('uid', '')}'")
            print(f"    uid2: '{trans.get('uid2', '')}'")
            print(f"    uid3: '{trans.get('uid3', '')}'")
            print(f"    uid4: '{trans.get('uid4', '')}'")
            print(f"    uid5: '{trans.get('uid5', '')}'")
    else:
        print("未找到 Mighty Paw 品牌的交易记录")
        
        # 显示所有品牌
        brands = set(t.get("merchant_name", "") for t in transactions)
        print(f"\n所有品牌列表: {sorted(brands)}")

print("\n" + "=" * 60)
print("查询完成")
print("=" * 60)

