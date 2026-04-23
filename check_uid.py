"""简化版：查看 Mighty Paw 品牌的 uid 字段"""
import requests
import json
from datetime import datetime, timedelta

# 读取配置
with open("config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

url = "https://app.partnerboost.com/api.php?mod=medium&op=transaction"
body = {
    "token": config.get("pb_token", ""),
    "begin_date": (datetime.now() - timedelta(days=60)).strftime("%Y-%m-%d"),
    "end_date": datetime.now().strftime("%Y-%m-%d"),
    "type": "json", "status": "All", "limit": 2000, "page": 1
}

print("正在查询...")
print(f"日期范围: {body['begin_date']} ~ {body['end_date']}")

try:
    r = requests.post(url, json=body, timeout=120)
    print(f"HTTP 状态码: {r.status_code}")
    data = r.json()
    
    # 处理不同的返回格式
    status = data.get("status", {})
    if isinstance(status, dict):
        code = status.get("code")
        msg = status.get("message", "")
    else:
        code = status
        msg = ""
    
    print(f"API 返回码: {code}")
    
    if code != 0:
        print(f"API 错误: {msg}")
        exit(1)
    
    inner_data = data.get("data")
    if inner_data is None:
        print(f"data 为空，完整返回: {json.dumps(data, ensure_ascii=False)[:500]}")
        exit(1)
    
    trans = inner_data.get("list", []) if isinstance(inner_data, dict) else []
    print(f"共 {len(trans)} 条交易")
except Exception as e:
    print(f"请求异常: {e}")
    import traceback
    traceback.print_exc()
    exit(1)

# 找 Mighty Paw
mp = [t for t in trans if "Mighty" in str(t.get("merchant_name", ""))]
print(f"\nMighty Paw 交易: {len(mp)} 条\n")

for i, t in enumerate(mp, 1):
    print(f"[{i}] ASIN={t.get('prod_id')}, 佣金=${t.get('sale_comm')}")
    print(f"    uid='{t.get('uid','')}', uid2='{t.get('uid2','')}', uid3='{t.get('uid3','')}'")
    print()

