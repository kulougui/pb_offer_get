"""测试 PartnerBoost 交易/佣金 API"""

import requests
import json
from datetime import datetime, timedelta

print('=' * 60)
print('测试 PartnerBoost 交易/佣金 API')
print('=' * 60)

url = 'https://app.partnerboost.com/api.php?mod=medium&op=transaction'
token = 'lrGBDS12Bt1nr5nH'

# 查询最近60天的交易 (API限制最多62天)
end_date = datetime.now().strftime('%Y-%m-%d')
begin_date = (datetime.now() - timedelta(days=60)).strftime('%Y-%m-%d')

print(f'查询日期范围: {begin_date} 至 {end_date}')

body = {
    'token': token,
    'begin_date': begin_date,
    'end_date': end_date,
    'type': 'json',
    'status': 'All',
    'limit': 20,
    'page': 1
}

print('\n请求参数:')
print(json.dumps(body, indent=2))

print('\n正在调用API...')
response = requests.post(url, json=body, headers={'Content-Type': 'application/json'}, timeout=30)

print(f'状态码: {response.status_code}')
data = response.json()

if data.get('status', {}).get('code') == 0:
    result = data.get('data', {})
    print(f'\n总页数: {result.get("total_page", 0)}')
    print(f'总交易数: {result.get("total_trans", 0)}')
    print(f'总条目数: {result.get("total_items", 0)}')
    
    transactions = result.get('list', [])
    if transactions:
        print(f'\n获取到 {len(transactions)} 条交易记录:')
        print('-' * 60)
        for t in transactions[:5]:  # 只显示前5条
            print(f"订单ID: {t.get('order_id')}")
            print(f"  品牌: {t.get('merchant_name')}")
            print(f"  交易时间: {t.get('order_time')}")
            print(f"  销售金额: {t.get('sale_amount')}")
            print(f"  佣金: {t.get('sale_comm')}")
            print(f"  佣金率: {t.get('comm_rate')}")
            print(f"  状态: {t.get('status')}")
            print(f"  UID: {t.get('uid')}")
            print(f"  产品ID: {t.get('prod_id')}")
            print(f"  验证日期: {t.get('validation_date')}")
            print()
        
        if len(transactions) > 5:
            print(f'... 还有 {len(transactions) - 5} 条记录')
    else:
        print('\n该时间范围内没有交易记录')
        
    # 打印完整的第一条记录结构（如果有）
    if transactions:
        print('\n' + '=' * 60)
        print('第一条交易记录的完整字段:')
        print('=' * 60)
        print(json.dumps(transactions[0], indent=2, ensure_ascii=False))
else:
    print(f'\n错误响应:')
    print(json.dumps(data, indent=2, ensure_ascii=False))

