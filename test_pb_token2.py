"""测试新Token的交易/佣金API"""
import requests
import json
from datetime import datetime, timedelta

print('=' * 60)
print('测试 Token: 5wXyGERfQ3rEQTdI')
print('=' * 60)

url = 'https://app.partnerboost.com/api.php?mod=medium&op=transaction'
token = '5wXyGERfQ3rEQTdI'

# 测试不同的日期范围
end_date = datetime.now().strftime('%Y-%m-%d')
begin_date = (datetime.now() - timedelta(days=60)).strftime('%Y-%m-%d')

print(f'\n【测试1】交易日期查询: {begin_date} 至 {end_date}')
body1 = {
    'token': token,
    'begin_date': begin_date,
    'end_date': end_date,
    'type': 'json',
    'status': 'All',
    'limit': 50,
    'page': 1
}
r1 = requests.post(url, json=body1)
data1 = r1.json()
print(f"状态: {data1.get('status')}")
print(f"总交易数: {data1.get('data', {}).get('total_trans', 0)}")
print(f"总条目: {data1.get('data', {}).get('total_items', 0)}")

# 测试验证日期查询
print(f'\n【测试2】验证日期查询: {begin_date} 至 {end_date}')
body2 = {
    'token': token,
    'validation_date_begin': begin_date,
    'validation_date_end': end_date,
    'type': 'json',
    'status': 'All',
    'limit': 50,
    'page': 1
}
r2 = requests.post(url, json=body2)
data2 = r2.json()
print(f"状态: {data2.get('status')}")
print(f"总交易数: {data2.get('data', {}).get('total_trans', 0)}")
print(f"总条目: {data2.get('data', {}).get('total_items', 0)}")

# 如果有数据，显示详情
transactions = data1.get('data', {}).get('list', []) or data2.get('data', {}).get('list', [])
if transactions:
    print(f'\n获取到 {len(transactions)} 条交易记录:')
    print('-' * 60)
    for t in transactions[:10]:
        print(f"订单ID: {t.get('order_id')}")
        print(f"  品牌: {t.get('merchant_name')}")
        print(f"  交易时间: {t.get('order_time')}")
        print(f"  销售金额: {t.get('sale_amount')}")
        print(f"  佣金: {t.get('sale_comm')}")
        print(f"  状态: {t.get('status')}")
        print(f"  UID: {t.get('uid')}")
        print()
    
    # 打印第一条完整记录
    print('\n完整记录结构:')
    print(json.dumps(transactions[0], indent=2, ensure_ascii=False))
else:
    print('\n该账号在此时间范围内没有交易记录')

print('\n' + '=' * 60)

