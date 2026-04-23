"""获取所有历史交易数据"""
import requests
import json
from datetime import datetime, timedelta

print('=' * 60)
print('获取所有历史交易数据')
print('=' * 60)

url = 'https://app.partnerboost.com/api.php?mod=medium&op=transaction'
token = '5wXyGERfQ3rEQTdI'

def get_transactions(begin_date, end_date, use_validation_date=False):
    """获取指定日期范围的交易"""
    if use_validation_date:
        body = {
            'token': token,
            'validation_date_begin': begin_date,
            'validation_date_end': end_date,
            'type': 'json',
            'status': 'All',
            'limit': 2000,
            'page': 1
        }
    else:
        body = {
            'token': token,
            'begin_date': begin_date,
            'end_date': end_date,
            'type': 'json',
            'status': 'All',
            'limit': 2000,
            'page': 1
        }
    
    all_transactions = []
    page = 1
    
    while True:
        body['page'] = page
        r = requests.post(url, json=body)
        data = r.json()
        
        if data.get('status', {}).get('code') != 0:
            print(f"  错误: {data.get('status')}")
            break
        
        transactions = data.get('data', {}).get('list', [])
        total_page = int(data.get('data', {}).get('total_page', 0))
        
        all_transactions.extend(transactions)
        
        if page >= total_page or not transactions:
            break
        page += 1
    
    return all_transactions

def get_all_history(use_validation_date=False):
    """分段获取所有历史数据（每次60天）"""
    all_transactions = []
    end_date = datetime.now()
    
    # 假设从2024年1月1日开始（可以根据实际调整）
    start_from = datetime(2024, 1, 1)
    
    print(f"\n查询方式: {'验证日期' if use_validation_date else '交易日期'}")
    print(f"查询范围: {start_from.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")
    print('-' * 40)
    
    current_end = end_date
    
    while current_end > start_from:
        current_begin = current_end - timedelta(days=60)
        if current_begin < start_from:
            current_begin = start_from
        
        begin_str = current_begin.strftime('%Y-%m-%d')
        end_str = current_end.strftime('%Y-%m-%d')
        
        print(f"  查询 {begin_str} 至 {end_str}...", end=' ')
        
        transactions = get_transactions(begin_str, end_str, use_validation_date)
        print(f"获取 {len(transactions)} 条")
        
        all_transactions.extend(transactions)
        
        # 移动到下一个时间段
        current_end = current_begin - timedelta(days=1)
    
    return all_transactions

# 使用交易日期获取所有数据
print("\n" + "=" * 60)
print("【方法1】使用交易日期查询")
print("=" * 60)
trans_by_order = get_all_history(use_validation_date=False)
print(f"\n交易日期查询总计: {len(trans_by_order)} 条")

# 使用验证日期获取所有数据
print("\n" + "=" * 60)
print("【方法2】使用验证日期查询")
print("=" * 60)
trans_by_validation = get_all_history(use_validation_date=True)
print(f"\n验证日期查询总计: {len(trans_by_validation)} 条")

# 统计汇总
print("\n" + "=" * 60)
print("汇总统计")
print("=" * 60)

def summarize(transactions, name):
    if not transactions:
        print(f"\n{name}: 无数据")
        return
    
    total_sales = sum(float(t.get('sale_amount', 0)) for t in transactions)
    total_comm = sum(float(t.get('sale_comm', 0)) for t in transactions)
    
    # 按状态统计
    status_count = {}
    for t in transactions:
        status = t.get('status', 'Unknown')
        status_count[status] = status_count.get(status, 0) + 1
    
    # 按品牌统计
    brand_count = {}
    brand_comm = {}
    for t in transactions:
        brand = t.get('merchant_name', 'Unknown')
        brand_count[brand] = brand_count.get(brand, 0) + 1
        brand_comm[brand] = brand_comm.get(brand, 0) + float(t.get('sale_comm', 0))
    
    print(f"\n{name}:")
    print(f"  总交易数: {len(transactions)}")
    print(f"  总销售额: ${total_sales:,.2f}")
    print(f"  总佣金: ${total_comm:,.2f}")
    print(f"\n  按状态分布:")
    for status, count in sorted(status_count.items()):
        print(f"    {status}: {count}")
    print(f"\n  按品牌分布 (Top 10):")
    sorted_brands = sorted(brand_comm.items(), key=lambda x: x[1], reverse=True)[:10]
    for brand, comm in sorted_brands:
        print(f"    {brand}: {brand_count[brand]}单, 佣金${comm:,.2f}")

summarize(trans_by_order, "交易日期查询结果")
summarize(trans_by_validation, "验证日期查询结果")

# 建议
print("\n" + "=" * 60)
print("建议")
print("=" * 60)
print("""
1. 如果想知道"某段时间内产生了多少订单" → 用【交易日期】
2. 如果想知道"某段时间内确认了多少佣金" → 用【验证日期】
3. 获取所有历史订单，建议使用【交易日期】，因为它反映订单的实际发生时间
4. 验证日期查询到的数量更多，因为包含了之前产生但最近才验证的订单
""")

