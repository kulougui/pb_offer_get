"""测试 Google Ads API"""

from google.ads.googleads.client import GoogleAdsClient
from google.oauth2 import service_account

print('=' * 60)
print('获取 MCC 下的子账户和广告系列信息')
print('=' * 60)

DEVELOPER_TOKEN = '1YsRjWGxV6XUdxtX8MiT3Q'
MCC_CUSTOMER_ID = '6885177935'
SERVICE_ACCOUNT_FILE = 'credentials/google_ads_service_account.json'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=['https://www.googleapis.com/auth/adwords']
)

client = GoogleAdsClient(
    credentials=credentials,
    developer_token=DEVELOPER_TOKEN,
    login_customer_id=MCC_CUSTOMER_ID,
    use_proto_plus=True
)

print('\n【1】获取MCC下的子账户列表')
print('-' * 40)

ga_service = client.get_service('GoogleAdsService')

query = """
    SELECT
        customer_client.client_customer,
        customer_client.level,
        customer_client.manager,
        customer_client.descriptive_name,
        customer_client.currency_code,
        customer_client.time_zone,
        customer_client.id
    FROM customer_client
    WHERE customer_client.level <= 1
"""

try:
    response = ga_service.search(customer_id=MCC_CUSTOMER_ID, query=query)
    
    sub_accounts = []
    for row in response:
        cc = row.customer_client
        account_info = {
            'id': cc.id,
            'name': cc.descriptive_name,
            'is_manager': cc.manager,
            'currency': cc.currency_code,
            'timezone': cc.time_zone
        }
        sub_accounts.append(account_info)
        print(f'账户ID: {cc.id}')
        print(f'  名称: {cc.descriptive_name}')
        print(f'  是否为经理账户: {cc.manager}')
        print(f'  货币: {cc.currency_code}')
        print(f'  时区: {cc.time_zone}')
        print()
    
    print(f'共找到 {len(sub_accounts)} 个账户')
    
except Exception as e:
    print(f'获取子账户错误: {e}')
    sub_accounts = []

print('\n【2】获取广告系列信息')
print('-' * 40)

for account in sub_accounts:
    if not account['is_manager']:
        customer_id = str(account['id'])
        print(f"\n账户 {account['name']} ({customer_id}) 的广告系列:")
        
        campaign_query = """
            SELECT
                campaign.id,
                campaign.name,
                campaign.status,
                campaign.advertising_channel_type,
                metrics.impressions,
                metrics.clicks,
                metrics.cost_micros
            FROM campaign
            WHERE campaign.status != 'REMOVED'
            ORDER BY campaign.id
            LIMIT 20
        """
        
        try:
            campaign_response = ga_service.search(customer_id=customer_id, query=campaign_query)
            
            campaign_count = 0
            for row in campaign_response:
                campaign_count += 1
                c = row.campaign
                m = row.metrics
                cost = m.cost_micros / 1000000 if m.cost_micros else 0
                print(f'  - {c.name}')
                print(f'    ID: {c.id}, 状态: {c.status.name}, 类型: {c.advertising_channel_type.name}')
                print(f'    展示: {m.impressions}, 点击: {m.clicks}, 花费: ${cost:.2f}')
            
            if campaign_count == 0:
                print('  (无广告系列)')
            else:
                print(f'  共 {campaign_count} 个广告系列')
                
        except Exception as e:
            print(f'  获取广告系列错误: {str(e)[:300]}')

print('\n' + '=' * 60)
print('测试完成')
print('=' * 60)

