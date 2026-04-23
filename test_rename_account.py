"""测试 Google Ads API - 修改账号名称"""

from google.ads.googleads.client import GoogleAdsClient
from google.oauth2 import service_account

print('=' * 60)
print('测试 Google Ads API - 修改账号名称')
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

# ============ 第一步：列出所有子账户及其当前名称 ============
print('\n【1】获取所有子账户及当前名称')
print('-' * 40)

ga_service = client.get_service('GoogleAdsService')

query = """
    SELECT 
        customer_client.id, 
        customer_client.descriptive_name, 
        customer_client.manager
    FROM customer_client
    WHERE customer_client.manager = FALSE
"""

response = ga_service.search(customer_id=MCC_CUSTOMER_ID, query=query)

accounts = []
for row in response:
    account_id = str(row.customer_client.id)
    name = row.customer_client.descriptive_name
    accounts.append((account_id, name))
    print(f"  账号ID: {account_id}, 当前名称: {name}")

print(f"\n共 {len(accounts)} 个子账户")

# ============ 第二步：测试修改账号名称 ============
print('\n【2】测试修改账号名称')
print('-' * 40)

def rename_account(client, account_id, new_name, mcc_id):
    """
    修改 Google Ads 账号的 descriptive_name
    
    参数:
        client: GoogleAdsClient 实例
        account_id: 要修改的账号ID (不带横杠)
        new_name: 新的账号名称
        mcc_id: MCC 账户 ID (用于 login_customer_id)
    
    返回:
        bool: 是否成功
    """
    try:
        customer_service = client.get_service('CustomerService')
        
        # 创建 Customer 操作
        customer_operation = client.get_type('CustomerOperation')
        customer = customer_operation.update
        
        # 设置资源名称（必须）
        customer.resource_name = f"customers/{account_id}"
        
        # 设置新名称
        customer.descriptive_name = new_name
        
        # 设置需要更新的字段（使用 protobuf 的 FieldMask）
        from google.protobuf import field_mask_pb2
        customer_operation.update_mask = field_mask_pb2.FieldMask(paths=['descriptive_name'])
        
        # 执行更新
        # 注意：需要使用目标账号ID作为 customer_id
        response = customer_service.mutate_customer(
            customer_id=account_id,
            operation=customer_operation
        )
        
        print(f"  ✓ 成功修改账号 {account_id} 的名称为: {new_name}")
        print(f"    返回的资源名: {response.result.resource_name}")
        return True
        
    except Exception as e:
        print(f"  ✗ 修改失败: {e}")
        return False


# 选择一个测试账号（使用列表中的第一个）
if accounts:
    test_account_id, current_name = accounts[0]
    
    print(f"\n选择测试账号: {test_account_id}")
    print(f"当前名称: {current_name}")
    
    # 询问用户是否要测试
    print("\n⚠️ 这是一个实际修改操作！")
    print("如果要测试，请取消下面的注释并设置新名称：")
    print()
    print('# NEW_NAME = "新的账号名称"')
    print('# rename_account(client, test_account_id, NEW_NAME, MCC_CUSTOMER_ID)')
    print()
    
    # ============ 实际测试修改账号名称 ============
    # 测试1: 先改成测试名称
    NEW_NAME = "A-ads159-TEST"
    print(f"\n测试: 将名称从 '{current_name}' 改为 '{NEW_NAME}'")
    success = rename_account(client, test_account_id, NEW_NAME, MCC_CUSTOMER_ID)
    
    if success:
        # 测试2: 再改回原名称
        print(f"\n恢复: 将名称从 '{NEW_NAME}' 改回 '{current_name}'")
        rename_account(client, test_account_id, current_name, MCC_CUSTOMER_ID)
    
else:
    print("没有找到子账户")

print('\n' + '=' * 60)
print('测试完成')
print('=' * 60)

