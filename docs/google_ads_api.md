# Google Ads API 调用文档

## 1. 凭证信息

| 配置项 | 值 |
|--------|-----|
| Developer Token | `1YsRjWGxV6XUdxtX8MiT3Q` |
| MCC Customer ID | `6885177935` (原格式: 688-517-7935) |
| 服务账号邮箱 | `mcc1service@pure-vehicle-469813-t3.iam.gserviceaccount.com` |
| 服务账号项目 | `pure-vehicle-469813-t3` |
| 服务账号密钥文件 | `credentials/google_ads_service_account.json` |

## 2. MCC账户结构

- **MCC名称**: mcc10
- **MCC ID**: 6885177935
- **子账户数量**: 44个（43个广告账户 + 1个MCC账户）
- **货币类型**: CNY (21个), USD (22个)
- **时区**: 主要是 Asia/Shanghai

## 3. API调用流程

### 3.1 创建客户端

```python
from google.ads.googleads.client import GoogleAdsClient
from google.oauth2 import service_account

DEVELOPER_TOKEN = '1YsRjWGxV6XUdxtX8MiT3Q'
MCC_CUSTOMER_ID = '6885177935'
SERVICE_ACCOUNT_FILE = 'credentials/google_ads_service_account.json'

# 创建凭证
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=['https://www.googleapis.com/auth/adwords']
)

# 创建客户端
client = GoogleAdsClient(
    credentials=credentials,
    developer_token=DEVELOPER_TOKEN,
    login_customer_id=MCC_CUSTOMER_ID,
    use_proto_plus=True
)
```

### 3.2 获取可访问的客户列表

```python
customer_service = client.get_service('CustomerService')
accessible_customers = customer_service.list_accessible_customers()

for resource_name in accessible_customers.resource_names:
    print(resource_name)  # 格式: customers/1234567890
```

### 3.3 获取MCC下的子账户

```python
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

response = ga_service.search(customer_id=MCC_CUSTOMER_ID, query=query)

for row in response:
    cc = row.customer_client
    print(f'ID: {cc.id}, 名称: {cc.descriptive_name}, 是否经理: {cc.manager}')
```

### 3.4 获取广告系列信息

```python
# 查询某个子账户的广告系列
customer_id = '1398464735'  # 子账户ID

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

response = ga_service.search(customer_id=customer_id, query=campaign_query)

for row in response:
    c = row.campaign
    m = row.metrics
    cost = m.cost_micros / 1000000  # 转换为美元
    print(f'{c.name}: 展示={m.impressions}, 点击={m.clicks}, 花费=${cost:.2f}')
```

### 3.5 获取指定日期范围的数据

```python
# 获取最近7天的数据
campaign_query = """
    SELECT
        campaign.id,
        campaign.name,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions
    FROM campaign
    WHERE campaign.status != 'REMOVED'
        AND segments.date DURING LAST_7_DAYS
"""

# 获取指定日期范围
campaign_query = """
    SELECT
        campaign.id,
        campaign.name,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros
    FROM campaign
    WHERE campaign.status != 'REMOVED'
        AND segments.date >= '2026-01-01'
        AND segments.date <= '2026-01-23'
"""
```

## 4. 子账户列表

### CNY账户 (21个)

| 账户ID | 名称 | 时区 |
|--------|------|------|
| 9754035027 | ads203 | Asia/Shanghai |
| 1292780744 | ads207 | Asia/Shanghai |
| 5420826805 | ads208 | Asia/Shanghai |
| 9279338766 | ads209 | Asia/Shanghai |
| 2432878496 | ads210 | Asia/Shanghai |
| 4137945297 | ads211 | Asia/Shanghai |
| 3807264093 | ads213 | Asia/Shanghai |
| 4410825286 | ads215 | Asia/Shanghai |
| 4358154496 | ads217 | Asia/Shanghai |
| 7622971758 | ads218 | Asia/Shanghai |
| 7222980164 | ads221 | Asia/Shanghai |
| 5064539813 | ads222 | Asia/Shanghai |
| 8107901442 | ads223 | Asia/Shanghai |
| 4407185284 | ads224 | Asia/Shanghai |
| 6569045288 | ads225 | Asia/Shanghai |
| 6550744092 | ads226 | Asia/Shanghai |
| 8699186497 | ads227 | Asia/Shanghai |
| 5879502896 | ads228 | Asia/Shanghai |
| 1544895039 | ads229 | Asia/Shanghai |
| 3700485048 | ads230 | Asia/Shanghai |
| 2998450899 | ads231 | Asia/Shanghai |

### USD账户 (22个)

| 账户ID | 名称 | 时区 |
|--------|------|------|
| 1398464735 | ads159 | Asia/Shanghai |
| 3731206198 | ads189 | Asia/Shanghai |
| 1061177746 | ads191 | America/Denver |
| 4971052593 | ads193 | Asia/Shanghai |
| 9877149156 | ads195 | Asia/Shanghai |
| 6062338015 | ads196 | Asia/Shanghai |
| 1163448635 | ads197 | Asia/Shanghai |
| 9133198723 | ads200 | Asia/Shanghai |
| 6012885444 | ads202 | Asia/Shanghai |
| 7783739900 | ads204 | Asia/Shanghai |
| 8266106729 | ads232 | Asia/Shanghai |
| 8627943090 | ads233 | Asia/Shanghai |
| 6877016435 | ads235 | Asia/Shanghai |
| 4451338143 | ads236 | Asia/Shanghai |
| 7860626288 | ads237 | Asia/Shanghai |
| 2015973042 | ads238 | Asia/Shanghai |
| 4173424190 | ads239 | Asia/Shanghai |
| 7176175548 | ads240 | Asia/Shanghai |
| 1457338297 | ads241 | Asia/Shanghai |
| 6200169125 | ads242 | Asia/Shanghai |
| 4036466897 | ads243 | Asia/Shanghai |
| 3214125458 | ads244 | Asia/Shanghai |

## 5. 完整Python类示例

```python
from google.ads.googleads.client import GoogleAdsClient
from google.oauth2 import service_account


class GoogleAdsAPI:
    def __init__(self, developer_token, mcc_customer_id, service_account_file):
        self.developer_token = developer_token
        self.mcc_customer_id = mcc_customer_id
        
        credentials = service_account.Credentials.from_service_account_file(
            service_account_file,
            scopes=['https://www.googleapis.com/auth/adwords']
        )
        
        self.client = GoogleAdsClient(
            credentials=credentials,
            developer_token=developer_token,
            login_customer_id=mcc_customer_id,
            use_proto_plus=True
        )
        
        self.ga_service = self.client.get_service('GoogleAdsService')
    
    def get_sub_accounts(self):
        """获取MCC下的所有子账户"""
        query = """
            SELECT
                customer_client.id,
                customer_client.descriptive_name,
                customer_client.manager,
                customer_client.currency_code,
                customer_client.time_zone
            FROM customer_client
            WHERE customer_client.level <= 1
        """
        
        response = self.ga_service.search(
            customer_id=self.mcc_customer_id, 
            query=query
        )
        
        accounts = []
        for row in response:
            cc = row.customer_client
            if not cc.manager:  # 排除经理账户
                accounts.append({
                    'id': str(cc.id),
                    'name': cc.descriptive_name,
                    'currency': cc.currency_code,
                    'timezone': cc.time_zone
                })
        return accounts
    
    def get_campaigns(self, customer_id):
        """获取指定账户的广告系列"""
        query = """
            SELECT
                campaign.id,
                campaign.name,
                campaign.status,
                campaign.advertising_channel_type,
                metrics.impressions,
                metrics.clicks,
                metrics.cost_micros,
                metrics.conversions
            FROM campaign
            WHERE campaign.status != 'REMOVED'
        """
        
        response = self.ga_service.search(
            customer_id=customer_id, 
            query=query
        )
        
        campaigns = []
        for row in response:
            c = row.campaign
            m = row.metrics
            campaigns.append({
                'id': str(c.id),
                'name': c.name,
                'status': c.status.name,
                'type': c.advertising_channel_type.name,
                'impressions': m.impressions,
                'clicks': m.clicks,
                'cost': m.cost_micros / 1000000,
                'conversions': m.conversions
            })
        return campaigns
    
    def get_all_campaigns(self):
        """获取所有子账户的广告系列"""
        accounts = self.get_sub_accounts()
        all_campaigns = []
        
        for account in accounts:
            campaigns = self.get_campaigns(account['id'])
            for campaign in campaigns:
                campaign['account_id'] = account['id']
                campaign['account_name'] = account['name']
                all_campaigns.append(campaign)
        
        return all_campaigns


# 使用示例
if __name__ == "__main__":
    api = GoogleAdsAPI(
        developer_token='1YsRjWGxV6XUdxtX8MiT3Q',
        mcc_customer_id='6885177935',
        service_account_file='credentials/google_ads_service_account.json'
    )
    
    # 获取所有子账户
    accounts = api.get_sub_accounts()
    print(f'子账户数量: {len(accounts)}')
    
    # 获取所有广告系列
    campaigns = api.get_all_campaigns()
    for c in campaigns:
        print(f"{c['account_name']}: {c['name']} - ${c['cost']:.2f}")
```

## 6. 常用GAQL查询

### 按日期获取广告系列数据
```sql
SELECT
    segments.date,
    campaign.id,
    campaign.name,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros
FROM campaign
WHERE segments.date DURING LAST_30_DAYS
ORDER BY segments.date DESC
```

### 获取广告组信息
```sql
SELECT
    ad_group.id,
    ad_group.name,
    ad_group.status,
    campaign.name,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros
FROM ad_group
WHERE campaign.status != 'REMOVED'
    AND ad_group.status != 'REMOVED'
```

### 获取关键词信息
```sql
SELECT
    ad_group_criterion.keyword.text,
    ad_group_criterion.keyword.match_type,
    metrics.impressions,
    metrics.clicks,
    metrics.cost_micros
FROM keyword_view
WHERE segments.date DURING LAST_7_DAYS
```

## 7. 注意事项

1. **费用单位**: `cost_micros` 的单位是微元(1/1000000)，需要除以1000000转换为美元
2. **Customer ID格式**: 不带横线的纯数字格式
3. **login_customer_id**: 使用MCC账户ID作为登录ID，可以访问其下所有子账户
4. **API限制**: 注意Google Ads API的配额限制

## 8. 待实现功能

- [ ] 根据广告系列名称匹配投放链接
- [ ] 集成到GUI界面
- [ ] 自动回填飞书表格中的广告数据

---

*文档创建时间：2026-01-23*
*最后更新：2026-01-23*

