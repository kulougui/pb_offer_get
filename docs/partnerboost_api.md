# PartnerBoost API 调用文档

## 1. 凭证信息

| 配置项 | 值 | 用途 |
|--------|-----|------|
| API Token (Offer) | `lrGBDS12Bt1nr5nH` | 获取Offer和投放链接 |
| API Token (交易) | `5wXyGERfQ3rEQTdI` | 获取交易/佣金数据 |
| API 基础URL | `https://app.partnerboost.com` | - |

## 2. API列表

| API | 用途 | 状态 |
|-----|------|------|
| 获取Offer列表 | 获取可推广的产品 | ✅ 已测试 |
| 获取投放链接 | 根据ASIN获取推广链接 | ✅ 已测试 |
| 获取交易/佣金 | 查询交易和佣金数据 | ✅ 已测试 |

---

## 3. 获取Offer列表 API

### 3.1 接口信息

| 项目 | 值 |
|------|-----|
| URL | `https://app.partnerboost.com/api/datafeed/get_fba_products` |
| 方法 | POST |
| Content-Type | application/json |

### 3.2 请求参数

```json
{
    "token": "lrGBDS12Bt1nr5nH",
    "page_size": 100,
    "page": 1,
    "default_filter": 0,
    "country_code": "",
    "brand_id": null,
    "sort": "",
    "asins": "",
    "relationship": 1,
    "is_original_currency": 0,
    "has_promo_code": 0,
    "has_acc": 0,
    "filter_sexual_wellness": 0
}
```

| 参数 | 类型 | 说明 |
|------|------|------|
| token | string | API令牌 |
| page_size | int | 每页数量 |
| page | int | 页码 |
| country_code | string | 国家代码 (如: US, DE) |
| brand_id | string | 品牌ID |
| asins | string | ASIN列表，逗号分隔 |
| relationship | int | 关系类型 (0或1) |
| has_promo_code | int | 是否有促销码 (0或1) |
| has_acc | int | 是否有ACC佣金 (0或1) |

### 3.3 响应示例

```json
{
    "status": {
        "code": 0,
        "msg": "success"
    },
    "data": {
        "list": [
            {
                "product_id": "xxx",
                "product_name": "产品名称",
                "image": "图片URL",
                "asin": "B07ZVKTP53",
                "discount": "0%",
                "commission": "10%",
                "category": "分类",
                "availability": "IN_STOCK",
                "rating": "4.6",
                "reviews": "33813",
                "url": "https://www.amazon.de/dp/B07ZVKTP53",
                "brand_name": "Anker",
                "country_code": "DE",
                "original_price": "$44.93",
                "discount_price": "$44.93",
                "currency": "USD"
            }
        ],
        "has_more": true
    }
}
```

---

## 4. 获取投放链接 API

### 4.1 接口信息

| 项目 | 值 |
|------|-----|
| URL | `https://app.partnerboost.com/api/datafeed/get_amazon_link_by_asin` |
| 方法 | POST |
| Content-Type | application/json |

### 4.2 请求参数

```json
{
    "token": "lrGBDS12Bt1nr5nH",
    "asins": "B0FV81ZXTC",
    "country_code": "DE",
    "uid": "",
    "return_partnerboost_link": 1
}
```

| 参数 | 类型 | 说明 |
|------|------|------|
| token | string | API令牌 |
| asins | string | ASIN码 |
| country_code | string | 国家代码 |
| uid | string | 自定义跟踪ID |
| return_partnerboost_link | int | 是否返回短链接 (1=是) |

### 4.3 响应示例

```json
{
    "status": {
        "code": 0,
        "msg": "success"
    },
    "data": [
        {
            "asin": "B0FV81ZXTC",
            "link": "https://www.amazon.de/dp/B0FV81ZXTC?maas=...",
            "partnerboost_link": "https://pboost.me/OWlo6Sd",
            "link_id": "xxx"
        }
    ],
    "error_list": []
}
```

---

## 5. 获取交易/佣金 API

### 5.1 接口信息

| 项目 | 值 |
|------|-----|
| URL | `https://app.partnerboost.com/api.php?mod=medium&op=transaction` |
| 方法 | POST |
| Content-Type | application/json |

### 5.2 两种查询方式对比

API支持两种查询方式：**交易日期查询** 和 **验证日期查询**

| 查询方式 | 参数 | 含义 | 使用场景 |
|----------|------|------|----------|
| **交易日期** | begin_date / end_date | 订单发生的时间（用户购买时间） | 查询某段时间内产生的订单 |
| **验证日期** | validation_date_begin / validation_date_end | 订单被审核确认的时间 | 查询某段时间内被验证的订单 |

**实际测试数据对比** (Token: 5wXyGERfQ3rEQTdI):

| 查询方式 | 总交易数 | 总销售额 | 总佣金 | 状态分布 |
|----------|----------|----------|--------|----------|
| 交易日期 | 226条 | $42,245.76 | $2,952.95 | Approved:205, Pending:17, Rejected:4 |
| 验证日期 | 205条 | $35,487.40 | $2,448.04 | 全部Approved |

**建议**：
- 获取所有历史订单 → 使用**交易日期**（包含所有状态，数据更完整）
- 查询已确认的佣金 → 使用**验证日期**（只返回已验证的订单）

### 5.3 请求参数

**使用交易日期查询：**
```json
{
    "token": "5wXyGERfQ3rEQTdI",
    "begin_date": "2025-11-24",
    "end_date": "2026-01-23",
    "type": "json",
    "status": "All",
    "limit": 2000,
    "page": 1
}
```

**使用验证日期查询：**
```json
{
    "token": "5wXyGERfQ3rEQTdI",
    "validation_date_begin": "2025-11-24",
    "validation_date_end": "2026-01-23",
    "type": "json",
    "status": "All",
    "limit": 2000,
    "page": 1
}
```

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| token | string | 是 | API令牌 |
| begin_date | string | 是* | 交易开始日期 (YYYY-MM-DD) |
| end_date | string | 是* | 交易结束日期 |
| validation_date_begin | string | 是* | 验证开始日期 |
| validation_date_end | string | 是* | 验证结束日期 |
| type | string | 否 | 返回格式 (json/xml) |
| status | string | 否 | 状态筛选 (Pending/Approved/Rejected/Normal/All) |
| limit | int | 否 | 每页数量 (最大2000) |
| page | int | 否 | 页码 |
| brand_id | string | 否 | 品牌ID |
| uid | string | 否 | 自定义跟踪ID |

*注：交易日期或验证日期二选一必填，查询时间跨度不能超过62天

### 5.4 响应示例

```json
{
    "status": {
        "code": 0,
        "msg": "Success"
    },
    "data": {
        "total_page": 1,
        "total_trans": "12",
        "total_items": "12",
        "limit": 2000,
        "list": [
            {
                "mcid": "amzdeanker",
                "merchant_name": "DE-Anker",
                "order_id": "C28T5-68LRM13QIB",
                "order_time": "1768968000",
                "sale_amount": "453.64",
                "sale_comm": "36.29",
                "status": "Pending",
                "prod_id": "B0D1XQ3Z56",
                "order_unit": "1",
                "uid": "",
                "click_ref": "pb_j5blun",
                "partnerboost_id": "770946991a9f3fce3066a4cd5a57ca6b",
                "comm_rate": "8%",
                "validation_date": "null",
                "brand_id": "126776",
                "last_update_time": "01-23-2026",
                "paid_status": 0
            }
        ]
    }
}
```

### 5.5 响应字段说明

| 字段 | 说明 |
|------|------|
| total_page | 总页数 |
| total_trans | 总交易数 |
| total_items | 总条目数 |
| order_id | 订单ID |
| order_time | 交易时间（Unix时间戳，需要转换） |
| sale_amount | 销售金额 |
| sale_comm | 佣金金额 |
| comm_rate | 佣金比例 |
| status | 状态 (Pending/Approved/Rejected) |
| prod_id | **产品ASIN**（可用于关联Offer） |
| uid | 自定义跟踪ID |
| validation_date | 验证日期 |
| paid_status | 是否已付款 (1=已付, 0=未付) |
| merchant_name | 品牌名称 |
| brand_id | 品牌ID |

### 5.6 错误码

| 错误码 | 说明 |
|--------|------|
| 0 | 成功 |
| 1000 | Publisher不存在 |
| 1001 | Token无效 |
| 1002 | 调用频率过高 |
| 1003 | 缺少必需参数或格式错误 |
| 1005 | uid超过200字符 |
| 1006 | 查询时间跨度超过62天 |

---

## 6. 获取所有历史交易数据

由于API限制每次查询最多62天，获取所有历史数据需要分段查询：

```python
from datetime import datetime, timedelta
import requests

def get_all_history_transactions(token, start_from=None):
    """
    获取所有历史交易数据（自动分段查询）
    
    Args:
        token: API Token
        start_from: 起始日期，默认从2024-01-01开始
    
    Returns:
        所有交易记录列表
    """
    url = 'https://app.partnerboost.com/api.php?mod=medium&op=transaction'
    
    if start_from is None:
        start_from = datetime(2024, 1, 1)
    
    all_transactions = []
    current_end = datetime.now()
    
    while current_end > start_from:
        # 每次查询60天
        current_begin = current_end - timedelta(days=60)
        if current_begin < start_from:
            current_begin = start_from
        
        begin_str = current_begin.strftime('%Y-%m-%d')
        end_str = current_end.strftime('%Y-%m-%d')
        
        # 分页获取该时间段的所有数据
        page = 1
        while True:
            body = {
                'token': token,
                'begin_date': begin_str,
                'end_date': end_str,
                'type': 'json',
                'status': 'All',
                'limit': 2000,
                'page': page
            }
            
            response = requests.post(url, json=body)
            data = response.json()
            
            if data.get('status', {}).get('code') != 0:
                break
            
            transactions = data.get('data', {}).get('list', [])
            total_page = int(data.get('data', {}).get('total_page', 0))
            
            all_transactions.extend(transactions)
            
            if page >= total_page or not transactions:
                break
            page += 1
        
        # 移动到下一个时间段
        current_end = current_begin - timedelta(days=1)
    
    return all_transactions


# 使用示例
transactions = get_all_history_transactions('5wXyGERfQ3rEQTdI')
print(f"总交易数: {len(transactions)}")

# 统计
total_sales = sum(float(t.get('sale_amount', 0)) for t in transactions)
total_comm = sum(float(t.get('sale_comm', 0)) for t in transactions)
print(f"总销售额: ${total_sales:,.2f}")
print(f"总佣金: ${total_comm:,.2f}")
```

---

## 7. 实际测试数据统计

使用Token `5wXyGERfQ3rEQTdI` 的历史数据统计：

### 7.1 总体数据

| 指标 | 数值 |
|------|------|
| 总交易数 | 226条 |
| 总销售额 | $42,245.76 |
| 总佣金 | $2,952.95 |
| 数据时间范围 | 2025-07-25 ~ 2026-01-23 |

### 7.2 按状态分布

| 状态 | 数量 | 说明 |
|------|------|------|
| Approved | 205 | 已确认 |
| Pending | 17 | 待审核 |
| Rejected | 4 | 已拒绝 |

### 7.3 按品牌分布 (Top 10)

| 品牌 | 订单数 | 佣金 |
|------|--------|------|
| eufy | 32 | $1,142.29 |
| Anker | 59 | $524.16 |
| DE-Anker | 36 | $509.89 |
| Beatbot Amazon | 8 | $391.56 |
| WOLFBOX | 6 | $54.94 |
| SPORTSROYALS | 2 | $38.98 |
| HONOR IT | 29 | $37.00 |
| LAD-US Comfee | 1 | $31.96 |
| RingConn-US | 2 | $30.04 |
| LAD-Italy Midea | 2 | $27.79 |

---

## 8. 完整Python类

```python
import requests
from datetime import datetime, timedelta


class PartnerBoostAPI:
    def __init__(self, token):
        self.token = token
        self.base_url = "https://app.partnerboost.com"
    
    def get_offers(self, page=1, page_size=100, country_code="", brand_id=None):
        """获取Offer列表"""
        url = f"{self.base_url}/api/datafeed/get_fba_products"
        body = {
            "token": self.token,
            "page_size": page_size,
            "page": page,
            "default_filter": 0,
            "country_code": country_code,
            "brand_id": brand_id,
            "sort": "",
            "asins": "",
            "relationship": 1,
            "is_original_currency": 0,
            "has_promo_code": 0,
            "has_acc": 0,
            "filter_sexual_wellness": 0
        }
        response = requests.post(url, json=body)
        return response.json()
    
    def get_all_offers(self, country_code=""):
        """获取所有Offer（自动翻页）"""
        all_offers = []
        page = 1
        has_more = True
        
        while has_more:
            result = self.get_offers(page=page, country_code=country_code)
            if result.get("status", {}).get("code") == 0:
                offers = result.get("data", {}).get("list", [])
                has_more = result.get("data", {}).get("has_more", False)
                all_offers.extend(offers)
                page += 1
            else:
                break
        
        return all_offers
    
    def get_tracking_link(self, asin, country_code, uid=""):
        """获取投放链接"""
        url = f"{self.base_url}/api/datafeed/get_amazon_link_by_asin"
        body = {
            "token": self.token,
            "asins": asin,
            "country_code": country_code,
            "uid": uid,
            "return_partnerboost_link": 1
        }
        response = requests.post(url, json=body)
        data = response.json()
        
        if data.get("status", {}).get("code") == 0:
            links = data.get("data", [])
            if links:
                return links[0].get("partnerboost_link", "")
        return ""
    
    def get_transactions(self, begin_date, end_date, status="All", page=1, limit=2000, use_validation_date=False):
        """
        获取交易/佣金数据
        
        Args:
            begin_date: 开始日期 (YYYY-MM-DD)
            end_date: 结束日期 (YYYY-MM-DD)
            status: 状态筛选 (All/Pending/Approved/Rejected)
            page: 页码
            limit: 每页数量 (最大2000)
            use_validation_date: 是否使用验证日期查询
        """
        url = f"{self.base_url}/api.php?mod=medium&op=transaction"
        
        if use_validation_date:
            body = {
                "token": self.token,
                "validation_date_begin": begin_date,
                "validation_date_end": end_date,
                "type": "json",
                "status": status,
                "limit": limit,
                "page": page
            }
        else:
            body = {
                "token": self.token,
                "begin_date": begin_date,
                "end_date": end_date,
                "type": "json",
                "status": status,
                "limit": limit,
                "page": page
            }
        
        response = requests.post(url, json=body)
        return response.json()
    
    def get_all_transactions(self, start_from=None, use_validation_date=False):
        """
        获取所有历史交易数据（自动分段+分页）
        
        Args:
            start_from: 起始日期 (datetime对象)，默认2024-01-01
            use_validation_date: 是否使用验证日期查询
        
        Returns:
            所有交易记录列表
        """
        if start_from is None:
            start_from = datetime(2024, 1, 1)
        
        all_transactions = []
        current_end = datetime.now()
        
        while current_end > start_from:
            current_begin = current_end - timedelta(days=60)
            if current_begin < start_from:
                current_begin = start_from
            
            begin_str = current_begin.strftime('%Y-%m-%d')
            end_str = current_end.strftime('%Y-%m-%d')
            
            # 分页获取
            page = 1
            while True:
                result = self.get_transactions(
                    begin_str, end_str, 
                    page=page, 
                    use_validation_date=use_validation_date
                )
                
                if result.get("status", {}).get("code") != 0:
                    break
                
                data = result.get("data", {})
                transactions = data.get("list", [])
                total_page = int(data.get("total_page", 0))
                
                all_transactions.extend(transactions)
                
                if page >= total_page or not transactions:
                    break
                page += 1
            
            current_end = current_begin - timedelta(days=1)
        
        return all_transactions
    
    def convert_order_time(self, timestamp):
        """将Unix时间戳转换为日期字符串"""
        try:
            return datetime.fromtimestamp(int(timestamp)).strftime('%Y-%m-%d %H:%M:%S')
        except:
            return timestamp


# 使用示例
if __name__ == "__main__":
    # Offer API
    offer_api = PartnerBoostAPI(token="lrGBDS12Bt1nr5nH")
    offers = offer_api.get_offers(page_size=10)
    print(f"获取到 {len(offers.get('data', {}).get('list', []))} 个offer")
    
    # 交易API
    trans_api = PartnerBoostAPI(token="5wXyGERfQ3rEQTdI")
    transactions = trans_api.get_all_transactions()
    print(f"总交易数: {len(transactions)}")
    
    # 统计
    total_comm = sum(float(t.get('sale_comm', 0)) for t in transactions)
    print(f"总佣金: ${total_comm:,.2f}")
```

---

## 9. 注意事项

1. **时间跨度限制**: 交易查询API最多支持62天的时间跨度
2. **分页限制**: 交易API每页最多返回2000条
3. **调用频率**: 避免频繁调用，否则可能返回错误码1002
4. **Token区分**: Offer API和交易API可能使用不同的Token
5. **时间戳转换**: order_time是Unix时间戳，需要转换为日期
6. **prod_id就是ASIN**: 可用于关联Offer数据

---

*文档创建时间：2026-01-23*
*最后更新：2026-01-23*
