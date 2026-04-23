# 飞书API调用文档

## 1. 应用凭证

| 配置项 | 值 |
|--------|-----|
| App ID | `cli_a8517363cf3bd013` |
| App Secret | `O4Sm3UNHjpykF9OZq3LroblsrCVYyQEp` |

## 2. 电子表格信息

| 配置项 | 值 |
|--------|-----|
| 电子表格Token | `KnJ1wphpBiVMrGkWl5ncUkMGnfe` |
| 电子表格标题 | Mediabuy项目管理 |
| 电子表格URL | https://my.feishu.cn/wiki/KnJ1wphpBiVMrGkWl5ncUkMGnfe |
| Offer工作表ID | `kPlW5z` |
| Offer工作表名称 | offer |
| 总行数 | 115 |
| 总列数 | 39 |

## 3. API调用流程

### 3.1 获取 tenant_access_token

**请求：**
```
POST https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal
Content-Type: application/json

{
    "app_id": "cli_a8517363cf3bd013",
    "app_secret": "O4Sm3UNHjpykF9OZq3LroblsrCVYyQEp"
}
```

**响应：**
```json
{
    "code": 0,
    "expire": 6890,
    "msg": "ok",
    "tenant_access_token": "t-g1041nfoRSPMJWNFFGRU7IE5BVKSUNFB4XX26MJT"
}
```

**注意事项：**
- token有效期约2小时（expire字段，单位秒）
- 需要在过期前刷新

### 3.2 获取电子表格信息

**请求：**
```
GET https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}
Authorization: Bearer {tenant_access_token}
```

**示例：**
```
GET https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/KnJ1wphpBiVMrGkWl5ncUkMGnfe
```

**响应：**
```json
{
    "code": 0,
    "data": {
        "spreadsheet": {
            "owner_id": "ou_3d80099c478de109665aa1cb07c84cbc",
            "title": "Mediabuy项目管理",
            "token": "KnJ1wphpBiVMrGkWl5ncUkMGnfe",
            "url": "https://my.feishu.cn/wiki/KnJ1wphpBiVMrGkWl5ncUkMGnfe"
        }
    },
    "msg": ""
}
```

### 3.3 获取工作表元数据

**请求：**
```
GET https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}
Authorization: Bearer {tenant_access_token}
```

**示例：**
```
GET https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/KnJ1wphpBiVMrGkWl5ncUkMGnfe/sheets/kPlW5z
```

**响应：**
```json
{
    "code": 0,
    "data": {
        "sheet": {
            "grid_properties": {
                "column_count": 39,
                "frozen_column_count": 0,
                "frozen_row_count": 0,
                "row_count": 115
            },
            "hidden": false,
            "index": 0,
            "resource_type": "sheet",
            "sheet_id": "kPlW5z",
            "title": "offer"
        }
    },
    "msg": "success"
}
```

### 3.4 读取工作表数据

**请求：**
```
GET https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range}
Authorization: Bearer {tenant_access_token}
```

**range格式：** `{sheet_id}!{起始单元格}:{结束单元格}`

**示例（读取前20行，A到Z列）：**
```
GET https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/KnJ1wphpBiVMrGkWl5ncUkMGnfe/values/kPlW5z!A1:Z20
```

**响应：**
```json
{
    "code": 0,
    "data": {
        "revision": 25230,
        "spreadsheetToken": "KnJ1wphpBiVMrGkWl5ncUkMGnfe",
        "valueRange": {
            "majorDimension": "ROWS",
            "range": "kPlW5z!A1:Z20",
            "revision": 25230,
            "values": [
                ["状态", "投放中的ads", "广告系列数量", ...],
                ["未测试", null, null, ...]
            ]
        }
    },
    "msg": "success"
}
```

## 4. Offer工作表结构

### 4.1 列定义

| 列号 | 列字母 | 列名 | 数据类型 | 说明 |
|------|--------|------|----------|------|
| 1 | A | 状态 | 文本 | 如：未测试、测试中、已上线 |
| 2 | B | 投放中的ads | 文本 | 广告相关信息 |
| 3 | C | 广告系列数量 | 数字 | Google Ads广告系列数量 |
| 4 | D | 广告系列总花费 | 数字 | 总花费金额 |
| 5 | E | 总佣金 | 数字 | 总佣金金额 |
| 6 | F | 投放链接 | URL对象 | pboost.me短链接 |
| 7 | G | 国家代码 | 文本 | 如：DE, US |
| 8 | H | 品牌名称 | 文本 | 如：Anker |
| 9 | I | 折扣价 | 数字 | 折扣后价格 |
| 10 | J | 佣金 | 数字 | 佣金比例（如0.08表示8%） |
| 11 | K | 产品名称 | 文本 | 产品标题 |
| 12 | L | ASIN | 文本 | 亚马逊ASIN码 |
| 13 | M | 库存状态 | 文本 | IN_STOCK / OUT_OF_STOCK |
| 14 | N | 产品链接 | URL对象 | 亚马逊产品链接 |
| 15 | O | 更新时间 | 数字 | 日期序列号 |
| 16 | P | 原价 | 数字 | 原始价格 |
| 17 | Q | 货币 | 文本 | USD |

### 4.2 特殊数据格式

**URL对象格式（投放链接、产品链接）：**
```json
[
    {
        "cellPosition": null,
        "link": "https://pboost.me/wWlM5QZ",
        "text": "https://pboost.me/wWlM5QZ",
        "type": "url"
    }
]
```

读取时需要从数组中提取 `link` 字段。

## 5. Python代码示例

```python
import requests

class FeishuAPI:
    def __init__(self, app_id, app_secret):
        self.app_id = app_id
        self.app_secret = app_secret
        self.base_url = "https://open.feishu.cn/open-apis"
        self.token = None
    
    def get_tenant_access_token(self):
        """获取 tenant_access_token"""
        url = f"{self.base_url}/auth/v3/tenant_access_token/internal"
        body = {
            "app_id": self.app_id,
            "app_secret": self.app_secret
        }
        response = requests.post(url, json=body)
        data = response.json()
        if data.get("code") == 0:
            self.token = data.get("tenant_access_token")
            return self.token
        return None
    
    def get_headers(self):
        """获取请求头"""
        return {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
    
    def get_spreadsheet_info(self, spreadsheet_token):
        """获取电子表格信息"""
        url = f"{self.base_url}/sheets/v3/spreadsheets/{spreadsheet_token}"
        response = requests.get(url, headers=self.get_headers())
        return response.json()
    
    def get_sheet_metadata(self, spreadsheet_token, sheet_id):
        """获取工作表元数据"""
        url = f"{self.base_url}/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{sheet_id}"
        response = requests.get(url, headers=self.get_headers())
        return response.json()
    
    def read_sheet_data(self, spreadsheet_token, range_str):
        """读取工作表数据"""
        url = f"{self.base_url}/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range_str}"
        response = requests.get(url, headers=self.get_headers())
        return response.json()
    
    def extract_url_from_cell(self, cell_value):
        """从单元格中提取URL"""
        if isinstance(cell_value, list) and len(cell_value) > 0:
            if isinstance(cell_value[0], dict) and cell_value[0].get("type") == "url":
                return cell_value[0].get("link", "")
        return cell_value if cell_value else ""


# 使用示例
if __name__ == "__main__":
    api = FeishuAPI(
        app_id="cli_a8517363cf3bd013",
        app_secret="O4Sm3UNHjpykF9OZq3LroblsrCVYyQEp"
    )
    
    # 获取token
    api.get_tenant_access_token()
    
    # 读取offer表格数据
    spreadsheet_token = "KnJ1wphpBiVMrGkWl5ncUkMGnfe"
    sheet_id = "kPlW5z"
    range_str = f"{sheet_id}!A1:Q115"
    
    data = api.read_sheet_data(spreadsheet_token, range_str)
    values = data.get("data", {}).get("valueRange", {}).get("values", [])
    
    # 第一行是表头
    headers = values[0] if values else []
    
    # 遍历数据行
    for row in values[1:]:
        status = row[0] if len(row) > 0 else ""
        tracking_link = api.extract_url_from_cell(row[5]) if len(row) > 5 else ""
        country_code = row[6] if len(row) > 6 else ""
        brand_name = row[7] if len(row) > 7 else ""
        asin = row[11] if len(row) > 11 else ""
        print(f"{asin} | {brand_name} | {country_code} | {status}")
```

## 6. 待实现功能

- [ ] 写入/更新工作表数据
- [ ] 批量更新多行数据
- [ ] 与Google MCC API集成，回填广告数据

---

*文档创建时间：2026-01-23*
*最后更新：2026-01-23*

