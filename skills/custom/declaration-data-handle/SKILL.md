---
name: declaration-data-handle
description: |
  触发场景：
  1. 用户发送包含报关资料的压缩包时
  2. 用户要求处理之前收到的压缩包时

  功能：解压报关资料压缩包，使用 markitdown 将文件转为 Markdown 格式，通过 Agent 分析提取数据，生成报关单 Excel 文件。
---

# 报关单生成器

## 完整流程

### 步骤 1: 复制文件到工作区

将报关资料压缩包移动到 `/mnt/user-data/workspace` 目录：
```bash
mv {原文件路径} /mnt/user-data/workspace/
```

### 步骤 2: 解压缩文件

使用系统命令解压缩文件到 `/tmp/declaration_extract` 目录：
- `.zip`: `unzip /mnt/user-data/workspace/{文件名} -d /tmp/declaration_extract`
- `.rar`: `unrar x /mnt/user-data/workspace/{文件名} /tmp/declaration_extract/`
- `.7z`: `7z x /mnt/user-data/workspace/{文件名} -o/tmp/declaration_extract`

### 步骤 3: Agent 分析每个文件

**逐个读取 Markdown 文件，让 Agent 按照以下 JSON Schema 分析：**

```json
{
  "FileName": "文件名",
  "FileBusinessType": "单证业务类型（枚举：正式报关单|报关单草单|发票|装箱单|报关委托书|企业信息|企业默认值|其他未知单证）",
  "DecMessage": {
    "DecHead": {
      "IEFlag": "进出口标志（进口|出口）",
      "BillNo": "提单号",
      "ContrNo": "合同号",
      "CutMode": "征免性质",
      "DistinatePort": "经停港/抵运港",
      "FeeCurr": "运费币制",
      "FeeMark": "总价标记（费率|单价|总价）",
      "FeeRate": "运费率",
      "GrossWet": "毛重",
      "InsurCurr": "保险费币制",
      "InsurMark": "保险费标记（费率|单价|总价）",
      "InsurRate": "保险费率",
      "ManualNo": "备案号",
      "NetWt": "净重",
      "NoteS": "备注",
      "OtherCurr": "杂费币制",
      "OtherMark": "杂费标记（费率|单价|总价）",
      "OtherRate": "杂费率",
      "OwnerCode": "消费使用/生产销售企业海关编码",
      "OwnerCiqCode": "消费使用/生产销售企业检验检疫编码",
      "OwnerName": "消费使用/生产销售企业名称",
      "PackNo": "件数",
      "TradeCountry": "启运国/抵运国(地区)",
      "TradeMode": "贸易方式",
      "TrafMode": "运输方式",
      "TrafName": "运输工具名称",
      "VoyNo": "航次号",
      "TransMode": "成交方式",
      "WrapType": "包装种类",
      "TradeArea": "贸易国别(地区)",
      "DespPort": "启运港",
      "OverseasConsignorAEOCode": "境外发货企业AEO代码",
      "OverseasConsigneeEname": "境外收货企业名称(外文)",
      "TradeCode": "境内收发货企业海关编码",
      "TradeCiqCode": "境内收发货企业检验检疫编码",
      "TradeName": "境内收发货企业名称",
      "TradeNameEn": "境内收发货企业名称（英文）"
    },
    "DecLists": {
      "DecList": [
        {
          "Description": "货物描述",
          "GNo": "商品序号（数字）",
          "ContrItem": "备案序号",
          "CodeTS": "HS编码",
          "DeclPrice": "申报单价",
          "OriginCountry": "原产国(地区)",
          "TradeCurr": "成交币制",
          "DeclTotal": "申报总价",
          "NetWt": "净重",
          "Volume": "体积",
          "Qty1": "数量1",
          "Unit1": "单位1",
          "Qty2": "数量2",
          "Unit2": "单位2",
          "DestinationCountry": "最终目的国(地区)",
          "District": "境内目的地/境内货源地",
          "DutyMode": "征减免税方式"
        }
      ]
    }
  }
}
```

**重要：**
- 严格从原文提取，不编造内容
- 数字保持原样，不转换格式
- 缺失的字段直接省略，不要留空字符串或 null
- 英文内容保持原样，不要翻译

### 步骤 4: 运行脚本整合生成报关单

将所有文件的分析结果保存为 JSON 数组，然后运行脚本：

```bash
python skills/custom/declaration-data-handle/scripts/main.py \
  --analysis-results /path/to/analysis_results.json \
  --output-dir /mnt/user-data/outputs
```

或者，如果分析结果在内存中，直接以 JSON 字符串形式传递：

```bash
python skills/custom/declaration-data-handle/scripts/main.py \
  --analysis-json '{JSON字符串}' \
  --output-dir /mnt/user-data/outputs
```

**分析结果格式说明：**

分析结果可以是以下任一格式：

1. **数组格式（推荐）**：
```json
[
  {
    "guid": "文件唯一标识符（可选，未提供时自动生成）",
    "analysis": {
      "FileName": "文件名",
      "FileBusinessType": "单证业务类型",
      "DecMessage": {...}
    }
  },
  {
    "FileName": "文件名",
    "FileBusinessType": "单证业务类型",
    "DecMessage": {...}
  }
]
```

2. **单个对象格式**：
```json
{
  "FileName": "文件名",
  "FileBusinessType": "单证业务类型",
  "DecMessage": {...}
}
```

**企业默认值处理：**
- 当 `FileBusinessType` 为 `'企业默认值'` 时，其商品数据会作为默认值应用到所有商品
- 对于每个商品，如果某个字段缺失，会自动从企业默认值中填充

**发票作为基准：**
- 当 `FileBusinessType` 为 `'发票'` 时，作为商品基准
- 其他单证的商品数据通过 `Description` 进行模糊匹配（相似度阈值 70%）合并到基准中

**合并数据特点：**
- 每个字段的值都保留来源文件的 GUID 溯源信息
- 支持多个来源的数据合并，相同字段值去重
- 支持下拉框选择多值，不同来源值用颜色标记（蓝色=单一来源，黑色=多来源一致，红色=多来源不一致）

参数说明：
- `--analysis-results`: 分析结果 JSON 文件路径
- `--analysis-json`: 分析结果 JSON 字符串（替代 --analysis-results）
- `--output-dir`: 输出目录，默认 `/mnt/user-data/outputs`
- `--with-source`: 可选，打包时包含原始文件
- `--template-path`: 可选，报关单模板路径

### 步骤 5: 发送文件给用户

将生成的报关单 Excel 或压缩包发送给用户。

## 注意事项

- 使用 conda Python 3.12 环境运行脚本
- 确保 markitdown 已安装（用于文件转 Markdown）
- 报关单模板默认路径：`skills/custom/declaration-data-handle/scripts/assets/报关单.xlsx`
- 向量数据库路径：`skills/custom/declaration-data-handle/scripts/chroma_db`（用于查询商品税号、品名、申报要素）
- 产品库数据：`skills/custom/declaration-data-handle/scripts/data/product_library.csv`

### 重建向量数据库

如果需要更新或重建向量数据库，运行：

```bash
cd skills/custom/declaration-data-handle/scripts/data
python build_chroma_db.py
```

脚本会：
1. 读取 `product_library.csv`
2. 将数据导入到 `../chroma_db/product_library` 集合
3. 使用 `英文描述`、`品名`、`申报要素`、`税号` 字段进行向量化
