---
name: declaration-data-handle
description: |-
  处理报关资料压缩包，解压后将文件转为 Markdown，通过 Agent 分析提取数据，最终生成报关单 Excel 文件。
  触发场景：
  1. 用户发送包含报关资料的压缩包时
  2. 用户要求处理之前收到的压缩包时
  依赖环境：
  - Python 依赖：torch, lancedb, sentence-transformers, openpyxl, fuzzywuzzy, pandas
  - 向量数据库：skills/custom/declaration-data-handle/scripts/lance_db/product_library.lance
---

# 报关单生成器

此技能用于处理报关资料，解压文件、分析内容、提取数据并生成报关单 Excel。

## Architecture

```
declaration-data-handle/
├── SKILL.md                          ← You are here. Core logic and flow.
└── scripts/
    ├── main.py                       ← 主脚本，整合数据并生成 Excel
    ├── assets/
    │   └── 报关单.xlsx               ← 报关单模板
    ├── data/
    │   ├── product_library.csv       ← 产品库数据
    │   └── build_lance_db.py         ← 重建向量数据库脚本
    └── lance_db/
        └── product_library.lance/    ← 向量数据库（LanceDB）
```

## Ground Rules

- **严格按步骤执行。** 一步一个脚印，不要跳过流程。
- **数据准确性优先。** 严格从原文提取，不编造内容，数字保持原样。
- **缺失字段直接省略。** 不要留空字符串或 null。
- **英文内容保持原样。** 不要翻译。

## Workflow Phases

| Phase | Goal | Key Actions |
|-------|------|-------------|
| **1. 准备文件** | 将压缩包移动到工作区并解压 | 移动文件到 `/mnt/user-data/workspace`，解压到 `/tmp/declaration_extract` |
| **2. 转换格式** | 将文件转为 Markdown | Agent 将文件转为 Markdown，**Excel 文件每个 sheet 单独转换为一个 MD 文件** |
| **3. Agent 分析** | 提取报关数据 | 逐个分析文件，按 JSON Schema 提取数据 |
| **4. 生成报关单** | 整合数据生成 Excel | 运行 main.py 脚本 |
| **5. 交付文件** | 发送给用户 | 将生成的 Excel 文件发送给用户 |

## JSON Schema for Analysis

每个文件分析后需按以下 JSON 格式输出：

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

## Analysis Results Format

分析结果可以是以下任一格式：

### 1. 数组格式（推荐）
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

### 2. 单个对象格式
```json
{
  "FileName": "文件名",
  "FileBusinessType": "单证业务类型",
  "DecMessage": {...}
}
```

## Data Merging Rules

### 企业默认值处理
- 当 `FileBusinessType` 为 `'企业默认值'` 时，其商品数据会作为默认值应用到所有商品
- 对于每个商品，如果某个字段缺失，会自动从企业默认值中填充

### 发票作为基准
- 当 `FileBusinessType` 为 `'发票'` 时，作为商品基准
- 其他单证的商品数据通过 `Description` 进行模糊匹配（相似度阈值 70%）合并到基准中

### 合并数据特点
- 每个字段的值都保留来源文件的 GUID 溯源信息
- 支持多个来源的数据合并，相同字段值去重
- 支持下拉框选择多值，不同来源值用颜色标记（蓝色=单一来源，黑色=多来源一致，红色=多来源不一致）

## Script Usage

运行 `scripts/main.py` 生成报关单：

```bash
python scripts/main.py \
  --analysis-results /path/to/analysis_results.json \
  --output-dir /mnt/user-data/outputs \
  --template-path scripts/assets/报关单.xlsx
```

或者使用 JSON 字符串：

```bash
python scripts/main.py \
  --analysis-json '[{"FileName": "xxx.pdf", ...}]' \
  --output-dir /mnt/user-data/outputs
```

**参数说明：**
| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--analysis-results` | 分析结果 JSON 文件路径 | - |
| `--analysis-json` | 分析结果 JSON 字符串（替代 --analysis-results） | - |
| `--output-dir` | 输出目录 | `/mnt/user-data/outputs` |
| `--template-path` | 报关单模板路径 | `scripts/assets/报关单.xlsx` |

**依赖检查：**
- 如果 `lancedb` 和 `sentence-transformers` 未安装，向量检索功能将被禁用
- 如果 `openpyxl` 未安装，将无法生成 Excel 文件

## Rebuild Vector Database

如果需要更新或重建向量数据库，运行：

```bash
cd scripts/data
python build_lance_db.py
```

脚本会：
1. 读取 `product_library.csv`
2. 将数据导入到 `../lance_db/product_library.lance` 数据库
3. 使用 `英文描述`、`品名`、`申报要素`、`税号` 字段进行向量化

**依赖：** `lancedb`, `sentence-transformers`
