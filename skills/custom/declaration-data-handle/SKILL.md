---
name: declaration-data-handle
description: 处理报关资料压缩包，解压后将文件转为 Markdown，通过 Agent 分析提取数据，最终生成报关单 Excel 文件。
---

# 报关单生成器

此技能用于处理报关资料，解压文件、分析内容、提取数据并生成报关单 Excel。

## Architecture

```
declaration-data-handle/
├── SKILL.md                          ← You are here. Core logic and flow.
└── scripts/
    ├── main.py                       ← 主脚本，整合数据并生成 Excel
    ├── check.py                      ← 文件检查脚本，校验文件数量
    ├── filetomd.py                   ← 文件批量转 Markdown 脚本（支持 Excel 分 Sheet）
    ├── assets/
    │   └── 报关单.xlsx               ← 报关单模板
    ├── data/
    │   ├── product_library.csv       ← 产品库数据
    │   └── build_lance_db.py         ← 重建向量数据库脚本
    └── lance_db/
        └── product_library.lance/    ← 向量数据库（LanceDB）
```

## Ground Rules

- **严格按步骤执行**： 一步一个脚印，不要跳过流程。
- **数据准确性优先**： 严格从原文提取，不编造内容，数字保持原样，不要翻译。
- **缺失字段直接省略**：不要留空字符串或 null。
- **每个文件按照结构尽可能地提取内容**：不要忽略任何字段。

## Workflow Phases

| Phase           | Goal                  | Key Actions                                                                                     |
| --------------- | --------------------- | ----------------------------------------------------------------------------------------------- |
| **1. 准备文件**     | 将压缩包解压                | 将压缩包解压                                                                                          |
| **2. 批量转换格式**   | 将解压目录下所有文件转为 Markdown | 执行 `filetomd.py` 批量转换此次解压的所有文件，输出 MD 文件保存在解压目录下                                                 |
| **3. 校验文件数量**   | 校验转换后的文件数量是否正确        | 运行 `check.py` 检查原始文件目录获取预期文件数量（含 Excel sheet），对比 markdown 目录下的 MD 文件数量，如不一致则重新执行 filetomd.py 转换 |
| **4. Agent 分析** | 提取报关数据                | 逐个分析 markdown/ 目录下的 MD 文件，按 JSON Schema 提取数据                                                    |
| **5. 生成报关单**    | 整合数据生成 Excel          | 运行 main.py 脚本                                                                                   |
| **6. 交付文件**     | 发送给用户                 | 将生成的 Excel 文件发送给用户                                                                              |

## Analysis Results Format

- `main.py` 脚本**只接受带外层包装的数组格式**，必须是数组，每个元素必须包含：
  - `guid`：文件唯一标识符（必填，上层调用系统负责生成）
  - `analysis`：单证分析结果对象（必填，不能为 null/非字典类型）

```json
[
  {
    "guid": "550e8400-e29b-41d4-a716-446655440000",
    "analysis": {
      "FileName": "{{file_name}}",
      "FileBusinessType": "单证业务类型（枚举：正式报关单|报关单草单|发票|装箱单|报关委托书|企业信息|企业默认值|其他未知单证）",
      "Note": "正式报关单不会在识别阶段出现",
      "DecMessage": {
        "DecHead": {
          "IEFlag": "进出口标志",
          "BillNo": "提单号",
          "ContrNo": "合同号",
          "CutMode": "征免性质",
          "DistinatePort": "经停港/抵运港",
          "FeeCurr": "运费币制",
          "FeeMark": "总价标记",
          "FeeRate": "运费率",
          "GrossWet": "毛重",
          "InsurCurr": "保险费币制",
          "InsurMark": "保险费标记",
          "InsurRate": "保险费率",
          "ManualNo": "备案号",
          "NetWt": "净重",
          "NoteS": "备注",
          "OtherCurr": "杂费币制",
          "OtherMark": "杂费标记",
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
              "GNo": "商品序号",
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
  }
]
```

### 字段详细规范（JSON Schema）

#### DecHead 字段约束

| 字段路径                       | 类型      | 特殊约束                                          |
| -------------------------- | ------- | --------------------------------------------- |
| `IEFlag`                   | string  | 枚举：进口 / 出口                                    |
| `BillNo`                   | string  | 提单号                                           |
| `ContrNo`                  | string  | 合同号，ContractNo 或 InvoiceNo 并存时 ContractNo 优先  |
| `CutMode`                  | string  | 征免性质，如：一般征税、其他法定                              |
| `DistinatePort`            | string  | 经停港/抵运港，仅港口名称，不含国家                            |
| `FeeCurr`                  | string  | 运费币制                                          |
| `FeeMark`                  | string  | 枚举：费率 / 单价 / 总价                               |
| `FeeRate`                  | number  | 最多两位小数                                        |
| `GrossWet`                 | number  | 毛重(KG)，最多六位小数                                 |
| `InsurCurr`                | string  | 保险费币制                                         |
| `InsurMark`                | string  | 枚举：费率 / 单价 / 总价                               |
| `InsurRate`                | number  | 最多两位小数                                        |
| `ManualNo`                 | string  | 备案号，12位，大写字母开头+11位数字，通常为保税手册/账册号              |
| `NetWt`                    | number  | 净重(KG)，最多六位小数                                 |
| `NoteS`                    | string  | 备注，仅从报关单草单的"标记唛码及备注"中获取                       |
| `OtherCurr`                | string  | 杂费币制                                          |
| `OtherMark`                | string  | 枚举：费率 / 单价 / 总价                               |
| `OtherRate`                | number  | 杂费率，可为负数，最多两位小数                               |
| `OwnerCode`                | string  | 消费使用/生产销售企业海关编码，18位社会信用代码                     |
| `OwnerCiqCode`             | string  | 消费使用/生产销售企业检验检疫编码，10位数字或字母                    |
| `OwnerName`                | string  | 消费使用/生产销售企业名称                                 |
| `PackNo`                   | integer | 件数（运输包装件数，非PCS数量）                             |
| `TradeCountry`             | string  | 启运国/抵运国(地区)，国家名称，如 Japan、中国                   |
| `TradeMode`                | string  | 贸易方式，如：一般贸易、进料对口                              |
| `TrafMode`                 | string  | 运输方式，如：水路运输、by sea                            |
| `TrafName`                 | string  | 运输工具代码及名称                                     |
| `VoyNo`                    | string  | 航次号                                           |
| `TransMode`                | string  | 成交方式，如：FOB、CIF                                |
| `WrapType`                 | string  | 包装种类，如：CTNS、PALLETS                           |
| `TradeArea`                | string  | 贸易国别(地区)，国家名称                                 |
| `DespPort`                 | string  | 启运港，仅港口名称                                     |
| `OverseasConsignorAEOCode` | string  | 境外发货企业AEO代码                                   |
| `OverseasConsigneeEname`   | string  | 境外收货企业名称(外文)                                  |
| `TradeCode`                | string  | 境内收发货企业海关编码，18位社会信用代码                         |
| `TradeCiqCode`             | string  | 境内收发货企业检验检疫编码，10位数字或字母                        |
| `TradeName`                | string  | 境内收发货企业名称，出口：通常位于文件的最上方，即售方，进口：通常位于文件表头部分，即买方 |
| `TradeNameEn`              | string  | 境内收发货企业名称（英文）                                 |

#### DecList 字段约束

| 字段路径                 | 类型      | 特殊约束                |
| -------------------- | ------- | ------------------- |
| `Description`        | string  | 货物描述，完整，不得缺失任何信息    |
| `GNo`                | integer | 商品顺序号，从1开始          |
| `ContrItem`          | string  | 备案序号                |
| `CodeTS`             | string  | HS编码                |
| `DeclPrice`          | number  | 申报单价，最多两位小数         |
| `OriginCountry`      | string  | 原产国(地区)，国家名称        |
| `TradeCurr`          | string  | 成交币制                |
| `DeclTotal`          | number  | 申报总价，最多两位小数         |
| `NetWt`              | number  | 净重(KG)，最多六位小数       |
| `Volume`             | number  | 体积(M³)，最多三位小数       |
| `Qty1`               | number  | Unit1对应的数量，最多六位小数   |
| `Unit1`              | string  | 单位1（净重、体积以外的单位）     |
| `Qty2`               | number  | Unit2对应的数量，最多六位小数   |
| `Unit2`              | string  | 单位2（净重、体积以外的单位）     |
| `DestinationCountry` | string  | 最终目的国(地区)，国家名称      |
| `District`           | string  | 境内目的地/境内货源地，如：上海其他等 |
| `DutyMode`           | string  | 征减免税方式，如：照章征税、全免等   |

### 校验规则

脚本启动时会严格校验输入的 JSON 文件：

- 解析失败（文件不存在、JSON 格式错误）→ 返回错误码 1
- 根元素必须是**数组**
- 数组中的**每个元素**代表一份单证分析结果，必须是**字典**
- 每个单证字典必须有 `guid` 字段（文件唯一标识符，由上层调用系统生成，不能缺失）
- 每个单证字典必须有 `analysis` 字段（单证分析结果对象，不能缺失）
- `analysis` 字段必须是**字典**类型（不能为 null 或其他类型）
- 任何一项校验不通过，脚本立即返回错误码 1 退出

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

**参数说明：**

| 参数                   | 说明                                   | 默认值                       |
| -------------------- | ------------------------------------ | ------------------------- |
| `--analysis-results` | 分析结果 JSON 文件路径（必填）                   | -                         |
| `--output-dir`       | 输出目录                                 | `/mnt/user-data/outputs`  |
| `--template-path`    | 报关单模板路径                              | `scripts/assets/报关单.xlsx` |

**依赖检查：**

- 如果 `lancedb` 和 `sentence-transformers` 未安装，向量检索功能将被禁用
- 如果 `openpyxl` 未安装，将无法生成 Excel 文件

## 批量转换 Markdown

使用 `filetomd.py` 脚本将解压目录下所有支持的文件批量转换为 Markdown：

### 使用 filetomd.py

```bash
# 批量转换整个解压目录，自动输出到 extract_dir/markdown/
python scripts/filetomd.py /tmp/declaration_extract

# 自定义输出目录
python scripts/filetomd.py /tmp/declaration_extract -o /tmp/my_md_output

# 单个文件转换
python scripts/filetomd.py /path/to/document.pdf
```

脚本特性：

- 自动识别所有支持格式（PDF、DOCX、XLSX、XLS、PPTX、图片等）
- Excel 文件自动将每个 Sheet 单独保存为独立的 MD 文件
- 输出目录自动创建，无需手动提前建立

## 校验文件数量

在执行完 filetomd.py 转换后，必须校验文件数量是否一致：

### 校验流程

1. 使用 `check.py` 检查解压后的原始文件目录，获取预期的文件数量（含 Excel sheet）
2. 统计 `markdown/` 子目录下转换后的 Markdown 文件实际数量
3. 对比两个数量是否一致
4. 如果一致，继续下一步；如果不一致，重新执行 filetomd.py 批量转换

### 使用 check.py

```bash
# 检查原始文件
python scripts/check.py /tmp/declaration_extract --json
```

返回的 JSON 中 `total_count` 字段即为预期的文件数量（Excel 的每个 sheet 算作一个文件）。

### 校验示例

```python
import os
import json
from pathlib import Path
from scripts.check import check_files

# 1. 运行批量转换
extract_dir = Path("/tmp/declaration_extract")
import subprocess
subprocess.run(["python", "scripts/filetomd.py", str(extract_dir)], check=True)

# 2. 检查原始文件获取预期数量
original_info = check_files(str(extract_dir))
expected_count = original_info["total_count"]

# 3. 统计 MD 文件
md_dir = extract_dir / "markdown"
md_count = len(list(md_dir.glob("*.md"))) if md_dir.exists() else 0

# 4. 对比校验
if expected_count == md_count:
    print("文件数量校验通过，继续下一步")
else:
    print(f"文件数量不一致：期望 {expected_count}，实际 {md_count}，重新转换格式")
```

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
