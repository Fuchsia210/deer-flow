***

name: hospital-report
description: 处理东方医院影像报告，将中文内容翻译成英文，并生成格式化的Word文档使用场景，1. 用户发送东方医院影像报告的图片时；2. 用户要求翻译之前收到的图片并生成Word文档时；3. 用户提供包含影像报告数据的JSON文件时。
----------------------------------------------------------------------------------------------------------------------------

# 东方医院影像报告Word版生成器

此技能用于处理东方医院影像报告，提取数据、翻译成英文并生成格式化的Word文档。

## Architecture

```
hospital-report/
├── SKILL.md                          ← You are here. Core logic and flow.
├── assets/
│   └── sample.docx                   ← Word文档模板
└── scripts/
    └── main.py                       ← 主脚本，生成Word文档
```

## Ground Rules

- **医学术语准确。** 使用专业医学领域术语进行翻译。
- **原文内容完整保留。** 中文内容必须完全遵照原文。
- **枚举值从指定选项选取。** 确保 EXAMINATION TYPE 和 GENDER 符合要求。
- **保持专业、规范的格式。**

## Workflow Phases

| Phase            | Goal          | Key Actions                                                                                 |
| ---------------- | ------------- | ------------------------------------------------------------------------------------------- |
| **1. 提取报告数据**    | 理解图片内容        | **先使用 view-image**工具处理图片，提取URL，**如果minimax的MCP图形理解工具可用，则将刚刚拿到的url作为入参，重新传入图形理解工具，再次解析图片内容** |
| **2. 整理JSON并翻译** | 生成标准中英文对照JSON | 整理内容为JSON格式，翻译成英文                                                                           |
| **3. 生成Word文档**  | 生成格式化Word     | 运行 main.py 脚本                                                                               |
| **4. 发送文档给用户**   | 交付结果          | 将生成的Word文档发送给用户                                                                             |

## JSON Structure

生成的JSON必须包含以下字段：

```json
{
  "chinese": {
    "EXAMINATION TYPE": "影像类型",
    "NAME": "患者姓名",
    "MRN": "放射检查编号",
    "DOB": "出生日期",
    "GENDER": "性别",
    "AGE": "年龄",
    "RADIOLOGY NO": "成像编号",
    "DATE": "日期",
    "EXAMINATION SITE": "检查部位",
    "Technique": "检查方法",
    "Findings": "放射学发现",
    "Impressions": "放射学诊断",
    "REPORTING PHYSICIAN": "报告医生"
  },
  "english": {
    "EXAMINATION TYPE": "影像类型（英文）",
    "NAME": "患者姓名（英文）",
    "MRN": "放射检查编号",
    "DOB": "出生日期",
    "GENDER": "性别（英文）",
    "AGE": "年龄",
    "RADIOLOGY NO": "成像编号",
    "DATE": "日期",
    "EXAMINATION SITE": "检查部位（英文）",
    "Technique": "检查方法（英文）",
    "Findings": "放射学发现（英文）",
    "Impressions": "放射学诊断（英文）",
    "REPORTING PHYSICIAN": "报告医生（英文）"
  }
}
```

### Field Notes

- `EXAMINATION TYPE`：英文从以下选项选取：X RAY；PLAIN MRI；CONTRAST ENHANCED MRI；PLAIN CT SCAN；CONTRAST ENHANCED CT SCAN
- `GENDER`：英文从以下选项选取：Male；Female
- `Impressions`：可以是字符串或字符串数组
- 除字段说明特别申明的以外，中文内容必须完全遵照原文，英文使用医学领域的专业术语

### JSON Example

```json
{
  "chinese": {
    "EXAMINATION TYPE": "MRI增强",
    "NAME": "张三",
    "MRN": "123456",
    "DOB": "1990-01-01",
    "GENDER": "男",
    "AGE": "36",
    "RADIOLOGY NO": "MR20230101001",
    "DATE": "2023-01-01",
    "EXAMINATION SITE": "头部",
    "Technique": "头颅MRI平扫，层厚5mm，层距lmm；横断面：T1WI、T2WI、FLAIR、DWI（b=0，b=1000）；矢状面：T1WI。",
    "Findings": "脑实质未见明显异常信号，脑室系统未见扩张，脑沟裂未见增宽",
    "Impressions": ["双侧小脑沟可疑T1W1高信号，请结合临床，建议SW1复查","轻度脑萎缩"],
    "REPORTING PHYSICIAN": "张三"
  },
  "english": {
    "EXAMINATION TYPE": "CONTRAST ENHANCED MRI",
    "NAME": "Zhang San",
    "MRN": "123456",
    "DOB": "1990-01-01",
    "GENDER": "Male",
    "AGE": "36",
    "RADIOLOGY NO": "23010101",
    "DATE": "2023-01-01",
    "EXAMINATION SITE": "Head",
    "Technique": "Head MRI plain scan, with a slice thickness of 5mm and a slice interval of 1mm; Axial section: T1WI, T2WI, FLAIR, DWI (b=0, b=1000); Sagittal section: T1WI.",
    "Findings": "No obvious abnormal signals in brain parenchyma, no ventricular system dilatation, no widening of cerebral sulci",
    "Impressions": ["Bilateral cerebellar sulci suspicious T1WI hyperintensity, please correlate clinically", "recommend SWI re-examination."],
    "REPORTING PHYSICIAN": "Zhang San"
  }
}
```

## Script Usage

使用 `scripts/main.py` 生成Word文档：

```bash
python skills/custom/hospital-report/scripts/main.py \
  --json-file /mnt/user-data/workspace/report.json \
  --output-dir /mnt/user-data/outputs/
```

**脚本说明：**

- `--json-file`：包含报告数据的JSON文件路径（建议使用 `/mnt/user-data/workspace/`）
- `--output-dir`：生成的Word文档保存目录（建议使用 `/mnt/user-data/outputs/`）
- 文件名格式：`{MRN}_{NAME}_{timestamp}.docx`

**脚本功能：**

- 基于 `assets/sample.docx` 模板创建文档
- 保留页眉页脚
- 使用英文数据填充报告内容
- 自动创建输出目录（如果不存在）

## Complete Execution Steps

1. **接收用户输入**：确认用户提供了影像报告图片或JSON文件
2. **提取/接收数据**：理解图片内容或直接使用JSON文件
3. **整理为标准格式**：确保JSON包含所有必需字段
4. **保存JSON文件**：将JSON数据保存到 `/mnt/user-data/workspace/` 目录
5. **调用脚本**：运行脚本生成Word文档
6. **获取输出**：脚本会打印生成的文档路径
7. **发送给用户**：将生成的Word文档发送给用户

## Notes

- 脚本需要 `python-docx` 库支持
- 确保 `assets/sample.docx` 模板文件存在
- 生成的Word文档使用英文数据，格式为标准的放射学报告

