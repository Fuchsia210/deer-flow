---
name: hospital-report
description: 处理东方医院影像报告，将中文内容翻译成英文，并生成格式化的Word文档,使用场景1. 用户发送东方医院影像报告的图片时;2. 用户要求翻译之前收到的图片并生成Word文档时;3. 用户提供包含影像报告数据的JSON文件时
---

# 东方医院影像报告Word版生成器

此技能用于处理东方医院影像报告，提取相关数据，翻译成英文，并生成格式化的Word文档。

## 工作流程

### 1. 提取报告数据

- 将刚刚上传图片的url作为输入传递给`MiniMax_understand_image` 读取图片信息

### 2. 整理json并翻译

- 将中文内容翻译成英文，使用医学领域的专业术语,
- 将提取的内容整理为标准中英文对照JSON格式（见下方JSON结构）

**注意事项：**

- 英文 `EXAMINATION TYPE` 从以下选项中选取：X RAY；PLAIN MRI；CONTRAST ENHANCED MRI；PLAIN CT SCAN；CONTRAST ENHANCED CT SCAN
- 英文 `GENDER` 从以下选项中选取：Male；Female
- 除字段说明特别申明的以外，中文内容必须完全遵照原文，英文使用医学领域的专业术语

### 3. 生成Word文档

使用 `scripts/main.py` 脚本生成Word文档：

```bash
python scripts/main.py --json-file <json_file_path> --output-dir <output_directory>
```

**脚本说明：**

- 脚本位于 `scripts/main.py`
- `--json-file`：包含报告数据的JSON文件路径（建议使用 `/mnt/user-data/workspace/`）
- `--output-dir`：生成的Word文档保存目录（建议使用 `/mnt/user-data/outputs/`）
- 文件名格式：`{MRN}_{NAME}_{timestamp}.docx`

### 4. 发送文档给用户

生成完成后，将Word文档发送给用户。

## JSON结构

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

### 字段说明

- `EXAMINATION TYPE`：影像类型，例如：X光，MRI增强等
- `NAME`：患者姓名
- `MRN`：放射检查编号
- `DOB`：出生日期
- `GENDER`：性别
- `AGE`：年龄
- `RADIOLOGY NO`：成像编号
- `DATE`：日期
- `EXAMINATION SITE`：检查部位，描述影像扫描的人体部位，例如：1. 胸部正位； 2. 胰腺；3. 头颅常规
- `Technique`：检查方法，描述检查是如何操作，可能只包含人体部位，也可能含有一些技术指标，例如：1. 胸部正位；2. 头颅MRI平扫，层厚5mm，层距lmm……；3. 取仰卧位，头先进机架……
- `Findings`：放射学发现，描述检查以后发现了什么现象，例如：1. 双侧大脑半球对称，……；2.肝脏形态，大小正常……
- `Impressions`：放射学诊断，描述发现的问题或者未发现问题，可以是字符串或字符串数组
- `REPORTING PHYSICIAN`：报告医生

### JSON样例

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

## 脚本使用说明

### scripts/main.py

此脚本接收JSON文件作为输入，生成格式化的Word文档。

**功能：**

- 基于 `assets/sample.docx` 模板创建文档
- 保留页眉页脚
- 使用英文数据填充报告内容
- 生成文件名格式：`{MRN}_{NAME}_{timestamp}.docx`

**使用方法：**

```bash
python scripts/main.py <json_file_path>
```

**配置：**

- `SAMPLE_DOCX_PATH`：模板文档路径（默认：`../assets/sample.docx`）

## 完整执行步骤

1. **接收用户输入**：确认用户提供了影像报告图片或JSON文件
2. **提取/接收数据**：使用MiniMax理解图片或直接使用JSON文件
3. **整理为标准格式**：确保JSON包含所有必需字段
4. **保存JSON文件**：将JSON数据保存到 `/mnt/user-data/workspace/` 目录
5. **调用脚本**：运行脚本生成Word文档
   ```bash
   python scripts/main.py \
     --json-file /mnt/user-data/workspace/report.json \
     --output-dir /mnt/user-data/outputs/
   ```
6. **获取输出**：脚本会打印生成的文档路径
7. **发送给用户**：将生成的Word文档发送给用户

## 注意事项

- 脚本需要 `python-docx` 库支持
- 确保 `assets/sample.docx` 模板文件存在
- 生成的Word文档使用英文数据，格式为标准的放射学报告
- 输出目录会通过命令行参数 `--output-dir` 指定，如果不存在会自动创建

