# Custom Sandbox Image

此目录包含用于预装依赖的自定义sandbox镜像配置。

## 目录结构

```
docker/image-custom/
├── Dockerfile           # 自定义sandbox镜像
├── README.md            # 本文件
├── build.sh             # 快速构建脚本
└── sandbox-config-example.yaml  # config.yaml配置示例
```

## 支持的Skills

本镜像预装了以下Skills所需的依赖：

### 1. hospital-report

处理东方医院影像报告，翻译并生成Word文档。

### 2. declaration-data-handle

处理报关资料压缩包，提取数据并生成报关单Excel。

## 预装的依赖

### 核心依赖

- `python-docx` - Word文档生成 (hospital-report)
- `openpyxl` - Excel文件处理 (declaration-data-handle)
- `Pillow` - 图片处理
- `lxml` - XML/HTML处理
- `pandas` - 数据处理

### 向量数据库依赖 (declaration-data-handle)

- `lancedb` - 向量数据库
- `sentence-transformers` - 文本嵌入模型
- `torch` - PyTorch深度学习框架

### 模糊匹配依赖 (declaration-data-handle)

- `fuzzywuzzy` - 模糊字符串匹配
- `python-Levenshtein` - Levenshtein距离算法

## 使用方法

### 步骤1：构建自定义镜像

```bash
cd /home/ayl13/deer-flow/docker/image-custom

# 使用清华镜像源构建（推荐国内用户）
docker build -t deer-flow-sandbox:custom \
  --build-arg PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple \
  .

# 或使用阿里云镜像源
docker build -t deer-flow-sandbox:custom \
  --build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple \
  .
```

### 步骤2：更新 config.yaml

编辑项目根目录的 `config.yaml` 文件，找到 `sandbox` 部分：

```yaml
sandbox:
  use: deerflow.community.aio_sandbox:AioSandboxProvider
  image: deer-flow-sandbox:custom  # 使用自定义镜像
  replicas: 3                         # 预热容器数量
  idle_timeout: 3600                  # 1小时空闲超时
  environment:
    PIP_INDEX_URL: https://pypi.tuna.tsinghua.edu.cn/simple
```

### 步骤3：重启服务

```bash
# 如果使用Docker开发模式
make docker-stop
make docker-start

# 如果使用本地开发模式
# 重启DeerFlow服务
```

## 验证

构建完成后，可以使用以下命令验证镜像：

```bash
# 验证 hospital-report 依赖
docker run --rm deer-flow-sandbox:custom python -c "import docx; print('python-docx OK')"

# 验证 declaration-data-handle 依赖
docker run --rm deer-flow-sandbox:custom python -c "import openpyxl; print('openpyxl OK')"
docker run --rm deer-flow-sandbox:custom python -c "import fuzzywuzzy; print('fuzzywuzzy OK')"
docker run --rm deer-flow-sandbox:custom python -c "import lancedb; print('lancedb OK')"
docker run --rm deer-flow-sandbox:custom python -c "import sentence_transformers; print('sentence-transformers OK')"
```

