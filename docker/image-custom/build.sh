#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
IMAGE_NAME="${IMAGE_NAME:-deer-flow-sandbox:custom}"

echo "=========================================="
echo "  Building Custom Sandbox Image"
echo "=========================================="
echo ""
echo "Image: $IMAGE_NAME"
echo "Directory: $SCRIPT_DIR"
echo ""

# 检测是否有自定义PIP_INDEX_URL
if [ -n "$PIP_INDEX_URL" ]; then
    echo "Using custom PIP_INDEX_URL: $PIP_INDEX_URL"
    BUILD_ARGS="--build-arg PIP_INDEX_URL=$PIP_INDEX_URL"
elif [ -f "$SCRIPT_DIR/../.env" ]; then
    # 从项目.env文件中读取PIP_INDEX_URL
    source "$SCRIPT_DIR/../.env" 2>/dev/null || true
    if [ -n "$PIP_INDEX_URL" ]; then
        echo "Using PIP_INDEX_URL from .env: $PIP_INDEX_URL"
        BUILD_ARGS="--build-arg PIP_INDEX_URL=$PIP_INDEX_URL"
    else
        echo "Using default PIP_INDEX_URL (Aliyun mirror)"
        BUILD_ARGS="--build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple"
    fi
else
    echo "Using default PIP_INDEX_URL (Aliyun mirror)"
    BUILD_ARGS="--build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple"
fi

MARKITDOWN_SOURCE="${MARKITDOWN_SOURCE:-/home/ayl13/markitdown}"
if [ -d "$MARKITDOWN_SOURCE" ]; then
    echo "Using local markitdown source: $MARKITDOWN_SOURCE"
    mkdir -p "$SCRIPT_DIR/markitdown-src"
    cp -r "$MARKITDOWN_SOURCE"/* "$SCRIPT_DIR/markitdown-src/"
fi

echo ""
echo "Building Docker image..."
docker build $BUILD_ARGS -t "$IMAGE_NAME" "$SCRIPT_DIR"

echo ""
echo "=========================================="
echo "  Build Complete!"
echo "=========================================="
echo ""
echo "To use this image, add the following to your config.yaml:"
echo ""
cat << 'EOF'
sandbox:
  use: deerflow.community.aio_sandbox:AioSandboxProvider
  image: $IMAGE_NAME
  replicas: 3
  idle_timeout: 3600
  environment:
    PIP_INDEX_URL: https://mirrors.aliyun.com/pypi/simple
EOF
echo ""

# 验证镜像
echo "Verifying image..."
echo ""

echo "Checking hospital-report dependencies..."
docker run --rm "$IMAGE_NAME" python -c "import docx; print('  ✓ python-docx')" 2>/dev/null || echo "  ⚠ python-docx not available"
docker run --rm "$IMAGE_NAME" python -c "from PIL import Image; print('  ✓ Pillow')" 2>/dev/null || echo "  ⚠ Pillow not available"

echo ""
echo "Checking declaration-data-handle dependencies..."
docker run --rm "$IMAGE_NAME" python -c "import openpyxl; print('  ✓ openpyxl')" 2>/dev/null || echo "  ⚠ openpyxl not available"
docker run --rm "$IMAGE_NAME" python -c "import pandas; print('  ✓ pandas')" 2>/dev/null || echo "  ⚠ pandas not available"
docker run --rm "$IMAGE_NAME" python -c "import fuzzywuzzy; print('  ✓ fuzzywuzzy')" 2>/dev/null || echo "  ⚠ fuzzywuzzy not available"
docker run --rm "$IMAGE_NAME" python -c "import lancedb; print('  ✓ lancedb')" 2>/dev/null || echo "  ⚠ lancedb not available"
docker run --rm "$IMAGE_NAME" python -c "import sentence_transformers; print('  ✓ sentence-transformers')" 2>/dev/null || echo "  ⚠ sentence-transformers not available"
docker run --rm "$IMAGE_NAME" python -c "import torch; print('  ✓ torch')" 2>/dev/null || echo "  ⚠ torch not available"
docker run --rm "$IMAGE_NAME" python -c "from markitdown import MarkItDown; print('  ✓ markitdown')" 2>/dev/null || echo "  ⚠ markitdown not available"
docker run --rm "$IMAGE_NAME" python -c "import xlrd; print('  ✓ xlrd')" 2>/dev/null || echo "  ⚠ xlrd not available"

echo ""
echo "Done!"
