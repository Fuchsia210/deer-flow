FROM enterprise-public-cn-beijing.cr.volces.com/vefaas-public/all-in-one-sandbox:latest

ARG PIP_INDEX_URL

ENV PIP_INDEX_URL=${PIP_INDEX_URL:-https://pypi.tuna.tsinghua.edu.cn/simple}

RUN echo "Installing pre-cached Python packages for hospital-report and declaration-data-handle skills..." && \
    pip install --no-cache-dir --index-url "$PIP_INDEX_URL" \
    python-docx \
    Pillow \
    lxml \
    openpyxl \
    pandas \
    fuzzywuzzy \
    python-Levenshtein \
    lancedb \
    sentence-transformers \
    torch \
    && echo "Pre-cached packages installed successfully"

RUN echo "Setting up initialization marker..." && \
    echo "Dependencies are pre-installed in this image" > /sandbox_ready.txt
