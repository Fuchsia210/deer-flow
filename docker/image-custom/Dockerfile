FROM enterprise-public-cn-beijing.cr.volces.com/vefaas-public/all-in-one-sandbox:latest

ARG PIP_INDEX_URL
ARG HF_ENDPOINT_URL

ENV PIP_INDEX_URL=${PIP_INDEX_URL:-https://pypi.tuna.tsinghua.edu.cn/simple}
ENV HF_ENDPOINT=${HF_ENDPOINT_URL:-https://hf-mirror.com}

RUN echo "Installing pre-cached Python packages for hospital-report and declaration-data-handle skills..." && \
    pip install --no-cache-dir --index-url "$PIP_INDEX_URL" \
    python-docx>=1.0.0 \
    Pillow>=10.0.0 \
    lxml>=4.9.0 \
    openpyxl>=3.1.0 \
    pandas>=2.0.0 \
    numpy>=1.24.0 \
    fuzzywuzzy==0.18.0 \
    python-Levenshtein==0.27.3 \
    RapidFuzz==3.14.5 \
    lancedb==0.30.2 \
    sentence-transformers==5.4.1 \
    torch==2.11.0 \
    xlrd>=2.0.0 \
    && echo "Pre-cached packages installed successfully"

COPY markitdown-src /home/ayl13/markitdown

RUN echo "Installing markitdown from local source..." && \
    pip install --no-cache-dir --index-url "$PIP_INDEX_URL" -e /home/ayl13/markitdown/packages/markitdown[all] && \
    echo "Markitdown installed successfully"

RUN echo "Pre-downloading sentence transformer model (using HF mirror)..." && \
    python -c "from sentence_transformers import SentenceTransformer; model = SentenceTransformer('all-MiniLM-L6-v2'); print('Model downloaded successfully')" && \
    echo "Model pre-cached successfully"

RUN echo "Setting up initialization marker..." && \
    echo "Dependencies are pre-installed in this image" > /sandbox_ready.txt
