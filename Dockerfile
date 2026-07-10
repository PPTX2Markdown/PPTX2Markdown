FROM python:3.11-slim

ARG TORCH_INDEX_URL=https://download.pytorch.org/whl/cu128

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HF_HOME=/root/.cache/huggingface \
    TORCH_INDEX_URL=${TORCH_INDEX_URL}

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    fonts-dejavu-core \
    libglib2.0-0 \
    libgl1 \
    libgomp1 \
    libsm6 \
    libxext6 \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /workspace

COPY . /workspace

RUN python -m pip install --upgrade pip setuptools wheel && \
    pip install torch torchvision --index-url ${TORCH_INDEX_URL} && \
    pip install ".[all]"

RUN python -c "import torch; print('torch', torch.__version__, 'cuda_build=', torch.version.cuda)" && \
    python -c "import pptx2markdown; print('pptx2markdown', pptx2markdown.__version__)" && \
    python -c "from transformers import AutoProcessor, Qwen2_5_VLForConditionalGeneration; print('qwen vl ok')" && \
    python -c "from surya.layout import LayoutPredictor; print('surya ok')"

CMD ["sleep", "infinity"]
