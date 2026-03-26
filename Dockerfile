FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HF_HOME=/root/.cache/huggingface

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

COPY requirements.txt /tmp/requirements.txt

RUN python -m pip install --upgrade pip setuptools wheel && \
    pip install torch==2.10.0 --index-url https://download.pytorch.org/whl/cpu && \
    pip install -r /tmp/requirements.txt

RUN python -c "import torch; print(torch.__version__)" && \
    python -c "import accelerate; print(accelerate.__version__)" && \
    python -c "import transformers; print(transformers.__version__)" && \
    python -c "import cv2; print(cv2.__version__)" && \
    python -c "from PIL import Image; print('PIL ok')" && \
    python -c "from transformers import AutoProcessor, Qwen2_5_VLForConditionalGeneration; print('qwen vl ok')" && \
    python -c "from surya.table_rec import TableRecPredictor; print('surya ok')"

CMD ["sleep", "infinity"]
