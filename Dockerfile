FROM python:3.10-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

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

RUN python -m pip install --upgrade pip && \
    pip install "transformers>=4.56.1,<5" && \
    pip install torch --index-url https://download.pytorch.org/whl/cpu && \
    pip install surya-ocr --no-deps && \
    pip install "pillow>=10.2.0,<11" && \
    pip install click einops filetype numpy opencv-python-headless platformdirs pre-commit pydantic-settings pypdfium2==4.30.0 python-dotenv requests

RUN python -c "import torch; print(torch.__version__)" && \
    python -c "import transformers; print(transformers.__version__)" && \
    python -c "import cv2; print(cv2.__version__)" && \
    python -c "from PIL import Image; print('PIL ok')" && \
    python -c "from surya.table_rec import TableRecPredictor; print('surya ok')"

CMD ["sleep", "infinity"]
