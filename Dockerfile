FROM python:3.11-slim

WORKDIR /app

# 依存関係のインストール
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    libxcb1 \
    libx11-6 \
    libsm6 \
    libxext6 \
    libxrender1 \
    tesseract-ocr \
    tesseract-ocr-jpn \
    libtesseract-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ソースコードをコピー
COPY . .

CMD ["tail", "-f", "/dev/null"]
