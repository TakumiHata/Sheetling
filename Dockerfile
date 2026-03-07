FROM python:3.11-slim

# 日本語フォント（PDF画像化時のフォールバック用）
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 依存関係のインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ソースコードをコピー
COPY . .

CMD ["tail", "-f", "/dev/null"]
