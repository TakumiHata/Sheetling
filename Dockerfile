FROM python:3.11-slim

# システムの依存関係をインストール
RUN apt-get update && apt-get install -y \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 依存関係のインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ソースコードをコピー（開発時はマウントされるがパスとして必要）
COPY . .

# アプリケーションの実行（デフォルトでは tty で待機させるため空にするか tail などを指定）
CMD ["tail", "-f", "/dev/null"]
