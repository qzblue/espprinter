# 使用輕量級 Python 3.11 映像
FROM python:3.11-slim

# 設定工作目錄
WORKDIR /app

# 安裝系統依賴 (如果需要的話，但這裏主要是為了 timezone)
# 設置時區為 澳門/台北 時間 (Asia/Macau)
RUN apt-get update && apt-get install -y tzdata && \
    ln -sf /usr/share/zoneinfo/Asia/Macau /etc/localtime && \
    echo "Asia/Macau" > /etc/timezone && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# 複製依賴文件
COPY requirements.txt .

# 安裝 Python 依賴
RUN pip install --no-cache-dir -r requirements.txt

# 複製應用程序代碼
COPY . .

# 創建 export 目錄
RUN mkdir -p /app/exports

# 暴露端口 5000
EXPOSE 5000

# 啟動應用
CMD ["gunicorn", "-w", "1", "-k", "gthread", "--threads", "8", "-b", "0.0.0.0:5000", "webapp:app"]


