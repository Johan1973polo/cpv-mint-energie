FROM python:3.13-slim

RUN apt-get update && apt-get install -y \
    libreoffice-writer \
    fonts-liberation \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p output uploads instance static

EXPOSE 10000

CMD ["gunicorn", "app_fusion:app", "--bind", "0.0.0.0:10000", "--workers", "1", "--timeout", "120"]
