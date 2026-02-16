FROM python:3.12-slim

# LibreOffice for PPTX -> PDF conversion
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    libreoffice-writer \
    libreoffice-calc \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV PYTHONUNBUFFERED=1

CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-10000}"]

