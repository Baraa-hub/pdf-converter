FROM python:3.11-slim
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-ara \
    tesseract-ocr-eng \
    poppler-utils \
    libmupdf-dev \
    mupdf \
    mupdf-tools \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8080
CMD gunicorn --bind 0.0.0.0:8080 --timeout 300 --workers 4 --threads 2 app:app
