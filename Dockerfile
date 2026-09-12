FROM python:3.12-slim

# Noto fonts give PDF output proper Devanagari / Bengali / Tamil / Arabic glyphs
RUN apt-get update && apt-get install -y --no-install-recommends \
        fonts-noto-core fonts-noto-ui-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY bot.py .

ENV WORKDIR=/data \
    PDF_FONT_PATH=/usr/share/fonts/truetype/noto/NotoSansDevanagari-Regular.ttf \
    PYTHONUNBUFFERED=1
VOLUME ["/data"]

CMD ["python", "bot.py"]
