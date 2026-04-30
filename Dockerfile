FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-core \
    libreoffice-calc \
    libreoffice-writer \
    libxinerama1 \
    libxrandr2 \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

# pango/cairo para weasyprint (fallback si libreoffice falla)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 libharfbuzz0b libpangoft2-1.0-0 \
    libpangocairo-1.0-0 libgdk-pixbuf-2.0-0 libcairo2 shared-mime-info \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD uvicorn eecc_server:app --host 0.0.0.0 --port ${PORT:-10000}
