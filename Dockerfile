FROM python:3.11-slim

# Install LibreOffice Writer (lighter than full libreoffice) for .doc -> .docx
# conversion in parsers/docx_parser.py. fonts-liberation provides decent
# fallback fonts so soffice produces readable output.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first to benefit from Docker layer caching
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source
COPY . .

# Render injects $PORT; app.py reads it. Expose a sensible default.
ENV PORT=10000
EXPOSE 10000

CMD ["python", "app.py"]
