FROM python:3.11-slim

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY ocr_correct.py diagnose_chunks.py verify_chunks.py ./

# Create default directories
RUN mkdir -p data output

ENV PYTHONUNBUFFERED=1

ENTRYPOINT ["python", "ocr_correct.py"]
CMD ["--help"]
