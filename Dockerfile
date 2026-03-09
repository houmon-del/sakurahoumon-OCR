FROM python:3.11

WORKDIR /app

# Install PyTorch CPU-only first (smaller than GPU version)
RUN pip install --no-cache-dir \
    torch torchvision --index-url https://download.pytorch.org/whl/cpu

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Pre-download all 4 yomitoku AI models into the image
RUN python -m yomitoku.cli.download_model

# Copy application code
COPY . .

# Create uploads directory
RUN mkdir -p uploads

# Thread control (prevent SIGABRT on Cloud Run)
ENV OMP_NUM_THREADS=2
ENV MKL_NUM_THREADS=2
ENV TORCH_NUM_THREADS=2
ENV ORT_NUM_THREADS=2
ENV ONNXRUNTIME_NUM_THREADS=2
ENV OPENBLAS_NUM_THREADS=2
ENV PYTHONUNBUFFERED=1

# Cloud Run sets PORT env var (default 8080)
ENV PORT=8080

EXPOSE ${PORT}

# Use gunicorn for production
CMD exec gunicorn --bind 0.0.0.0:${PORT} --workers 1 --threads 2 --timeout 300 app:app
