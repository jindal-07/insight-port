FROM python:3.12-slim

WORKDIR /app

# Install dependencies first so Docker layer caching skips this on code changes
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code
COPY server.py facebook_export.py build_pivot_sheet.py ./
COPY frontend ./frontend

# HF Spaces runs containers as a non-root user with uid 1000
RUN useradd -m -u 1000 user
USER user
ENV HOME=/home/user \
    PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8

EXPOSE 7860

# Single worker is REQUIRED: export state (running job, log, in-memory ZIP)
# lives in the process. Threads handle concurrent polling/downloads.
# Binds to $PORT when the platform injects one (Render), else 7860 (HF Spaces).
CMD gunicorn -w 1 --threads 8 --timeout 120 -b 0.0.0.0:${PORT:-7860} server:app
