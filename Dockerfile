# KPMG Capital Analysis Statement Generator — FastAPI backend + static React frontend
# Build:  docker build -t capital-statement-kpmg .
# Run:    docker run -p 8080:8080 -e GEMINI_API_KEY=xxx capital-statement-kpmg
#
# Designed for Google Cloud Run / Cloud Build: listens on $PORT (Cloud Run
# injects this; defaults to 8080 for local `docker run`).

FROM python:3.11-slim

# Prevent .pyc files and force stdout/stderr to be unbuffered so logs show up
# immediately in Cloud Logging.
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Install Python dependencies first so this layer is cached across builds
# that only change application code.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Application code (see .dockerignore for what's excluded — docs, media,
# node_modules, local venv, etc. are not needed to run the service).
COPY . .

# Run as a non-root user.
RUN useradd --create-home --shell /bin/bash appuser \
    && chown -R appuser:appuser /app
USER appuser

# Cloud Run sets $PORT at runtime; default to 8080 for local `docker run`.
ENV PORT=8080
EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD python -c "import os,urllib.request; urllib.request.urlopen(f'http://127.0.0.1:{os.environ.get(\"PORT\",\"8080\")}/api/health').read()" || exit 1

# Shell form so $PORT is expanded at container start.
CMD uvicorn api:app --host 0.0.0.0 --port ${PORT}
