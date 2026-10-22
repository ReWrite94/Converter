FROM python:3.12-slim

ARG DEBIAN_FRONTEND=noninteractive
ENV PYTHONUNBUFFERED=1

# ── System packages ──────────────────────────────────────────────────────────
RUN apt-get update -o Acquire::Retries=5 && apt-get install -y -o Acquire::Retries=5 --no-install-recommends \
      libreoffice-writer libreoffice-impress libreoffice-calc libreoffice-core \
      fontconfig fonts-dejavu-core fonts-liberation curl wget \
  && apt-get clean && rm -rf /var/lib/apt/lists/*

# ── Python deps ─────────────────────────────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── App ─────────────────────────────────────────────────────────────────────
WORKDIR /srv
COPY app.py security.py ./

# Run as non-root: app code is owned by root and read-only for the service.
# DOC_PDF_API_TOKEN must be set at startup (docker run -e …).
RUN useradd --create-home --uid 10001 app
USER app

EXPOSE 8080
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8080"]
