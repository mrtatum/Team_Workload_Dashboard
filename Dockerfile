# syntax=docker/dockerfile:1
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 PORT=8501

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY server.py onedrive_source.py holidays.txt ./
COPY static ./static

RUN useradd -m appuser && chown -R appuser /app
USER appuser

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request,sys; sys.exit(0) if urllib.request.urlopen('http://localhost:8501/api/health').status==200 else sys.exit(1)"

# 2 workers, 4 threads — plenty for an internal dashboard; 120s timeout for OneDrive pulls.
CMD ["gunicorn", "-b", "0.0.0.0:8501", "-w", "2", "--threads", "4", "-t", "120", "server:app"]
