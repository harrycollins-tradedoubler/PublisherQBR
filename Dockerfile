# syntax=docker/dockerfile:1
FROM python:3.12-slim AS backend
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1
WORKDIR /app/backend
COPY backend/requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY backend/app ./app
EXPOSE 8008
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8008/api/health', timeout=3)"
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8008"]

FROM node:22-alpine AS publisher-qbr-local-runner
ENV NODE_ENV=production \
    PORT=3020 \
    PUBLISHER_QBR_LOCAL_RUNNER_HOST=0.0.0.0 \
    PUBLISHER_QBR_ALLOW_REMOTE_LOCAL_WEBHOOK=1 \
    PUBLISHER_QBR_PPTX_SERVICE_URL=http://host.docker.internal:3010 \
    PUBLISHER_QBR_DEBUG_DIR=/app/publisher-qbr-local-runner/debug-runs
WORKDIR /app/publisher-qbr-local-runner
COPY publisher-qbr-local-runner/package*.json ./
COPY publisher-qbr-local-runner/ ./
RUN mkdir -p /app/publisher-qbr-local-runner/debug-runs
EXPOSE 3020
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:3020/health').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"
CMD ["npm", "start"]
