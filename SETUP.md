# Publisher QBR Setup

This setup guide covers the active local Publisher QBR workflow:

```text
Chrome extension -> local Publisher QBR runner -> qbr-pptx-service
```

## Prerequisites

- Node.js 22 or newer.
- Chrome or another Chromium browser that can load unpacked extensions.
- Access to TD admin credentials or an admin bearer token for publisher impersonation.
- The adjacent `qbr-pptx-service/` service available locally.

## 1. Start qbr-pptx-service

```powershell
cd qbr-pptx-service
npm install
npm start
```

Verify:

```powershell
curl http://127.0.0.1:3010/health
```

The runner sends generated deck payloads to:

```text
POST http://127.0.0.1:3010/generate
```

## 2. Start the Publisher QBR runner

```powershell
cd publisher-qbr-local-runner
npm install
npm start
```

Verify:

```powershell
curl http://127.0.0.1:3020/health
```

Active webhook:

```text
http://127.0.0.1:3020/webhook-local/publisher-qbr-v5-competitor-weekly-chart-20260505
```

## 3. Load the Chrome extension

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Load unpacked extension from `publisher-qbr-chrome-extension-prototype/`.
4. In the extension, use the local runner webhook URL above.
5. Save TD connection settings.
6. Impersonate the primary publisher.
7. Add up to four comparison publishers if needed.
8. Submit the Publisher QBR request.

The extension stores TD tokens only in the Chrome service worker session and includes them in the request sent to the local runner.

## Environment

Create `.env` from `.env.example` when using Docker helpers or shell-provided defaults:

```powershell
Copy-Item .env.example .env
```

Important values:

- `PUBLISHER_QBR_API_KEY`
- `PUBLISHER_QBR_PPTX_SERVICE_URL`
- `QBR_PPTX_SERVICE_URL`
- `PUBLISHER_QBR_LOCAL_RUNNER_PORT`
- `PUBLISHER_QBR_AGENT_MODE`
- `OPENAI_API_KEY`
- `PUBLISHER_QBR_DEBUG_DIR`

## Historical Material

n8n workflows and Cloudflare tunnel instructions are retired. They are not active setup steps. Historical workflow exports are archived under `workflows/archive/n8n-retired-2026-06/`.
