# Publisher QBR

Publisher QBR is the local workspace for generating Publisher QBR PowerPoint reports.

The active runtime is:

```text
Chrome extension -> local Publisher QBR runner -> qbr-pptx-service
```

n8n workflows and Cloudflare tunnels are retired from the active Publisher QBR runtime. They are kept only as historical recovery material under `workflows/archive/`.

## Active Modules

- `publisher-qbr-chrome-extension-prototype/` - active Chrome extension UI. It handles TD admin login, publisher impersonation, request assembly, and submission to the local runner.
- `publisher-qbr-local-runner/` - active local Node.js workflow runner. It accepts the extension webhook request on `127.0.0.1:3020`, fetches TD data, builds the Publisher QBR payload, and calls the PPTX service.
- `qbr-pptx-service/` - active adjacent PowerPoint generation service. It receives `/generate` calls from the runner. Treat it as its own service dependency and do not copy its implementation into this repo.
- `backend/` and `frontend/` - legacy app-hub UI/API modules retained for compatibility and reference. They are not required for the active extension-to-runner workflow.
- `workflows/archive/` - retired workflow exports retained only for reference/recovery.

`publisher-qbr-service/` and `publisher-qbr-service_donotuse/` are not active runtime services. The old Publisher-local PPTX implementation was retired because it no longer has the current generator code.

## Repository Boundary

Treat this folder as the Publisher QBR repository.

Do not use unrelated projects as fallback context. If `ai-agent-agenthub/` exists, leave it alone. `qbr-pptx-service/` is an active adjacent dependency for the local Publisher QBR flow, but it should be modified only when a task explicitly targets that service.

If Git commands fail because `.git` metadata is missing or unexpected, stop and report the state instead of operating on a nested project.

## Local Runtime

Start the active PPTX service from `qbr-pptx-service/` first:

```powershell
cd qbr-pptx-service
npm install
npm start
```

Expected service endpoints:

```text
GET  http://127.0.0.1:3010/health
POST http://127.0.0.1:3010/generate
```

Then start the Publisher runner:

```powershell
cd publisher-qbr-local-runner
npm install
npm start
```

Expected runner endpoints:

```text
GET  http://127.0.0.1:3020/health
POST http://127.0.0.1:3020/webhook-local/publisher-qbr-v5-competitor-weekly-chart-20260505
```

Load the extension:

1. Open `chrome://extensions`.
2. Enable Developer mode.
3. Load unpacked extension from `publisher-qbr-chrome-extension-prototype/`.
4. Confirm the Publisher QBR webhook URL is the local runner URL above.
5. Enter TD credentials or an admin bearer token, impersonate the primary publisher, add comparison publishers if needed, and submit.

## Docker Helper

Docker can run the local runner, but it does not run `qbr-pptx-service` from this repo. Start `qbr-pptx-service` separately on the host at `127.0.0.1:3010`, then run:

```powershell
docker compose up publisher-qbr-local-runner
```

The compose runner calls the host PPTX service through:

```text
http://host.docker.internal:3010
```

## Environment Variables

- `PUBLISHER_QBR_API_KEY` - shared key sent by the runner to `qbr-pptx-service`.
- `PUBLISHER_QBR_PPTX_SERVICE_URL` or `QBR_PPTX_SERVICE_URL` - PPTX service base URL. Local default: `http://127.0.0.1:3010`.
- `PUBLISHER_QBR_LOCAL_RUNNER_PORT` - runner port. Default: `3020`.
- `PUBLISHER_QBR_AGENT_MODE` - set to `deterministic` to skip live model narrative generation.
- `OPENAI_API_KEY` - optional. Enables live model narrative generation when deterministic mode is not set.
- `PUBLISHER_QBR_DEBUG_DIR` - optional sanitized runner debug artifact directory.

Backend-only legacy settings such as `QBR_AGENT_WEBHOOK_URL` are retained for the old app-hub compatibility flow and are not required for the active extension workflow.

## Validation

Run the active checks from the repository root:

```powershell
cd publisher-qbr-local-runner
npm test
```

```powershell
node publisher-qbr-chrome-extension-prototype/tests/background-payload.test.js
```

Optional compatibility checks:

```powershell
cd frontend
npm run build
```

```powershell
cd backend
python -m unittest discover tests
```

## Retired Systems

n8n, Cloudflare tunnels, old remote webhook URLs, and the retired Publisher-local PPTX service are not active setup, deployment, routing, or runtime requirements. Old workflow exports live under `workflows/archive/n8n-retired-2026-06/` for reference only.
