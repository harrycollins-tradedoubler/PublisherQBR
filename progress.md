# Publisher QBR Progress

## Current Runtime

Active:

```text
Chrome extension -> local Publisher QBR runner -> qbr-pptx-service
```

Retired from active runtime:

- n8n workflow execution
- Cloudflare tunnel routing
- `publisher-qbr-service` / `publisher-qbr-service_donotuse`

## Current Status

- Chrome extension source is active in `publisher-qbr-chrome-extension-prototype/`.
- Local runner source is active in `publisher-qbr-local-runner/`.
- PowerPoint generation is handled by adjacent `qbr-pptx-service/`.
- Old n8n workflow exports are archived under `workflows/archive/n8n-retired-2026-06/`.

## Recent Cleanup

- Removed stale unrelated training-project and agent-specific documentation from active docs.
- Reframed backend/frontend app-hub modules as legacy compatibility, not the active Publisher QBR flow.
- Documented local runner webhook on `127.0.0.1:3020`.
- Documented `qbr-pptx-service` `/generate` dependency on port `3010`.
- Removed tracked generated extension ZIP and retired Publisher-local PPTX service files.

## Open Follow-Ups

- Keep extension and runner tests passing as the request payload evolves.
- Keep stale n8n/Cloudflare references confined to the archive or explicit compatibility code.

