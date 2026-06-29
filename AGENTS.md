# AI Agent Instructions

## Repository Boundary

Treat this folder as the Publisher QBR repository.

Active Publisher QBR runtime:

```text
Chrome extension -> local Publisher QBR runner -> qbr-pptx-service
```

`qbr-pptx-service/` is the active adjacent PPTX generation dependency. Inspect it only when needed to confirm its public runtime contract. Do not copy its implementation into Publisher QBR and do not modify it unless the user explicitly asks to work on that service.

Do not read, modify, stage, commit, or use unrelated project paths as fallback context. If `ai-agent-agenthub/` exists, leave it alone.

If Git commands fail because `.git` metadata is missing or unexpected, stop and report the state instead of operating on a nested project.

## Workflow

- Keep changes scoped to Publisher QBR unless the user explicitly expands scope.
- Preserve unrelated user changes.
- Do not stage or commit unless explicitly asked.
- Prefer existing project patterns in `publisher-qbr-chrome-extension-prototype/`, `publisher-qbr-local-runner/`, `backend/`, `frontend/`, and `scripts/`.
- For JavaScript, TypeScript, React, and app changes, run the relevant build or test command before finishing when practical.

## Retired Systems

n8n and Cloudflare tunnels are historical only. Do not describe them as active setup, active deployment, default routing, or required local runtime.

Old workflow exports live under `workflows/archive/` for reference/recovery only. Do not edit archived workflow JSON unless the user explicitly asks to restore or inspect retired workflow material.

`publisher-qbr-service` and `publisher-qbr-service_donotuse` are retired/outdated Publisher-local PPTX service implementations and are not active runtime dependencies.
