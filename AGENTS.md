# AI Agent Instructions

## Repository Boundary

Treat this folder as the Publisher QBR repository only.

Do not read, modify, stage, commit, or use these paths as part of Publisher QBR work:

- `ai-agent-agenthub/`
- `qbr-pptx-service/`

Those paths belong to separate Advertiser QBR / Agentic RAG Masterclass work and must not be used as fallback context for this repository.

The external folder `C:\Users\harcol\Workflows\qbr-pptx-service` may be inspected only as read-only reference material for design patterns when the user explicitly asks for that context. Never modify files in that folder, copy implementation into Publisher QBR, stage it, or commit it from this repository.

If Git commands fail because `.git` metadata is missing or unexpected, stop and report the state instead of operating on a nested project.

## Workflow

- Keep changes scoped to Publisher QBR.
- Prefer existing project patterns in `backend/`, `frontend/`, `n8n-sync/`, and `scripts/`.
- Before deleting or resetting anything, create a backup branch or confirm that the work is already safely committed.
- For JavaScript, TypeScript, React, and app changes, run the relevant build or test command before finishing when practical.
- For n8n workflow edits, use the `n8nac-skills` schema and validation workflow before changing node JSON.
