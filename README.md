# Publisher QBR

Publisher QBR is the Tradedoubler workspace for creating and operating the Publisher QBR agent experience.

The current application is split into:

- `frontend/` - Vite, React, and TypeScript UI for the agent hub and QBR request form.
- `backend/` - FastAPI service for agent routing, TD auth helpers, QBR job status, and report downloads.
- `n8n-sync/` - version-controlled n8n workflow JSON used by the Publisher QBR automation.
- `publisher-qbr-service/` - required Publisher QBR PowerPoint generation service.
- `publisher-qbr/` - placeholder/project notes for Publisher QBR-specific assets.
- `scripts/` - local helper scripts for starting and stopping the active backend/frontend services.

## Repository Boundary

This repository is for Publisher QBR only.

Do not add, edit, stage, or commit Advertiser QBR, Agentic RAG Masterclass, or unrelated service code in this repository. In particular, `qbr-pptx-service/` is not part of this Publisher QBR repo.

The external folder `C:\Users\harcol\Workflows\qbr-pptx-service` may be used as read-only reference material for design patterns, but it must never be modified, staged, or committed from this repository.

If a local folder named `ai-agent-agenthub/` exists inside this workspace, treat it as a separate project. Do not read from it, modify it, stage it, or use it as a fallback Git repository.

Do not remove `publisher-qbr-service/`. It is the Publisher QBR presentation service used to generate editable PowerPoint reports and is separate from the advertiser `qbr-pptx-service`.

## Local Development

Backend:

```powershell
cd backend
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Frontend:

```powershell
cd frontend
npm install
npm run dev
```

Publisher PowerPoint service:

```powershell
cd publisher-qbr-service
npm install
npm start
```

Default local URLs:

- Backend: `http://localhost:8000`
- Frontend: `http://localhost:5173`
- Publisher PowerPoint service: `http://localhost:3010`

## Configuration

The backend reads environment variables from `backend/.env`.

Important settings:

- `DATABASE_URL` or Neon API settings for backend persistence.
- `QBR_AGENT_WEBHOOK_URL` for the Publisher QBR n8n webhook.
- Tradedoubler URL settings used by the TD auth routes.

Environment files are intentionally ignored by Git.

## Working Safely

Before large cleanup or repo synchronization work:

```powershell
git status --short --branch
git branch -vv
```

If local work exists, create a backup branch before resetting or cleaning:

```powershell
git switch -c backup/local-work-before-cleanup
git add -A
git commit -m "Backup local work before cleanup"
```

Only use destructive commands such as `git reset --hard` or `git clean -fd` after the work is backed up.

## Checks

Frontend checks:

```powershell
cd frontend
npm run build
```

Backend smoke check:

```powershell
cd backend
uvicorn app.main:app --reload
```
