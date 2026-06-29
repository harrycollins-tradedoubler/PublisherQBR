# Publisher QBR Local Runner

Active local Node.js implementation of the Publisher QBR workflow.

Runtime path:

```text
Chrome extension -> local Publisher QBR runner -> qbr-pptx-service
```

The runner replaces the retired n8n workflow for active local Publisher QBR generation. It keeps the old webhook request/response shape only as a compatibility contract with the extension and tests.

## Run Locally

Start `qbr-pptx-service` first on port `3010`, then run:

```powershell
cd publisher-qbr-local-runner
npm install
npm start
```

Runner health:

```text
GET http://127.0.0.1:3020/health
```

Extension webhook:

```text
POST http://127.0.0.1:3020/webhook-local/publisher-qbr-v5-competitor-weekly-chart-20260505
```

The runner posts final deck payloads to:

```text
http://127.0.0.1:3010/generate
```

Override with:

```powershell
$env:PUBLISHER_QBR_LOCAL_RUNNER_PORT = "3020"
$env:PUBLISHER_QBR_PPTX_SERVICE_URL = "http://127.0.0.1:3010"
$env:PUBLISHER_QBR_API_KEY = "your-local-generator-key"
$env:OPENAI_API_KEY = "your-openai-key"
```

`OPENAI_API_KEY` is optional. If it is not set, or if `PUBLISHER_QBR_AGENT_MODE=deterministic`, the runner builds deterministic Publisher QBR narrative from computed workflow tables.

## Docker Helper

The Docker runner expects `qbr-pptx-service` to be running on the host:

```powershell
docker compose up publisher-qbr-local-runner
```

Compose sets:

```text
PUBLISHER_QBR_PPTX_SERVICE_URL=http://host.docker.internal:3010
```

## Contract

The local route accepts the extension request shape:

- `message: "QBR_REQUEST {...}"`
- `payload`
- `qbr_payload`
- `td_tokens`
- `publisherSessions`

It returns a webhook-compatible JSON projection containing the generated PPTX URL when generation succeeds.

## Tests

```powershell
cd publisher-qbr-local-runner
npm test
```
