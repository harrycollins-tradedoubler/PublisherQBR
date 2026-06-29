# Publisher QBR Product Requirements

## Overview

Publisher QBR generates editable Publisher performance review decks from TD publisher data. The active workflow starts in a Chrome extension, runs analysis locally in the Publisher QBR runner, and delegates PowerPoint generation to `qbr-pptx-service`.

## Active Runtime

```text
Chrome extension -> local Publisher QBR runner -> qbr-pptx-service
```

The active workflow does not require n8n, Cloudflare tunnels, hosted Publisher services, or the retired `publisher-qbr-service` implementation.

## Users

- Commercial or account teams preparing Publisher QBR decks.
- Operators testing Publisher QBR generation locally with TD impersonation.
- Engineers maintaining the local runner and extension request contract.

## Core Requirements

- The Chrome extension collects TD connection settings, primary publisher, site ID, optional comparison publishers, language, currency, and reporting period.
- The extension obtains TD impersonation tokens and sends the request to the local runner.
- The local runner accepts the extension-compatible webhook payload on `127.0.0.1:3020`.
- The runner fetches current and comparison TD data, builds Publisher QBR tables, produces deterministic narrative when no model key is configured, and posts the final deck payload to `qbr-pptx-service`.
- `qbr-pptx-service` returns a downloadable editable `.pptx` link or equivalent generated file response.

## Out Of Scope

- Reintroducing n8n as an active runtime.
- Reintroducing Cloudflare tunnels as active local setup.
- Maintaining the retired `publisher-qbr-service` implementation.
- Copying `qbr-pptx-service` code into this repository.

## Acceptance Criteria

- A user can load the extension and submit a Publisher QBR request to the local runner.
- The local runner can run without `OPENAI_API_KEY` in deterministic mode.
- The runner calls `qbr-pptx-service` on the configured `/generate` endpoint.
- Active docs do not present n8n, Cloudflare tunnels, or `publisher-qbr-service` as required runtime.
- Historical workflow exports are archived and clearly labeled as recovery-only.
