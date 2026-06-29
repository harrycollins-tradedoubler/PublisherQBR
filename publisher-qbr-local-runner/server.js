const http = require("node:http");

const { runPublisherQbrWorkflow } = require("./lib/publisherQbrRunner");
const { redactSensitive } = require("./lib/publisherQbrAgent");

const PORT = Number(process.env.PORT || process.env.PUBLISHER_QBR_LOCAL_RUNNER_PORT || 3020);
const HOST = process.env.HOST || process.env.PUBLISHER_QBR_LOCAL_RUNNER_HOST || "127.0.0.1";
const LOCAL_WEBHOOK_PATH = "/webhook-local/publisher-qbr-v5-competitor-weekly-chart-20260505";

function json(res, status, payload) {
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Accept, Authorization"
  });
  res.end(JSON.stringify(payload));
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

function isLoopbackAddress(address) {
  const value = String(address || "").toLowerCase();
  return value === "127.0.0.1"
    || value === "::1"
    || value === "::ffff:127.0.0.1"
    || value === "localhost";
}

function envFlag(name) {
  return /^(1|true|yes|on)$/i.test(String(process.env[name] || ""));
}

function createLocalRequestGuard(options = {}) {
  const allowRemoteLocalWebhook = options.allowRemoteLocalWebhook === undefined
    ? envFlag("PUBLISHER_QBR_ALLOW_REMOTE_LOCAL_WEBHOOK")
    : Boolean(options.allowRemoteLocalWebhook);
  return (req) => allowRemoteLocalWebhook || isLoopbackAddress(req.socket.remoteAddress);
}

function redactedErrorMessage(error) {
  const message = error instanceof Error ? error.message : String(error);
  return String(redactSensitive({ message }).message || "Unknown error");
}

function createServer(options = {}) {
  const runner = options.runner || runPublisherQbrWorkflow;
  const isLocalRequest = options.isLocalRequest || createLocalRequestGuard({
    allowRemoteLocalWebhook: options.allowRemoteLocalWebhook
  });

  return http.createServer(async (req, res) => {
    const requestUrl = new URL(req.url || "/", "http://localhost");

    if (req.method === "OPTIONS") {
      json(res, 204, {});
      return;
    }

    if (req.method === "GET" && requestUrl.pathname === "/health") {
      json(res, 200, { ok: true, service: "publisher-qbr-local-runner" });
      return;
    }

    if (requestUrl.pathname === LOCAL_WEBHOOK_PATH) {
      if (req.method !== "POST") {
        json(res, 405, { success: false, message: "Method not allowed." });
        return;
      }
      if (!isLocalRequest(req)) {
        json(res, 403, { success: false, message: "Local webhook is localhost-only." });
        return;
      }

      try {
        const body = await readBody(req);
        const result = await runner({ body, req });
        json(res, 200, result);
      } catch (error) {
        const detail = redactedErrorMessage(error);
        console.error("publisher_qbr_local_runner_error", redactSensitive({
          message: detail,
          stack: error instanceof Error ? error.stack : ""
        }));
        json(res, 500, {
          success: false,
          provider: "publisher-qbr-local-runner",
          message: `Publisher QBR local workflow failed: ${detail}`,
          error: detail
        });
      }
      return;
    }

    json(res, 404, { success: false, message: "Not found." });
  });
}

if (require.main === module) {
  const server = createServer();
  server.listen(PORT, HOST, () => {
    console.log(`Publisher QBR local runner listening on http://${HOST}:${PORT}`);
    console.log(`Local webhook: http://${HOST}:${PORT}${LOCAL_WEBHOOK_PATH}`);
  });
}

module.exports = {
  createServer,
  createLocalRequestGuard,
  isLoopbackAddress,
  LOCAL_WEBHOOK_PATH
};
