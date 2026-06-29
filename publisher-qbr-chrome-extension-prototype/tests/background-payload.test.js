const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function jsonResponse(data) {
  return {
    ok: true,
    status: 200,
    async text() {
      return JSON.stringify(data);
    }
  };
}

function loadBackground(fetchImpl) {
  const code = fs.readFileSync(path.join(__dirname, "..", "background.js"), "utf8");
  const context = {
    chrome: {
      action: { onClicked: { addListener() {} } },
      tabs: { create() {} },
      runtime: {
        getURL(value) {
          return value;
        },
        onMessage: { addListener() {} }
      }
    },
    AbortController,
    clearTimeout,
    console,
    crypto: {
      randomUUID() {
        return "test-thread-id";
      }
    },
    fetch: fetchImpl,
    setTimeout
  };

  vm.createContext(context);
  vm.runInContext(code, context);
  return context;
}

async function run() {
  const calls = [];
  let webhookBody = null;
  const context = loadBackground(async (url, init) => {
    calls.push({ url, init });

    if (String(url).includes("/admin/impersonate")) {
      const username = new URL(String(url)).searchParams.get("username");
      return jsonResponse({ access_token: `token-for-${username}` });
    }

    if (String(url) === "https://example.test/webhook") {
      webhookBody = JSON.parse(init.body);
      return jsonResponse({ ok: true, reportUrl: "https://example.test/report.pptx" });
    }

    throw new Error(`Unexpected URL ${url}`);
  });

  const response = await context.handleMessage({
    type: "SUBMIT_PUBLISHER_QBR",
    bearerToken: "admin-token",
    cfg: {
      qbrWebhookUrl: "https://example.test/webhook"
    },
    payload: {
      clientUsername: "primary-publisher",
      sourceID: "100",
      siteID: "100",
      comparisonPublishers: [1, 2, 3, 4].map((index) => ({
        label: `Publisher ${index}`,
        clientUsername: `comparison-${index}`,
        sourceID: `20${index}`,
        siteID: `20${index}`
      }))
    }
  });

  assert.equal(response.ok, true);
  assert.equal(calls.filter((call) => String(call.url).includes("/admin/impersonate")).length, 5);

  const payload = webhookBody.payload;
  assert.equal(payload.td_tokens.impersonate_access_token, "token-for-primary-publisher");
  assert.equal(payload.publisherSessions.length, 5);
  assert.equal(webhookBody.publisherSessions.length, 5);
  assert.deepEqual(Object.keys(payload.tdTokensByPublisher), [
    "primary-publisher",
    "comparison-1",
    "comparison-2",
    "comparison-3",
    "comparison-4"
  ]);

  assert.deepEqual(
    payload.competitorPublishers.map((publisher) => publisher.td_tokens.impersonate_access_token),
    ["token-for-comparison-1", "token-for-comparison-2", "token-for-comparison-3", "token-for-comparison-4"]
  );
  assert.deepEqual(
    payload.comparisonPublishers.map((publisher) => publisher.td_tokens.impersonate_access_token),
    ["token-for-comparison-1", "token-for-comparison-2", "token-for-comparison-3", "token-for-comparison-4"]
  );
  assert.deepEqual(
    payload.competitors.map((publisher) => publisher.td_tokens.impersonate_access_token),
    ["token-for-comparison-1", "token-for-comparison-2", "token-for-comparison-3", "token-for-comparison-4"]
  );
}

run()
  .then(() => {
    console.log("background payload token routing test passed");
  })
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
