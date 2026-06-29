const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const { createServer, createLocalRequestGuard, LOCAL_WEBHOOK_PATH } = require("../server");
const {
  createPublisherQbrAgent,
  parseAgentJson,
  redactSensitive
} = require("../lib/publisherQbrAgent");
const {
  buildMetricsUrl,
  extractIncomingPayload,
  normalizeWorkflowInput,
  assertPublisherPptxService,
  projectN8nResponse,
  runPublisherQbrWorkflow
} = require("../lib/publisherQbrRunner");

function listen(server) {
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      resolve(`http://127.0.0.1:${address.port}`);
    });
  });
}

function close(server) {
  return new Promise((resolve, reject) => {
    server.close((error) => (error ? reject(error) : resolve()));
  });
}

async function readJson(response) {
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

function extensionBody(payload = {}) {
  return {
    message: `QBR_REQUEST ${JSON.stringify(payload)}`,
    thread_id: "thread-1",
    payload,
    qbr_payload: payload,
    td_tokens: {
      user_access_token: "primary-user-token",
      impersonate_access_token: "primary-impersonate-token"
    },
    publisherSessions: []
  };
}

function samplePayload(overrides = {}) {
  return {
    type: "PUBLISHER_QBR_REQUEST",
    clientUsername: "publisher-a",
    sourceID: "123",
    fromDate: "20260101",
    toDate: "20260331",
    languageCode: "EN",
    currencyCode: "EUR",
    publisherExportEndpoint: "https://td.example/statistics/export",
    digitalWalletEndpoint: "https://td.example/payments/digitalwallets",
    programStatusEndpoint: "https://td.example/programs",
    comparisonPublishers: [
      {
        label: "Competitor 1",
        clientUsername: "publisher-b",
        sourceID: "456",
        td_tokens: {
          impersonate_access_token: "competitor-impersonate-token"
        }
      }
    ],
    ...overrides
  };
}

test("local webhook route returns the n8n-compatible response shape", async () => {
  const server = createServer({
    isLocalRequest: () => true,
    runner: async ({ body }) => {
      assert.equal(body.payload.clientUsername, "publisher-a");
      return projectN8nResponse({
        provider: "publisher-qbr-pptx",
        message: "Editable Publisher QBR PowerPoint generated successfully.",
        presentation_id: "deck-1",
        pptx_url: "http://127.0.0.1:3010/files/deck.pptx",
        file_name: "deck.pptx",
        slide_count: 18,
        theme: "TD"
      });
    }
  });
  const root = await listen(server);

  try {
    const response = await fetch(`${root}${LOCAL_WEBHOOK_PATH}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(extensionBody(samplePayload()))
    });
    const body = await readJson(response);

    assert.equal(response.status, 200);
    assert.deepEqual(Object.keys(body), [
      "success",
      "provider",
      "message",
      "generation_status",
      "generation_id",
      "presentation_id",
      "presentation_url",
      "edit_url",
      "pptx_url",
      "gap_analysis_report_url",
      "gap_analysis_report_file_name",
      "file_name",
      "theme",
      "slide_count",
      "error"
    ]);
    assert.equal(body.success, true);
    assert.equal(body.pptx_url, "http://127.0.0.1:3010/files/deck.pptx");
  } finally {
    await close(server);
  }
});

test("local webhook rejects non-loopback callers", async () => {
  const server = createServer({
    isLocalRequest: () => false,
    runner: async () => {
      throw new Error("runner should not be called");
    }
  });
  const root = await listen(server);

  try {
    const response = await fetch(`${root}${LOCAL_WEBHOOK_PATH}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(extensionBody(samplePayload()))
    });
    const body = await readJson(response);

    assert.equal(response.status, 403);
    assert.equal(body.success, false);
    assert.match(body.message, /localhost-only/i);
  } finally {
    await close(server);
  }
});


test("local webhook guard can be relaxed for Docker loopback port publishing", () => {
  const req = { socket: { remoteAddress: "172.18.0.1" } };

  assert.equal(createLocalRequestGuard()(req), false);
  assert.equal(createLocalRequestGuard({ allowRemoteLocalWebhook: true })(req), true);
});
test("local webhook exposes redacted failure details", async () => {
  const server = createServer({
    isLocalRequest: () => true,
    runner: async () => {
      throw new Error("OpenAI failed with token sk-secret-value");
    }
  });
  const root = await listen(server);

  try {
    const response = await fetch(`${root}${LOCAL_WEBHOOK_PATH}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(extensionBody(samplePayload()))
    });
    const body = await readJson(response);

    assert.equal(response.status, 500);
    assert.equal(body.success, false);
    assert.match(body.message, /Publisher QBR local workflow failed: OpenAI failed/);
    assert.equal(JSON.stringify(body).includes("sk-secret-value"), false);
  } finally {
    await close(server);
  }
});

test("PPTX service health guard rejects the retired Publisher-local service", async () => {
  await assert.rejects(
    () => assertPublisherPptxService("http://127.0.0.1:3010", {
      fetch: async () => new Response(JSON.stringify({ ok: true, service: "publisher-qbr-service" }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      })
    }),
    /expected publisher-qbr-pptx-service/i
  );
});
test("payload extraction supports extension and QBR_REQUEST shapes", () => {
  const payload = samplePayload();
  assert.deepEqual(extractIncomingPayload({ payload }).payload, payload);
  assert.deepEqual(extractIncomingPayload({ qbr_payload: payload }).payload, payload);
  assert.deepEqual(extractIncomingPayload({ body: { payload } }).payload, payload);
  assert.deepEqual(
    extractIncomingPayload({ message: `QBR_REQUEST ${JSON.stringify(payload)}` }).payload,
    payload
  );
  assert.deepEqual(
    extractIncomingPayload({ message: `QBR_REQUEST ${JSON.stringify(payload)}`, td_tokens: { impersonate_access_token: "tok" } }).tdTokens,
    { impersonate_access_token: "tok" }
  );
});

test("agent validates strict model JSON and renders markdown analysis", async () => {
  const calls = [];
  const agent = createPublisherQbrAgent({
    modelClient: {
      complete: async (request) => {
        calls.push(request);
        return {
          content: JSON.stringify({
            reportingPeriod: ["Current period EUR values compared YoY."],
            kpiHighlights: ["Publisher Commission EUR 10 vs EUR 8, +25% YoY."],
            programLevelAnalysis: [{ title: "Program A growth", description: "Program A increased commission from EUR 5 to EUR 8." }],
            moversAndShakers: ["Program A publisher commission increased YoY."],
            risksAndDependencies: [{ title: "Commission concentration", description: "Program A contributes most commission." }]
          })
        };
      }
    },
    logger: { info() {}, warn() {}, error() {} }
  });

  const output = await agent.run({
    dataForAI: "{\"tables\":{}}",
    payload: { qbrFocus: "General performance review", languageName: "English" }
  });

  assert.equal(calls.length, 1);
  assert.equal(calls[0].model, "gpt-5.4-mini");
  assert.match(calls[0].system, /Tradedoubler publisher reporting assistant/);
  assert.match(calls[0].user, /Publisher Performance/);
  assert.match(output.markdown, /## Publisher Performance/);
  assert.match(output.markdown, /### KPI Highlights/);
});

test("agent rejects malformed or schema-invalid model output", async () => {
  const malformed = createPublisherQbrAgent({
    modelClient: { complete: async () => ({ content: "not json" }) },
    fallbackOnInvalidOutput: false,
    logger: { info() {}, warn() {}, error() {} }
  });
  await assert.rejects(() => malformed.run({ dataForAI: "{}", payload: {} }), /valid JSON/i);

  const invalidShape = createPublisherQbrAgent({
    modelClient: { complete: async () => ({ content: JSON.stringify({ kpiHighlights: [] }) }) },
    fallbackOnInvalidOutput: false,
    logger: { info() {}, warn() {}, error() {} }
  });
  await assert.rejects(() => invalidShape.run({ dataForAI: "{}", payload: {} }), /reportingPeriod/i);
});

test("agent falls back to deterministic output when live model returns invalid JSON", async () => {
  const warnings = [];
  const agent = createPublisherQbrAgent({
    modelClient: { complete: async () => ({ content: "## Publisher Performance\n- prose instead of JSON" }) },
    logger: {
      info() {},
      warn(event, details) {
        warnings.push({ event, details });
      },
      error() {}
    }
  });

  const output = await agent.run({
    dataForAI: JSON.stringify({
      tables: {
        programLevelBreakdown: [{ programName: "Primary Program", publisherCommission: 10 }]
      },
      diagnostics: { currentRows: 1, previousRows: 1, competitorPublishers: 0 }
    }),
    payload: {
      languageName: "English",
      reportingPeriod: "2026-01-01 to 2026-03-31",
      comparisonPeriod: "2025-01-01 to 2025-03-31"
    }
  });

  assert.equal(output.deterministic, true);
  assert.match(output.markdown, /Primary Program/);
  assert.equal(warnings.some((entry) => entry.event === "publisher_qbr_agent_invalid_output_fallback"), true);
});

test("agent can run deterministically without OPENAI_API_KEY", async () => {
  const previousKey = process.env.OPENAI_API_KEY;
  delete process.env.OPENAI_API_KEY;
  const agent = createPublisherQbrAgent({
    logger: { info() {}, warn() {}, error() {} }
  });

  try {
    const output = await agent.run({
      dataForAI: JSON.stringify({
        payload: {
          reportingPeriod: "2026-01-01 to 2026-03-31",
          comparisonPeriod: "2025-01-01 to 2025-03-31",
          currencyCode: "EUR"
        },
        tables: {
          kpiSummaryTable: [
            { Period: "Recent", Rows: "2" },
            { Period: "Previous", Rows: "1" }
          ],
          programLevelBreakdown: [
            { programName: "Primary Program", publisherCommission: 10 }
          ],
          riskDependenciesTable: [
            { Issue: "Concentration", Analysis: "One program dominates commission." }
          ]
        },
        diagnostics: {
          currentRows: 2,
          previousRows: 1,
          competitorPublishers: 1
        }
      }),
      payload: {
        languageName: "English",
        reportingPeriod: "2026-01-01 to 2026-03-31",
        comparisonPeriod: "2025-01-01 to 2025-03-31",
        qbrFocus: "General performance review"
      }
    });

    assert.match(output.markdown, /## Publisher Performance/);
    assert.match(output.markdown, /2026-01-01 to 2026-03-31/);
    assert.match(output.markdown, /Primary Program/);
    assert.equal(output.deterministic, true);
  } finally {
    if (previousKey) process.env.OPENAI_API_KEY = previousKey;
  }
});

test("agent enforces allowlisted tools, typed input, max iterations, and redacted logging", async () => {
  const logEntries = [];
  const agent = createPublisherQbrAgent({
    maxIterations: 2,
    tools: {
      echoMetric: {
        description: "Echo a metric",
        schema: {
          type: "object",
          required: ["metric"],
          properties: { metric: { type: "string" } }
        },
        execute: async ({ metric }) => ({ metric })
      }
    },
    modelClient: {
      complete: async (_request, context) => {
        if (context.iteration === 1) {
          return { toolCall: { name: "echoMetric", input: { metric: "commission" } } };
        }
        return {
          content: JSON.stringify({
            reportingPeriod: ["Period ok."],
            kpiHighlights: ["Highlight ok."],
            programLevelAnalysis: [],
            moversAndShakers: [],
            risksAndDependencies: []
          })
        };
      }
    },
    logger: {
      info(event, details) {
        logEntries.push({ event, details });
      },
      warn() {},
      error() {}
    }
  });

  const result = await agent.run({
    dataForAI: "{}",
    payload: {
      languageName: "English",
      td_tokens: { impersonate_access_token: "secret-token" }
    }
  });

  assert.equal(result.toolResults.length, 1);
  assert.equal(result.toolResults[0].result.metric, "commission");
  assert.equal(JSON.stringify(logEntries).includes("secret-token"), false);
  assert.equal(JSON.stringify(redactSensitive({ Authorization: "Bearer secret-token", nested: { apiKey: "abc" } })).includes("secret-token"), false);

  const badTool = createPublisherQbrAgent({
    tools: {},
    modelClient: { complete: async () => ({ toolCall: { name: "notAllowed", input: {} } }) },
    logger: { info() {}, warn() {}, error() {} }
  });
  await assert.rejects(() => badTool.run({ dataForAI: "{}", payload: {} }), /not allowlisted/i);

  const badInput = createPublisherQbrAgent({
    tools: {
      echoMetric: {
        schema: { type: "object", required: ["metric"], properties: { metric: { type: "string" } } },
        execute: async () => ({})
      }
    },
    modelClient: { complete: async () => ({ toolCall: { name: "echoMetric", input: { metric: 1 } } }) },
    logger: { info() {}, warn() {}, error() {} }
  });
  await assert.rejects(() => badInput.run({ dataForAI: "{}", payload: {} }), /metric/i);

  const looping = createPublisherQbrAgent({
    maxIterations: 1,
    tools: {
      echoMetric: {
        schema: { type: "object", required: ["metric"], properties: { metric: { type: "string" } } },
        execute: async () => ({})
      }
    },
    modelClient: { complete: async () => ({ toolCall: { name: "echoMetric", input: { metric: "x" } } }) },
    logger: { info() {}, warn() {}, error() {} }
  });
  await assert.rejects(() => looping.run({ dataForAI: "{}", payload: {} }), /maximum agent iterations/i);
});

test("runner uses TD auth headers, pagination, agent output, and PPTX projection", async () => {
  const calls = [];
  const fetchImpl = async (url, init = {}) => {
    calls.push({ url: String(url), init });
    const parsed = new URL(String(url));

    if (parsed.pathname.endsWith("/statistics/export")) {
      const offset = Number(parsed.searchParams.get("offset") || "0");
      const sourceId = parsed.searchParams.get("sourceId");
      return new Response(JSON.stringify({
        items: [
          {
            programId: sourceId === "456" ? "C1" : "P1",
            programName: sourceId === "456" ? "Competitor Program" : "Primary Program",
            clicks: offset === 0 ? 10 : 5,
            sales: 2,
            orderValue: 100,
            publisherCommission: 10
          }
        ],
        total: offset === 0 ? 2 : 2,
        limit: 1,
        offset
      }), { status: 200, headers: { "Content-Type": "application/json" } });
    }

    if (parsed.pathname.endsWith("/digitalwallets")) {
      return new Response(JSON.stringify({ items: [{ programId: "P1", amount: 3 }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    }

    if (parsed.pathname.endsWith("/programs")) {
      return new Response(JSON.stringify({ items: [{ id: "P1", name: "Primary Program", statusId: 3 }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    }

    if (parsed.pathname.endsWith("/health")) {
        return new Response(JSON.stringify({ ok: true, service: "publisher-qbr-pptx-service", provider: "publisher-qbr-pptx" }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      if (parsed.pathname.endsWith("/generate")) {
      const body = JSON.parse(init.body);
      assert.equal(body.analysisLevel, "publisher_program");
      assert.equal(JSON.stringify(body).includes("advertiser_qbr"), true);
      assert.equal(JSON.stringify(body).includes(":3011"), false);
      return new Response(JSON.stringify({
        success: true,
        provider: "publisher-qbr-pptx",
        message: "Generated.",
        presentation_id: "deck-123",
        pptx_url: "http://127.0.0.1:3010/files/deck.pptx",
        file_name: "deck.pptx",
        slide_count: 18,
        theme: "TD"
      }), { status: 200, headers: { "Content-Type": "application/json" } });
    }

    throw new Error(`Unexpected URL ${url}`);
  };

  const result = await runPublisherQbrWorkflow(extensionBody(samplePayload({
    publisherPptxApiUrl: "http://127.0.0.1:3010"
  })), {
    fetch: fetchImpl,
    modelClient: {
      complete: async () => ({
        content: JSON.stringify({
          reportingPeriod: ["2026 Q1 compared to 2025 Q1."],
          kpiHighlights: ["Commission EUR 10 vs EUR 8."],
          programLevelAnalysis: [{ title: "Primary Program", description: "Commission growth confirmed." }],
          moversAndShakers: ["Primary Program moved up."],
          risksAndDependencies: []
        })
      })
    },
    logger: { info() {}, warn() {}, error() {} }
  });

  assert.equal(result.success, true);
  assert.equal(result.pptx_url, "http://127.0.0.1:3010/files/deck.pptx");

  const primaryCalls = calls.filter((call) => call.url.includes("/statistics/export") && call.url.includes("sourceId=123"));
  const competitorCalls = calls.filter((call) => call.url.includes("/statistics/export") && call.url.includes("sourceId=456"));
  assert.ok(primaryCalls.length >= 2);
  assert.ok(competitorCalls.length >= 2);
  assert.equal(primaryCalls[0].init.headers.Authorization, "Bearer primary-impersonate-token");
  assert.equal(competitorCalls[0].init.headers.Authorization, "Bearer competitor-impersonate-token");
});

test("statistics export URLs match n8n defaults and do not force a one-row limit", () => {
  const normalized = normalizeWorkflowInput(extensionBody(samplePayload()));
  const url = buildMetricsUrl(
    normalized.payload,
    normalized.payload.fromDate,
    normalized.payload.toDate,
    normalized.payload.primarySourceId
  );

  assert.equal(url.searchParams.get("reportType"), "program");
  assert.equal(url.searchParams.get("intervalType"), "day");
  assert.equal(url.searchParams.get("sourceId"), "123");
  assert.equal(url.searchParams.has("limit"), false);
  assert.equal(url.searchParams.has("offset"), false);
});

test("runner can write sanitized debug artifacts for the final PPTX payload", async () => {
  const debugDir = await fs.mkdtemp(path.join(os.tmpdir(), "publisher-qbr-local-runner-"));
  const payload = samplePayload({
    publisherPptxApiUrl: "http://127.0.0.1:3010",
    td_tokens: { impersonate_access_token: "secret-primary-token" }
  });

  const result = await runPublisherQbrWorkflow(extensionBody(payload), {
    debugDir,
    fetch: async (url, init = {}) => {
      const parsed = new URL(String(url));
      if (parsed.pathname.endsWith("/health")) {
        return new Response(JSON.stringify({ ok: true, service: "publisher-qbr-pptx-service", provider: "publisher-qbr-pptx" }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      if (parsed.pathname.endsWith("/generate")) {
        return new Response(JSON.stringify({
          success: true,
          provider: "publisher-qbr-pptx",
          pptx_url: "http://127.0.0.1:3010/files/deck.pptx"
        }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      return new Response(JSON.stringify({ items: [{ programId: "P1", publisherCommission: 10 }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    },
    modelClient: {
      complete: async () => ({
        content: JSON.stringify({
          reportingPeriod: ["ok"],
          kpiHighlights: ["ok"],
          programLevelAnalysis: [],
          moversAndShakers: [],
          risksAndDependencies: []
        })
      })
    },
    logger: { info() {}, warn() {}, error() {} }
  });

  assert.equal(result.success, true);
  const files = await fs.readdir(debugDir);
  const debugFile = files.find((file) => file.endsWith(".json"));
  assert.ok(debugFile);
  const debugJson = JSON.parse(await fs.readFile(path.join(debugDir, debugFile), "utf8"));
  assert.equal(debugJson.finalPptxPayload.analysisLevel, "publisher_program");
  assert.equal(JSON.stringify(debugJson).includes("secret-primary-token"), false);

  await fs.rm(debugDir, { recursive: true, force: true });
});

test("runner builds publisher-native aggregate tables for the PPTX service", async () => {
  let finalPayload;
  const result = await runPublisherQbrWorkflow(extensionBody(samplePayload({
    publisherPptxApiUrl: "http://127.0.0.1:3010"
  })), {
    fetch: async (url, init = {}) => {
      const parsed = new URL(String(url));
      const sourceId = parsed.searchParams.get("sourceId");

      if (parsed.pathname.endsWith("/health")) {
        return new Response(JSON.stringify({ ok: true, service: "publisher-qbr-pptx-service", provider: "publisher-qbr-pptx" }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      if (parsed.pathname.endsWith("/generate")) {
        finalPayload = JSON.parse(init.body);
        return new Response(JSON.stringify({
          success: true,
          provider: "publisher-qbr-pptx",
          pptx_url: "http://127.0.0.1:3010/files/deck.pptx"
        }), { status: 200, headers: { "Content-Type": "application/json" } });
      }

      if (parsed.pathname.endsWith("/digitalwallets")) {
        return new Response(JSON.stringify({ items: [{ programId: "P1", programName: "Program One", amount: 5 }] }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      if (parsed.pathname.endsWith("/programs")) {
        return new Response(JSON.stringify({ items: [{ id: "P1", name: "Program One", statusId: 3 }] }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      const isPrevious = parsed.searchParams.get("fromDate") === "20250101";
      const rows = sourceId === "456"
        ? [
            { date: "2026-01-05T00:00:00.000+0100", programId: "P1", programName: "Program One", clicks: 15, sales: 3, orderValue: 60, commission: 12 },
            { date: "2026-01-12T00:00:00.000+0100", programId: "P3", programName: "Competitor Gap Program", clicks: 9, sales: 2, orderValue: 40, commission: 8 }
          ]
        : [
            { date: isPrevious ? "2025-01-06T00:00:00.000+0100" : "2026-01-05T00:00:00.000+0100", programId: "P1", programName: "Program One", clicks: isPrevious ? 10 : 20, sales: isPrevious ? 1 : 4, orderValue: isPrevious ? 25 : 100, commission: isPrevious ? 5 : 20 },
            { date: isPrevious ? "2025-01-13T00:00:00.000+0100" : "2026-01-12T00:00:00.000+0100", programId: "P2", programName: "Program Two", clicks: 7, sales: 0, orderValue: 0, commission: 0 }
          ];
      return new Response(JSON.stringify({ items: rows }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    },
    modelClient: {
      complete: async () => ({
        content: JSON.stringify({
          reportingPeriod: ["ok"],
          kpiHighlights: ["ok"],
          programLevelAnalysis: [],
          moversAndShakers: [],
          risksAndDependencies: []
        })
      })
    },
    logger: { info() {}, warn() {}, error() {} }
  });

  assert.equal(result.success, true);
  const varianceRow = finalPayload.publisherTables.kpiSummaryTable.find((row) => row.Period === "% Variance");
  assert.equal(varianceRow.Conversions, "+300.0%");
  assert.equal(varianceRow["Conversion Rate"], "+151.9%");
  assert.equal(varianceRow["Publisher Commission"], "+300.0%");
  assert.equal(varianceRow["Digital Wallet"], "0.0%");
  assert.equal(varianceRow["Total Earnings"], "+150.0%");
  assert.equal(finalPayload.publisherTables.kpiSummaryTable[0].Conversions, "4");
  assert.equal(finalPayload.publisherTables.kpiSummaryTable[0]["Publisher Commission"], "â‚¬20");
  assert.deepEqual(Object.keys(finalPayload.publisherTables.programLevelBreakdown[0]), [
    "Program ID",
    "Program Name",
    "Publisher Commission",
    "Digital Wallet",
    "Total Earnings",
    "Conversions",
    "Order Value",
    "Publisher Commission YoY %",
    "Earnings YoY %",
    "Conversions YoY %"
  ]);
  assert.equal(finalPayload.publisherTables.programLevelBreakdown[0]["Program Name"], "Program One");
  assert.ok(finalPayload.publisherTables.publisherPerformanceSummaryTable.some((row) => row.Metric === "Publisher Commission"));
  assert.ok(finalPayload.publisherTables.competitorAnalysisTable.some((row) => row.Publisher === "Primary"));
  assert.ok(finalPayload.publisherTables.moversShakersCommissionChart.some((row) => row.Direction === "Up"));
  assert.ok(Object.prototype.hasOwnProperty.call(finalPayload.publisherTables.moversShakersCommissionChart[0], "Chart Label"));
  assert.ok(Object.prototype.hasOwnProperty.call(finalPayload.publisherTables.moversShakersCommissionChart[0], "Publisher Commission Change"));
  assert.ok(finalPayload.publisherTables.programActivationSnapshotTable.some((row) => row.Metric === "Joined programs"));
  assert.ok(finalPayload.publisherTables.programGapAnalysisSummaryTable.some((row) => row.Metric === "Gap programs"));
  assert.ok(finalPayload.publisherTables.programGapAnalysisSummaryTable.some((row) => row.Metric === "Activation opportunities"));
  assert.ok(finalPayload.publisherTables.programGapAnalysisSummaryTable.some((row) => row.Metric === "Competitor pub comm opportunity"));
  assert.ok(finalPayload.publisherTables.programGapAnalysisByTypeTable.some((row) => row["Gap Type"] === "Application"));
  assert.deepEqual(Object.keys(finalPayload.publisherTables.competitorSharePubCommChart[0]), [
    "Competitor Group Summary",
    "Your Site",
    "Comp. A",
    "Comp. B",
    "Comp. C",
    "Comp. D"
  ]);
  assert.equal(finalPayload.publisherTables.competitorSharePubCommChart[0]["Competitor Group Summary"], "Publisher Commission PP");
  assert.match(finalPayload.publisherTables.competitorSharePubCommChart[0]["Your Site"], /^â‚¬/);
  assert.match(finalPayload.publisherTables.competitorSharePubCommChart[0]["Comp. A"], /^â‚¬/);
  assert.deepEqual(Object.keys(finalPayload.publisherTables.competitorWeeklyPubCommChart[0]), [
    "Week",
    "Primary",
    "Comp. A",
    "Comp. B",
    "Comp. C",
    "Comp. D"
  ]);
  assert.equal(finalPayload.publisherTables.competitorWeeklyPubCommChart[0].Primary, 20);
  assert.equal(finalPayload.publisherTables.competitorWeeklyPubCommChart[0]["Comp. A"], 12);
  assert.ok(Object.prototype.hasOwnProperty.call(finalPayload.publisherTables.programGapAnalysisTable[0], "Program"));
  assert.ok(Object.prototype.hasOwnProperty.call(finalPayload.publisherTables.programGapAnalysisTable[0], "Comp. A"));
  assert.deepEqual(Object.keys(finalPayload.publisherTables.topProgramsCompetitorPerformanceTable[0]), [
    "Program Name",
    "publisher-a",
    "Comp. A",
    "Comp. B",
    "Comp. C",
    "Comp. D"
  ]);
  assert.equal(finalPayload.publisherTables.topProgramsCompetitorPerformanceTable[0]["publisher-a"], "63%");
  assert.equal(finalPayload.publisherTables.topProgramsCompetitorPerformanceTable[0]["Comp. A"], "37%");
  assert.equal(finalPayload.slideTableBindings.kpi_summary_table, "kpiSummaryTable + kpiVarianceColorHintsTable");
});

test("runner keeps ten up and ten down publisher commission movers for slide 7", async () => {
  let finalPayload;
  const metricRows = (isPrevious) => {
    const rows = [];
    for (let index = 1; index <= 12; index += 1) {
      rows.push({
        programId: `UP${index}`,
        programName: `Up Program ${index}`,
        commission: isPrevious ? 1 : 100 + index
      });
      rows.push({
        programId: `DOWN${index}`,
        programName: `Down Program ${index}`,
        commission: isPrevious ? 100 + index : 0
      });
    }
    return rows;
  };

  await runPublisherQbrWorkflow(extensionBody(samplePayload({
    comparisonPublishers: [],
    publisherPptxApiUrl: "http://127.0.0.1:3010"
  })), {
    fetch: async (url, init = {}) => {
      const parsed = new URL(String(url));
      if (parsed.pathname.endsWith("/health")) {
        return new Response(JSON.stringify({ ok: true, service: "publisher-qbr-pptx-service", provider: "publisher-qbr-pptx" }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      if (parsed.pathname.endsWith("/generate")) {
        finalPayload = JSON.parse(init.body);
        return new Response(JSON.stringify({
          success: true,
          provider: "publisher-qbr-pptx",
          pptx_url: "http://127.0.0.1:3010/files/deck.pptx"
        }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      if (parsed.pathname.endsWith("/digitalwallets") || parsed.pathname.endsWith("/programs")) {
        return new Response(JSON.stringify({ items: [] }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      return new Response(JSON.stringify({ items: metricRows(parsed.searchParams.get("fromDate") === "20250101") }), {
        status: 200,
        headers: { "Content-Type": "application/json" }
      });
    },
    modelClient: {
      complete: async () => ({
        content: JSON.stringify({
          reportingPeriod: ["ok"],
          kpiHighlights: ["ok"],
          programLevelAnalysis: [],
          moversAndShakers: [],
          risksAndDependencies: []
        })
      })
    },
    logger: { info() {}, warn() {}, error() {} }
  });

  const movers = finalPayload.publisherTables.moversShakersCommissionChart;
  assert.equal(movers.length, 20);
  assert.equal(movers.slice(0, 10).every((row) => row.Direction === "Up"), true);
  assert.equal(movers.slice(10).every((row) => row.Direction === "Down"), true);
});

test("runner sends the full gap table so slide 9 is not derived from only top ten rows", async () => {
  let finalPayload;
  const competitorRows = Array.from({ length: 12 }, (_, index) => ({
    programId: `GAP${index + 1}`,
    programName: `Gap Program ${index + 1}`,
    commission: 100 + index
  }));

  await runPublisherQbrWorkflow(extensionBody(samplePayload({
    publisherPptxApiUrl: "http://127.0.0.1:3010"
  })), {
    fetch: async (url, init = {}) => {
      const parsed = new URL(String(url));
      const sourceId = parsed.searchParams.get("sourceId");
      if (parsed.pathname.endsWith("/health")) {
        return new Response(JSON.stringify({ ok: true, service: "publisher-qbr-pptx-service", provider: "publisher-qbr-pptx" }), {
          status: 200,
          headers: { "Content-Type": "application/json" }
        });
      }

      if (parsed.pathname.endsWith("/generate")) {
        finalPayload = JSON.parse(init.body);
        return new Response(JSON.stringify({
          success: true,
          provider: "publisher-qbr-pptx",
          pptx_url: "http://127.0.0.1:3010/files/deck.pptx"
        }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      if (parsed.pathname.endsWith("/digitalwallets")) {
        return new Response(JSON.stringify({ items: [] }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      if (parsed.pathname.endsWith("/programs")) {
        return new Response(JSON.stringify({ items: [] }), { status: 200, headers: { "Content-Type": "application/json" } });
      }
      return new Response(JSON.stringify({
        items: sourceId === "456" ? competitorRows : []
      }), { status: 200, headers: { "Content-Type": "application/json" } });
    },
    modelClient: {
      complete: async () => ({
        content: JSON.stringify({
          reportingPeriod: ["ok"],
          kpiHighlights: ["ok"],
          programLevelAnalysis: [],
          moversAndShakers: [],
          risksAndDependencies: []
        })
      })
    },
    logger: { info() {}, warn() {}, error() {} }
  });

  assert.equal(finalPayload.publisherTables.programGapAnalysisTable.length, 12);
  assert.equal(finalPayload.publisherTables.programGapAnalysisSummaryTable.find((row) => row.Metric === "Gap programs").Value, "12");
});

test("normalization protects publisher payload boundaries", () => {
  const normalized = normalizeWorkflowInput(extensionBody(samplePayload({
    presentonTemplateId: "advertiser-template",
    publisherPptxApiUrl: "http://127.0.0.1:3011"
  })));
  assert.equal(normalized.payload.analysisLevel, "publisher_program");
  assert.equal(normalized.payload.primarySourceId, "123");
  assert.equal(normalized.payload.competitorPublishers.length, 1);
});



