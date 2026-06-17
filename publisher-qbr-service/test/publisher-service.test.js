const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

process.env.QBR_AUTO_TRANSLATE = "0";

const { generatePresentation } = require("../lib/generator");
const { createServer } = require("../server");

function listen(server) {
  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, () => {
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

function publisherPayload(overrides = {}) {
  return {
    analysisLevel: "publisher_program",
    client: "Publisher Boundary Test",
    deckTitle: "Publisher QBR - Boundary Test",
    outputFileName: "Publisher Boundary Test.pptx",
    fullContent: true,
    includeAppendix: false,
    languageCode: "EN",
    currencyCode: "EUR",
    reportingPeriod: "2026-01-01 to 2026-03-31",
    comparisonPeriod: "2025-01-01 to 2025-03-31",
    programYoYTable: [
      { Row: "Recent", Clicks: "1000", Conversions: "100", "Conversion Rate": "10%", "Total Order Value": "EUR 5000", "Publisher Commission": "EUR 500", "Digital Wallet": "EUR 20", "Total Earnings": "EUR 520" },
      { Row: "Previous", Clicks: "900", Conversions: "90", "Conversion Rate": "10%", "Total Order Value": "EUR 4050", "Publisher Commission": "EUR 450", "Digital Wallet": "EUR 15", "Total Earnings": "EUR 465" },
      { Row: "Difference", Clicks: "100", Conversions: "10", "Conversion Rate": "0%", "Total Order Value": "EUR 950", "Publisher Commission": "EUR 50", "Digital Wallet": "EUR 5", "Total Earnings": "EUR 55" },
      { Row: "% Variance", Clicks: "11.1%", Conversions: "11.1%", "Conversion Rate": "0%", "Total Order Value": "23.5%", "Publisher Commission": "11.1%", "Digital Wallet": "33.3%", "Total Earnings": "11.8%" }
    ],
    publisherTables: {
      publisherPerformanceSummary: [
        { Publisher: "Publisher A", Segment: "Cashback", "Order Value": "EUR 3000", "Current Sales": "60", "Sales YoY %": "20%" }
      ],
      programLevelBreakdown: [
        { Program: "Program A", "Program ID": "123", "Current OV": "EUR 3000", "OV YoY %": "20%", "Current Sales": "60", "Sales YoY %": "20%" }
      ],
      brandNewPrograms: [
        { Program: "New Program", "Program ID": "456", "Current OV": "EUR 500", "Current Sales": "10" }
      ]
    },
    ...overrides
  };
}

test("publisher service defaults to the local publisher tunnel port", async () => {
  const serverSource = await fs.readFile(path.join(__dirname, "..", "server.js"), "utf8");
  assert.match(serverSource, /process\.env\.PORT\s*\|\|\s*3010/);
});

test("publisher QBR payload produces publisher template slides", async () => {
  const result = await generatePresentation(publisherPayload());
  const titles = result.deckSpec.slides.map((slide) => slide.title);

  assert.equal(result.deckSpec.metadata.analysisLevel, "publisher_program");
  assert.equal(titles.length, 18);
  assert.ok(titles.includes("Program Performance: Executive Summary"));
  assert.ok(titles.includes("KPI Summary Table: Conversions & Earnings"));
  assert.ok(titles.includes("Program-Level Analysis: Publisher Commission"));
  assert.ok(titles.includes("Growth opportunity in the competitor gap"));
  assert.ok(titles.includes("Top 10 competitor-funded gaps"));
  assert.ok(titles.includes("Program Connection Status (1/2)"));
  assert.ok(titles.includes("Program Connection Status (2/2)"));
  assert.ok(titles.includes("Share Within Competitor Group"));
  assert.ok(titles.includes("KPI Highlights & Business Implications"));
  assert.ok(titles.includes("Risks & Dependencies"));
  assert.ok(titles.includes("Top 10 New Programs"));
  assert.doesNotMatch(JSON.stringify(result.deckSpec), /Affiliate Program|Publisher Performance Summary|Brand New Programs|Average Order Value \(AOV\)|\bAOV\b|\bROI\b/);
});

test("publisher QBR payload follows the current workflow slide blueprint", async () => {
  const result = await generatePresentation(publisherPayload({
    client: "digidip",
    slideBlueprint: [
      { slide: 1, key: "cover", title: "Performance Review Cover" },
      { slide: 2, key: "reporting_period", title: "Reporting Period" },
      { slide: 3, key: "kpi_tiles", title: "Program Performance: Executive Summary" },
      { slide: 4, key: "kpi_summary_table", title: "KPI Summary Table: Conversions & Earnings" },
      { slide: 5, key: "program_activation_snapshot", title: "Activation Status" },
      { slide: 6, key: "program_level_analysis", title: "Program-Level Analysis: Publisher Commission" },
      { slide: 7, key: "movers_commission_chart", title: "Movers & Shakers: Publisher Commission" },
      { slide: 8, key: "brand_new_programs", title: "Top 10 New Programs" },
      { slide: 9, key: "growth_opportunity_gap", title: "Growth opportunity in the competitor gap" },
      { slide: 10, key: "top_competitor_funded_gaps", title: "Top 10 competitor-funded gaps" },
      { slide: 11, key: "program_connection_status_1", title: "Program Connection Status (1/2)" },
      { slide: 12, key: "program_connection_status_2", title: "Program Connection Status (2/2)" },
      { slide: 13, key: "competitor_analysis", title: "Competitor Analysis" },
      { slide: 14, key: "share_within_competitor_group", title: "Share Within Competitor Group" },
      { slide: 15, key: "top_programs_competitor_performance", title: "Top Programs Competitor Performance" },
      { slide: 16, key: "kpi_highlights", title: "KPI Highlights & Business Implications" },
      { slide: 17, key: "risks_dependencies", title: "Risks & Dependencies" },
      { slide: 18, key: "thank_you", title: "Thank You" }
    ],
    publisherTables: {
      ...publisherPayload().publisherTables,
      programActivationSnapshotTable: [
        { Metric: "Joined programs", Total: "100", New: "12", "New %": "12%" }
      ],
      competitorAnalysisTable: [
        { Publisher: "Primary", "Publisher Commission": "EUR 1000", "Publisher Commission PP": "EUR 800" }
      ],
      topProgramsCompetitorPerformanceTable: [
        { "Program ID": "123", "Program Name": "Program A", Primary: "45%", "Comp. A": "20%", "Comp. B": "15%", "Comp. C": "10%", "Comp. D": "10%" }
      ],
      programConnectionStatusTable: [
        { "Program ID": "123", "Program Name": "Program A", "Connection Status": "Active" }
      ],
      programGapAnalysisSummaryTable: [
        { Metric: "Gap programs", Value: "104", Detail: "Programs requiring activation" }
      ],
      programGapAnalysisTable: [
        { Rank: "1", "Program / ID": "Wightlink 248069", "Primary status": "Accepted", "Competitor signal": "1/4 competitors earning", "Pub Comm - Specified Sites Value": "EUR 18410" }
      ],
      competitorSharePubCommChart: [
        { Publisher: "Your Site", "Publisher Commission PP": "40%", "Publisher Commission": "50%" }
      ],
      riskDependenciesTable: [
        { Issue: "Commission concentration", Analysis: "Program A contributes a high share of commission.", Action: "Monitor next cycle." }
      ]
    }
  }));
  const titles = result.deckSpec.slides.map((slide) => slide.title);
  const allText = JSON.stringify(result.deckSpec);

  assert.deepEqual(titles, [
    "digidip performance review",
    "Reporting Period",
    "Program Performance: Executive Summary",
    "KPI Summary Table: Conversions & Earnings",
    "Activation Status",
    "Program-Level Analysis: Publisher Commission",
    "Movers & Shakers: Publisher Commission",
    "Top 10 New Programs",
    "Growth opportunity in the competitor gap",
    "Top 10 competitor-funded gaps",
    "Program Connection Status (1/2)",
    "Program Connection Status (2/2)",
    "Competitor Analysis",
    "Share Within Competitor Group",
    "Top Programs Competitor Performance",
    "KPI Highlights & Business Implications",
    "Risks & Dependencies",
    "digidip - Thank you."
  ]);
  assert.doesNotMatch(allText, /Affiliate Program Quarterly Business Review|affiliate program's performance|Publisher Performance Summary|Brand New Programs|Average Order Value \(AOV\)|ROI/);
});

test("advertiser-style payload is rejected by the publisher service", async () => {
  const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "publisher-qbr-service-"));
  let called = false;
  const server = createServer({
    apiKey: "test-key",
    outputDir,
    scheduleDeletion: false,
    generatePresentation: async () => {
      called = true;
      throw new Error("should not route advertiser payload to publisher generator");
    }
  });

  const root = await listen(server);
  try {
    const response = await fetch(`${root}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": "test-key" },
      body: JSON.stringify({ client: "Advertiser QBR", programYoYTable: [] })
    });
    const body = await response.json();
    assert.equal(response.status, 400);
    assert.equal(body.success, false);
    assert.match(body.message, /publisher_program/i);
    assert.equal(called, false);
  } finally {
    await close(server);
    await fs.rm(outputDir, { recursive: true, force: true });
  }
});

test("publisher payload carrying advertiser template fields is rejected before generation", async () => {
  const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "publisher-qbr-service-"));
  let called = false;
  const server = createServer({
    apiKey: "test-key",
    outputDir,
    scheduleDeletion: false,
    generatePresentation: async () => {
      called = true;
      throw new Error("should not generate with advertiser template fields");
    }
  });

  const root = await listen(server);
  try {
    const response = await fetch(`${root}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": "test-key" },
      body: JSON.stringify(publisherPayload({
        presentonTemplateId: "advertiser-qbr-template",
        templateFamily: "advertiser_qbr"
      }))
    });
    const body = await response.json();
    assert.equal(response.status, 400);
    assert.equal(body.success, false);
    assert.match(body.message, /advertiser template/i);
    assert.equal(called, false);
  } finally {
    await close(server);
    await fs.rm(outputDir, { recursive: true, force: true });
  }
});

test("publisher payload pointing at advertiser PPTX service port is rejected before generation", async () => {
  const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "publisher-qbr-service-"));
  let called = false;
  const server = createServer({
    apiKey: "test-key",
    outputDir,
    scheduleDeletion: false,
    generatePresentation: async () => {
      called = true;
      throw new Error("should not generate with advertiser service URL");
    }
  });

  const root = await listen(server);
  try {
    const response = await fetch(`${root}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": "test-key" },
      body: JSON.stringify(publisherPayload({
        publisherPptxApiUrl: "http://127.0.0.1:3011",
        presentonApiUrl: "http://127.0.0.1:3011"
      }))
    });
    const body = await response.json();
    assert.equal(response.status, 400);
    assert.equal(body.success, false);
    assert.match(body.message, /port 3011/i);
    assert.equal(called, false);
  } finally {
    await close(server);
    await fs.rm(outputDir, { recursive: true, force: true });
  }
});

test("two publisher runs with the same outputFileName create separate files and signed URLs", async () => {
  const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "publisher-qbr-service-"));
  const server = createServer({
    apiKey: "test-key",
    downloadTokenSecret: "download-secret",
    outputDir,
    scheduleDeletion: false,
    generatePresentation: async () => ({
      normalized: { debug: false },
      deckSpec: {
        metadata: { requestId: "publisher-request", analysisLevel: "publisher_program" },
        slides: [{ title: "Publisher Performance Summary" }],
        theme: { name: "TD" }
      },
      buffer: Buffer.from(`pptx-${Date.now()}-${Math.random()}`),
      fileName: "same-name.pptx"
    })
  });

  const root = await listen(server);
  try {
    const request = () => fetch(`${root}/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": "test-key" },
      body: JSON.stringify(publisherPayload({ outputFileName: "same-name.pptx" }))
    });

    const first = await request();
    const second = await request();
    assert.equal(first.status, 200);
    assert.equal(second.status, 200);

    const firstBody = await first.json();
    const secondBody = await second.json();
    assert.equal(firstBody.file_name, "same-name.pptx");
    assert.match(secondBody.file_name, /^same-name_[0-9a-f-]+\.pptx$/i);
    assert.notEqual(firstBody.file_name, secondBody.file_name);
    assert.match(firstBody.pptx_url, new RegExp(`/files/${firstBody.file_name.replace(".", "\\.")}\\?expires=`));
    assert.match(secondBody.pptx_url, new RegExp(`/files/${secondBody.file_name.replace(".", "\\.")}\\?expires=`));

    const files = await fs.readdir(outputDir);
    assert.ok(files.includes(firstBody.file_name));
    assert.ok(files.includes(secondBody.file_name));
  } finally {
    await close(server);
    await fs.rm(outputDir, { recursive: true, force: true });
  }
});
