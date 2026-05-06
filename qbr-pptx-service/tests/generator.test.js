const assert = require("node:assert/strict");

const { generatePresentation } = require("../lib/generator");

async function buildDeckSpec(overrides = {}) {
  const payload = {
    client: "topcashback",
    deckTitle: "QBR - topcashback",
    reportingPeriod: "2026-01-16 to 2026-04-15",
    comparisonPeriod: "2025-01-16 to 2025-04-15",
    languageCode: "EN",
    languageName: "English",
    currencyCode: "GBP",
    qbrFocus: "General performance review",
    programYoYTable: [
      {
        Row: "Recent",
        Conversions: "125",
        "Conv Rate": "5.0%",
        Clicks: "2,500",
        "Earnings per Click": "GBP 0.40",
        "Earnings per Commission": "1.33x",
        "Order Value": "£12,500",
        "Publisher Commission": "£750",
        "Digital Wallet": "£250",
        "Total Earnings": "£1,000",
        "Active Programs": "4"
      },
      {
        Row: "Previous",
        Conversions: "100",
        "Conv Rate": "4.0%",
        Clicks: "2,000",
        "Earnings per Click": "GBP 0.40",
        "Earnings per Commission": "1.14x",
        "Order Value": "£10,000",
        "Publisher Commission": "£700",
        "Digital Wallet": "£100",
        "Total Earnings": "£800",
        "Active Programs": "4"
      },
      {
        Row: "Difference",
        Conversions: "25",
        "Conv Rate": "1.0%",
        Clicks: "500",
        "Earnings per Click": "GBP 0.00",
        "Earnings per Commission": "0.19x",
        "Order Value": "£2,500",
        "Publisher Commission": "£50",
        "Digital Wallet": "£150",
        "Total Earnings": "£200",
        "Active Programs": "0"
      },
      {
        Row: "Variance",
        Conversions: "+25.0%",
        "Conv Rate": "+25.0%",
        Clicks: "+25.0%",
        "Earnings per Click": "0.0%",
        "Earnings per Commission": "+16.7%",
        "Order Value": "+25.0%",
        "Publisher Commission": "+7.1%",
        "Digital Wallet": "+150.0%",
        "Total Earnings": "+25.0%",
        "Active Programs": "0.0%"
      }
    ],
    publisherTables: {
      segmentSummary: [
        {
          Segment: "Cashback",
          "Total Sales": "125",
          "Sales YoY %": "+25.0%",
          "Total OV": "£12,500",
          "OV YoY %": "+25.0%",
          Publishers: "4"
        }
      ]
    },
    ...overrides
  };

  const result = await generatePresentation(payload);
  return result.deckSpec;
}

async function coverUsesPublisherPerformanceReviewTitle() {
  const deckSpec = await buildDeckSpec();
  assert.equal(deckSpec.slides[0].title, "topcashback performance review");
  assert.doesNotMatch(deckSpec.slides[0].title, /affiliate program/i);
  assert.doesNotMatch(deckSpec.slides[0].summary, /affiliate program/i);
  assert.doesNotMatch(JSON.stringify(deckSpec), /affiliate program/i);
}

async function executiveSummaryUsesRequestedPublisherMetrics() {
  const deckSpec = await buildDeckSpec();
  const slide = deckSpec.slides.find((item) => item.id === "executive-summary");
  const labels = slide.kpis.map((card) => card.label);

  assert.deepEqual(labels, [
    "Conversions",
    "Total Order Value",
    "Total Earnings",
    "Conversion Rate",
    "Publisher Commission"
  ]);
  assert(labels.every((label) => !/roi|aov|average order value|sales/i.test(label)));
}

async function kpiSummaryTableUsesMetricRowsAndAddsRequestedMetrics() {
  const deckSpec = await buildDeckSpec();
  const slide = deckSpec.slides.find((item) => item.id === "kpi-volume-conversion");
  const table = slide.tables[0];
  const metricNames = table.rows.map((row) => row[0]);

  assert.deepEqual(table.columns, ["Metric", "Recent", "Previous", "Difference", "% Variance"]);
  assert.deepEqual(metricNames, [
    "Conversions",
    "Conversion Rate",
    "Clicks",
    "Earnings per Click",
    "Earnings per Commission",
    "Total Order Value",
    "Publisher Commission",
    "Digital Wallet",
    "Total Earnings",
    "Active Programs"
  ]);
  assert(metricNames.every((label) => !/roi|aov|average order value/i.test(label)));
}

async function programLevelAnalysisUsesPublisherCommissionHierarchy() {
  const deckSpec = await buildDeckSpec({
    programScopeTable: [
      {
        "Program ID": "1111",
        "Program Name": "Lower Commission Retailer",
        Conversions: "90",
        "Order Value": "£9,000",
        "Publisher Commission": "£300",
        "Digital Wallet": "£120",
        "Total Earnings": "£420",
        "Publisher Commission YoY %": "+5.0%",
        "Earnings YoY %": "+8.0%",
        "Conversions YoY %": "+20.0%"
      },
      {
        "Program ID": "2222",
        "Program Name": "Higher Commission Retailer",
        Conversions: "40",
        "Order Value": "£4,000",
        "Publisher Commission": "£700",
        "Digital Wallet": "£25",
        "Total Earnings": "£725",
        "Publisher Commission YoY %": "-2.0%",
        "Earnings YoY %": "+1.0%",
        "Conversions YoY %": "-10.0%"
      }
    ]
  });
  const slide = deckSpec.slides.find((item) => item.id === "kpi-cost-roi");
  const table = slide.tables[0];

  assert.deepEqual(table.columns, [
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
  assert.equal(table.rows[0][0], "2222");
  assert.equal(table.rows[0][1], "Higher Commission Retailer");
  assert.equal(table.rows[0][2], "£700");
  assert.equal(table.rows[1][0], "1111");
  assert.equal(table.rows[1][1], "Lower Commission Retailer");
  assert(table.columns.every((label) => !/^clicks$/i.test(label)));
  assert(table.columns.every((label) => !/clicks yoy|^sales$|sales yoy|^commission$/i.test(label)));
}

async function competitorAnalysisSlideUsesAnonymousComparisonTable() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      competitorAnalysisTable: [
        {
          Metric: "Programs with Pub Comm",
          "Your Site": "180",
          "Publisher 1": "115",
          "Publisher 2": "119",
          "Publisher 3": "36",
          "Publisher 4": "46",
          "Distinct comp. prog. #": "150"
        },
        {
          Metric: "of which you have 0 Pub Comm",
          "Your Site": "n/a",
          "Publisher 1": "8",
          "Publisher 2": "15",
          "Publisher 3": "4",
          "Publisher 4": "3",
          "Distinct comp. prog. #": "24"
        },
        {
          Metric: "Pub Comm of the above",
          "Your Site": "n/a",
          "Publisher 1": "£951",
          "Publisher 2": "£4,439",
          "Publisher 3": "£59",
          "Publisher 4": "£154",
          "Distinct comp. prog. #": "£5,497"
        }
      ]
    }
  });
  const slide = deckSpec.slides.find((item) => item.id === "competitor-analysis");
  const table = slide.tables[0];

  assert.equal(slide.title, "Competitor Analysis");
  assert.deepEqual(table.columns, [
    "Competitor Group Summary",
    "Your Site",
    "Publisher 1",
    "Publisher 2",
    "Publisher 3",
    "Publisher 4",
    "Distinct comp. prog. #"
  ]);
  assert.deepEqual(table.rows[2], [
    "Pub Comm of the above",
    "n/a",
    "£951",
    "£4,439",
    "£59",
    "£154",
    "£5,497"
  ]);
}

async function competitorAnalysisSlideIncludesWeeklyComboChart() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      competitorAnalysisTable: [
        {
          Metric: "Programs with Pub Comm",
          "Your Site": "2",
          "Publisher 1": "1",
          "Publisher 2": "1",
          "Publisher 3": "1",
          "Publisher 4": "1",
          "Distinct comp. prog. #": "4"
        }
      ],
      competitorWeeklyPubCommChart: [
        {
          Week: "2026-01-05",
          "Your Site": "£80,000",
          "Publisher 1": "£18,500",
          "Publisher 2": "£12,000",
          "Publisher 3": "£1,800",
          "Publisher 4": "£2,400"
        },
        {
          Week: "2026-01-12",
          "Your Site": "£75,000",
          "Publisher 1": "£21,000",
          "Publisher 2": "£13,500",
          "Publisher 3": "£2,100",
          "Publisher 4": "£2,600"
        }
      ]
    }
  });
  const slide = deckSpec.slides.find((item) => item.id === "competitor-analysis");

  assert.equal(slide.chart.type, "weekly-pub-comm-combo");
  assert.deepEqual(slide.chart.categories, ["2026-01-05", "2026-01-12"]);
  assert.equal(slide.chart.series[0].renderAs, "bar");
  assert.equal(slide.chart.series[0].label, "Your Site");
  assert.equal(slide.chart.series[1].renderAs, "line");
  assert.equal(slide.chart.series[1].label, "Publisher 1");
  assert.deepEqual(slide.chart.series[0].values, [80000, 75000]);
  assert.deepEqual(slide.chart.series[1].values, [18500, 21000]);
}

async function run() {
  const tests = [
    coverUsesPublisherPerformanceReviewTitle,
    executiveSummaryUsesRequestedPublisherMetrics,
    kpiSummaryTableUsesMetricRowsAndAddsRequestedMetrics,
    programLevelAnalysisUsesPublisherCommissionHierarchy,
    competitorAnalysisSlideUsesAnonymousComparisonTable,
    competitorAnalysisSlideIncludesWeeklyComboChart
  ];

  for (const test of tests) {
    await test();
    console.log(`ok - ${test.name}`);
  }
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
