const assert = require("node:assert/strict");
const JSZip = require("jszip");

const { generatePresentation } = require("../lib/generator");

async function buildPresentationResult(overrides = {}) {
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

  return generatePresentation(payload);
}

async function buildDeckSpec(overrides = {}) {
  const result = await buildPresentationResult(overrides);
  return result.deckSpec;
}

async function coverUsesPublisherPerformanceReviewTitle() {
  const deckSpec = await buildDeckSpec();
  assert.equal(deckSpec.slides[0].title, "Topcashback Performance Review");
  assert.doesNotMatch(deckSpec.slides[0].title, /affiliate program/i);
  assert.doesNotMatch(deckSpec.slides[0].summary, /affiliate program/i);
  assert.doesNotMatch(JSON.stringify(deckSpec), /affiliate program/i);
}

async function renderedCoverUsesWhiteTdLogoAsset() {
  const result = await buildPresentationResult();
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file("ppt/slides/slide1.xml").async("string");
  const slideRels = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string");
  const mediaFiles = Object.keys(zip.files).filter((name) => /^ppt\/media\/image[-\d]+\.png$/.test(name));
  const imageRels = [...slideRels.matchAll(/Target="\.\.\/media\/image[-\d]+\.png"/g)];

  assert.match(slideXml, /<p:pic/);
  assert(imageRels.length >= 2);
  assert(mediaFiles.length >= 2);
  assert.doesNotMatch(slideXml, /<a:t>td<\/a:t>|tradedoubler/i);
}

function slideHasImageExtent(slideXml, widthInches, heightInches) {
  const emuPerInch = 914400;
  return [...slideXml.matchAll(/<a:ext cx="(\d+)" cy="(\d+)"/g)]
    .some((match) => (
      Math.abs(Number(match[1]) - Math.round(widthInches * emuPerInch)) <= 8
        && Math.abs(Number(match[2]) - Math.round(heightInches * emuPerInch)) <= 8
    ));
}

async function renderedCoverUsesCyanWireframeAndNoUtilityPills() {
  const result = await buildPresentationResult();
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file("ppt/slides/slide1.xml").async("string");
  const slideRels = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string");
  const imageRels = [...slideRels.matchAll(/Target="\.\.\/media\/image[-\d]+\.png"/g)];

  assert(imageRels.length >= 3);
  assert.equal(slideHasImageExtent(slideXml, 4.62, 4.62), true);
  assert.doesNotMatch(slideXml, /QBR Report|2026-01-16 to 2026-04-15|Analysis/);
}

async function renderedThankYouUsesLargerCyanWireframeAndCompactQuestionBubble() {
  const result = await buildPresentationResult();
  const zip = await JSZip.loadAsync(result.buffer);
  const thankYouSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "thank-you") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${thankYouSlideIndex}.xml`).async("string");
  const slideRels = await zip.file(`ppt/slides/_rels/slide${thankYouSlideIndex}.xml.rels`).async("string");
  const imageRels = [...slideRels.matchAll(/Target="\.\.\/media\/image[-\d]+\.png"/g)];

  assert(imageRels.length >= 2);
  assert.equal(slideHasImageExtent(slideXml, 5.86, 5.86), true);
  assert.match(slideXml, /Any Questions\?/);
  assert.doesNotMatch(slideXml, /<a:ext cx="11018520" cy="1024128"/);
}

async function executiveSummaryUsesRequestedPublisherMetrics() {
  const deckSpec = await buildDeckSpec();
  const slide = deckSpec.slides.find((item) => item.id === "executive-summary");
  const labels = slide.kpis.map((card) => card.label);

  assert.deepEqual(labels, [
    "Conversions",
    "Total Order Value",
    "Publisher Commission",
    "Digital Wallet",
    "Total Earnings",
    "Conversion Rate"
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

async function programActivationSnapshotSlideFollowsKpiSummary() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      programActivationSnapshotTable: [
        { Metric: "Joined programs", Total: "337", New: "148", "New %": "44%" },
        { Metric: "With clicks", Total: "257", New: "82", "New %": "32%" },
        { Metric: "With pub commission", Total: "180", New: "50", "New %": "28%" },
        { Metric: "Inactive", Total: "157", New: "98", "New %": "62%" }
      ]
    }
  });
  const ids = deckSpec.slides.map((slide) => slide.id);
  const kpiIndex = ids.indexOf("kpi-volume-conversion");
  const activationIndex = ids.indexOf("program-activation-snapshot");
  const programAnalysisIndex = ids.indexOf("kpi-cost-roi");
  const slide = deckSpec.slides[activationIndex];

  assert.notEqual(activationIndex, -1);
  assert.equal(activationIndex, kpiIndex + 1);
  assert.equal(programAnalysisIndex, activationIndex + 1);
  assert.equal(slide.kind, "program-activation-snapshot");
  assert.deepEqual(slide.activationSnapshot.map((item) => item.label), [
    "Joined programs",
    "With clicks",
    "Pub commission",
    "Inactive"
  ]);
  assert.equal(slide.activationSnapshot[0].total, "337");
  assert.equal(slide.activationSnapshot[3].newPercent, "62%");
  assert(slide.activationSnapshot.every((item) => !("tag" in item)));
}

async function renderedActivationSnapshotUsesTdLineGridStyle() {
  const result = await buildPresentationResult({
    publisherTables: {
      programActivationSnapshotTable: [
        { Metric: "Joined programs", Total: "337", New: "148", "New %": "44%" },
        { Metric: "With clicks", Total: "257", New: "82", "New %": "32%" },
        { Metric: "With pub commission", Total: "180", New: "50", "New %": "28%" },
        { Metric: "Inactive", Total: "157", New: "98", "New %": "62%" }
      ]
    }
  });
  const activationSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "program-activation-snapshot") + 1;
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file(`ppt/slides/slide${activationSlideIndex}.xml`).async("string");

  assert.match(slideXml, /Activation Status/);
  assert.match(slideXml, /Joined programs/);
  assert.match(slideXml, /With clicks/);
  assert.match(slideXml, /Pub commission/);
  assert.match(slideXml, /Inactive/);
  assert.match(slideXml, /<a:srgbClr val="FFFFFF"/);
  assert.match(slideXml, /algn="ctr"/);
  assert.doesNotMatch(slideXml, /PUBLISHER KPI FOLLOW-UP|Base|Traffic|Earning|Watch/);
  assert.doesNotMatch(slideXml, /Current period|Primary publisher/);
  assert.doesNotMatch(slideXml, /071E5C|123D8F/);
  assert.doesNotMatch(slideXml, /\sb="1"/);
}

async function renderedActivationSnapshotUsesBrandedFifthElementDivider() {
  const result = await buildPresentationResult({
    publisherTables: {
      programActivationSnapshotTable: [
        { Metric: "Joined programs", Total: "337", New: "148", "New %": "44%" },
        { Metric: "With clicks", Total: "257", New: "82", "New %": "32%" },
        { Metric: "With pub commission", Total: "180", New: "50", "New %": "28%" },
        { Metric: "Inactive", Total: "157", New: "98", "New %": "62%" }
      ]
    }
  });
  const activationSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "program-activation-snapshot") + 1;
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file(`ppt/slides/slide${activationSlideIndex}.xml`).async("string");
  const slideRels = await zip.file(`ppt/slides/_rels/slide${activationSlideIndex}.xml.rels`).async("string");
  const imageRels = [...slideRels.matchAll(/Target="\.\.\/media\/image[-\d]+\.png"/g)];
  const extents = [...slideXml.matchAll(/<a:ext cx="(\d+)" cy="(\d+)"/g)]
    .map((match) => ({ cx: Number(match[1]), cy: Number(match[2]) }));
  const emuPerInch = 914400;
  const hasWideThinDivider = extents.some((extent) => (
    Math.abs(extent.cx - Math.round(4.72 * emuPerInch)) <= 8
      && Math.abs(extent.cy - Math.round(3.34 * emuPerInch)) <= 8
  ));

  assert(imageRels.length >= 2);
  assert.equal(hasWideThinDivider, true);
  assert.doesNotMatch(slideXml, /prst="line"/);
}

async function renderedSlideUtilityTextIsRemovedAndTableTitlesAreCentered() {
  const result = await buildPresentationResult({
    programScopeTable: [
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
  const zip = await JSZip.loadAsync(result.buffer);
  const reportingSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "reporting-period") + 1;
  const programSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "kpi-cost-roi") + 1;
  const thankYouSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "thank-you") + 1;
  const reportingXml = await zip.file(`ppt/slides/slide${reportingSlideIndex}.xml`).async("string");
  const programXml = await zip.file(`ppt/slides/slide${programSlideIndex}.xml`).async("string");
  const thankYouXml = await zip.file(`ppt/slides/slide${thankYouSlideIndex}.xml`).async("string");

  assert.doesNotMatch(reportingXml, /Selected period from the QBR request/i);
  assert.match(programXml, /Per-program view ordered by publisher commission\./);
  assert.match(programXml, /algn="ctr"/);
  assert.match(thankYouXml, /Any Questions\?/);
  assert.doesNotMatch(thankYouXml, /TD Publisher Performance Review|2026-01-16 to 2026-04-15/);
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
  assert.equal(table.rows[0][3], "£25");
  assert.equal(table.rows[0][4], "£725");
  assert.equal(table.rows[1][0], "1111");
  assert.equal(table.rows[1][1], "Lower Commission Retailer");
  assert(table.columns.every((label) => !/^clicks$/i.test(label)));
  assert(table.columns.every((label) => !/clicks yoy|^sales$|sales yoy|^commission$/i.test(label)));
}

async function topNewProgramsIncludesWalletAndSortsByTotalEarnings() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      brandNewProgramsTable: [
        {
          "Program ID": "1111",
          "Program Name": "Lower Earning Retailer",
          Conversions: "10",
          "Order Value": "£5,000",
          "Publisher Commission": "£500",
          "Digital Wallet": "£10",
          "Total Earnings": "£510"
        },
        {
          "Program ID": "2222",
          "Program Name": "Higher Earning Retailer",
          Conversions: "4",
          "Order Value": "£2,000",
          "Publisher Commission": "£300",
          "Digital Wallet": "£450",
          "Total Earnings": "£750"
        }
      ]
    }
  });
  const slide = deckSpec.slides.find((item) => item.id === "brand-new-publishers");
  const table = slide.tables[0];

  assert.deepEqual(table.columns, [
    "Program ID",
    "Program Name",
    "Conversions",
    "Order Value",
    "Publisher Commission",
    "Digital Wallet",
    "Total Earnings"
  ]);
  assert.equal(table.rows[0][0], "2222");
  assert.equal(table.rows[0][1], "Higher Earning Retailer");
  assert.equal(table.rows[0][5], "£450");
  assert.equal(table.rows[0][6], "£750");
}

async function moversShakersUsesPublisherCommissionBarChart() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      moversShakersCommissionChart: [
        {
          "Program Name": "Growth Retailer",
          "Publisher Commission Change": "Â£8,000",
          "Publisher Commission Change Value": "8000",
          "Publisher Commission YoY %": "+80.0%"
        },
        {
          "Program Name": "Decline Retailer",
          "Publisher Commission Change": "-Â£5,000",
          "Publisher Commission Change Value": "-5000",
          "Publisher Commission YoY %": "-50.0%"
        }
      ]
    }
  });
  const slide = deckSpec.slides.find((item) => item.id === "movers-shakers-publisher-commission");

  assert.equal(slide.kind, "movers-bar-chart");
  assert.equal(slide.chart.type, "movers-commission-bar");
  assert.deepEqual(slide.chart.rows.map((row) => row.label), ["Growth Retailer", "Decline Retailer"]);
  assert.equal(deckSpec.slides.some((item) => item.id === "movers-shakers-clicks"), false);
  assert.equal(deckSpec.slides.some((item) => item.id === "movers-shakers-sales"), false);
  assert.equal(deckSpec.slides.some((item) => item.id === "movers-shakers-ov"), false);
}

async function moversShakersDisplaysDeclinersSmallestLossFirst() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      moversShakersCommissionChart: [
        {
          "Program Name": "Growth Retailer",
          "Publisher Commission Change": "GBP 8,000",
          "Publisher Commission Change Value": "8000",
          "Publisher Commission YoY %": "+80.0%"
        },
        {
          "Program Name": "Morrisons Grocery",
          "Publisher Commission Change": "-GBP 251,731",
          "Publisher Commission Change Value": "-251731",
          "Publisher Commission YoY %": "-54.9%"
        },
        {
          "Program Name": "Travel Republic",
          "Publisher Commission Change": "-GBP 94,575",
          "Publisher Commission Change Value": "-94575",
          "Publisher Commission YoY %": "-100.0%"
        },
        {
          "Program Name": "Lenstore UK Contact Lenses",
          "Publisher Commission Change": "-GBP 8,932",
          "Publisher Commission Change Value": "-8932",
          "Publisher Commission YoY %": "-52.5%"
        }
      ]
    }
  });
  const slide = deckSpec.slides.find((item) => item.id === "movers-shakers-publisher-commission");

  assert.deepEqual(slide.chart.rows.map((row) => row.label), [
    "Growth Retailer",
    "Lenstore UK Contact Lenses",
    "Travel Republic",
    "Morrisons Grocery"
  ]);
}

async function renderedMoversShakersNegativeBarsUseRedFill() {
  const result = await buildPresentationResult({
    themeOverrides: {
      colors: {
        accentAlt: "#D93025"
      }
    },
    publisherTables: {
      moversShakersCommissionChart: [
        {
          "Program Name": "Growth Retailer",
          "Publisher Commission Change": "GBP 8,000",
          "Publisher Commission Change Value": "8000",
          "Publisher Commission YoY %": "+80.0%"
        },
        {
          "Program Name": "Decline Retailer",
          "Publisher Commission Change": "-GBP 5,000",
          "Publisher Commission Change Value": "-5000",
          "Publisher Commission YoY %": "-50.0%"
        }
      ]
    }
  });
  const moversSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "movers-shakers-publisher-commission") + 1;
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file(`ppt/slides/slide${moversSlideIndex}.xml`).async("string");

  assert.match(slideXml, /<a:srgbClr val="D93025"/);
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

async function productionPublisherRequestKeepsFullPublisherSlideSet() {
  const deckSpec = await buildDeckSpec({
    slideBlueprint: [
      { slide: 1, key: "cover", title: "Performance Review Cover" },
      { slide: 2, key: "reporting_period", title: "Reporting Period" },
      { slide: 3, key: "kpi_tiles", title: "Executive KPI Tiles" },
      { slide: 4, key: "kpi_summary_table", title: "KPI Summary Table" },
      { slide: 5, key: "program_level_analysis", title: "Program Level Analysis" },
      { slide: 6, key: "movers_commission_chart", title: "Movers & Shakers - Publisher Commission" },
      { slide: 7, key: "kpi_highlights", title: "KPI Highlights" },
      { slide: 8, key: "competitor_analysis", title: "Competitor Analysis" },
      { slide: 9, key: "brand_new_programs", title: "Top 10 Newly Activated Programs" },
      { slide: 10, key: "sales_growth_risk_dependencies", title: "Conversion Growth Signals & Risk and Dependencies" },
      { slide: 11, key: "thank_you", title: "Thank You" }
    ],
    targetSlides: 11,
    publisherTables: {
      moversShakersCommissionChart: [
        {
          "Program Name": "Growth Retailer",
          "Publisher Commission Change": "GBP 8,000",
          "Publisher Commission Change Value": "8000",
          "Publisher Commission YoY %": "+80.0%"
        }
      ],
      competitorAnalysisTable: [
        {
          "Competitor Group Summary": "Programs with Pub Comm",
          "Your Site": "180",
          "Publisher 1": "115",
          "Publisher 2": "119",
          "Publisher 3": "36",
          "Publisher 4": "46",
          "Distinct comp. prog. #": "150"
        }
      ],
      competitorWeeklyPubCommChart: [
        {
          Week: "2026-01-05",
          "Your Site": "GBP 80,000",
          "Publisher 1": "GBP 18,500"
        },
        {
          Week: "2026-01-12",
          "Your Site": "GBP 75,000",
          "Publisher 1": "GBP 21,000"
        }
      ],
      topProgramsCompetitorPerformanceTable: [
        {
          "Program Name": "Morrisons Grocery",
          TopCashback: "66%",
          "Comp. A": "8%",
          "Comp. B": "22%",
          "Comp. C": "4%",
          "Comp. D": "0%"
        }
      ],
      salesGrowthSignalsTable: [
        {
          "Program Name": "Growth Retailer",
          "Conversions YoY Change": "25",
          "Conversions YoY %": "+25.0%",
          "Total Earnings YoY Change": "GBP 200",
          "Total Earnings YoY %": "+25.0%"
        }
      ],
      riskDependenciesTable: [
        {
          "Program Name": "Decline Retailer",
          "Risk Type": "YoY decline",
          Evidence: "Conversions -20, total earnings -GBP 500",
          Priority: "High"
        }
      ],
      brandNewProgramsTable: [
        {
          "Program ID": "3333",
          "Program Name": "New Retailer",
          Conversions: "10",
          "Order Value": "GBP 1,000",
          "Publisher Commission": "GBP 100",
          "Digital Wallet": "GBP 0",
          "Total Earnings": "GBP 100"
        }
      ]
    }
  });
  const ids = deckSpec.slides.map((slide) => slide.id);

  assert(ids.includes("executive-summary"));
  assert(ids.includes("movers-shakers-publisher-commission"));
  assert(ids.includes("competitor-analysis"));
  assert(ids.includes("top-programs-competitor-performance"));
  assert.equal(ids.includes("sales-growth-signals"), false);
  assert(ids.includes("risks-dependencies"));
  assert(ids.includes("brand-new-publishers"));
  assert(ids.includes("thank-you"));
  assert.deepEqual(
    ids.slice(ids.indexOf("competitor-analysis"), ids.indexOf("brand-new-publishers") + 1),
    ["competitor-analysis", "top-programs-competitor-performance", "brand-new-publishers"]
  );
  assert.deepEqual(ids.slice(-2), ["risks-dependencies", "thank-you"]);

  const competitorPerformanceSlide = deckSpec.slides.find((slide) => slide.id === "top-programs-competitor-performance");
  const table = competitorPerformanceSlide.tables[0];
  assert.equal(table.primaryHighlightColumn, 1);
  assert.deepEqual(table.columns, ["Program Name", "TopCashback", "Comp. A", "Comp. B", "Comp. C", "Comp. D"]);
  assert.deepEqual(table.rows[0], ["Morrisons Grocery", "66%", "8%", "22%", "4%", "0%"]);
}

async function publisherQbrAnalysisUsesProgramLanguage() {
  const deckSpec = await buildDeckSpec({
    publisherOverviewObservations: [
      "This publisher was the biggest contributor to the decline and should be reviewed."
    ],
    publisherTables: {
      topDecliningPublishers: [
        {
          Publisher: "Morrisons Grocery",
          Segment: "Cashback",
          "Sales YoY %": "-54.9%",
          "OV YoY Change": "-GBP 251,731",
          "OV YoY %": "-54.9%"
        }
      ]
    }
  });
  const kpiSlide = deckSpec.slides.find((slide) => slide.id === "kpi-highlights");
  const riskSlide = deckSpec.slides.find((slide) => slide.id === "risks-dependencies");
  const analysisText = [
    ...kpiSlide.bullets,
    ...riskSlide.tables[0].rows.flat()
  ].join(" ");
  const withoutMetricName = analysisText.replace(/publisher commission/gi, "");

  assert.match(analysisText, /program/i);
  assert.doesNotMatch(withoutMetricName, /\bpublishers?\b/i);
}

async function run() {
  const tests = [
    coverUsesPublisherPerformanceReviewTitle,
    renderedCoverUsesWhiteTdLogoAsset,
    renderedCoverUsesCyanWireframeAndNoUtilityPills,
    renderedThankYouUsesLargerCyanWireframeAndCompactQuestionBubble,
    executiveSummaryUsesRequestedPublisherMetrics,
    kpiSummaryTableUsesMetricRowsAndAddsRequestedMetrics,
    programActivationSnapshotSlideFollowsKpiSummary,
    renderedActivationSnapshotUsesTdLineGridStyle,
    renderedActivationSnapshotUsesBrandedFifthElementDivider,
    renderedSlideUtilityTextIsRemovedAndTableTitlesAreCentered,
    programLevelAnalysisUsesPublisherCommissionHierarchy,
    topNewProgramsIncludesWalletAndSortsByTotalEarnings,
    moversShakersUsesPublisherCommissionBarChart,
    moversShakersDisplaysDeclinersSmallestLossFirst,
    renderedMoversShakersNegativeBarsUseRedFill,
    competitorAnalysisSlideUsesAnonymousComparisonTable,
    competitorAnalysisSlideIncludesWeeklyComboChart,
    productionPublisherRequestKeepsFullPublisherSlideSet,
    publisherQbrAnalysisUsesProgramLanguage
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
