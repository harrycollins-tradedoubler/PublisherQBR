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

function slideHasFilledShapeExtent(slideXml, widthInches, heightInches, color) {
  const emuPerInch = 914400;
  const shapes = slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
  return shapes.some((shape) => {
    const extent = shape.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!extent) return false;
    const widthMatches = Math.abs(Number(extent[1]) - Math.round(widthInches * emuPerInch)) <= 8;
    const heightMatches = Math.abs(Number(extent[2]) - Math.round(heightInches * emuPerInch)) <= 8;
    return widthMatches && heightMatches && shape.includes(`<a:srgbClr val="${color}"`);
  });
}

function slideHasNonHorizontalLineColor(slideXml, color) {
  const shapes = slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
  return shapes.some((shape) => {
    if (!shape.includes('prst="line"') || !shape.includes(`<a:srgbClr val="${color}"`)) return false;
    const extent = shape.match(/<a:ext cx="(-?\d+)" cy="(-?\d+)"/);
    return extent && Math.abs(Number(extent[2])) > 0;
  });
}

function slideHasEllipseColor(slideXml, color) {
  const shapes = slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
  return shapes.some((shape) => shape.includes('prst="ellipse"') && shape.includes(`<a:srgbClr val="${color}"`));
}

async function chartXmlsForSlide(zip, slideIndex) {
  const relsPath = `ppt/slides/_rels/slide${slideIndex}.xml.rels`;
  const relsFile = zip.file(relsPath);
  if (!relsFile) return [];
  const relsXml = await relsFile.async("string");
  const chartTargets = [...relsXml.matchAll(/<Relationship[^>]*Type="[^"]*\/chart"[^>]*Target="([^"]+)"/g)]
    .map((match) => match[1])
    .map((target) => {
      if (target.startsWith("/")) return target.replace(/^\//, "");
      return target.replace(/^\.\.\//, "ppt/");
    });
  return Promise.all(chartTargets.map((target) => zip.file(target).async("string")));
}

function chartSeriesXmlByName(chartXml, name) {
  const series = chartXml.match(/<c:ser>[\s\S]*?<\/c:ser>/g) || [];
  return series.find((item) => item.includes(`<c:v>${name}</c:v>`)) || "";
}

function chartSeriesLineColor(chartXml, name) {
  const seriesXml = chartSeriesXmlByName(chartXml, name);
  const match = seriesXml.match(/<c:spPr>[\s\S]*?<a:ln[\s\S]*?<a:srgbClr val="([A-Fa-f0-9]{6})"/);
  return match ? match[1].toUpperCase() : "";
}

async function renderedCoverUsesCyanWireframeAndNoUtilityPills() {
  const result = await buildPresentationResult();
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file("ppt/slides/slide1.xml").async("string");
  const slideRels = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string");
  const imageRels = [...slideRels.matchAll(/Target="\.\.\/media\/image[-\d]+\.png"/g)];

  assert(imageRels.length >= 3);
  assert.equal(slideHasImageExtent(slideXml, 4.4, 4.4), true);
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
  assert.equal(slideHasImageExtent(slideXml, 6.9, 6.9), true);
  assert.match(slideXml, /Topcashback - Thank you\./);
  assert.doesNotMatch(slideXml, /topcashback - Thank you\./);
  assert.match(slideXml, /Any Questions\?/);
  assert.match(slideXml, /<a:pPr algn="ctr"/);
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

async function executiveSummaryDescribesFullPublisherProgramScope() {
  const deckSpec = await buildDeckSpec({
    analysisProgramIds: ["101", "102", "103", "104", "105", "106", "107", "108", "109", "110", "111", "112"]
  });
  const slide = deckSpec.slides.find((item) => item.id === "executive-summary");

  assert.match(slide.summary, /^Across all programs connected to topcashback, performance was mixed/i);
  assert.doesNotMatch(slide.summary, /selected programs|Across 12 programs/i);
}

async function executiveSummaryCorrectsProvidedSelectedProgramScope() {
  const deckSpec = await buildDeckSpec({
    executiveSummaryText: "Across 12 selected programs, performance was mixed in Jan 2026 - Apr 2026. Conversions moved -6.9%."
  });
  const slide = deckSpec.slides.find((item) => item.id === "executive-summary");

  assert.match(slide.summary, /^Across all programs connected to topcashback, performance was mixed/i);
  assert.doesNotMatch(slide.summary, /12 selected programs|selected programs/i);
}

async function executiveSummaryNormalizesDirectionalVarianceText() {
  const deckSpec = await buildDeckSpec({
    programYoYTable: [
      { Row: "Recent", Conversions: "140,846", "Order Value": "£22,309,036" },
      { Row: "Previous", Conversions: "151,294", "Order Value": "£31,966,933" },
      { Row: "Difference", Conversions: "-10,448", "Order Value": "-£9,657,897" },
      { Row: "Variance", Conversions: "▼ -6.9%", "Order Value": "â–¼ -30.2%" }
    ]
  });
  const slide = deckSpec.slides.find((item) => item.id === "executive-summary");
  const conversions = slide.kpis.find((card) => card.label === "Conversions");
  const orderValue = slide.kpis.find((card) => card.label === "Total Order Value");

  assert.equal(conversions.summary, "140,846 vs 151,294 PY -6.9%");
  assert.equal(conversions.delta, "-6.9%");
  assert.equal(orderValue.summary, "£22,309,036 vs £31,966,933 PY -30.2%");
  assert.equal(orderValue.delta, "-30.2%");
  assert.doesNotMatch(JSON.stringify(slide), /�|â|Â|▲|▼|↗|↘/);
}

async function renderedExecutiveSummaryDoesNotUseInvalidRichTextParagraphs() {
  const result = await buildPresentationResult();
  const zip = await JSZip.loadAsync(result.buffer);
  const executiveSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "executive-summary") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${executiveSlideIndex}.xml`).async("string");
  const paragraphs = slideXml.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];
  const invalidRichTextParagraphs = paragraphs.filter((paragraph) => /<a:r>[\s\S]*<a:pPr/.test(paragraph));

  assert.equal(invalidRichTextParagraphs.length, 0);
}

function assertColoredText(slideXml, text, color) {
  const escaped = text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`<a:srgbClr val="${color}"[\\s\\S]{0,600}<a:t>${escaped}<\\/a:t>`);
  assert.match(slideXml, pattern);
}

async function renderedExecutiveSummaryKpiDeltasUseTrafficLightColors() {
  const result = await buildPresentationResult({
    programYoYTable: [
      {
        Row: "Recent",
        Conversions: "93,107",
        "Conv Rate": "25.7%",
        Clicks: "2,500",
        "Earnings per Click": "GBP 0.40",
        "Earnings per Commission": "1.33x",
        "Order Value": "Â£15,913,142",
        "Publisher Commission": "Â£706,981",
        "Digital Wallet": "Â£67,578",
        "Total Earnings": "Â£774,559",
        "Active Programs": "4"
      },
      {
        Row: "Previous",
        Conversions: "98,729",
        "Conv Rate": "23.4%",
        Clicks: "2,000",
        "Earnings per Click": "GBP 0.40",
        "Earnings per Commission": "1.14x",
        "Order Value": "Â£15,913,142",
        "Publisher Commission": "Â£966,077",
        "Digital Wallet": "Â£158,760",
        "Total Earnings": "Â£1,124,837",
        "Active Programs": "4"
      },
      {
        Row: "Difference",
        Conversions: "-5,622",
        "Conv Rate": "2.3pp",
        Clicks: "500",
        "Earnings per Click": "GBP 0.00",
        "Earnings per Commission": "0.19x",
        "Order Value": "Â£0",
        "Publisher Commission": "-Â£259,096",
        "Digital Wallet": "-Â£91,182",
        "Total Earnings": "-Â£350,278",
        "Active Programs": "0"
      },
      {
        Row: "Variance",
        Conversions: "-5.7%",
        "Conv Rate": "+9.8%",
        Clicks: "0.0%",
        "Earnings per Click": "0.0%",
        "Earnings per Commission": "+16.7%",
        "Order Value": "0.0%",
        "Publisher Commission": "-26.8%",
        "Digital Wallet": "-57.4%",
        "Total Earnings": "-31.1%",
        "Active Programs": "0.0%"
      }
    ]
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const executiveSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "executive-summary") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${executiveSlideIndex}.xml`).async("string");

  assertColoredText(slideXml, "-5.7%", "EB5757");
  assertColoredText(slideXml, "0.0%", "F2C94C");
  assertColoredText(slideXml, "+9.8%", "57A66C");
}

async function renderedExecutiveSummaryKpiDeltasStayColoredWhenLocalizedSummaryDiffers() {
  const result = await buildPresentationResult({
    languageCode: "PL",
    languageName: "Polish",
    programYoYTable: [
      {
        Row: "Recent",
        Conversions: "93,107",
        "Conv Rate": "25.7%",
        Clicks: "2,500",
        "Earnings per Click": "GBP 0.40",
        "Order Value": "£15,913,142",
        "Publisher Commission": "£706,981",
        "Digital Wallet": "£67,578",
        "Total Earnings": "£774,559",
        "Active Programs": "4"
      },
      {
        Row: "Previous",
        Conversions: "98,729",
        "Conv Rate": "23.4%",
        Clicks: "2,000",
        "Earnings per Click": "GBP 0.40",
        "Order Value": "£24,652,989",
        "Publisher Commission": "£966,077",
        "Digital Wallet": "£158,760",
        "Total Earnings": "£1,124,837",
        "Active Programs": "4"
      },
      {
        Row: "Difference",
        Conversions: "-5,622",
        "Conv Rate": "2.3pp",
        Clicks: "500",
        "Earnings per Click": "GBP 0.00",
        "Order Value": "-£8,739,847",
        "Publisher Commission": "-£259,096",
        "Digital Wallet": "-£91,182",
        "Total Earnings": "-£350,278",
        "Active Programs": "0"
      },
      {
        Row: "Variance",
        Conversions: "-5.7%",
        "Conv Rate": "+9.8%",
        Clicks: "0.0%",
        "Earnings per Click": "0.0%",
        "Order Value": "-35.5%",
        "Publisher Commission": "-26.8%",
        "Digital Wallet": "-57.4%",
        "Total Earnings": "-31.1%",
        "Active Programs": "0.0%"
      }
    ]
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const executiveSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "executive-summary") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${executiveSlideIndex}.xml`).async("string");

  assert.match(slideXml, /<a:srgbClr val="EB5757"[\s\S]{0,700}<a:t>-\d+(?:[.,]\d+)?%<\/a:t>/);
  assert.match(slideXml, /<a:srgbClr val="57A66C"[\s\S]{0,700}<a:t>\+\d+(?:[.,]\d+)?%<\/a:t>/);
}

async function kpiSummaryTableUsesMetricRowsAndAddsRequestedMetrics() {
  const deckSpec = await buildDeckSpec();
  const slide = deckSpec.slides.find((item) => item.id === "kpi-volume-conversion");
  const table = slide.tables[0];
  const metricNames = table.rows.map((row) => row[0]);
  const earningsPerConversion = table.rows.find((row) => row[0] === "Earnings per Conversion");

  assert.deepEqual(table.columns, ["Metric", "Recent", "Previous", "Difference", "% Variance"]);
  assert.deepEqual(metricNames, [
    "Conversions",
    "Conversion Rate",
    "Clicks",
    "Earnings per Click",
    "Earnings per Conversion",
    "Total Order Value",
    "Publisher Commission",
    "Digital Wallet",
    "Total Earnings",
    "Active Programs"
  ]);
  assert(metricNames.every((label) => !/roi|aov|average order value/i.test(label)));
  assert.deepEqual(earningsPerConversion, ["Earnings per Conversion", "£6.00", "£7.00", "-£1.00", "-14.3%"]);
}

async function renderedTablesApplyTrafficLightColorsToVarianceCells() {
  const result = await buildPresentationResult({
    programYoYTable: [
      {
        Row: "Recent",
        Conversions: "93,107",
        "Conv Rate": "25.7%",
        Clicks: "2,500",
        "Earnings per Click": "GBP 0.40",
        "Order Value": "£15,913,142",
        "Publisher Commission": "£706,981",
        "Digital Wallet": "£67,578",
        "Total Earnings": "£774,559",
        "Active Programs": "4"
      },
      {
        Row: "Previous",
        Conversions: "98,729",
        "Conv Rate": "23.4%",
        Clicks: "2,000",
        "Earnings per Click": "GBP 0.40",
        "Order Value": "£15,913,142",
        "Publisher Commission": "£966,077",
        "Digital Wallet": "£158,760",
        "Total Earnings": "£1,124,837",
        "Active Programs": "4"
      },
      {
        Row: "Difference",
        Conversions: "-5,622",
        "Conv Rate": "2.3pp",
        Clicks: "500",
        "Earnings per Click": "GBP 0.00",
        "Order Value": "£0",
        "Publisher Commission": "-£259,096",
        "Digital Wallet": "-£91,182",
        "Total Earnings": "-£350,278",
        "Active Programs": "0"
      },
      {
        Row: "Variance",
        Conversions: "-5.7%",
        "Conv Rate": "+9.8%",
        Clicks: "0.0%",
        "Earnings per Click": "0.0%",
        "Order Value": "0.0%",
        "Publisher Commission": "-26.8%",
        "Digital Wallet": "-57.4%",
        "Total Earnings": "-31.1%",
        "Active Programs": "0.0%"
      }
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
  const zip = await JSZip.loadAsync(result.buffer);
  const kpiSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "kpi-volume-conversion") + 1;
  const kpiXml = await zip.file(`ppt/slides/slide${kpiSlideIndex}.xml`).async("string");

  assertColoredText(kpiXml, "-5.7%", "EB5757");
  assertColoredText(kpiXml, "0.0%", "F2C94C");
  assertColoredText(kpiXml, "+9.8%", "57A66C");
}

async function renderedTablesCenterHeadersAndValues() {
  const result = await buildPresentationResult();
  const zip = await JSZip.loadAsync(result.buffer);
  let tableCount = 0;

  for (let index = 1; index <= result.deckSpec.slides.length; index += 1) {
    const slideXml = await zip.file(`ppt/slides/slide${index}.xml`).async("string");
    const tableBlocks = slideXml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/g) || [];
    tableBlocks.forEach((tableXml) => {
      tableCount += 1;
      assert.match(tableXml, /algn="ctr"/);
      assert.doesNotMatch(tableXml, /algn="l"|algn="r"/);
    });
  }

  assert(tableCount > 0);
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

async function renderedActivationSnapshotUsesOriginalLineDivider() {
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
  const extents = [...slideXml.matchAll(/<a:ext cx="(\d+)" cy="(\d+)"/g)]
    .map((match) => ({ cx: Number(match[1]), cy: Number(match[2]) }));
  const emuPerInch = 914400;
  const hasAssetDivider = extents.some((extent) => (
    Math.abs(extent.cx - Math.round(4.72 * emuPerInch)) <= 8
      && Math.abs(extent.cy - Math.round(3.34 * emuPerInch)) <= 8
  ));

  assert.equal(hasAssetDivider, false);
  assert.match(slideXml, /prst="line"/);
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

  assert.equal(slide.chart.type, "weekly-pub-comm-line");
  assert.deepEqual(slide.chart.categories, ["2026-01-05", "2026-01-12"]);
  assert.equal(slide.chart.series[0].renderAs, "line");
  assert.equal(slide.chart.series[0].label, "Your Site");
  assert.equal(slide.chart.series[1].renderAs, "line");
  assert.equal(slide.chart.series[1].label, "Publisher 1");
  assert.deepEqual(slide.chart.series[0].values, [80000, 75000]);
  assert.deepEqual(slide.chart.series[1].values, [18500, 21000]);
}

async function competitorWeeklyComboChartPreservesMissingValuesWithoutZeroing() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      competitorAnalysisTable: [
        {
          Metric: "Programs with Pub Comm",
          "Your Site": "2",
          "Publisher 1": "1"
        }
      ],
      competitorWeeklyPubCommChart: [
        {
          Week: "2026-01-05",
          "Your Site": "GBP 1,000",
          "Publisher 1": "GBP 500",
          "Publisher 2": ""
        },
        {
          Week: "2026-01-12",
          "Your Site": "",
          "Publisher 1": "n/a",
          "Publisher 2": "GBP 250"
        },
        {
          Week: "2026-01-19",
          "Your Site": "GBP 3,000",
          "Publisher 1": "GBP 750",
          "Publisher 2": "-"
        }
      ]
    }
  });
  const slide = deckSpec.slides.find((item) => item.id === "competitor-analysis");

  assert.deepEqual(slide.chart.series[0].values, [1000, null, 3000]);
  assert.deepEqual(slide.chart.series[1].values, [500, null, 750]);
  assert.deepEqual(slide.chart.series[2].values, [null, 250, null]);
}

async function renderedCompetitorWeeklyChartMasksTemplateArtwork() {
  const result = await buildPresentationResult({
    publisherTables: {
      competitorAnalysisTable: [
        {
          Metric: "Programs with Pub Comm",
          "Your Site": "2",
          "Publisher 1": "1"
        }
      ],
      competitorWeeklyPubCommChart: [
        {
          Week: "2026-01-05",
          "Your Site": "GBP 1,000",
          "Publisher 1": "GBP 500"
        },
        {
          Week: "2026-01-12",
          "Your Site": "GBP 2,000",
          "Publisher 1": "GBP 750"
        }
      ]
    }
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const competitorSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "competitor-analysis") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${competitorSlideIndex}.xml`).async("string");

  assert.equal(slideHasFilledShapeExtent(slideXml, 12.18, 3.42, "F3F4F6"), true);
  assert.equal(slideHasFilledShapeExtent(slideXml, 11.52, 2.64, "F3F4F6"), true);
}

async function renderedCompetitorWeeklyChartConnectsAcrossMissingLineValues() {
  const result = await buildPresentationResult({
    publisherTables: {
      competitorAnalysisTable: [
        {
          Metric: "Programs with Pub Comm",
          "Your Site": "2",
          "Publisher 1": "1"
        }
      ],
      competitorWeeklyPubCommChart: [
        {
          Week: "2026-01-05",
          "Your Site": "GBP 1,000",
          "Publisher 1": "GBP 500"
        },
        {
          Week: "2026-01-12",
          "Your Site": "GBP 1,500",
          "Publisher 1": ""
        },
        {
          Week: "2026-01-19",
          "Your Site": "GBP 2,000",
          "Publisher 1": "GBP 750"
        }
      ]
    }
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const competitorSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "competitor-analysis") + 1;
  const chartXmls = await chartXmlsForSlide(zip, competitorSlideIndex);
  const chartXml = chartXmls.join("\n");

  assert.equal(chartXmls.length, 1);
  assert.match(chartXml, /<c:lineChart>/);
  assert.match(chartXml, /<c:dispBlanksAs val="span"\/>/);
  assert.match(chartXml, /<c:pt idx="1"><c:v><\/c:v><\/c:pt>/);
}

async function renderedCompetitorWeeklyChartUsesNativeContinuousLineChart() {
  const result = await buildPresentationResult({
    publisherTables: {
      competitorAnalysisTable: [
        {
          Metric: "Programs with Pub Comm",
          "Your Site": "2",
          "Publisher 1": "1",
          "Publisher 2": "1",
          "Publisher 3": "1",
          "Publisher 4": "1"
        }
      ],
      competitorWeeklyPubCommChart: [
        {
          Week: "2026-01-05",
          "Your Site": "GBP 1,000",
          "Publisher 1": "GBP 500",
          "Publisher 2": "GBP 300",
          "Publisher 3": "GBP 200",
          "Publisher 4": "GBP 100"
        },
        {
          Week: "2026-01-12",
          "Your Site": "GBP 1,500",
          "Publisher 1": "",
          "Publisher 2": "GBP 350",
          "Publisher 3": "GBP 250",
          "Publisher 4": "GBP 120"
        },
        {
          Week: "2026-01-19",
          "Your Site": "GBP 2,000",
          "Publisher 1": "GBP 750",
          "Publisher 2": "GBP 400",
          "Publisher 3": "GBP 260",
          "Publisher 4": "GBP 130"
        }
      ]
    }
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const competitorSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "competitor-analysis") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${competitorSlideIndex}.xml`).async("string");
  const chartXmls = await chartXmlsForSlide(zip, competitorSlideIndex);
  const chartXml = chartXmls.join("\n");

  assert.equal(chartXmls.length, 1);
  assert.match(chartXml, /<c:lineChart>/);
  assert.match(chartXml, /<c:dispBlanksAs val="span"\/>/);
  assert.match(chartXml, /<c:symbol val="none"\/>/);
  assert.doesNotMatch(chartXml, /<c:barChart>/);
  assert.equal(chartSeriesLineColor(chartXml, "Publisher 4"), "EB5757");
  assert.doesNotMatch(slideXml, /prst="line"[\s\S]{0,500}F28E2B/);
  assert.equal(slideHasEllipseColor(slideXml, "F28E2B"), false);
}

async function topNewProgramsMovesImmediatelyAfterMoversAndCompetitorShareFollowsAnalysis() {
  const deckSpec = await buildDeckSpec({
    publisherTables: {
      competitorAnalysisTable: [
        {
          "Competitor Group Summary": "Pub Comm of the above PP",
          "Your Site": "630",
          "Publisher 1": "180",
          "Publisher 2": "110",
          "Publisher 3": "30",
          "Publisher 4": "40",
          "Distinct comp. prog. #": "24"
        },
        {
          "Competitor Group Summary": "Pub Comm of the above",
          "Your Site": "640",
          "Publisher 1": "140",
          "Publisher 2": "120",
          "Publisher 3": "50",
          "Publisher 4": "50",
          "Distinct comp. prog. #": "24"
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
  const moversIndex = ids.indexOf("movers-shakers-publisher-commission");
  const brandNewIndex = ids.indexOf("brand-new-publishers");
  const competitorAnalysisIndex = ids.indexOf("competitor-analysis");
  const shareIndex = ids.indexOf("competitor-share-pub-comm");
  const shareSlide = deckSpec.slides[shareIndex];

  assert.equal(brandNewIndex, moversIndex + 1);
  assert.equal(competitorAnalysisIndex, 8);
  assert.equal(shareIndex, competitorAnalysisIndex + 1);
  assert.equal(shareSlide.kind, "competitor-share-bar-chart");
  assert.deepEqual(shareSlide.chart.categories, ["Your Site", "Comp. A", "Comp. B", "Comp. C", "Comp. D"]);
  assert.deepEqual(shareSlide.chart.series.map((series) => series.label), ["% of Site Group PP", "% of Site Group"]);
  assert.deepEqual(shareSlide.chart.series[0].values, [63.6, 18.2, 11.1, 3, 4]);
  assert.deepEqual(shareSlide.chart.series[1].values, [64, 14, 12, 5, 5]);
}

async function programConnectionStatusSlideFollowsTopNewProgramsAndRendersKey() {
  const result = await buildPresentationResult({
    programStatusCreatedFromDate: "2025-11-19",
    publisherTables: {
      brandNewProgramsTable: [
        {
          "Program ID": "392254",
          "Program Name": "Everpress",
          Conversions: "1",
          "Order Value": "GBP 100",
          "Publisher Commission": "GBP 10",
          "Digital Wallet": "GBP 0",
          "Total Earnings": "GBP 10"
        }
      ],
      programConnectionStatusTable: [
        {
          "Program ID": "392254",
          "Program Name": "Everpress",
          "Status ID": "3",
          "Connection Status": "Accepted"
        },
        {
          "Program ID": "392637",
          "Program Name": "Atlanta Braves",
          "Status ID": "0",
          "Connection Status": "Not Connected"
        },
        {
          "Program ID": "392847",
          "Program Name": "Bosch DIY UK",
          "Status ID": "2",
          "Connection Status": "Hold UC"
        },
        {
          "Program ID": "397866",
          "Program Name": "York - MSc Computer Science",
          "Status ID": "0",
          "Connection Status": "Not Connected"
        }
      ]
    }
  });
  const ids = result.deckSpec.slides.map((slide) => slide.id);
  const statusIndex = ids.indexOf("program-connection-status");
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file(`ppt/slides/slide${statusIndex + 1}.xml`).async("string");

  assert.equal(statusIndex, ids.indexOf("brand-new-publishers") + 1);
  assert.equal(ids[statusIndex + 1], "competitor-analysis");
  assert.deepEqual(
    result.deckSpec.slides[statusIndex].programConnectionStatus.rows.map((row) => row.programName),
    ["Everpress", "Bosch DIY UK", "Atlanta Braves", "York - MSc Computer Science"]
  );
  assert.deepEqual(
    result.deckSpec.slides[statusIndex].programConnectionStatus.rows.map((row) => row.status),
    ["Accepted", "Hold Under Consideration", "Not Connected", "Not Connected"]
  );
  assert.match(slideXml, /Program Connection Status/);
  assert.doesNotMatch(slideXml, /Programs created since the three-month lookback cutoff/);
  assert.match(slideXml, /4 programs reviewed/);
  assert.match(slideXml, /Created from 19 November 2025/);
  assert.match(slideXml, /Everpress/);
  assert.match(slideXml, /Atlanta Braves/);
  assert.match(slideXml, /Bosch DIY UK/);
  assert.match(slideXml, /York - MSc Computer Science/);
  assert.doesNotMatch(slideXml, /York - MSc Computer\.\.\./);
  assert.match(slideXml, /0 Not Conn\./);
  assert.match(slideXml, /2 Hold UC/);
  assert.match(slideXml, /3 Accepted/);
  assert.match(slideXml, /<a:srgbClr val="57A66C"/);
  assert.match(slideXml, /<a:srgbClr val="8A94A6"/);
  assert.match(slideXml, /<a:srgbClr val="F2C94C"/);
}

async function gapAnalysisImpactSlideUsesAssetUnderlayAndOmitsBottomBreakdown() {
  const result = await buildPresentationResult({
    client: "TopCashBack",
    publisherTables: {
      programGapAnalysisTable: [
        {
          "Program Name": "Philips Hue UK - AFF",
          "Program ID": "354193",
          TopCashBack: "Accepted",
          "Comp A": "Pub Comm",
          "Comp B": "Pub Comm",
          "Comp C": "Pub Comm",
          "Comp D": "Pub Comm",
          "Competitor Pub Comm": "£3,500",
          "Gap Type": "Activation",
          "Recommended Action": "Activate"
        },
        {
          "Program Name": "SupremeCBD",
          "Program ID": "335495",
          TopCashBack: "No Connection",
          "Comp A": "Pub Comm",
          "Comp B": "No Connection",
          "Comp C": "No Connection",
          "Comp D": "No Connection",
          "Competitor Pub Comm": "£1,000",
          "Gap Type": "Application",
          "Recommended Action": "Apply"
        },
        {
          "Program Name": "HP Store",
          "Program ID": "21701",
          TopCashBack: "Clicks",
          "Comp A": "Pub Comm",
          "Comp B": "Clicks",
          "Comp C": "No Connection",
          "Comp D": "Denied",
          "Competitor Pub Comm": "£600",
          "Gap Type": "Click Leakage",
          "Recommended Action": "Fix tracking / conversion"
        }
      ],
      programConnectionStatusTable: [
        {
          "Program ID": "354193",
          "Program Name": "Philips Hue UK - AFF",
          "Status ID": "3",
          "Connection Status": "Accepted"
        }
      ]
    }
  });
  const slideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "gap-analysis-impact") + 1;
  const slide = result.deckSpec.slides[slideIndex - 1];
  const ids = result.deckSpec.slides.map((item) => item.id);
  const zip = await JSZip.loadAsync(result.buffer);
  const slideXml = await zip.file(`ppt/slides/slide${slideIndex}.xml`).async("string");
  const slideRels = await zip.file(`ppt/slides/_rels/slide${slideIndex}.xml.rels`).async("string");

  assert(slideIndex > 0);
  assert.equal(ids[slideIndex - 2], "brand-new-publishers");
  assert.equal(ids[slideIndex], "gap-analysis-register");
  assert.equal(slide.kind, "gap-analysis-impact");
  assert.match(slide.title, /Growth opportunity/i);
  assert.equal(slide.kpis[0].value, "£5.1k");
  assert.equal(slide.kpis[1].value, "3");
  assert.equal(slide.kpis[2].value, "1");
  assert.match(slideXml, /<a:t>£5\.1k<\/a:t>/);
  assert.match(slideRels, /Target="\.\.\/media\/image[-\d]+\.png"/);
  assert.equal(slideHasImageExtent(slideXml, 3.15, 3.15), false);
  assert.equal(slideHasImageExtent(slideXml, 4.65, 4.65), true);
  assert.doesNotMatch(slideXml, /<a:t>Activation<\/a:t>/);
  assert.doesNotMatch(slideXml, /<a:t>Application<\/a:t>/);
  assert.doesNotMatch(slideXml, /<a:t>Click leakage<\/a:t>/);
  assert.doesNotMatch(slideXml, /<a:t>Recovery<\/a:t>/);
}

async function gapAnalysisSectionUsesThreeColumnRegisterAndPaginatesLargeLists() {
  const gapRows = Array.from({ length: 80 }, (_, index) => {
    const status = index % 4 === 0
      ? "Accepted"
      : index % 4 === 1
        ? "Clicks"
        : index % 4 === 2
          ? "No Connection"
          : "Ended";
    const gapType = status === "Accepted"
      ? "Activation"
      : status === "Clicks"
        ? "Click Leakage"
        : status === "No Connection"
          ? "Application"
          : "Recovery";
    return {
      "Program Name": index === 0 ? "Philips Hue UK - AFF" : `Gap Program ${String(index + 1).padStart(2, "0")}`,
      "Program ID": String(350000 + index),
      TopCashBack: status,
      "Comp A": "Pub Comm",
      "Comp B": index % 3 === 0 ? "Pub Comm" : "Accepted",
      "Comp C": index % 5 === 0 ? "Pub Comm" : "No Connection",
      "Comp D": index % 7 === 0 ? "Pub Comm" : "Denied",
      "Competitor Pub Comm": `GBP ${3500 - index * 75}`,
      "Gap Type": gapType,
      "Recommended Action": gapType === "Activation" ? "Activate accepted program" : "Build outreach plan"
    };
  });
  const result = await buildPresentationResult({
    client: "TopCashBack",
    publisherTables: {
      programGapAnalysisTable: gapRows,
      brandNewProgramsTable: [
        {
          "Program ID": "392254",
          "Program Name": "Everpress",
          Conversions: "1",
          "Order Value": "GBP 100",
          "Publisher Commission": "GBP 10",
          "Digital Wallet": "GBP 0",
          "Total Earnings": "GBP 10"
        }
      ],
      programConnectionStatusTable: [
        {
          "Program ID": "354193",
          "Program Name": "Philips Hue UK - AFF",
          "Status ID": "3",
          "Connection Status": "Accepted"
        }
      ]
    }
  });
  const ids = result.deckSpec.slides.map((slide) => slide.id);
  const impactIndex = ids.indexOf("gap-analysis-impact");
  const zip = await JSZip.loadAsync(result.buffer);
  const registerXml = await zip.file(`ppt/slides/slide${ids.indexOf("gap-analysis-register") + 1}.xml`).async("string");
  const registerPage2Xml = await zip.file(`ppt/slides/slide${ids.indexOf("gap-analysis-register-2") + 1}.xml`).async("string");
  const registerSlide = result.deckSpec.slides.find((slide) => slide.id === "gap-analysis-register");
  const registerSlide2 = result.deckSpec.slides.find((slide) => slide.id === "gap-analysis-register-2");

  assert.deepEqual(ids.slice(impactIndex, impactIndex + 4), [
    "gap-analysis-impact",
    "gap-analysis-register",
    "gap-analysis-register-2",
    "program-connection-status"
  ]);
  assert.equal(registerSlide.kind, "gap-analysis-register");
  assert.equal(registerSlide.gapRegister.rows.length, 66);
  assert.equal(registerSlide2.gapRegister.rows.length, 14);
  assert.equal(registerSlide.gapRegister.columnCount, 3);
  assert.equal(registerSlide.gapRegister.totalRows, 80);
  assert.equal(ids.includes("gap-analysis-priority-programs"), false);
  assert.equal(ids.includes("gap-analysis-portfolio"), false);
  assert.equal(ids.includes("gap-analysis-detail"), false);
  assert.match(registerXml, /Gap Analysis Register \(1\/2\)/);
  assert.match(registerXml, /Program \/ ID \/ Status \/ Pub Comm/);
  assert.match(registerXml, /Philips Hue UK - AFF/);
  assert.match(registerXml, /350000/);
  assert.match(registerXml, /£3,500/);
  assert.match(registerPage2Xml, /Gap Analysis Register \(2\/2\)/);
  assert.match(registerPage2Xml, /Gap Program 80/);
  assert.doesNotMatch(registerXml, /Gap Program 80/);
  assert.doesNotMatch(registerXml, /Priority programs to close the gap/);
}

async function programConnectionStatusPaginatesLargeProgramLists() {
  const statusRows = Array.from({ length: 121 }, (_, index) => ({
    "Program ID": String(400000 + index),
    "Program Name": `Program ${index + 1}`,
    "Status ID": index % 3 === 0 ? "0" : "3",
    "Connection Status": index % 3 === 0 ? "Not Connected" : "Accepted"
  }));
  const deckSpec = await buildDeckSpec({
    programStatusCreatedFromDate: "2025-11-19",
    publisherTables: {
      brandNewProgramsTable: [
        {
          "Program ID": "392254",
          "Program Name": "Everpress",
          Conversions: "1",
          "Order Value": "GBP 100",
          "Publisher Commission": "GBP 10",
          "Digital Wallet": "GBP 0",
          "Total Earnings": "GBP 10"
        }
      ],
      programConnectionStatusTable: statusRows
    }
  });
  const statusSlides = deckSpec.slides.filter((slide) => slide.id.startsWith("program-connection-status"));

  assert.equal(statusSlides.length, 2);
  assert.equal(statusSlides[0].title, "Program Connection Status (1/2)");
  assert.equal(statusSlides[1].title, "Program Connection Status (2/2)");
  assert.equal(statusSlides[0].programConnectionStatus.rows.length, 84);
  assert.equal(statusSlides[1].programConnectionStatus.rows.length, 37);
  assert.equal(statusSlides[0].programConnectionStatus.totalRows, 121);
  assert.equal(statusSlides[0].programConnectionStatus.cutoffDate, "2025-11-19");
}

async function renderedCompetitorShareChartUsesThemeBarsAndLabels() {
  const result = await buildPresentationResult({
    publisherTables: {
      competitorAnalysisTable: [
        {
          "Competitor Group Summary": "Pub Comm of the above PP",
          "Your Site": "630",
          "Publisher 1": "180",
          "Publisher 2": "110",
          "Publisher 3": "30",
          "Publisher 4": "40"
        },
        {
          "Competitor Group Summary": "Pub Comm of the above",
          "Your Site": "640",
          "Publisher 1": "140",
          "Publisher 2": "120",
          "Publisher 3": "50",
          "Publisher 4": "50"
        }
      ]
    }
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const shareSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "competitor-share-pub-comm") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${shareSlideIndex}.xml`).async("string");

  assert.match(slideXml, /Share of publisher commission within competitor group/);
  assert.match(slideXml, /Your Site/);
  assert.match(slideXml, /Comp\. A/);
  assert.match(slideXml, /% of Site Group PP/);
  assert.match(slideXml, /% of Site Group/);
  assert.match(slideXml, /<a:srgbClr val="2F6FF2"/);
  assert.match(slideXml, /64%/);
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
  assert(ids.includes("competitor-share-pub-comm"));
  assert(ids.includes("top-programs-competitor-performance"));
  assert.equal(ids.includes("sales-growth-signals"), false);
  assert(ids.includes("risks-dependencies"));
  assert(ids.includes("brand-new-publishers"));
  assert(ids.includes("thank-you"));
  assert.deepEqual(
    ids.slice(ids.indexOf("movers-shakers-publisher-commission"), ids.indexOf("top-programs-competitor-performance") + 1),
    [
      "movers-shakers-publisher-commission",
      "brand-new-publishers",
      "competitor-analysis",
      "competitor-share-pub-comm",
      "top-programs-competitor-performance"
    ]
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

async function riskDependenciesUseEvidenceBackedMitigations() {
  const deckSpec = await buildDeckSpec({
    publisherOverviewObservations: [
      "10 brand-new programs were activated, contributing £0 in combined OV."
    ],
    publisherTables: {
      riskDependenciesTable: [
        {
          "Program Name": "Decline Retailer",
          "Risk Type": "YoY decline",
          Evidence: "Conversions -20, total earnings -GBP 500",
          Priority: "High"
        },
        {
          "Program Name": "Traffic Retailer",
          "Risk Type": "High traffic, zero conversions",
          Evidence: "Traffic events 2,500, conversions 0",
          Priority: "Medium"
        }
      ],
      brandNewProgramsTable: [
        {
          "Program ID": "111",
          "Program Name": "New Zero One",
          Conversions: "0",
          "Order Value": "GBP 0",
          "Publisher Commission": "GBP 0",
          "Digital Wallet": "GBP 0",
          "Total Earnings": "GBP 0"
        },
        {
          "Program ID": "222",
          "Program Name": "New Zero Two",
          Conversions: "0",
          "Order Value": "GBP 0",
          "Publisher Commission": "GBP 0",
          "Digital Wallet": "GBP 0",
          "Total Earnings": "GBP 0"
        }
      ]
    }
  });

  const riskSlide = deckSpec.slides.find((slide) => slide.id === "risks-dependencies");
  const table = riskSlide.tables[0];
  const rowText = table.rows.map((row) => row.join(" ")).join(" ");

  assert.deepEqual(table.columns, ["Risk / rationale"]);
  assert.match(rowText, /Decline Retailer.*Conversions -20, total earnings -GBP 500/i);
  assert.match(rowText, /Traffic Retailer.*Traffic events 2,500, conversions 0/i);
  assert.match(rowText, /New Zero One, New Zero Two/i);
  assert.match(rowText, /Action:/i);
  assert.match(rowText, /small controlled traffic sample|visible launch placement/i);
  assert.doesNotMatch(rowText, /confirm the root cause|assign an owner|agree a recovery action/i);
  assert.doesNotMatch(rowText, /Order value mix volatility.*brand-new programs were activated/i);
}

async function riskDependenciesFallbackRowsKeepSpecificAnalysisEvidence() {
  const deckSpec = await buildDeckSpec({
    publisherOverviewObservations: [
      "Publisher commission concentration: Currys generated Â£180,000 of the Â£240,000 commission total, creating dependency if its cashback placement drops.",
      "Conversion pressure: Argos traffic increased 34% but conversions fell 18%, indicating landing-page or tracking leakage."
    ],
    publisherTables: {}
  });
  const riskSlide = deckSpec.slides.find((slide) => slide.id === "risks-dependencies");
  const rowText = riskSlide.tables[0].rows.map((row) => row.join(" ")).join(" ");

  assert.match(rowText, /Currys generated £180,000 of the £240,000 commission total/i);
  assert.match(rowText, /Argos traffic increased 34% but conversions fell 18%/i);
  assert.match(rowText, /Protect the leading program's placement/i);
  assert.match(rowText, /Audit the affected program's landing path/i);
  assert(rowText.length < 1100);
  assert.doesNotMatch(rowText, /Program concentration risk High Reduce reliance on the exposed program set/i);
  assert.doesNotMatch(rowText, /confirm the root cause|assign an owner|agree a recovery action/i);
}

async function renderedRiskDependenciesUsesLightTileLayoutWithoutGenericColumns() {
  const result = await buildPresentationResult({
    publisherOverviewObservations: [
      "Publisher commission concentration: Currys generated £180,000 of the £240,000 commission total, creating dependency if its cashback placement drops.",
      "Conversion pressure: Argos traffic increased 34% but conversions fell 18%, indicating landing-page or tracking leakage.",
      "Order value mix volatility: Skimlinks added +£590,200 in OV but lower-value programs declined.",
      "Activation dependency: two new programs show 0 order value and 0 conversions."
    ],
    publisherTables: {}
  });
  const zip = await JSZip.loadAsync(result.buffer);
  const riskSlideIndex = result.deckSpec.slides.findIndex((slide) => slide.id === "risks-dependencies") + 1;
  const slideXml = await zip.file(`ppt/slides/slide${riskSlideIndex}.xml`).async("string");

  assert.doesNotMatch(slideXml, /<a:tbl>/);
  assert.doesNotMatch(slideXml, /<a:t>Impact<\/a:t>|<a:t>Mitigation<\/a:t>/);
  assert.doesNotMatch(slideXml, /…|\.\.\./);
  assert.match(slideXml, /<a:srgbClr val="F3F4F6"/);
  assert.match(slideXml, /<a:srgbClr val="D7E4FF"/);
  assert.match(slideXml, /<a:srgbClr val="2F6FF2"/);
  assert.doesNotMatch(slideXml, /<p:bgPr><a:solidFill><a:srgbClr val="2F6FF2"/);
  assert.doesNotMatch(slideXml, /<a:t>01<\/a:t>|<a:t>02<\/a:t>|<a:t>03<\/a:t>/);
  assert.match(slideXml, /<a:pPr algn="ctr"/);
  assert.match(slideXml, /<a:t>Analysis<\/a:t>/);
  assert.match(slideXml, /<a:t>Action<\/a:t>/);
  assert.match(slideXml, /Protect the leading program&apos;s placement|Audit the affected program&apos;s landing path|Prioritise higher-basket placements/i);
  assert.doesNotMatch(slideXml, /confirm the root cause|assign an owner|agree a recovery action/i);
  assert.match(slideXml, /Currys generated £180,000/);
  assert.match(slideXml, /Argos traffic increased 34%/);
}

async function run() {
  const tests = [
    coverUsesPublisherPerformanceReviewTitle,
    renderedCoverUsesWhiteTdLogoAsset,
    renderedCoverUsesCyanWireframeAndNoUtilityPills,
    renderedThankYouUsesLargerCyanWireframeAndCompactQuestionBubble,
    executiveSummaryUsesRequestedPublisherMetrics,
    executiveSummaryDescribesFullPublisherProgramScope,
    executiveSummaryCorrectsProvidedSelectedProgramScope,
    executiveSummaryNormalizesDirectionalVarianceText,
    renderedExecutiveSummaryDoesNotUseInvalidRichTextParagraphs,
    renderedExecutiveSummaryKpiDeltasUseTrafficLightColors,
    renderedExecutiveSummaryKpiDeltasStayColoredWhenLocalizedSummaryDiffers,
    kpiSummaryTableUsesMetricRowsAndAddsRequestedMetrics,
    renderedTablesApplyTrafficLightColorsToVarianceCells,
    renderedTablesCenterHeadersAndValues,
    programActivationSnapshotSlideFollowsKpiSummary,
    renderedActivationSnapshotUsesTdLineGridStyle,
    renderedActivationSnapshotUsesOriginalLineDivider,
    renderedSlideUtilityTextIsRemovedAndTableTitlesAreCentered,
    programLevelAnalysisUsesPublisherCommissionHierarchy,
    topNewProgramsIncludesWalletAndSortsByTotalEarnings,
    moversShakersUsesPublisherCommissionBarChart,
    moversShakersDisplaysDeclinersSmallestLossFirst,
    renderedMoversShakersNegativeBarsUseRedFill,
    competitorAnalysisSlideUsesAnonymousComparisonTable,
    competitorAnalysisSlideIncludesWeeklyComboChart,
    competitorWeeklyComboChartPreservesMissingValuesWithoutZeroing,
    renderedCompetitorWeeklyChartMasksTemplateArtwork,
    renderedCompetitorWeeklyChartConnectsAcrossMissingLineValues,
    renderedCompetitorWeeklyChartUsesNativeContinuousLineChart,
    topNewProgramsMovesImmediatelyAfterMoversAndCompetitorShareFollowsAnalysis,
    gapAnalysisImpactSlideUsesAssetUnderlayAndOmitsBottomBreakdown,
    gapAnalysisSectionUsesThreeColumnRegisterAndPaginatesLargeLists,
    programConnectionStatusSlideFollowsTopNewProgramsAndRendersKey,
    programConnectionStatusPaginatesLargeProgramLists,
    renderedCompetitorShareChartUsesThemeBarsAndLabels,
    productionPublisherRequestKeepsFullPublisherSlideSet,
    publisherQbrAnalysisUsesProgramLanguage,
    riskDependenciesUseEvidenceBackedMitigations,
    riskDependenciesFallbackRowsKeepSpecificAnalysisEvidence,
    renderedRiskDependenciesUsesLightTileLayoutWithoutGenericColumns
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
