const fs = require("node:fs/promises");
const fsSync = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const https = require("node:https");

function loadPackage(name) {
  try {
    return require(name);
  } catch (error) {
    const fallback = path.join(__dirname, "..", "..", "presentation-ai", "node_modules", name);
    return require(fallback);
  }
}

const PptxGenJS = loadPackage("pptxgenjs");
const JSZip = loadPackage("jszip");
const TEMPLATE_BLUE_BG_PATH = path.join(__dirname, "..", "assets", "qbr-bg-blue.png");
const TEMPLATE_LIGHT_BG_PATH = path.join(__dirname, "..", "assets", "qbr-bg-light.png");
const TD_LOGO_WHITE_PATH = path.join(__dirname, "..", "assets", "td-logo-white.png");
const TD_FIFTH_ELEMENT_WHITE_PATH = path.join(__dirname, "..", "assets", "fifth-element-white.png");
const TD_FIFTH_ELEMENT_WIREFRAME_CYAN_PATH = path.join(__dirname, "..", "assets", "fifth-element-wireframe-cyan.png");
const HAS_TEMPLATE_BLUE_BG = fsSync.existsSync(TEMPLATE_BLUE_BG_PATH);
const HAS_TEMPLATE_LIGHT_BG = fsSync.existsSync(TEMPLATE_LIGHT_BG_PATH);
const HAS_TD_LOGO_WHITE = fsSync.existsSync(TD_LOGO_WHITE_PATH);
const HAS_TD_FIFTH_ELEMENT_WHITE = fsSync.existsSync(TD_FIFTH_ELEMENT_WHITE_PATH);
const HAS_TD_FIFTH_ELEMENT_WIREFRAME_CYAN = fsSync.existsSync(TD_FIFTH_ELEMENT_WIREFRAME_CYAN_PATH);
const KPI_ICON_PATHS = {
  sales: path.join(__dirname, "..", "assets", "kpi-icon-sales.png"),
  ordervalue: path.join(__dirname, "..", "assets", "kpi-icon-ordervalue.png"),
  aov: path.join(__dirname, "..", "assets", "kpi-icon-aov.png"),
  convrate: path.join(__dirname, "..", "assets", "kpi-icon-convrate.png"),
  roi: path.join(__dirname, "..", "assets", "kpi-icon-roi.png"),
  commission: path.join(__dirname, "..", "assets", "kpi-icon-roi.png")
};
const HAS_KPI_ICON = Object.fromEntries(
  Object.entries(KPI_ICON_PATHS).map(([key, filePath]) => [key, fsSync.existsSync(filePath)])
);

const TEXT_REPLACEMENTS = [
  [/Ã‚Â£|ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â£|Ãâ€œÃ¢â‚¬Å¡Ãâ€™ÃË†|ÃË†/g, "\u00A3"],
  [/Ã‚â‚¬|ÃƒÆ’Ã‚Â¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬|Ãâ€œÃ¢â‚¬Å¡Ãâ€™Ã‚Â¤/g, "\u20AC"],
  [/Ã‚Â¥/g, "\u00A5"],
  [/Ã‚/g, ""],
  [/Ã‚Â /g, " "],
  [/Ã¢â‚¬â€œ|Ã¢â‚¬â€/g, "-"],
  [/Ã¢â‚¬Ëœ|Ã¢â‚¬â„¢/g, "'"],
  [/Ã¢â‚¬Å“|Ã¢â‚¬ï¿½/g, '"'],
  [/zÃ…â€š|zÃƒÆ’Ã¢â‚¬Â¦ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡/g, "z\u0142"]
];
const DIRECTION_MARK_REGEX = /[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE\u2191-\u2199]/g;
const MOJIBAKE_DIRECTION_MARK_REGEX = /(?:\u00E2\u2013[\u00B2\u00B3\u00B4\u00BC\u00BD\u00BE]|\u00E2\u2020[\u2018\u201C\u201D\u02DC])/g;

const TABLE_KEY_MAP = {
  top10increase: "topGrowthPublishers",
  topgrowthpublishers: "topGrowthPublishers",
  top10decrease: "topDecliningPublishers",
  topdecliningpublishers: "topDecliningPublishers",
  top10byov: "topCurrentPerformers",
  topcurrentperformers: "topCurrentPerformers",
  segmentsummary: "segmentSnapshot",
  moversshakerssales: "moversSales",
  moversshakersov: "moversOrderValue",
  moversshakersclicks: "moversClicks",
  moversshakersaov: "moversAov",
  moversshakerscommissionchart: "moversCommissionChart",
  moversshakerscommission: "moversCommission",
  moverscommissionchart: "moversCommissionChart",
  moverscommission: "moversCommission",
  brandnewtop: "brandNewPublishers",
  brandnewprogramstable: "brandNewPublishers",
  brandnewprograms: "brandNewPublishers",
  programconnectionstatustable: "programConnectionStatusTable",
  programconnectionstatus: "programConnectionStatusTable",
  newprogramconnectionstatustable: "programConnectionStatusTable",
  newprogramconnectionstatus: "programConnectionStatusTable",
  programgapanalysistable: "programGapAnalysisTable",
  programgapanalysis: "programGapAnalysisTable",
  gapanalysistable: "programGapAnalysisTable",
  gapanalysis: "programGapAnalysisTable",
  programgapanalysissummarytable: "programGapAnalysisSummaryTable",
  programgapanalysissummary: "programGapAnalysisSummaryTable",
  newemergingtop: "newEmergingPublishers",
  stoppedactivitytop: "stoppedActivity",
  newpublisherprospects: "newPublisherProspects",
  competitoranalysistable: "competitorAnalysisTable",
  competitorgroupsummary: "competitorAnalysisTable",
  competitorsharepubcommchart: "competitorSharePubCommChart",
  competitorsharepublishercommissionchart: "competitorSharePubCommChart",
  competitorsharechart: "competitorSharePubCommChart",
  competitorweeklypubcommchart: "competitorWeeklyPubCommChart",
  weeklypubcommchart: "competitorWeeklyPubCommChart",
  programactivationsnapshottable: "programActivationSnapshotTable",
  programactivationsnapshot: "programActivationSnapshotTable",
  activationsnapshottable: "programActivationSnapshotTable",
  activationsnapshot: "programActivationSnapshotTable",
  topprogramscompetitorperformancetable: "topProgramsCompetitorPerformanceTable",
  topprogramscompetitorperformance: "topProgramsCompetitorPerformanceTable",
  competitorperformancetable: "topProgramsCompetitorPerformanceTable",
  riskdependenciestable: "riskDependenciesTable",
  risksdependenciestable: "riskDependenciesTable",
  riskdependencies: "riskDependenciesTable",
  risksdependencies: "riskDependenciesTable"
};

const PROGRAM_BREAKDOWN_COLUMNS = [
  { label: "Program", aliases: ["Program", "Program Name", "ProgramName", "Name"] },
  { label: "Publisher Commission", aliases: ["Publisher Commission", "Current Publisher Commission", "Commission", "Current Commission"] },
  { label: "Digital Wallet", aliases: ["Digital Wallet", "Digital Wallets", "DigitalWallet", "DigitalWallets"] },
  { label: "Total Earnings", aliases: ["Total Earnings", "Total Earning", "TotalEarnings"] },
  { label: "Conversions", aliases: ["Conversions", "Current Conversions", "Sales", "Current Sales", "CurrentSales"] },
  { label: "Order Value", aliases: ["Order Value", "Total Order Value", "Current OV", "Current Order Value", "CurrentOrderValue", "CurrentOV"] },
  { label: "Publisher Commission YoY %", aliases: ["Publisher Commission YoY %", "Commission YoY %", "Publisher Commission % YoY", "Commission % YoY"] },
  { label: "Earnings YoY %", aliases: ["Earnings YoY %", "Total Earnings YoY %", "Total Earnings % YoY", "YoY %"] },
  { label: "Conversions YoY %", aliases: ["Conversions YoY %", "Current Conversions YoY %", "Sales YoY %", "Current Sales YoY %", "Sales % YoY", "SalesYoY%"] }
];

const DEFAULT_THEME = {
  id: "td-default",
  name: "TD",
  companyName: "Tradedoubler",
  logoText: "Tradedoubler",
  fonts: {
    heading: "Aptos",
    body: "Aptos",
    mono: "Aptos"
  },
  colors: {
    ink: "#2F333B",
    paper: "#F3F4F6",
    canvas: "#E8EDF9",
    accent: "#2F6FF2",
    accentAlt: "#EB5757",
    success: "#57A66C",
    warning: "#F2C94C",
    highlight: "#AFC4F5",
    muted: "#5B6372",
    border: "#D8DCE5"
  }
};

const LANGUAGE_LOCALE_MAP = {
  EN: "en-GB",
  FR: "fr-FR",
  NL: "nl-NL",
  DE: "de-DE",
  IT: "it-IT",
  NO: "nb-NO",
  SV: "sv-SE",
  DA: "da-DK",
  FI: "fi-FI",
  ES: "es-ES",
  PL: "pl-PL"
};

const LANGUAGE_TRANSLATION_TARGET_MAP = {
  EN: "en",
  FR: "fr",
  NL: "nl",
  DE: "de",
  IT: "it",
  NO: "no",
  SV: "sv",
  DA: "da",
  FI: "fi",
  ES: "es",
  PL: "pl"
};

const DEFAULT_UI_LABELS = {
  qbrReport: "QBR Report",
  anyQuestions: "Any Questions?",
  thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
  currentPeriod: "Current Period",
  comparisonPeriodYoy: "Comparison Period (YoY)",
  basisYoy: "Basis: Year-over-Year (YoY)",
  publisherActivityBySegment: "Publisher Activity by Segment",
  keyObservations: "Key Observations",
  reportingPeriodPrefix: "Reporting Period",
  dataAsOfPrefix: "Data as of",
  comparisonPeriodPrefix: "Comparison Period",
  allFiguresStatement: "All figures are reported in {currency} unless otherwise stated. YoY variance is calculated as Current Period vs Comparison Period.",
  analysisTagSuffix: "Analysis",
  segmentSignalUnavailable: "Segment signal not available.",
  detailedMovementUnavailable: "Detailed movement not available from this extract.",
  kpiSignalGeneric: "KPI Signal",
  kpiDriverUnavailable: "Driver not confirmed from available KPI data.",
  kpiDetailUnavailable: "Detail not available from current extract.",
  kpiTitleConversionRateImprovement: "Conversion Rate Improvement",
  kpiTitleSalesVolumePressure: "Sales Volume Pressure",
  kpiTitleAovGrowthOffset: "AOV Growth Partially Offsetting Volume Decline",
  kpiTitleRisingCpa: "Rising CPA",
  kpiTitleRoiTrend: "ROI Trend"
};

const UI_LABELS_BY_LANGUAGE = {
  FR: {
    qbrReport: "Rapport QBR",
    anyQuestions: "Des questions ?",
    thankYouSubtitleTemplate: "Revue de performance publisher TD - {period}",
    currentPeriod: "Période actuelle",
    comparisonPeriodYoy: "Période de comparaison (YoY)",
    basisYoy: "Référence : glissement annuel (YoY)",
    publisherActivityBySegment: "Activité des éditeurs par segment",
    keyObservations: "Observations clés",
    reportingPeriodPrefix: "Période de reporting",
    dataAsOfPrefix: "Données au",
    comparisonPeriodPrefix: "Période de comparaison",
    allFiguresStatement: "Toutes les valeurs sont présentées en {currency}, sauf indication contraire. La variation YoY est calculée entre la période actuelle et la période de comparaison.",
    analysisTagSuffix: "Analyse"
  },
  NL: {
    qbrReport: "QBR-rapport",
    anyQuestions: "Vragen?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Huidige periode",
    comparisonPeriodYoy: "Vergelijkingsperiode (YoY)",
    basisYoy: "Basis: jaar-op-jaar (YoY)",
    publisherActivityBySegment: "Publisheractiviteit per segment",
    keyObservations: "Belangrijkste observaties",
    reportingPeriodPrefix: "Rapportageperiode",
    dataAsOfPrefix: "Gegevens per",
    comparisonPeriodPrefix: "Vergelijkingsperiode",
    allFiguresStatement: "Alle cijfers worden gerapporteerd in {currency}, tenzij anders vermeld. De YoY-variantie wordt berekend als huidige periode versus vergelijkingsperiode.",
    analysisTagSuffix: "Analyse"
  },
  DE: {
    qbrReport: "QBR-Bericht",
    anyQuestions: "Fragen?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Aktueller Zeitraum",
    comparisonPeriodYoy: "Vergleichszeitraum (YoY)",
    basisYoy: "Basis: Jahr-über-Jahr (YoY)",
    publisherActivityBySegment: "Publisher-Aktivität nach Segment",
    keyObservations: "Wichtigste Erkenntnisse",
    reportingPeriodPrefix: "Berichtszeitraum",
    dataAsOfPrefix: "Datenstand",
    comparisonPeriodPrefix: "Vergleichszeitraum",
    allFiguresStatement: "Alle Werte werden in {currency} angegeben, sofern nicht anders vermerkt. Die YoY-Abweichung wird als aktueller Zeitraum gegenüber Vergleichszeitraum berechnet.",
    analysisTagSuffix: "Analyse"
  },
  IT: {
    qbrReport: "Report QBR",
    anyQuestions: "Domande?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Periodo corrente",
    comparisonPeriodYoy: "Periodo di confronto (YoY)",
    basisYoy: "Base: anno su anno (YoY)",
    publisherActivityBySegment: "Attività publisher per segmento",
    keyObservations: "Osservazioni chiave",
    reportingPeriodPrefix: "Periodo di reporting",
    dataAsOfPrefix: "Dati al",
    comparisonPeriodPrefix: "Periodo di confronto",
    allFiguresStatement: "Tutti i valori sono riportati in {currency}, salvo diversa indicazione. La variazione YoY è calcolata come periodo corrente vs periodo di confronto.",
    analysisTagSuffix: "Analisi"
  },
  NO: {
    qbrReport: "QBR-rapport",
    anyQuestions: "Spørsmål?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Gjeldende periode",
    comparisonPeriodYoy: "Sammenligningsperiode (YoY)",
    basisYoy: "Grunnlag: år-over-år (YoY)",
    publisherActivityBySegment: "Publisheraktivitet etter segment",
    keyObservations: "Nøkkelobservasjoner",
    reportingPeriodPrefix: "Rapporteringsperiode",
    dataAsOfPrefix: "Data per",
    comparisonPeriodPrefix: "Sammenligningsperiode",
    allFiguresStatement: "Alle tall er oppgitt i {currency}, med mindre annet er angitt. YoY-variansen er beregnet som gjeldende periode mot sammenligningsperioden.",
    analysisTagSuffix: "Analyse"
  },
  SV: {
    qbrReport: "QBR-rapport",
    anyQuestions: "Några frågor?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Aktuell period",
    comparisonPeriodYoy: "Jämförelseperiod (YoY)",
    basisYoy: "Grund: år över år (YoY)",
    publisherActivityBySegment: "Publisheraktivitet per segment",
    keyObservations: "Viktiga observationer",
    reportingPeriodPrefix: "Rapporteringsperiod",
    dataAsOfPrefix: "Data per",
    comparisonPeriodPrefix: "Jämförelseperiod",
    allFiguresStatement: "Alla siffror rapporteras i {currency} om inget annat anges. YoY-variansen beräknas som aktuell period jämfört med jämförelseperiod.",
    analysisTagSuffix: "Analys"
  },
  DA: {
    qbrReport: "QBR-rapport",
    anyQuestions: "Nogen spørgsmål?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Aktuel periode",
    comparisonPeriodYoy: "Sammenligningsperiode (YoY)",
    basisYoy: "Grundlag: år-til-år (YoY)",
    publisherActivityBySegment: "Publisheraktivitet efter segment",
    keyObservations: "Nøgleobservationer",
    reportingPeriodPrefix: "Rapporteringsperiode",
    dataAsOfPrefix: "Data pr.",
    comparisonPeriodPrefix: "Sammenligningsperiode",
    allFiguresStatement: "Alle tal rapporteres i {currency}, medmindre andet er angivet. YoY-variansen beregnes som aktuel periode versus sammenligningsperiode.",
    analysisTagSuffix: "Analyse"
  },
  FI: {
    qbrReport: "QBR-raportti",
    anyQuestions: "Kysymyksiä?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Nykyinen jakso",
    comparisonPeriodYoy: "Vertailujakso (YoY)",
    basisYoy: "Perusta: vuosi vuodelta (YoY)",
    publisherActivityBySegment: "Julkaisija-aktiivisuus segmenteittäin",
    keyObservations: "Keskeiset havainnot",
    reportingPeriodPrefix: "Raportointijakso",
    dataAsOfPrefix: "Tiedot päivältä",
    comparisonPeriodPrefix: "Vertailujakso",
    allFiguresStatement: "Kaikki luvut raportoidaan valuutassa {currency}, ellei toisin mainita. YoY-vaihtelu lasketaan nykyisen jakson ja vertailujakson välillä.",
    analysisTagSuffix: "Analyysi"
  },
  ES: {
    qbrReport: "Informe QBR",
    anyQuestions: "¿Preguntas?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Período actual",
    comparisonPeriodYoy: "Período de comparación (YoY)",
    basisYoy: "Base: interanual (YoY)",
    publisherActivityBySegment: "Actividad de publishers por segmento",
    keyObservations: "Observaciones clave",
    reportingPeriodPrefix: "Período del informe",
    dataAsOfPrefix: "Datos a fecha de",
    comparisonPeriodPrefix: "Período de comparación",
    allFiguresStatement: "Todas las cifras se presentan en {currency}, salvo que se indique lo contrario. La variación YoY se calcula como período actual frente a período de comparación.",
    analysisTagSuffix: "Análisis"
  },
  PL: {
    qbrReport: "Raport QBR",
    anyQuestions: "Pytania?",
    thankYouSubtitleTemplate: "TD Publisher Performance Review - {period}",
    currentPeriod: "Bieżący okres",
    comparisonPeriodYoy: "Okres porównawczy (r/r)",
    basisYoy: "Podstawa: rok do roku (r/r)",
    publisherActivityBySegment: "Aktywność wydawców według segmentu",
    keyObservations: "Kluczowe obserwacje",
    reportingPeriodPrefix: "Okres raportowania",
    dataAsOfPrefix: "Dane na dzień",
    comparisonPeriodPrefix: "Okres porównawczy",
    allFiguresStatement: "Wszystkie wartości raportowane są w walucie {currency}, o ile nie wskazano inaczej. Zmiana r/r jest liczona jako bieżący okres względem okresu porównawczego.",
    analysisTagSuffix: "Analiza",
    segmentSignalUnavailable: "Sygnał segmentu jest niedostępny.",
    detailedMovementUnavailable: "Szczegółowy opis zmian nie jest dostępny w tym wyciągu.",
    kpiSignalGeneric: "Sygnał KPI",
    kpiDriverUnavailable: "Brak potwierdzonego czynnika na podstawie dostępnych danych KPI.",
    kpiDetailUnavailable: "Szczegóły nie są dostępne w bieżącym wyciągu.",
    kpiTitleConversionRateImprovement: "Poprawa współczynnika konwersji",
    kpiTitleSalesVolumePressure: "Presja na wolumen sprzedaży",
    kpiTitleAovGrowthOffset: "Wzrost AOV częściowo kompensujący spadek wolumenu",
    kpiTitleRisingCpa: "Wzrost CPA",
    kpiTitleRoiTrend: "Trend ROI"
  }
};

const TRANSLATE_TIMEOUT_MS = Math.max(700, Number(process.env.QBR_TRANSLATE_TIMEOUT_MS || 1500));
const TRANSLATE_CONCURRENCY = Math.max(1, Number(process.env.QBR_TRANSLATE_CONCURRENCY || 12));
const TRANSLATE_MAX_TEXTS = Math.max(200, Number(process.env.QBR_TRANSLATE_MAX_TEXTS || 1200));
const AUTO_TRANSLATE_ENABLED = !/^(0|false|off)$/i.test(String(process.env.QBR_AUTO_TRANSLATE || "true"));

function cleanText(value, fallback = "") {
  const raw = String(value ?? fallback);
  const repaired = TEXT_REPLACEMENTS.reduce((text, [pattern, replacement]) => text.replace(pattern, replacement), raw);
  const repairedCurrency = repaired
    .replace(/Â£/g, "\u00A3")
    .replace(/â‚¬/g, "\u20AC")
    .replace(/Â¥/g, "\u00A5")
    .replace(/Â /g, " ")
    .replace(/Â/g, "")
    .replace(/â€“|â€”/g, "-")
    .replace(/â€˜|â€™/g, "'")
    .replace(/â€œ|â€�/g, '"')
    .replace(/â€¦/g, "...")
    .replace(/zÅ‚/g, "z\u0142");
  const xmlSafe = repairedCurrency
    .replace(MOJIBAKE_DIRECTION_MARK_REGEX, "")
    .replace(DIRECTION_MARK_REGEX, "")
    .replace(/\uFFFD/g, "");
  return xmlSafe.replace(/\s+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

function cleanInlineText(value, fallback = "") {
  return cleanText(value, fallback).replace(/\s+/g, " ").trim();
}

function titleCaseWords(value) {
  return cleanInlineText(value)
    .split(" ")
    .map((word) => word ? word[0].toUpperCase() + word.slice(1) : word)
    .join(" ");
}

function capitalizeFirstLetter(value) {
  const text = cleanInlineText(value || "");
  if (!text) return "";
  return `${text[0].toUpperCase()}${text.slice(1)}`;
}

function publisherAnalysisToProgramContext(value) {
  let text = cleanInlineText(value || "");
  if (!text) return "";

  const protectedPhrases = [];
  text = text.replace(/\b(publisher commission|publ commission|pub comm)\b/gi, (match) => {
    const token = `__QBR_METRIC_${protectedPhrases.length}__`;
    protectedPhrases.push([token, match]);
    return token;
  });

  text = text
    .replace(/\b[Pp]ublishers\b/g, (match) => (match[0] === "P" ? "Programs" : "programs"))
    .replace(/\b[Pp]ublisher\b/g, (match) => (match[0] === "P" ? "Program" : "program"));

  protectedPhrases.forEach(([token, phrase]) => {
    text = text.replaceAll(token, phrase);
  });

  return cleanInlineText(text);
}

function normalizeLanguageCode(value) {
  const code = cleanInlineText(value || "EN").toUpperCase();
  return LANGUAGE_TRANSLATION_TARGET_MAP[code] ? code : "EN";
}

function localeForLanguageCode(languageCode) {
  return LANGUAGE_LOCALE_MAP[normalizeLanguageCode(languageCode)] || "en-GB";
}

function uiLabelsForLanguage(languageCode) {
  const code = normalizeLanguageCode(languageCode);
  return {
    ...DEFAULT_UI_LABELS,
    ...(UI_LABELS_BY_LANGUAGE[code] || {})
  };
}

function uiLabel(deck, key, fallback) {
  const labels = deck?.metadata?.uiLabels || {};
  return cleanInlineText(labels[key] || fallback || "");
}

function shouldTranslateText(value) {
  const text = cleanInlineText(value || "");
  if (!text) return false;
  if (text.length < 2 || text.length > 2400) return false;
  if (/^https?:\/\//i.test(text)) return false;
  if (/^[\d\s.,:+\-\u2013\u2014/%()\u00A3\u20AC$z\u0142kr]+$/i.test(text)) return false;
  return /\p{L}/u.test(text);
}

function fetchJsonWithHttps(url, timeoutMs) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, {
      headers: {
        "User-Agent": "qbr-pptx-service/1.0"
      }
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => {
        const body = Buffer.concat(chunks).toString("utf8");
        if (response.statusCode < 200 || response.statusCode >= 300) {
          reject(new Error(`Unexpected status ${response.statusCode}`));
          return;
        }
        try {
          resolve(JSON.parse(body));
        } catch (error) {
          reject(error);
        }
      });
    });
    request.on("error", reject);
    request.setTimeout(timeoutMs, () => {
      request.destroy(new Error("Request timed out"));
    });
  });
}
async function translateWithGoogle(text, targetLang) {
  if (!AUTO_TRANSLATE_ENABLED) return text;
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${encodeURIComponent(targetLang)}&dt=t&q=${encodeURIComponent(text)}`;
  const canUseFetch = typeof fetch === "function" && typeof AbortController === "function";
  const controller = canUseFetch ? new AbortController() : null;
  const timer = canUseFetch
    ? setTimeout(() => controller.abort(), TRANSLATE_TIMEOUT_MS)
    : null;
  try {
    const data = canUseFetch
      ? await (async () => {
        const response = await fetch(url, { signal: controller.signal });
        if (!response.ok) return null;
        return response.json();
      })()
      : await fetchJsonWithHttps(url, TRANSLATE_TIMEOUT_MS);
    if (!Array.isArray(data) || !Array.isArray(data[0])) return text;
    const translated = data[0]
      .map((part) => (Array.isArray(part) && typeof part[0] === "string" ? part[0] : ""))
      .join("")
      .trim();
    return translated || text;
  } catch (_) {
    return text;
  } finally {
    if (timer) clearTimeout(timer);
  }
}
function collectDeckTextRefs(deckSpec) {
  const refs = [];
  const addRef = (container, key, priority = 2) => {
    if (!container) return;
    const value = container[key];
    if (typeof value !== "string") return;
    if (!shouldTranslateText(value)) return;
    refs.push({ container, key, text: value, priority });
  };

  for (const slide of deckSpec.slides || []) {
    ["title", "subtitle", "headline", "summary", "callout", "footerNote"].forEach((key) => addRef(slide, key, 4));

    if (Array.isArray(slide.bullets)) {
      slide.bullets.forEach((_, idx) => addRef(slide.bullets, idx, 4));
    }

    if (Array.isArray(slide.signals)) {
      slide.signals.forEach((signal) => {
        addRef(signal, "title", 4);
        addRef(signal, "detail", 4);
      });
    }

    if (Array.isArray(slide.kpis)) {
      slide.kpis.forEach((kpi) => {
        addRef(kpi, "label", 4);
        addRef(kpi, "summary", 4);
      });
    }

    if (Array.isArray(slide.tables)) {
      slide.tables.forEach((table) => {
        addRef(table, "title", 3);
        if (Array.isArray(table.columns)) {
          table.columns.forEach((_, idx) => addRef(table.columns, idx, 3));
        }
        if (Array.isArray(table.rows)) {
          table.rows.forEach((row) => {
            if (!Array.isArray(row)) return;
            row.forEach((cell, idx) => {
              if (typeof cell !== "string") return;
              const trimmed = cleanInlineText(cell);
              if (!shouldTranslateText(trimmed)) return;
              refs.push({ container: row, key: idx, text: cell, priority: 2 });
            });
          });
        }
      });
    }
  }

  return refs;
}

async function localizeDeckSpec(deckSpec, languageCode) {
  const code = normalizeLanguageCode(languageCode);
  const locale = localeForLanguageCode(code);
  deckSpec.metadata.locale = locale;
  deckSpec.metadata.uiLabels = uiLabelsForLanguage(code);

  const targetLang = LANGUAGE_TRANSLATION_TARGET_MAP[code] || "en";
  if (targetLang === "en") return deckSpec;

  const refs = collectDeckTextRefs(deckSpec);
  if (!refs.length) return deckSpec;

  const cache = new Map();
  const rankedTexts = new Map();
  refs.forEach((ref) => {
    const key = cleanText(ref.text || "");
    if (!key) return;
    const existing = rankedTexts.get(key);
    const currentPriority = Number(ref.priority || 1);
    if (!existing || currentPriority > existing.priority) {
      rankedTexts.set(key, { text: key, priority: currentPriority });
    }
  });
  const uniqueTexts = Array.from(rankedTexts.values())
    .sort((a, b) => (b.priority - a.priority) || (b.text.length - a.text.length))
    .slice(0, TRANSLATE_MAX_TEXTS)
    .map((entry) => entry.text);
  const translatableTexts = [];
  for (const text of uniqueTexts) {
    if (!shouldTranslateText(text)) {
      cache.set(text, text);
    } else {
      translatableTexts.push(text);
    }
  }

  if (translatableTexts.length) {
    let nextIndex = 0;
    const workerCount = Math.min(TRANSLATE_CONCURRENCY, translatableTexts.length);
    const workers = Array.from({ length: workerCount }, async () => {
      while (nextIndex < translatableTexts.length) {
        const idx = nextIndex;
        nextIndex += 1;
        const text = translatableTexts[idx];
        const translated = await translateWithGoogle(text, targetLang);
        cache.set(text, cleanText(translated || text));
      }
    });
    await Promise.all(workers);
  }

  refs.forEach((ref) => {
    const key = cleanText(ref.text || "");
    const translated = cache.get(key);
    if (translated) ref.container[ref.key] = translated;
  });

  return deckSpec;
}

function normalizeHex(value, fallback) {
  const compact = String(value ?? "").trim().replace(/^#/, "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(compact) ? `#${compact}` : fallback;
}

function normalizeTableKey(key) {
  const compact = cleanInlineText(key).toLowerCase().replace(/[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE]/g, "");
  return TABLE_KEY_MAP[compact] || compact || "table";
}

function titleFromKey(key) {
  return key
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/_/g, " ")
    .replace(/\b\w/g, (m) => m.toUpperCase())
    .trim();
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;

  let text = cleanInlineText(value);
  if (!text) return null;
  text = text.replace(/[\u00A3\u20AC$\u00A5]|z\u0142|kr|\b(?:GBP|EUR|USD|AUD|PLN|SEK|NOK|DKK|ISK)\b/gi, "").replace(/\s+/g, "");
  const isPercent = text.endsWith("%");
  text = text.replace(/%/g, "").replace(/[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE]/g, "");

  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");
  if (lastComma !== -1 && lastDot !== -1) {
    text = lastComma > lastDot ? text.replace(/\./g, "").replace(",", ".") : text.replace(/,/g, "");
  } else if (lastComma !== -1) {
    const decimals = text.length - lastComma - 1;
    text = decimals >= 1 && decimals <= 2 ? text.replace(",", ".") : text.replace(/,/g, "");
  } else if (lastDot !== -1) {
    const decimals = text.length - lastDot - 1;
    if (!(decimals >= 1 && decimals <= 4)) text = text.replace(/\./g, "");
  }

  const numeric = Number(text.replace(/[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE]/g, ""));
  if (!Number.isFinite(numeric)) return null;
  return isPercent ? numeric : numeric;
}

function detectUnit(label, sample) {
  const key = cleanInlineText(label).toLowerCase();
  const value = cleanInlineText(sample);
  if (value.includes("%") || key.includes("rate") || key.includes("variance")) return "percent";
  if (/[\u00A3\u20AC$\u00A5]|z\u0142|kr/i.test(value) || key.includes("value") || key.includes("commission") || key.includes("cpa")) return "currency";
  if (key.includes("roi")) return "ratio";
  if (/^\d+([,.]\d+)?$/.test(value.replace(/[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE]/g, ""))) return "number";
  return "text";
}

function parseSections(text) {
  const normalized = cleanText(text);
  if (!normalized) return [];

  const lines = normalized.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const sections = [];
  let currentTitle = "Summary";
  let bullets = [];
  let paragraphs = [];

  function flush() {
    if (bullets.length === 0 && paragraphs.length === 0 && sections.length > 0) return;
    sections.push({ title: currentTitle, bullets, paragraphs });
    bullets = [];
    paragraphs = [];
  }

  for (const line of lines) {
    const heading = line.match(/^#{2,6}\s+(.+)$/);
    if (heading) {
      if (bullets.length > 0 || paragraphs.length > 0 || sections.length === 0) flush();
      currentTitle = cleanInlineText(heading[1], "Summary");
      continue;
    }
    if (/^[-*]\s+/.test(line)) {
      bullets.push(cleanInlineText(line.replace(/^[-*]\s+/, "")));
      continue;
    }
    if (/^\d+\.\s+/.test(line)) {
      bullets.push(cleanInlineText(line.replace(/^\d+\.\s+/, "")));
      continue;
    }
    paragraphs.push(cleanInlineText(line));
  }

  if (bullets.length > 0 || paragraphs.length > 0 || sections.length === 0) flush();
  return sections.filter((section) => section.bullets.length || section.paragraphs.length);
}

function normalizeRows(rows) {
  if (!Array.isArray(rows)) return [];
  return rows
    .filter((row) => row && typeof row === "object" && !Array.isArray(row))
    .map((row) => {
      const normalized = {};
      for (const [key, value] of Object.entries(row)) {
        normalized[cleanInlineText(key, "Value")] = cleanInlineText(value, "-") || "-";
      }
      return normalized;
    });
}

function normalizeTables(input) {
  if (!input || typeof input !== "object" || Array.isArray(input)) return {};
  const tables = {};
  for (const [rawKey, rawRows] of Object.entries(input)) {
    const key = normalizeTableKey(rawKey);
    const rows = normalizeRows(rawRows);
    if (!rows.length) continue;
    tables[key] = {
      key,
      title: titleFromKey(key),
      columns: Object.keys(rows[0] || {}),
      rows
    };
  }
  return tables;
}

function normalizeProgramScopeTable(input) {
  const rowsInput = Array.isArray(input)
    ? input
    : input && typeof input === "object" && Array.isArray(input.rows)
      ? input.rows
      : [];

  const rows = normalizeRows(rowsInput);
  if (!rows.length) return null;

  const findValue = (row, aliases) => {
    for (const alias of aliases) {
      if (Object.prototype.hasOwnProperty.call(row, alias)) return row[alias];
      const matchedKey = Object.keys(row).find((key) => key.toLowerCase() === alias.toLowerCase());
      if (matchedKey) return row[matchedKey];
    }
    return "-";
  };

  const normalizedRows = rows
    .map((row) => PROGRAM_BREAKDOWN_COLUMNS.map((column) => findValue(row, column.aliases)))
    .filter((row) => row.some((value) => value && value !== "-"));

  if (!normalizedRows.length) return null;

  return {
    title: cleanInlineText((input && input.title) || "Program-Level Breakdown"),
    columns: PROGRAM_BREAKDOWN_COLUMNS.map((column) => column.label),
    rows: normalizedRows
  };
}

function formatMoneyAmount(value, currencyCode, locale = "en-GB", { signed = false, decimals = 2 } = {}) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "-";
  const symbol = getCurrencySymbol(currencyCode);
  const n = Number(value);
  const abs = Math.abs(n).toLocaleString(locale || "en-GB", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
  if (!signed) return `${symbol}${abs}`;
  if (n > 0) return `+${symbol}${abs}`;
  if (n < 0) return `-${symbol}${abs}`;
  return `${symbol}${abs}`;
}

function buildEarningsPerConversionMetric(metricMap, currencyCode, locale) {
  const conversions = metricMap.conversions || metricMap.sales;
  const publisherCommission = metricMap.publcommission || metricMap.publishercommission;
  if (!conversions || !publisherCommission) return null;

  const perConversion = (commission, conversionCount) => {
    const commissionValue = Number(commission);
    const conversionValue = Number(conversionCount);
    if (!Number.isFinite(commissionValue) || !Number.isFinite(conversionValue) || conversionValue === 0) return null;
    return commissionValue / conversionValue;
  };

  const currentValue = perConversion(publisherCommission.currentValue, conversions.currentValue);
  const previousValue = perConversion(publisherCommission.previousValue, conversions.previousValue);
  const differenceValue = currentValue !== null && previousValue !== null ? currentValue - previousValue : null;
  const varianceValue = differenceValue !== null && previousValue ? (differenceValue / Math.abs(previousValue)) * 100 : null;

  return {
    key: "earningsperconversion",
    label: "Earnings per Conversion",
    current: formatMoneyAmount(currentValue, currencyCode, locale),
    previous: formatMoneyAmount(previousValue, currencyCode, locale),
    difference: formatMoneyAmount(differenceValue, currencyCode, locale, { signed: true }),
    variance: varianceValue === null ? "-" : formatSignedPercent(varianceValue, locale, 1),
    currentValue,
    previousValue,
    differenceValue,
    varianceValue,
    unit: "currency"
  };
}

function normalizeMetrics(programYoYTable, currencyCode = "GBP", locale = "en-GB") {
  const rows = normalizeRows(programYoYTable);
  if (!rows.length) return { metrics: [], metricMap: {} };
  const recent = rows.find((row) => String(row.Row || "").toLowerCase().includes("recent")) || rows[0];
  const previous = rows.find((row) => String(row.Row || "").toLowerCase().includes("previous")) || rows[1];
  const difference = rows.find((row) => String(row.Row || "").toLowerCase().includes("difference")) || rows[2];
  const variance = rows.find((row) => String(row.Row || "").toLowerCase().includes("variance")) || rows[3];

  const columns = Object.keys(recent || {}).filter((column) => !["row", "metric"].includes(column.toLowerCase()));
  const metrics = columns.map((column) => {
    const current = recent ? recent[column] : undefined;
    const previousValue = previous ? previous[column] : undefined;
    const diff = difference ? difference[column] : undefined;
    const varianceValue = variance ? variance[column] : undefined;
    return {
      key: column.replace(/\s+/g, "").toLowerCase(),
      label: column,
      current,
      previous: previousValue,
      difference: diff,
      variance: varianceValue,
      currentValue: parseNumber(current),
      previousValue: parseNumber(previousValue),
      differenceValue: parseNumber(diff),
      varianceValue: parseNumber(varianceValue),
      unit: detectUnit(column, current || diff || varianceValue || previousValue || "")
    };
  });
  const metricMap = Object.fromEntries(metrics.map((metric) => [metric.key, metric]));
  const aliasMetric = (alias, sourceKey, label) => {
    if (!metricMap[alias] && metricMap[sourceKey]) {
      metricMap[alias] = {
        ...metricMap[sourceKey],
        key: alias,
        label
      };
    }
  };

  aliasMetric("conversions", "sales", "Conversions");
  aliasMetric("totalearnings", "totalearning", "Total Earnings");
  aliasMetric("totalearnings", "totalcommission", "Total Earnings");
  aliasMetric("digitalwallet", "digitalwallets", "Digital Wallet");
  aliasMetric("publcommission", "publishercommission", "Publisher Commission");
  aliasMetric("earningsperclick", "epc", "Earnings per Click");
  aliasMetric("earningspercommission", "earningsperpublcommission", "Earnings per Commission");
  aliasMetric("earningspercommission", "earningsperpublishercommission", "Earnings per Commission");

  const earningsPerConversion = buildEarningsPerConversionMetric(metricMap, currencyCode, locale);
  if (earningsPerConversion) {
    metricMap.earningsperconversion = earningsPerConversion;
  }

  return {
    metrics,
    metricMap
  };
}

function normalizeRecommendations(recommendations) {
  if (!Array.isArray(recommendations)) return [];
  return recommendations
    .map((item) => {
      if (typeof item === "string") return cleanInlineText(item);
      if (item && typeof item === "object") return cleanInlineText(item.text || item.title || item.body || "");
      return "";
    })
    .filter(Boolean);
}

function normalizeSignalItems(signals) {
  if (!Array.isArray(signals)) return [];
  return signals
    .map((item) => {
      if (typeof item === "string") {
        const text = cleanInlineText(item);
        if (!text) return null;
        const split = text.split(/\s*:\s*/);
        if (split.length >= 2) {
          return { title: cleanInlineText(split.shift()), detail: cleanInlineText(split.join(": ")) };
        }
        return { title: text, detail: "" };
      }
      if (!item || typeof item !== "object") return null;
      const title = cleanInlineText(item.title || item.heading || item.label || "");
      const detail = cleanInlineText(item.detail || item.body || item.text || "");
      if (!title && !detail) return null;
      return { title: title || detail, detail: detail && detail !== title ? detail : "" };
    })
    .filter(Boolean);
}

function normalizeIdList(input) {
  if (Array.isArray(input)) {
    return input.map((item) => cleanInlineText(item)).filter(Boolean);
  }
  if (typeof input === "string") {
    return input
      .split(",")
      .map((item) => cleanInlineText(item))
      .filter(Boolean);
  }
  return [];
}

function normalizeStringList(input) {
  if (Array.isArray(input)) {
    return input.map((item) => cleanInlineText(item)).filter(Boolean);
  }
  if (typeof input === "string") {
    return input
      .split(/\r?\n|[;]+/)
      .map((item) => cleanInlineText(item))
      .filter(Boolean);
  }
  return [];
}

function normalizeThemeFont(fontName, fallback) {
  const cleaned = cleanInlineText(fontName || fallback);
  return /^instrument\s+sans$/i.test(cleaned) ? "Aptos" : cleaned;
}

function resolveTheme(themeName, overrides) {
  const colors = (overrides && overrides.colors) || {};
  const fonts = (overrides && overrides.fonts) || {};

  return {
    ...DEFAULT_THEME,
    name: cleanInlineText((overrides && overrides.themeName) || themeName || DEFAULT_THEME.name),
    companyName: cleanInlineText((overrides && overrides.companyName) || DEFAULT_THEME.companyName),
    logoText: cleanInlineText((overrides && overrides.logoText) || (overrides && overrides.companyName) || DEFAULT_THEME.logoText),
    fonts: {
      heading: normalizeThemeFont(fonts.heading, DEFAULT_THEME.fonts.heading),
      body: normalizeThemeFont(fonts.body, DEFAULT_THEME.fonts.body),
      mono: normalizeThemeFont(fonts.mono, DEFAULT_THEME.fonts.mono)
    },
    colors: {
      ink: normalizeHex(colors.ink, DEFAULT_THEME.colors.ink),
      paper: normalizeHex(colors.paper, DEFAULT_THEME.colors.paper),
      canvas: normalizeHex(colors.canvas, DEFAULT_THEME.colors.canvas),
      accent: normalizeHex(colors.accent, DEFAULT_THEME.colors.accent),
      accentAlt: normalizeHex(colors.accentAlt, DEFAULT_THEME.colors.accentAlt),
      success: normalizeHex(colors.success, DEFAULT_THEME.colors.success),
      warning: normalizeHex(colors.warning, DEFAULT_THEME.colors.warning),
      highlight: normalizeHex(colors.highlight, DEFAULT_THEME.colors.highlight),
      muted: normalizeHex(colors.muted, DEFAULT_THEME.colors.muted),
      border: normalizeHex(colors.border, DEFAULT_THEME.colors.border)
    }
  };
}

function normalizePayload(payload) {
  const nestedPayload = payload && typeof payload.payload === "object" && payload.payload
    ? payload.payload
    : {};
  const client = cleanInlineText(payload.client || payload.clientName || "Client");
  const deckTitle = cleanInlineText(payload.deckTitle || `QBR - ${client}`);
  const reportingPeriod = cleanInlineText(payload.reportingPeriod || "Reporting period not provided");
  const comparisonPeriod = cleanInlineText(payload.comparisonPeriod || "Comparison period not provided");
  const qbrFocus = cleanInlineText(payload.qbrFocus || "General performance review");
  const qbrFocusDetail = cleanInlineText(payload.qbrFocusDetail || "");
  const languageCode = normalizeLanguageCode(payload.languageCode || "EN");
  const languageName = cleanInlineText(payload.languageName || "English");
  const locale = localeForLanguageCode(languageCode);
  const currencyCode = cleanInlineText(payload.currencyCode || "EUR").toUpperCase();
  const programOutput = cleanText(payload.programOutput || "");
  const publisherAnalysis = cleanText(payload.publisherAnalysis || "");
  const executiveSummaryText = cleanInlineText(
    payload.executiveSummaryText || payload.programExecutiveSummaryText || ""
  );
  const publisherOverviewObservations = normalizeStringList(
    payload.publisherOverviewObservations || payload.publisherKeyObservations || payload.keyObservations
  );
  const salesGrowthSignals = normalizeSignalItems(
    payload.salesGrowthSignals || payload.salesGrowthSignalBullets || payload.salesGrowthAnalysis
  );
  const rawProgramScopeTable = payload.programScopeTable || payload.programLevelBreakdown || payload.programBreakdownTable;
  const scopeRowsForIds = Array.isArray(rawProgramScopeTable)
    ? rawProgramScopeTable
    : rawProgramScopeTable && typeof rawProgramScopeTable === "object" && Array.isArray(rawProgramScopeTable.rows)
      ? rawProgramScopeTable.rows
      : [];
  const scopeDerivedProgramIds = Array.from(new Set(
    scopeRowsForIds
      .map((row) => {
        if (!row || typeof row !== "object" || Array.isArray(row)) return "";
        const key = Object.keys(row).find((candidate) => {
          const k = cleanInlineText(candidate).toLowerCase().replace(/\s+/g, "");
          return k === "programid" || k === "id";
        });
        return key ? cleanInlineText(row[key]) : "";
      })
      .filter(Boolean)
  ));

  const explicitAnalysisProgramIds = Array.from(new Set([
    ...normalizeIdList(payload.analysisProgramIds),
    ...normalizeIdList(payload.publisherProgramIds),
    ...normalizeIdList(payload.programIds),
    ...normalizeIdList(nestedPayload.analysisProgramIds),
    ...normalizeIdList(nestedPayload.publisherProgramIds),
    ...normalizeIdList(nestedPayload.programIds)
  ].filter(Boolean)));

  const fallbackProgramIds = Array.from(new Set([
    cleanInlineText(payload.programId || nestedPayload.programId || ""),
    cleanInlineText(payload.publisherProgramId || nestedPayload.publisherProgramId || "")
  ].filter(Boolean)));

  const analysisProgramIds = explicitAnalysisProgramIds.length
    ? Array.from(new Set([...explicitAnalysisProgramIds, ...scopeDerivedProgramIds]))
    : (scopeDerivedProgramIds.length ? scopeDerivedProgramIds : fallbackProgramIds);
  const tables = normalizeTables(payload.publisherTables || {});
  const { metrics, metricMap } = normalizeMetrics(payload.programYoYTable || [], currencyCode, locale);
  const programScopeTable = (
    Array.isArray(rawProgramScopeTable)
    || (rawProgramScopeTable && typeof rawProgramScopeTable === "object" && Array.isArray(rawProgramScopeTable.rows))
  )
    ? rawProgramScopeTable
    : normalizeProgramScopeTable(rawProgramScopeTable);

  return {
    requestId: cleanInlineText(payload.requestId || `qbr-${Date.now()}`),
    client,
    deckTitle,
    themeName: cleanInlineText(payload.themeName || "TD"),
    themeOverrides: payload.themeOverrides,
    reportingPeriod,
    comparisonPeriod,
    qbrFocus,
    qbrFocusDetail,
    languageCode,
    languageName,
    locale,
    currencyCode,
    programStatusCreatedFromDate: cleanInlineText(
      payload.programStatusCreatedFromDate ||
      payload.programStatusCreatedFromDateRaw ||
      nestedPayload.programStatusCreatedFromDate ||
      nestedPayload.programStatusCreatedFromDateRaw ||
      ""
    ),
    fullContent: payload.fullContent !== false,
    includeAppendix: payload.includeAppendix === true,
    debug: payload.debug === true,
    outputFileName: cleanInlineText(payload.outputFileName || ""),
    recommendations: normalizeRecommendations(payload.recommendations),
    programSections: parseSections(programOutput),
    publisherSections: parseSections(publisherAnalysis),
    executiveSummaryText,
    publisherOverviewObservations,
    salesGrowthSignals,
    analysisProgramIds,
    programOutput,
    publisherAnalysis,
    metrics,
    metricMap,
    tables,
    programScopeTable
  };
}

function trend(metric) {
  if (!metric || metric.varianceValue === null || metric.varianceValue === undefined) return "na";
  if (metric.varianceValue > 0.2) return "up";
  if (metric.varianceValue < -0.2) return "down";
  return "flat";
}

function cleanDeltaText(value) {
  const text = cleanInlineText(value || "");
  if (!text) return "";
  return text
    .replace(MOJIBAKE_DIRECTION_MARK_REGEX, "")
    .replace(DIRECTION_MARK_REGEX, "")
    .replace(/\uFFFD/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function metricCard(metric) {
  if (!metric || !metric.current) return null;
  const comparison = metric.previous
    ? `${metric.current} vs ${metric.previous} PY`
    : `${metric.current}`;
  const varianceText = cleanDeltaText(metric.variance);
  const summary = varianceText ? `${comparison} ${varianceText}` : comparison;
  return {
    label: metric.label,
    value: metric.current,
    previous: metric.previous || "",
    summary,
    delta: varianceText,
    trend: trend(metric)
  };
}

function getCurrencySymbol(code) {
  const c = cleanInlineText(code || "").toUpperCase();
  if (c === "GBP") return "\u00A3";
  if (c === "EUR") return "\u20AC";
  if (c === "USD") return "$";
  if (c === "AUD") return "A$";
  if (c === "PLN") return "z\u0142";
  if (["SEK", "NOK", "DKK", "ISK"].includes(c)) return "kr";
  return "";
}

function formatSignedMoney(value, currencyCode, locale = "en-GB") {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "N/A";
  const symbol = getCurrencySymbol(currencyCode);
  const n = Number(value);
  const abs = Math.abs(n);
  const rounded = abs >= 1000 ? Math.round(abs) : Number(abs.toFixed(0));
  const txt = Number(rounded).toLocaleString(locale || "en-GB");
  const sign = n >= 0 ? "+" : "-";
  return `${sign}${symbol}${txt}`;
}

function defaultBullets(input) {
  const bullets = input.programSections.flatMap((section) => section.bullets);
  if (bullets.length) return bullets.slice(0, 5);
  return input.publisherSections.flatMap((section) => section.bullets).slice(0, 5);
}

function buildHeadline(input) {
  const conversions = input.metricMap.conversions || input.metricMap.sales;
  const conversion = input.metricMap.convrate;
  const orderValue = input.metricMap.ordervalue;
  const totalEarnings = input.metricMap.totalearnings || input.metricMap.totalcommission;
  const publisherCommission = input.metricMap.publcommission || input.metricMap.publishercommission;

  if ((conversions && conversions.varianceValue > 0) && (totalEarnings && totalEarnings.varianceValue > 0)) return "Conversions and earnings both improved year on year";
  if ((conversion && conversion.varianceValue > 0) && (conversions && conversions.varianceValue < 0)) return "Conversion efficiency improved despite softer volume";
  if ((publisherCommission && publisherCommission.varianceValue > 0) && (totalEarnings && totalEarnings.varianceValue > 0)) return "Publisher earnings improved across commission and wallet income";
  if ((orderValue && orderValue.varianceValue < 0) && (conversion && conversion.varianceValue > 0)) return "Efficiency improved, but value generation remained under pressure";
  if (input.qbrFocus) return `${input.qbrFocus} remains the primary QBR focus`;
  return "Performance was mixed across volume and value measures";
}

function parsePeriodRange(reportingPeriod, locale = "en-GB") {
  const text = cleanInlineText(reportingPeriod || "");
  const match = text.match(/(\d{4}-\d{2}-\d{2})\s*(?:to|\u2013|-)\s*(\d{4}-\d{2}-\d{2})/i);
  if (!match) return text || "the current period";

  const start = new Date(`${match[1]}T00:00:00Z`);
  const end = new Date(`${match[2]}T00:00:00Z`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return text || "the current period";

  const options = { month: "short", year: "numeric", timeZone: "UTC" };
  const startLabel = start.toLocaleString(locale, options);
  const endLabel = end.toLocaleString(locale, options);
  return `${startLabel} \u2013 ${endLabel}`;
}

function parseIsoPeriod(periodText) {
  const text = cleanInlineText(periodText || "");
  const match = text.match(/(\d{4}-\d{2}-\d{2})\s*(?:to|\u2013|-)\s*(\d{4}-\d{2}-\d{2})/i);
  if (!match) return null;

  const start = new Date(`${match[1]}T00:00:00Z`);
  const end = new Date(`${match[2]}T00:00:00Z`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return null;
  return { start, end, startRaw: match[1], endRaw: match[2] };
}

function formatLongDate(date, locale = "en-GB") {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return date.toLocaleDateString(locale, {
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "UTC"
  });
}

function formatCompactDate(date, locale = "en-GB") {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return date.toLocaleDateString(locale, {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    timeZone: "UTC"
  }).replace(/\//g, "");
}

function formatPeriodForSlide(periodText, locale = "en-GB") {
  const parsed = parseIsoPeriod(periodText);
  if (!parsed) return cleanInlineText(periodText || "Not specified");
  return `${formatLongDate(parsed.start, locale)} \u2013 ${formatLongDate(parsed.end, locale)}`;
}

function buildCoverPeriodTag(periodText, locale = "en-GB") {
  const parsed = parseIsoPeriod(periodText);
  if (!parsed) return "PERIOD";
  return `${formatCompactDate(parsed.start, locale)}-${formatCompactDate(parsed.end, locale)}`;
}

function movementVerb(metric, positive = "increased", negative = "decreased") {
  if (!metric || metric.varianceValue === null || metric.varianceValue === undefined || Number.isNaN(Number(metric.varianceValue))) {
    return "changed";
  }
  if (Number(metric.varianceValue) > 0) return positive;
  if (Number(metric.varianceValue) < 0) return negative;
  return "was flat";
}

function buildExecutiveSummaryText(input) {
  const publisherLabel = cleanInlineText(input.client || "the publisher");
  const connectedProgramScope = `all programs connected to ${publisherLabel}`;

  const providedSummary = cleanDeltaText(input.executiveSummaryText || "");
  if (providedSummary) {
    const scopeCorrectedSummary = providedSummary
      .replace(/\bAcross\s+\d+\s+selected\s+programs\b/i, `Across ${connectedProgramScope}`)
      .replace(/\bAcross\s+selected\s+programs\b/i, `Across ${connectedProgramScope}`);
    const providedLooksFullPublisherScope = /all programs?|connected programs?|connected to|combined|portfolio/i.test(scopeCorrectedSummary);
    if (providedLooksFullPublisherScope) return scopeCorrectedSummary;
  }

  const m = input.metricMap || {};
  const conversions = m.conversions || m.sales || {};
  const conv = m.convrate || {};
  const ov = m.ordervalue || {};
  const publisherCommission = m.publcommission || m.publishercommission || {};
  const digitalWallet = m.digitalwallet || {};
  const totalEarnings = m.totalearnings || m.totalcommission || {};

  const periodLabel = parsePeriodRange(input.reportingPeriod, input.locale);
  const openingLine = `Across ${connectedProgramScope}, performance was mixed in ${periodLabel}.`;

  return cleanDeltaText(
    `${openingLine} Conversions moved ${cleanDeltaText(conversions.variance) || "N/A"} to ${conversions.current || "-"} and conversion rate moved ${cleanDeltaText(conv.variance) || "N/A"} to ${conv.current || "-"}. Publisher commission ${movementVerb(publisherCommission)} ${cleanDeltaText(publisherCommission.variance) || "N/A"} to ${publisherCommission.current || "-"}, while digital wallet earnings ${movementVerb(digitalWallet)} ${cleanDeltaText(digitalWallet.variance) || "N/A"} to ${digitalWallet.current || "-"}. Total earnings ended at ${totalEarnings.current || "-"} (${cleanDeltaText(totalEarnings.variance) || "N/A"}) and total order value ${movementVerb(ov)} ${cleanDeltaText(ov.variance) || "N/A"} to ${ov.current || "-"}.`
  );
}

function buildMetricRows(metricMap, keys) {
  return keys
    .map(([key, label]) => {
      const metric = metricMap[key];
      if (!metric) return null;
      return [
        label,
        metric.current || "-",
        metric.previous || "-",
        metric.difference || "-",
        metric.variance || "-"
      ];
    })
    .filter(Boolean);
}

function buildProgramBreakdownTable(input) {
  const targetColumns = [
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
  ];
  const selectedProgramIds = new Set(
    (Array.isArray(input.analysisProgramIds) ? input.analysisProgramIds : [])
      .map((value) => cleanInlineText(value))
      .filter(Boolean)
  );
  const canonicalizeProgramId = (value) => {
    const raw = cleanInlineText(value);
    if (!raw || raw === "-") return "";
    if (/^\d+$/.test(raw)) return String(Number(raw));
    return raw.toLowerCase();
  };
  const selectedProgramIdsCanonical = new Set(
    Array.from(selectedProgramIds).map((value) => canonicalizeProgramId(value)).filter(Boolean)
  );
  const selectedProgramIdSingle = selectedProgramIds.size === 1 ? Array.from(selectedProgramIds)[0] : "";

  function isSelectedProgramId(programIdValue) {
    if (!selectedProgramIds.size) return true;
    const candidate = cleanInlineText(programIdValue);
    if (!candidate || candidate === "-") return false;
    if (selectedProgramIds.has(candidate)) return true;
    const canonical = canonicalizeProgramId(candidate);
    return canonical ? selectedProgramIdsCanonical.has(canonical) : false;
  }

  function rowHasMetrics(row) {
    const cells = Array.isArray(row) ? row.slice(2) : [];
    return cells.some((cell) => {
      const value = cleanInlineText(cell);
      return value && value !== "-";
    });
  }

  function firstObjectValue(obj, aliases) {
    if (!obj || typeof obj !== "object") return "-";
    const directKeys = Object.keys(obj);
    const byLower = Object.fromEntries(directKeys.map((k) => [k.toLowerCase(), k]));
    for (const alias of aliases) {
      const key = byLower[String(alias).toLowerCase()];
      if (!key) continue;
      const value = obj[key];
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "-";
  }

  function firstRowCell(row, idx, aliases) {
    for (const alias of aliases) {
      const key = String(alias).toLowerCase();
      const colIndex = idx[key];
      if (colIndex === undefined) continue;
      const value = row[colIndex];
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "-";
  }

  function entry(programId, row) {
    return {
      programId: cleanInlineText(programId),
      row,
      publisherCommission: parseNumber(row[2]) || 0,
      programName: cleanInlineText(row[1])
    };
  }

  function sortEntries(entries) {
    return entries
      .slice()
      .sort((a, b) => {
        if (a.publisherCommission !== b.publisherCommission) return b.publisherCommission - a.publisherCommission;
        return a.programName.localeCompare(b.programName);
      });
  }

  function placeholderRows() {
    return Array.from(selectedProgramIds).map((id) => [id, "-", "-", "-", "-", "-", "-", "-", "-", "-"]);
  }

  const scope = input.programScopeTable;
  if (Array.isArray(scope) && scope.length && typeof scope[0] === "object" && !Array.isArray(scope[0])) {
    const mappedEntries = scope
      .map((row) => {
        let programId = firstObjectValue(row, ["Program ID", "ProgramId", "ProgramID", "ID"]);
        if ((!programId || cleanInlineText(programId) === "-") && selectedProgramIdSingle) {
          programId = selectedProgramIdSingle;
        }
        const programName = firstObjectValue(row, ["Program Name", "Program", "ProgramName", "Name"]) || programId;
        const publisherCommission = firstObjectValue(row, ["Publisher Commission", "Current Publisher Commission", "Commission", "Current Commission"]);
        const digitalWallet = firstObjectValue(row, ["Digital Wallet", "Digital Wallets", "DigitalWallet", "DigitalWallets"]);
        const totalEarnings = firstObjectValue(row, ["Total Earnings", "Total Earning", "TotalEarnings"]);
        const conversions = firstObjectValue(row, ["Conversions", "Current Conversions", "Sales", "Current Sales"]);
        const orderValue = firstObjectValue(row, ["Order Value", "Total Order Value", "Current OV", "Current Order Value"]);
        const publisherCommissionYoy = firstObjectValue(row, ["Publisher Commission YoY %", "Commission YoY %", "Publisher Commission % YoY", "Commission % YoY"]);
        const earningsYoy = firstObjectValue(row, ["Earnings YoY %", "Total Earnings YoY %", "Total Earnings % YoY", "YoY %"]);
        const conversionsYoy = firstObjectValue(row, ["Conversions YoY %", "Current Conversions YoY %", "Sales YoY %", "Current Sales YoY %"]);
        return entry(programId, [
          programId,
          programName,
          publisherCommission,
          digitalWallet,
          totalEarnings,
          conversions,
          orderValue,
          publisherCommissionYoy,
          earningsYoy,
          conversionsYoy
        ]);
      });
    const rows = mappedEntries.filter((item) => isSelectedProgramId(item.programId));
    const fallbackRows = mappedEntries.filter((item) => rowHasMetrics(item.row));

    if (!rows.length && selectedProgramIds.size && fallbackRows.length) {
      return {
        title: "Program-Level Breakdown",
        columns: targetColumns,
        rows: sortEntries(fallbackRows).map((item) => item.row),
        dense: false
      };
    }

    if (!rows.length && selectedProgramIds.size) {
      return {
        title: "Program-Level Breakdown",
        columns: targetColumns,
        rows: placeholderRows(),
        dense: false
      };
    }

    return {
      title: "Program-Level Breakdown",
      columns: targetColumns,
      rows: sortEntries(rows).map((item) => item.row),
      dense: false
    };
  }

  if (scope && Array.isArray(scope.rows) && scope.rows.length) {
    const idx = Object.fromEntries((scope.columns || []).map((col, i) => [cleanInlineText(col).toLowerCase(), i]));
    const mappedEntries = scope.rows
      .map((row) => {
        let programId = firstRowCell(row, idx, ["program id", "programid", "id"]);
        if ((!programId || cleanInlineText(programId) === "-") && selectedProgramIdSingle) {
          programId = selectedProgramIdSingle;
        }
        const programName = firstRowCell(row, idx, ["program name", "program", "programname", "name"]) || programId;
        const publisherCommission = firstRowCell(row, idx, ["publisher commission", "current publisher commission", "commission", "current commission"]);
        const digitalWallet = firstRowCell(row, idx, ["digital wallet", "digital wallets", "digitalwallet", "digitalwallets"]);
        const totalEarnings = firstRowCell(row, idx, ["total earnings", "total earning", "totalearnings"]);
        const conversions = firstRowCell(row, idx, ["conversions", "current conversions", "sales", "current sales"]);
        const orderValue = firstRowCell(row, idx, ["order value", "total order value", "current ov", "current order value"]);
        const publisherCommissionYoy = firstRowCell(row, idx, ["publisher commission yoy %", "commission yoy %", "publisher commission % yoy", "commission % yoy"]);
        const earningsYoy = firstRowCell(row, idx, ["earnings yoy %", "total earnings yoy %", "total earnings % yoy", "yoy %"]);
        const conversionsYoy = firstRowCell(row, idx, ["conversions yoy %", "current conversions yoy %", "sales yoy %", "current sales yoy %"]);
        return entry(programId, [
          programId,
          programName,
          publisherCommission,
          digitalWallet,
          totalEarnings,
          conversions,
          orderValue,
          publisherCommissionYoy,
          earningsYoy,
          conversionsYoy
        ]);
      });
    const rows = mappedEntries.filter((item) => isSelectedProgramId(item.programId));
    const fallbackRows = mappedEntries.filter((item) => rowHasMetrics(item.row));

    if (!rows.length && selectedProgramIds.size && fallbackRows.length) {
      return {
        title: "Program-Level Breakdown",
        columns: targetColumns,
        rows: sortEntries(fallbackRows).map((item) => item.row),
        dense: false
      };
    }

    if (!rows.length && selectedProgramIds.size) {
      return {
        title: "Program-Level Breakdown",
        columns: targetColumns,
        rows: placeholderRows(),
        dense: false
      };
    }

    return {
      title: "Program-Level Breakdown",
      columns: targetColumns,
      rows: sortEntries(rows).map((item) => item.row),
      dense: false
    };
  }

  if (Array.isArray(input.analysisProgramIds) && input.analysisProgramIds.length) {
    return {
      title: "Program-Level Breakdown",
      columns: targetColumns,
      rows: input.analysisProgramIds.map((id) => [id, "-", "-", "-", "-", "-", "-", "-", "-", "-"]),
      dense: false
    };
  }

  return {
    title: "Program-Level Breakdown",
    columns: targetColumns,
    rows: [["-", "-", "-", "-", "-", "-", "-", "-", "-", "-"]],
    dense: false
  };
}

function buildMoversCommissionBarChart(table, locale = "en-GB") {
  const rows = table && Array.isArray(table.rows) ? table.rows : [];
  const parsedRows = rows
    .map((row) => {
      const label = compactLabel(readTableCell(row, ["Chart Label", "Program Name", "Program", "Publisher"]), 48);
      const value = parseNumber(readTableCell(row, [
        "Publisher Commission Change Value",
        "Publisher Commission Change",
        "YoY Change"
      ])) ?? 0;
      const display = cleanInlineText(readTableCell(row, [
        "Publisher Commission Change",
        "YoY Change",
        "Change"
      ]) || value.toLocaleString(locale));
      const pct = cleanInlineText(readTableCell(row, [
        "Publisher Commission YoY %",
        "YoY %",
        "Variance"
      ]));

      return {
        label,
        value,
        display,
        pct,
        direction: value > 0 ? "up" : value < 0 ? "down" : "flat"
      };
    })
    .filter((row) => row.label && Number.isFinite(row.value) && row.value !== 0);

  const up = parsedRows
    .filter((row) => row.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, 10);
  const down = parsedRows
    .filter((row) => row.value < 0)
    .sort((a, b) => a.value - b.value)
    .slice(0, 10)
    .reverse();
  const chartRows = up.concat(down);

  return {
    type: "movers-commission-bar",
    title: "Partners with most change in publisher commission period-on-period",
    rows: chartRows.length ? chartRows : [{
      label: "No qualifying program movement",
      value: 0,
      display: "-",
      pct: "",
      direction: "flat"
    }]
  };
}

function tableRows(table, limit = 5) {
  if (!table || !table.rows || !table.rows.length) return null;
  const columns = Array.isArray(table.columns) ? [...table.columns] : [];
  const rows = table.rows.slice(0, limit).map((row) => columns.map((column) => row[column] || "-"));
  return { columns, rows };
}

function tableRowsWithRank(table, limit = 5) {
  const base = tableRows(table, limit);
  if (!base) return null;
  const ranked = base.columns[0] && base.columns[0].toLowerCase() === "rank";
  if (ranked) return base;

  return {
    columns: ["Rank", ...base.columns],
    rows: base.rows.map((row, index) => [String(index + 1), ...row])
  };
}

function tableOrPlaceholder(table, title, columns, placeholderRows = 5) {
  const ranked = tableRowsWithRank(table, placeholderRows);
  if (ranked && ranked.rows.length) {
    return {
      title,
      columns: ranked.columns,
      rows: ranked.rows
    };
  }

  const row = columns.map((column) => (column.toLowerCase() === "rank" ? "1" : "-"));
  return {
    title,
    columns,
    rows: [row]
  };
}

function tableOrPlaceholderNoRank(table, title, columns, placeholderRows = 5) {
  const base = tableRows(table, placeholderRows);
  if (base && base.rows.length) {
    return {
      title,
      columns: base.columns,
      rows: base.rows
    };
  }

  return {
    title,
    columns,
    rows: [columns.map(() => "-")]
  };
}

function buildCompetitorAnalysisTable(table) {
  const fallbackColumns = [
    "Competitor Group Summary",
    "Your Site",
    "Publisher 1",
    "Publisher 2",
    "Publisher 3",
    "Publisher 4",
    "Distinct comp. prog. #"
  ];
  if (!table || !Array.isArray(table.rows) || !table.rows.length) {
    return {
      title: "Competitor Group Summary",
      columns: fallbackColumns,
      rows: [fallbackColumns.map(() => "-")]
    };
  }

  const sourceColumns = Array.isArray(table.columns) ? table.columns : Object.keys(table.rows[0] || {});
  const metricColumn = sourceColumns.find((column) => /^competitor group summary$/i.test(column))
    || sourceColumns.find((column) => /^metric$/i.test(column))
    || sourceColumns[0]
    || "Metric";
  const valueColumns = sourceColumns.filter((column) => column !== metricColumn && column !== "_columns");
  const columns = ["Competitor Group Summary", ...valueColumns];
  const rows = table.rows.slice(0, 3).map((row) => [
    row[metricColumn] || "-",
    ...valueColumns.map((column) => row[column] || "-")
  ]);

  return {
    title: "Competitor Group Summary",
    columns,
    rows
  };
}

function competitorDisplayLabel(label, index) {
  const cleaned = cleanInlineText(label || "");
  if (index === 0 || /^(your site|primary publisher|primary site)$/i.test(cleaned)) return "Your Site";
  const letter = String.fromCharCode(64 + index);
  if (/^publisher\s*\d+$/i.test(cleaned) || /^competitor\s*\d+$/i.test(cleaned)) return `Comp. ${letter}`;
  return compactLabel(cleaned || `Comp. ${letter}`, 14);
}

function findCompetitorShareRow(rows, mode = "current") {
  const candidates = rows.filter((row) => {
    const metric = cleanInlineText(row.metric || "").toLowerCase();
    if (!/(pub|publisher).*(comm|commission)|commission/.test(metric)) return false;
    if (/distinct|program\s*#|count|with\s+pub|programs\s+with/i.test(metric)) return false;
    return true;
  });

  if (mode === "previous") {
    return candidates.find((row) => /(previous|prior|comparison|baseline|\bpy\b|\bpp\b)/i.test(row.metric)) || null;
  }

  return candidates.find((row) => !/(previous|prior|comparison|baseline|\bpy\b|\bpp\b)/i.test(row.metric))
    || candidates[0]
    || null;
}

function buildCompetitorSharePubCommChart(shareTable, analysisTable) {
  const source = shareTable && Array.isArray(shareTable.rows) && shareTable.rows.length
    ? shareTable
    : analysisTable;
  const fallbackColumns = ["Your Site", "Comp. A", "Comp. B", "Comp. C", "Comp. D"];
  if (!source || !Array.isArray(source.rows) || !source.rows.length) {
    return {
      type: "competitor-share-pub-comm",
      title: "Share of publisher commission within competitor group",
      categories: fallbackColumns,
      series: []
    };
  }

  const sourceColumns = Array.isArray(source.columns) ? source.columns : Object.keys(source.rows[0] || {});
  const metricColumn = sourceColumns.find((column) => /^competitor group summary$/i.test(cleanInlineText(column)))
    || sourceColumns.find((column) => /^metric$/i.test(cleanInlineText(column)))
    || sourceColumns[0]
    || "Metric";
  const valueColumns = sourceColumns
    .filter((column) => column !== metricColumn && column !== "_columns")
    .filter((column) => !/distinct|program\s*#|count/i.test(cleanInlineText(column)))
    .slice(0, 5);

  const normalizedRows = source.rows.map((row) => ({
    metric: row[metricColumn] || "",
    values: valueColumns.map((column) => parseNumber(row[column]) || 0)
  }));
  const rowsFromColumns = valueColumns.length && source.rows.some((row) =>
    valueColumns.some((column) => parseNumber(row[column]) !== null)
  );
  if (!rowsFromColumns) {
    return {
      type: "competitor-share-pub-comm",
      title: "Share of publisher commission within competitor group",
      categories: fallbackColumns,
      series: []
    };
  }

  const toShareValues = (values) => {
    const total = values.reduce((sum, value) => sum + Math.max(0, Number(value) || 0), 0);
    if (total <= 0) return values.map(() => 0);
    return values.map((value) => Number((((Math.max(0, Number(value) || 0) / total) * 100).toFixed(1))));
  };
  const current = findCompetitorShareRow(normalizedRows, "current");
  const previous = findCompetitorShareRow(normalizedRows, "previous");
  const series = [];
  if (previous) {
    series.push({
      label: "% of Site Group PP",
      values: toShareValues(previous.values)
    });
  }
  if (current) {
    series.push({
      label: "% of Site Group",
      values: toShareValues(current.values)
    });
  }

  return {
    type: "competitor-share-pub-comm",
    title: "Share of publisher commission within competitor group",
    categories: valueColumns.map((column, index) => competitorDisplayLabel(column, index)),
    series
  };
}

function buildTopProgramsCompetitorPerformanceTable(table) {
  const fallbackColumns = ["Program Name", "Primary Publisher", "Comp. A", "Comp. B", "Comp. C", "Comp. D"];
  if (!table || !Array.isArray(table.rows) || !table.rows.length) {
    return {
      title: "Top Programs Competitor Performance",
      columns: fallbackColumns,
      rows: [fallbackColumns.map(() => "-")],
      colAlign: ["left", "center", "center", "center", "center", "center"],
      colW: [3.9, 1.65, 1.35, 1.35, 1.35, 1.35],
      primaryHighlightColumn: 1,
      dense: true
    };
  }

  const sourceColumns = Array.isArray(table.columns) ? table.columns : Object.keys(table.rows[0] || {});
  const programColumn = sourceColumns.find((column) => /^program\s*name$/i.test(cleanInlineText(column)))
    || sourceColumns.find((column) => /program/i.test(cleanInlineText(column)))
    || sourceColumns[0]
    || "Program Name";
  const valueColumns = sourceColumns.filter((column) => column !== programColumn && column !== "_columns").slice(0, 5);
  const columns = ["Program Name", ...valueColumns];
  const rows = table.rows.slice(0, 10).map((row) => [
    cleanInlineText(row[programColumn] || "-"),
    ...valueColumns.map((column) => cleanInlineText(row[column] || "0%"))
  ]);

  return {
    title: "Top Programs Competitor Performance",
    columns,
    rows,
    colAlign: ["left", ...valueColumns.map(() => "center")],
    colW: [3.9, 1.65, 1.35, 1.35, 1.35, 1.35].slice(0, columns.length),
    primaryHighlightColumn: 1,
    dense: true
  };
}

function buildWeeklyPubCommComboChart(table) {
  if (!table || !Array.isArray(table.rows) || !table.rows.length) return null;

  const rawColumns = Array.isArray(table.columns) ? table.columns : Object.keys(table.rows[0] || {});
  const categoryColumn = rawColumns.find((column) => /^(week|date|period)$/i.test(cleanInlineText(column)))
    || rawColumns[0]
    || "Week";
  const seriesColumns = rawColumns.filter((column) => column !== categoryColumn && column !== "_columns");
  if (!seriesColumns.length) return null;

  const categories = table.rows.map((row) => cleanInlineText(row[categoryColumn] || "-"));
  const series = seriesColumns.map((column, index) => ({
    label: cleanInlineText(column || `Series ${index + 1}`),
    renderAs: "line",
    values: table.rows.map((row) => {
      const value = parseNumber(row[column]);
      return value === null ? null : value;
    })
  }));

  return {
    type: "weekly-pub-comm-line",
    title: "Publ comm by week",
    categories,
    series
  };
}

function buildPublisherOverviewBullets(input) {
  const tidyObservationLine = (line) => {
    let text = publisherAnalysisToProgramContext(line || "");
    text = text.replace(/^key observations?\s*[:\-]\s*/i, "");
    text = text
      .replace(/\s*-\s*I\s+recorded\b/gi, " recorded")
      .replace(/\s*-\s*I\s+drove\b/gi, " drove")
      .replace(/\s*-\s*I\s+account(?:s)?\s+for\b/gi, " accounts for")
      .replace(/\s+/g, " ")
      .trim();
    if (text && !/[.!?]$/.test(text)) text = `${text}.`;
    return text;
  };

  const cleanPublisherName = (value) =>
    cleanInlineText(value || "")
      .replace(/\s*-\s*I$/i, "")
      .replace(/\s+/g, " ")
      .trim();

  const isNarrativeCandidate = (line) => {
    const text = tidyObservationLine(line);
    if (!text || text.length < 30 || text.length > 260) return false;
    if (/\bsite\s*id\b/i.test(text)) return false;
    if (/\bcurrent sales:\b|\bcurrent ov:\b|\bov yoy change:\b|\bsales yoy %:\b/i.test(text)) return false;
    if (/^(voucher|cashback|other|content|css)\s*[-—]/i.test(text)) return false;
    if (/\btotal sales:\b|\btotal ov:\b|\bpublishers:\b/i.test(text)) return false;
    if (/^\d{4}-\d{2}-\d{2}\s+to\s+\d{4}-\d{2}-\d{2}/i.test(text)) return false;
    if (/\|\s/.test(text)) return false;
    return true;
  };

  const pickNarrativeBullets = (lines, limit = 4) => {
    const seen = new Set();
    const out = [];
    for (const line of lines || []) {
      const text = tidyObservationLine(line);
      if (!isNarrativeCandidate(text)) continue;
      const key = text.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(text);
      if (out.length >= limit) break;
    }
    return out;
  };

  if (input.publisherOverviewObservations && input.publisherOverviewObservations.length) {
    const curated = pickNarrativeBullets(input.publisherOverviewObservations, 4);
    if (curated.length) return curated;
  }

  const keyObservationSection = (input.publisherSections || []).find((section) =>
    /key observations?|publisher performance overview/i.test(cleanInlineText(section.title).toLowerCase())
  );
  const sectionBullets = pickNarrativeBullets(
    [
      ...(keyObservationSection?.bullets || []),
      ...(keyObservationSection?.paragraphs || [])
    ],
    4
  );
  if (sectionBullets.length) return sectionBullets;

  const growth = input.tables.topGrowthPublishers;
  const decline = input.tables.topDecliningPublishers;
  const current = input.tables.topCurrentPerformers;
  const brandNew = input.tables.brandNewPublishers;
  const segment = input.tables.segmentSnapshot;

  const obs = [];
  const topGrowthRow = growth?.rows?.[0];
  if (topGrowthRow) {
    const pub = cleanPublisherName(topGrowthRow["Program Name"] || topGrowthRow.Program || topGrowthRow.Publisher || "Top growth program");
    const seg = cleanInlineText(topGrowthRow.Segment || "N/A");
    const ovDelta = cleanInlineText(topGrowthRow["OV YoY Change"] || "N/A");
    const ovPct = cleanInlineText(topGrowthRow["OV YoY %"] || "N/A");
    obs.push(`${pub} was the strongest YoY uplift program (${seg}), adding ${ovDelta} in OV (${ovPct}).`);
  }

  const topDeclineRow = decline?.rows?.[0];
  if (topDeclineRow) {
    const pub = cleanPublisherName(topDeclineRow["Program Name"] || topDeclineRow.Program || topDeclineRow.Publisher || "Top declining program");
    const seg = cleanInlineText(topDeclineRow.Segment || "N/A");
    const ovDelta = cleanInlineText(topDeclineRow["OV YoY Change"] || "N/A");
    const ovPct = cleanInlineText(topDeclineRow["OV YoY %"] || "N/A");
    obs.push(`${pub} was the largest declining program (${seg}), with OV movement ${ovDelta} (${ovPct}).`);
  }

  if (brandNew?.rows?.length) {
    const totalOv = brandNew.rows.reduce((sum, row) => sum + (parseNumber(row["Current OV"]) || 0), 0);
    obs.push(`${brandNew.rows.length} brand-new programs were activated, contributing ${getCurrencySymbol(input.currencyCode)}${Math.round(totalOv).toLocaleString(input.locale || "en-GB")} in combined OV.`);
  }

  if (current?.rows?.length) {
    const top2 = current.rows.slice(0, 2).map((row) => ({
      name: cleanPublisherName(row["Program Name"] || row.Program || row.Publisher || "Program"),
      ov: parseNumber(row["Order Value"] || row["Current OV"] || row["Current Order Value"])
    }));
    if (top2.length === 2) {
      const top2Ov = top2.reduce((sum, item) => sum + (item.ov || 0), 0);
      const totalOv = (segment?.rows || []).reduce((sum, row) => sum + (parseNumber(row["Total OV"]) || 0), 0);
      const share = totalOv > 0 ? ` (${((top2Ov / totalOv) * 100).toFixed(1)}% of programme OV)` : "";
      obs.push(`Program concentration remains high: ${top2[0].name} and ${top2[1].name} account for ${getCurrencySymbol(input.currencyCode)}${Math.round(top2Ov).toLocaleString(input.locale || "en-GB")}${share}.`);
    }
  }

  const computed = pickNarrativeBullets(obs, 4).map(tidyObservationLine);
  if (computed.length) return computed;

  return [
    "Driver not confirmed from available segment and program data.",
    "No stable top program concentration signal available in current extract.",
    "New program contribution could not be quantified from available data.",
    "Review source program tables to confirm movement drivers."
  ];
}

function buildSegmentPerformanceBlocks(input) {
  const segment = input.tables.segmentSnapshot;
  if (!segment || !Array.isArray(segment.rows) || !segment.rows.length) {
    return [
      "Segment-level trend not available from current data extract.",
      "Use segment table to confirm YoY movement drivers before actioning.",
      "Cross-check top programs per segment for concentration effects."
    ];
  }

  const cleanPublisherLabel = (value) =>
    cleanInlineText(value || "")
      .replace(/\s*-\s*I$/i, "")
      .replace(/\s*-\s*I(?=\s|$)/gi, "")
      .replace(/\s+/g, " ")
      .trim();

  const iconBySegment = {
    voucher: "[V]",
    cashback: "[C]",
    other: "[O]",
    content: "[T]",
    css: "[CSS]"
  };

  const aiNarrativeCandidates = (input.publisherSections || [])
    .filter((section) =>
      /category snapshot|segment snapshot|publisher segment performance|confirmed changes|implications/i
        .test(cleanInlineText(section.title).toLowerCase())
    )
    .flatMap((section) => [...(section.bullets || []), ...(section.paragraphs || [])])
    .map((line) => cleanInlineText(line))
    .filter((line) => line.length >= 35 && line.length <= 420)
    .filter((line) => !/\bsite\s*id\b/i.test(line))
    .filter((line) => !/^\d{4}-\d{2}-\d{2}\s+to\s+\d{4}-\d{2}-\d{2}/i.test(line));

  const growthRows = (input.tables.topGrowthPublishers?.rows || []).map((row) => ({
    segment: cleanInlineText(row.Segment || ""),
    publisher: cleanPublisherLabel(row.Publisher || ""),
    salesCurrent: cleanInlineText(row["Current Sales"] || ""),
    salesPct: cleanInlineText(row["Sales YoY %"] || ""),
    ovDelta: cleanInlineText(row["OV YoY Change"] || ""),
    ovPct: cleanInlineText(row["OV YoY %"] || "")
  }));

  const declineRows = (input.tables.topDecliningPublishers?.rows || []).map((row) => ({
    segment: cleanInlineText(row.Segment || ""),
    publisher: cleanPublisherLabel(row.Publisher || ""),
    salesCurrent: cleanInlineText(row["Current Sales"] || ""),
    salesPct: cleanInlineText(row["Sales YoY %"] || ""),
    ovDelta: cleanInlineText(row["OV YoY Change"] || ""),
    ovPct: cleanInlineText(row["OV YoY %"] || "")
  }));
  const currentRows = (input.tables.topCurrentPerformers?.rows || []).map((row) => ({
    segment: cleanInlineText(row.Segment || ""),
    publisher: cleanPublisherLabel(row.Publisher || ""),
    ov: cleanInlineText(row["Order Value"] || row["Current OV"] || ""),
    sales: cleanInlineText(row["Current Sales"] || "")
  }));

  const preferredOrder = ["Voucher", "Cashback", "Other", "Content", "CSS"];
  const rows = segment.rows
    .map((row) => ({
      segment: cleanInlineText(row.Segment || "Segment"),
      publishers: cleanInlineText(row.Publishers || "N/A"),
      totalOv: cleanInlineText(row["Total OV"] || "-"),
      totalSales: cleanInlineText(row["Total Sales"] || "-"),
      ovYoy: cleanInlineText(row["OV YoY %"] || "N/A"),
      salesYoy: cleanInlineText(row["Sales YoY %"] || "N/A")
    }))
    .sort((a, b) => {
      const ai = preferredOrder.indexOf(a.segment);
      const bi = preferredOrder.indexOf(b.segment);
      if (ai === -1 && bi === -1) return 0;
      if (ai === -1) return 1;
      if (bi === -1) return -1;
      return ai - bi;
    });

  const clampDetail = (value, maxChars = 420) => {
    const text = cleanInlineText(value || "");
    if (!text || text.length <= maxChars) return text;
    return `${text.slice(0, maxChars - 1).trimEnd()}…`;
  };

  return rows.slice(0, 5).map((row) => {
    const icon = iconBySegment[row.segment.toLowerCase()] || "\u25AA";
    const growthForSegment = growthRows.find((item) => item.segment.toLowerCase() === row.segment.toLowerCase());
    const declineForSegment = declineRows.find((item) => item.segment.toLowerCase() === row.segment.toLowerCase());
    const topCurrentInSegment = currentRows
      .filter((item) => item.segment.toLowerCase() === row.segment.toLowerCase())
      .slice(0, 2);
    const movementParts = [];
    if (growthForSegment) {
      movementParts.push(`${growthForSegment.publisher} is the dominant growth driver, delivering ${growthForSegment.ovDelta} OV YoY (${growthForSegment.ovPct}).`);
    }
    if (declineForSegment) {
      movementParts.push(`${declineForSegment.publisher} is the primary drag, with ${declineForSegment.ovDelta} OV YoY (${declineForSegment.ovPct}).`);
    }
    if (topCurrentInSegment.length) {
      const contributorLine = topCurrentInSegment
        .map((item) => `${item.publisher}${item.ov ? ` (${item.ov})` : ""}`)
        .join(" and ");
      movementParts.push(`Leading current contributors include ${contributorLine}.`);
    }
    const movementLine = movementParts.join(" ");
    const defaultDetail = `${row.totalOv} total OV | ${row.publishers} active publishers | Sales: ${row.totalSales} (${row.salesYoy} YoY).${movementLine ? ` ${movementLine}` : ""}`;
    const aiDetail = aiNarrativeCandidates.find((line) =>
      new RegExp(`\\b${row.segment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(line)
    );
    let detail = aiDetail && aiDetail.length > defaultDetail.length * 0.6
      ? `${defaultDetail} ${aiDetail}`
      : defaultDetail;
    detail = detail
      .replace(/\s*-\s*I\s+is\s+the\s+primary\s+drag/gi, " is the primary drag")
      .replace(/\s*-\s*I\s+/gi, " ")
      .replace(/\s+/g, " ")
      .trim();
    detail = clampDetail(detail, 420);
    return `${icon} ${row.segment} - ${row.ovYoy} OV YoY\n${detail}`;
  });
}

function formatSignedCount(value, locale = "en-GB") {
  if (value === null || value === undefined) return "N/A";
  const n = Number(value);
  if (!Number.isFinite(n)) return "N/A";
  const rounded = Math.round(n);
  const abs = Math.abs(rounded).toLocaleString(locale || "en-GB");
  return `${rounded >= 0 ? "+" : "-"}${abs}`;
}

function buildSalesGrowthSignals(input) {
  if (Array.isArray(input.salesGrowthSignals) && input.salesGrowthSignals.length) {
    return input.salesGrowthSignals.slice(0, 5);
  }

  const m = input.metricMap || {};
  const segmentRows = input.tables.segmentSnapshot?.rows || [];
  const growthRows = input.tables.topGrowthPublishers?.rows || [];
  const declineRows = input.tables.topDecliningPublishers?.rows || [];
  const aovRows = input.tables.moversAov?.rows || [];
  const clickRows = input.tables.moversClicks?.rows || [];
  const currentRows = input.tables.topCurrentPerformers?.rows || [];

  const voucherRow = segmentRows.find((row) => cleanInlineText(row.Segment).toLowerCase() === "voucher");
  const cashbackRow = segmentRows.find((row) => cleanInlineText(row.Segment).toLowerCase() === "cashback");

  const topVoucherGrowth = growthRows
    .filter((row) => cleanInlineText(row.Segment).toLowerCase() === "voucher")
    .sort((a, b) => (parseNumber(b["OV YoY Change"]) || 0) - (parseNumber(a["OV YoY Change"]) || 0))[0];

  const topAovUps = aovRows
    .filter((row) => cleanInlineText(row.Direction).toLowerCase() === "up" || (parseNumber(row["YoY Change"]) || 0) > 0)
    .sort((a, b) => (parseNumber(b["YoY Change"]) || 0) - (parseNumber(a["YoY Change"]) || 0))
    .slice(0, 3);

  const topCashbackPublishers = currentRows
    .filter((row) => cleanInlineText(row.Segment).toLowerCase() === "cashback")
    .sort((a, b) => (parseNumber(b["Current Sales"]) || 0) - (parseNumber(a["Current Sales"]) || 0))
    .slice(0, 2);

  const topClickDecliners = clickRows
    .filter((row) => cleanInlineText(row.Direction).toLowerCase() === "down" || (parseNumber(row["YoY Change"]) || 0) < 0)
    .sort((a, b) => (parseNumber(a["YoY Change"]) || 0) - (parseNumber(b["YoY Change"]) || 0))
    .slice(0, 2);

  const clickLossAbs = Math.abs(parseNumber(m.clicks?.difference) || 0);
  const top2LossAbs = topClickDecliners.reduce((sum, row) => sum + Math.abs(parseNumber(row["YoY Change"]) || 0), 0);
  const top2Share = clickLossAbs > 0 ? `${((top2LossAbs / clickLossAbs) * 100).toFixed(0)}%` : "N/A";

  const signals = [
    {
      title: "Voucher Segment: Highest YoY Sales Growth",
      detail: voucherRow && topVoucherGrowth
        ? `The Voucher segment recorded ${cleanInlineText(voucherRow["Sales YoY %"] || "N/A")} sales growth and ${cleanInlineText(voucherRow["OV YoY %"] || "N/A")} OV growth YoY. ${cleanInlineText(topVoucherGrowth.Publisher || "Top voucher publisher")} delivered ${cleanInlineText(topVoucherGrowth["Current Sales"] || "N/A")} sales (${cleanInlineText(topVoucherGrowth["Sales YoY %"] || "N/A")}) and ${cleanInlineText(topVoucherGrowth["OV YoY Change"] || "N/A")} in OV (${cleanInlineText(topVoucherGrowth["OV YoY %"] || "N/A")}) year-over-year.`
        : "Voucher growth signal is not fully available in the current extract."
    },
    {
      title: `Conversion Rate: ${directionWord(m.convrate?.varianceValue) === "increased" ? "Improved" : "Moved"} to ${cleanInlineText(m.convrate?.current || "N/A")}`,
      detail: `Programme conversion rate moved from ${cleanInlineText(m.convrate?.previous || "N/A")} to ${cleanInlineText(m.convrate?.current || "N/A")} (${cleanInlineText(m.convrate?.variance || "N/A")}). Sales changed ${formatSignedCount(parseNumber(m.sales?.difference), input.locale)} while clicks changed ${formatSignedCount(parseNumber(m.clicks?.difference), input.locale)}, indicating the quality shift in converting traffic.`
    },
    {
      title: "AOV Growth Across Multiple Publishers",
      detail: topAovUps.length
        ? `Programme AOV moved ${cleanInlineText(m.aov?.variance || "N/A")} to ${cleanInlineText(m.aov?.current || "N/A")}. Largest AOV uplifts came from ${topAovUps.map((row) => `${cleanInlineText(row.Publisher || "Publisher")} (${cleanInlineText(row["YoY Change"] || "N/A")}, ${cleanInlineText(row["YoY %"] || "N/A")})`).join(", ")}.`
        : `Programme AOV moved ${cleanInlineText(m.aov?.variance || "N/A")} to ${cleanInlineText(m.aov?.current || "N/A")}.`
    },
    {
      title: "Cashback Segment: Largest Volume Base with Sales Decline",
      detail: cashbackRow
        ? `Cashback accounts for ${cleanInlineText(cashbackRow["Total OV"] || "N/A")} in OV (${cleanInlineText(cashbackRow["OV YoY %"] || "N/A")}) across ${cleanInlineText(cashbackRow.Publishers || "N/A")} publishers. ${topCashbackPublishers.length ? topCashbackPublishers.map((row) => `${cleanInlineText(row.Publisher || "Publisher")} (${cleanInlineText(row["Current Sales"] || "N/A")} sales, ${cleanInlineText(row["Sales YoY %"] || "N/A")})`).join(" and ") : "Top cashback contributors remain concentrated in a small group"} are the primary contributors by sales count.`
        : "Cashback segment-level signal is not fully available in the current extract."
    },
    {
      title: "Click Volume Decline Concentrated in Two Publishers",
      detail: topClickDecliners.length === 2
        ? `Total clicks changed ${cleanInlineText(m.clicks?.variance || "N/A")} (${cleanInlineText(m.clicks?.difference || "N/A")}). ${cleanInlineText(topClickDecliners[0].Publisher || "Publisher 1")} contributed ${cleanInlineText(topClickDecliners[0]["YoY Change"] || "N/A")} (${cleanInlineText(topClickDecliners[0]["YoY %"] || "N/A")}) and ${cleanInlineText(topClickDecliners[1].Publisher || "Publisher 2")} contributed ${cleanInlineText(topClickDecliners[1]["YoY Change"] || "N/A")} (${cleanInlineText(topClickDecliners[1]["YoY %"] || "N/A")}), together representing approximately ${top2Share} of total click loss.`
        : "Top click decline concentration could not be confirmed from available movers data."
    }
  ];

  return signals.slice(0, 5);
}

function readTableCell(row, aliases) {
  if (!row || typeof row !== "object") return "";
  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, alias)) return cleanInlineText(row[alias] || "");
    const key = Object.keys(row).find((candidate) => cleanInlineText(candidate).toLowerCase() === alias.toLowerCase());
    if (key) return cleanInlineText(row[key] || "");
  }
  return "";
}

function compactLabel(value, maxLen = 36) {
  const text = cleanInlineText(value || "");
  if (!text) return "";
  if (text.length <= maxLen) return text;
  return `${text.slice(0, maxLen - 1).trimEnd()}\u2026`;
}

function formatSignedPercent(value, locale = "en-GB", decimals = 1) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "N/A";
  const n = Number(value);
  const abs = Math.abs(n).toLocaleString(locale || "en-GB", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
  if (n > 0) return `+${abs}%`;
  if (n < 0) return `-${abs}%`;
  return `${abs}%`;
}

function buildDirectionalMoversTable(table, title, columns, upCount = 5, downCount = 5, locale = "en-GB") {
  const fallback = {
    title,
    columns,
    colW: [3.5, 1.8, 2.1, 2.1, 1.5],
    rows: [
      ["Top 5 Up (by YoY %)", "", "", "", ""],
      ["No qualifying publishers", "", "", "", ""],
      ["Top 5 Down (by YoY %)", "", "", "", ""],
      ["No qualifying publishers", "", "", "", ""]
    ]
  };
  if (!table || !Array.isArray(table.rows) || !table.rows.length) return fallback;

  const rows = table.rows.map((row) => {
    const publisher = compactLabel(readTableCell(row, ["Publisher"]), 36);
    const siteId = readTableCell(row, ["Site ID", "SiteID", "Site Id"]);
    const current = readTableCell(row, [columns[2]]);
    const yoyChange = readTableCell(row, ["YoY Change", "Sales YoY Change", "OV YoY Change", "Clicks YoY Change"]);
    const yoyPct = readTableCell(row, ["YoY %", "Sales YoY %", "OV YoY %", "Clicks YoY %"]);
    const direction = readTableCell(row, ["Direction"]).toLowerCase();
    const numericChange = parseNumber(yoyChange) ?? 0;
    const numericPct = parseNumber(yoyPct);
    const hasPct = Number.isFinite(numericPct);
    const rawPctText = cleanInlineText(yoyPct || "");
    const pctDisplayFallback = /^(n\/a|na|null|undefined|-)?$/i.test(rawPctText) ? "-" : rawPctText;

    const normalizedDirection = /(up|increase|increased|positive|growth|grow)/.test(direction)
      ? "up"
      : /(down|decrease|decreased|negative|decline|loss)/.test(direction)
        ? "down"
        : "";

    const directionalValue = hasPct ? numericPct : numericChange;
    const derivedDirection = normalizedDirection
      ? normalizedDirection
      : directionalValue < 0
        ? "down"
        : directionalValue > 0
          ? "up"
          : "";

    return {
      publisher,
      siteId,
      current,
      yoyChange,
      yoyPct: hasPct ? formatSignedPercent(numericPct, locale, 1) : pctDisplayFallback,
      direction: derivedDirection,
      numericChange,
      numericPct,
      hasPct
    };
  });
  const eligibleRows = rows.filter((row) =>
    row.hasPct
    && Number.isFinite(row.numericPct)
    && Math.abs(row.numericPct) > 0
    && cleanInlineText(row.publisher || "")
  );

  const up = eligibleRows
    .filter((row) => row.numericPct > 0)
    .sort((a, b) => {
      if ((b.numericPct || 0) !== (a.numericPct || 0)) return (b.numericPct || 0) - (a.numericPct || 0);
      return (b.numericChange || 0) - (a.numericChange || 0);
    })
    .slice(0, upCount);

  const down = eligibleRows
    .filter((row) => row.numericPct < 0)
    .sort((a, b) => {
      if ((a.numericPct || 0) !== (b.numericPct || 0)) return (a.numericPct || 0) - (b.numericPct || 0);
      return (a.numericChange || 0) - (b.numericChange || 0);
    })
    .slice(0, downCount);

  const outputRows = [];
  outputRows.push([`Top ${upCount} Up (by YoY %)`, "", "", "", ""]);
  if (up.length) {
    up.forEach((row) => outputRows.push([row.publisher || "", row.siteId || "", row.current || "", row.yoyChange || "", row.yoyPct || ""]));
  } else {
    outputRows.push(["No qualifying publishers", "", "", "", ""]);
  }

  outputRows.push([`Top ${downCount} Down (by YoY %)`, "", "", "", ""]);
  if (down.length) {
    down.forEach((row) => outputRows.push([row.publisher || "", row.siteId || "", row.current || "", row.yoyChange || "", row.yoyPct || ""]));
  } else {
    outputRows.push(["No qualifying publishers", "", "", "", ""]);
  }

  return {
    title,
    columns,
    colW: [3.5, 1.8, 2.1, 2.1, 1.5],
    rows: outputRows
  };
}

function buildTopNewProgramsTable(table) {
  const columns = [
    "Program ID",
    "Program Name",
    "Conversions",
    "Order Value",
    "Publisher Commission",
    "Digital Wallet",
    "Total Earnings"
  ];
  const rows = table && Array.isArray(table.rows)
    ? table.rows
      .slice()
      .sort((a, b) => {
        const byTotalEarnings = (parseNumber(readTableCell(b, ["Total Earnings", "Total Earning", "TotalEarnings"])) || 0)
          - (parseNumber(readTableCell(a, ["Total Earnings", "Total Earning", "TotalEarnings"])) || 0);
        if (byTotalEarnings !== 0) return byTotalEarnings;
        const byPublisherCommission = (parseNumber(readTableCell(b, ["Publisher Commission", "Current Publisher Commission", "Commission"])) || 0)
          - (parseNumber(readTableCell(a, ["Publisher Commission", "Current Publisher Commission", "Commission"])) || 0);
        if (byPublisherCommission !== 0) return byPublisherCommission;
        return cleanInlineText(readTableCell(a, ["Program Name", "Program", "Name"]))
          .localeCompare(cleanInlineText(readTableCell(b, ["Program Name", "Program", "Name"])));
      })
      .slice(0, 10)
      .map((row) => columns.map((column) => readTableCell(row, [column]) || "-"))
    : [];

  return {
    title: "Top 10 New Programs",
    columns,
    rows: rows.length ? rows : [columns.map(() => "-")],
    dense: false
  };
}

const CONNECTION_STATUS_META = [
  { id: 3, label: "Accepted", shortLabel: "Accepted", color: "#57A66C", sortRank: 0 },
  { id: 1, label: "Under Consideration", shortLabel: "Under Cons.", color: "#AFC4F5", sortRank: 1, aliases: ["Under consideration"] },
  { id: 2, label: "Hold Under Consideration", shortLabel: "Hold UC", color: "#F2C94C", sortRank: 2, aliases: ["Hold UC"] },
  { id: 0, label: "Not Connected", shortLabel: "Not Conn.", color: "#8A94A6", sortRank: 3 },
  { id: 6, label: "Hold Accepted", shortLabel: "Hold Acc.", color: "#F2994A", sortRank: 4 },
  { id: 9, label: "Ending", shortLabel: "Ending", color: "#9B51E0", sortRank: 5 },
  { id: 4, label: "Ended", shortLabel: "Ended", color: "#5B6372", sortRank: 6 },
  { id: 5, label: "Denied", shortLabel: "Denied", color: "#EB5757", sortRank: 7, aliases: ["Rejected"] }
];

const CONNECTION_STATUS_BY_ID = new Map(CONNECTION_STATUS_META.map((item) => [String(item.id), item]));
const CONNECTION_STATUS_BY_LABEL = new Map(
  CONNECTION_STATUS_META.flatMap((item) => [item.label, item.shortLabel, ...(item.aliases || [])]
    .filter(Boolean)
    .map((label) => [label.toLowerCase(), item]))
);

function connectionStatusMeta(statusId, statusText) {
  const idKey = cleanInlineText(statusId || "");
  if (CONNECTION_STATUS_BY_ID.has(idKey)) return CONNECTION_STATUS_BY_ID.get(idKey);
  const textKey = cleanInlineText(statusText || "").toLowerCase();
  if (CONNECTION_STATUS_BY_LABEL.has(textKey)) return CONNECTION_STATUS_BY_LABEL.get(textKey);
  return { id: idKey || "-", label: cleanInlineText(statusText || "Unknown"), shortLabel: cleanInlineText(statusText || "Unknown"), color: "#8A94A6" };
}

function buildProgramConnectionStatusTable(table) {
  const sourceRows = table && Array.isArray(table.rows) ? table.rows : [];
  const rows = sourceRows
    .map((row) => {
      const programId = readTableCell(row, ["Program ID", "ProgramID", "Program Id", "ID", "id"]);
      const programName = readTableCell(row, ["Program Name", "Program", "Name", "name"]);
      const statusId = readTableCell(row, ["Status ID", "StatusId", "statusId", "Connection Status ID"]);
      const status = readTableCell(row, ["Connection Status", "Status", "Status Name", "statusName"]);
      const createdDate = readTableCell(row, ["Created Date", "Created", "createdDate"]);
      const meta = connectionStatusMeta(statusId, status);
      return {
        programId: cleanInlineText(programId || ""),
        programName: cleanInlineText(programName || ""),
        statusId: cleanInlineText(statusId || String(meta.id || "")),
        status: meta.label,
        shortStatus: meta.shortLabel,
        color: meta.color,
        sortRank: Number.isFinite(meta.sortRank) ? meta.sortRank : 99,
        createdDate: cleanInlineText(createdDate || "")
      };
    })
    .filter((row) => row.programId || row.programName)
    .sort((a, b) => {
      if (a.sortRank !== b.sortRank) return a.sortRank - b.sortRank;
      const byName = String(a.programName || "").localeCompare(String(b.programName || ""), "en", { sensitivity: "base" });
      if (byName !== 0) return byName;
      return String(a.programId || "").localeCompare(String(b.programId || ""), "en", { sensitivity: "base", numeric: true });
    });

  const counts = {};
  rows.forEach((row) => {
    counts[row.status] = (counts[row.status] || 0) + 1;
  });

  return {
    title: "Program Connection Status",
    rows,
    counts,
    key: CONNECTION_STATUS_META
  };
}

function formatCompactCurrency(value, currencyCode, locale = "en-GB") {
  const symbol = getCurrencySymbol(currencyCode);
  const n = Number(value);
  if (!Number.isFinite(n)) return `${symbol}0`;
  const abs = Math.abs(n);
  const sign = n < 0 ? "-" : "";
  if (abs >= 1000000) {
    return `${sign}${symbol}${(abs / 1000000).toLocaleString(locale, {
      minimumFractionDigits: abs % 1000000 === 0 ? 0 : 1,
      maximumFractionDigits: 1
    })}m`;
  }
  if (abs >= 1000) {
    return `${sign}${symbol}${(abs / 1000).toLocaleString(locale, {
      minimumFractionDigits: abs % 1000 === 0 ? 0 : 1,
      maximumFractionDigits: 1
    })}k`;
  }
  return formatMoneyAmount(n, currencyCode, locale, { decimals: 0 });
}

function buildProgramGapAnalysisImpact(table, input) {
  const sourceRows = table && Array.isArray(table.rows) ? table.rows : [];
  const rows = sourceRows
    .map((row) => {
      const competitorPubComm = parseNumber(readTableCell(row, [
        "Competitor Pub Comm",
        "Competitor Commission Opportunity",
        "Pub Comm - Specified Sites",
        "Opportunity",
        "Value"
      ])) || 0;
      const gapType = readTableCell(row, ["Gap Type", "Type", "Bucket"]) || "Review";
      const action = readTableCell(row, ["Recommended Action", "Action", "Next Step"]) || "Review";
      const primaryStatus = readTableCell(row, [
        input.client,
        "Primary Publisher",
        "Primary Publisher Status",
        "VoucherCodes",
        "TopCashBack",
        "Status"
      ]) || "Review";
      const programName = readTableCell(row, ["Program Name", "Program", "Name"]) || "Program";
      return {
        programName,
        primaryStatus,
        competitorPubComm,
        gapType,
        action
      };
    })
    .filter((row) => row.programName && (row.competitorPubComm > 0 || row.primaryStatus || row.gapType));

  if (!rows.length) return null;

  const totalOpportunity = rows.reduce((sum, row) => sum + row.competitorPubComm, 0);
  const activationRows = rows.filter((row) => /activation/i.test(row.gapType) || /^accepted$/i.test(row.primaryStatus));
  const applicationRows = rows.filter((row) => /application/i.test(row.gapType) || /no connection/i.test(row.primaryStatus));
  const clickRows = rows.filter((row) => /click/i.test(row.gapType) || /^clicks$/i.test(row.primaryStatus));
  const recoveryRows = rows.filter((row) => /recovery/i.test(row.gapType) || /denied|ended|hold/i.test(row.primaryStatus));
  const activationValue = activationRows.reduce((sum, row) => sum + row.competitorPubComm, 0);
  const leadingGap = rows
    .slice()
    .sort((a, b) => b.competitorPubComm - a.competitorPubComm)[0];

  return {
    totalOpportunity,
    totalOpportunityDisplay: formatCompactCurrency(totalOpportunity, input.currencyCode, input.locale),
    totalPrograms: rows.length,
    activationPrograms: activationRows.length,
    applicationPrograms: applicationRows.length,
    clickPrograms: clickRows.length,
    recoveryPrograms: recoveryRows.length,
    activationValueDisplay: formatCompactCurrency(activationValue, input.currencyCode, input.locale),
    leadingProgram: compactLabel(leadingGap.programName, 34),
    leadingProgramValue: formatCompactCurrency(leadingGap.competitorPubComm, input.currencyCode, input.locale)
  };
}

function gapPriorityRank(row) {
  const gap = cleanInlineText(row.gapType || "").toLowerCase();
  const status = cleanInlineText(row.primaryStatus || "").toLowerCase();
  if (/activation/.test(gap) || /^accepted$/.test(status)) return 0;
  if (/click/.test(gap) || /^clicks$/.test(status)) return 1;
  if (/application/.test(gap) || /no connection/.test(status)) return 2;
  if (/recovery/.test(gap) || /denied|ended|hold/.test(status)) return 3;
  return 4;
}

function formatGapRegisterOpportunity(rawValue, numericValue, currencyCode, locale) {
  const raw = cleanInlineText(rawValue || "");
  const symbol = getCurrencySymbol(currencyCode);
  if (/^</.test(raw)) {
    const compact = raw
      .replace(/\bGBP\b/gi, symbol)
      .replace(/\s+/g, "")
      .replace(new RegExp(`^<${symbol}?`, "i"), `<${symbol}`);
    return compact || `<${symbol}100`;
  }
  const n = Number(numericValue);
  if (Number.isFinite(n) && n > 0) {
    return formatMoneyAmount(n, currencyCode, locale, { decimals: 0 });
  }
  return `${symbol}-`;
}

function buildProgramGapAnalysis(table, input) {
  const sourceRows = table && Array.isArray(table.rows) ? table.rows : [];
  const columns = table && Array.isArray(table.columns) ? table.columns : Object.keys(sourceRows[0] || {});
  const primaryAliases = [
    input.client,
    "Primary Publisher",
    "Primary Publisher Status",
    "VoucherCodes",
    "TopCashBack",
    "Status"
  ].map((item) => cleanInlineText(item).toLowerCase()).filter(Boolean);
  const knownColumns = new Set([
    "program name",
    "program",
    "name",
    "program id",
    "programid",
    "id",
    "competitor pub comm",
    "competitor commission opportunity",
    "pub comm - specified sites",
    "opportunity",
    "value",
    "gap type",
    "type",
    "bucket",
    "recommended action",
    "action",
    "next step",
    ...primaryAliases
  ]);
  const competitorColumns = columns
    .filter((column) => !knownColumns.has(cleanInlineText(column).toLowerCase()))
    .slice(0, 4);

  const rows = sourceRows
    .map((row) => {
      const rawCompetitorPubComm = readTableCell(row, [
        "Competitor Pub Comm",
        "Competitor Commission Opportunity",
        "Pub Comm - Specified Sites",
        "Opportunity",
        "Value"
      ]);
      const competitorPubComm = parseNumber(readTableCell(row, [
        "Competitor Pub Comm",
        "Competitor Commission Opportunity",
        "Pub Comm - Specified Sites",
        "Opportunity",
        "Value"
      ])) || 0;
      const programName = readTableCell(row, ["Program Name", "Program", "Name"]) || "Program";
      const programId = readTableCell(row, ["Program ID", "ProgramID", "Program Id", "ID", "id"]);
      const primaryStatus = readTableCell(row, [
        input.client,
        "Primary Publisher",
        "Primary Publisher Status",
        "VoucherCodes",
        "TopCashBack",
        "Status"
      ]) || "Review";
      const gapType = readTableCell(row, ["Gap Type", "Type", "Bucket"]) || "Review";
      const action = readTableCell(row, ["Recommended Action", "Action", "Next Step"]) || "Review";
      const competitorStatuses = competitorColumns
        .map((column) => readTableCell(row, [column]))
        .filter(Boolean);
      const earningCompetitors = competitorStatuses.filter((status) => /^pub comm$/i.test(status)).length;
      return {
        programName: cleanInlineText(programName),
        programId: cleanInlineText(programId),
        primaryStatus: cleanInlineText(primaryStatus),
        competitorPubComm,
        competitorPubCommDisplay: formatCompactCurrency(competitorPubComm, input.currencyCode, input.locale),
        registerOpportunityDisplay: formatGapRegisterOpportunity(rawCompetitorPubComm, competitorPubComm, input.currencyCode, input.locale),
        gapType: cleanInlineText(gapType),
        action: cleanInlineText(action),
        competitorSignal: competitorStatuses.length
          ? `${earningCompetitors}/${competitorStatuses.length} competitors earning`
          : (competitorPubComm > 0 ? "Competitor earning signal" : "No competitor earning signal"),
        priorityRank: 0
      };
    })
    .filter((row) => row.programName && (row.competitorPubComm > 0 || row.primaryStatus || row.gapType))
    .map((row) => ({ ...row, priorityRank: gapPriorityRank(row) }))
    .sort((a, b) => {
      if (a.priorityRank !== b.priorityRank) return a.priorityRank - b.priorityRank;
      if (b.competitorPubComm !== a.competitorPubComm) return b.competitorPubComm - a.competitorPubComm;
      return a.programName.localeCompare(b.programName, "en", { sensitivity: "base" });
    });

  if (!rows.length) return null;

  const revenueRows = rows
    .filter((row) => row.competitorPubComm > 0)
    .slice()
    .sort((a, b) => {
      if (b.competitorPubComm !== a.competitorPubComm) return b.competitorPubComm - a.competitorPubComm;
      if (a.priorityRank !== b.priorityRank) return a.priorityRank - b.priorityRank;
      return a.programName.localeCompare(b.programName, "en", { sensitivity: "base" });
    });
  const topRevenueRows = revenueRows.slice(0, 10);
  const topRevenueKeys = new Set(topRevenueRows.map((row) => `${row.programId || ""}::${row.programName || ""}`));
  const reportRows = rows.filter((row) => !topRevenueKeys.has(`${row.programId || ""}::${row.programName || ""}`));

  return {
    impact: buildProgramGapAnalysisImpact(table, input),
    rows,
    topRevenueRows,
    reportRows
  };
}

function paginateProgramGapAnalysisRegister(programGapAnalysis) {
  const rows = Array.isArray(programGapAnalysis?.rows) ? programGapAnalysis.rows : [];
  if (!rows.length) return [];
  const columnCount = 3;
  const rowsPerColumn = 22;
  const pageSize = columnCount * rowsPerColumn;
  const pageCount = Math.max(1, Math.ceil(rows.length / pageSize));
  return Array.from({ length: pageCount }, (_, pageIndex) => {
    const start = pageIndex * pageSize;
    return {
      rows: rows.slice(start, start + pageSize),
      totalRows: rows.length,
      pageIndex,
      pageCount,
      columnCount,
      rowsPerColumn
    };
  });
}

function programConnectionStatusColumnCount(rowCount) {
  if (rowCount > 100) return 6;
  if (rowCount > 72) return 5;
  return 4;
}

function programConnectionStatusRowsPerSlide(rowCount) {
  const columns = programConnectionStatusColumnCount(rowCount);
  return columns * 14;
}

function paginateProgramConnectionStatus(programConnectionStatus, cutoffDate) {
  const rows = Array.isArray(programConnectionStatus.rows) ? programConnectionStatus.rows : [];
  if (!rows.length) return [];

  const pageSize = programConnectionStatusRowsPerSlide(rows.length);
  const pageCount = Math.max(1, Math.ceil(rows.length / pageSize));
  return Array.from({ length: pageCount }, (_, pageIndex) => {
    const start = pageIndex * pageSize;
    return {
      ...programConnectionStatus,
      rows: rows.slice(start, start + pageSize),
      totalRows: rows.length,
      pageIndex,
      pageCount,
      cutoffDate: cleanInlineText(cutoffDate || "")
    };
  });
}

function buildProgramActivationSnapshot(table) {
  const rows = table && Array.isArray(table.rows) ? table.rows : [];
  const labelMap = new Map([
    ["joinedprograms", "Joined programs"],
    ["acceptedprograms", "Joined programs"],
    ["acceptances", "Joined programs"],
    ["withclicks", "With clicks"],
    ["wclicks", "With clicks"],
    ["programswithclicks", "With clicks"],
    ["withpubcommission", "Pub commission"],
    ["withpublishercommission", "Pub commission"],
    ["wpubcomm", "Pub commission"],
    ["wpublcomm", "Pub commission"],
    ["programswithpubcommission", "Pub commission"],
    ["inactive", "Inactive"],
    ["inactiveprograms", "Inactive"]
  ]);
  const fallback = [
    { label: "Joined programs", total: "-", newCount: "-", newPercent: "-" },
    { label: "With clicks", total: "-", newCount: "-", newPercent: "-" },
    { label: "Pub commission", total: "-", newCount: "-", newPercent: "-" },
    { label: "Inactive", total: "-", newCount: "-", newPercent: "-" }
  ];
  const byLabel = new Map();

  rows.forEach((row) => {
    const rawLabel = readTableCell(row, ["Metric", "Label", "Status", "Bucket", "Program Status"]);
    const normalized = cleanInlineText(rawLabel).toLowerCase().replace(/[^a-z0-9]/g, "");
    const label = labelMap.get(normalized) || cleanInlineText(rawLabel);
    if (!label) return;
    byLabel.set(label, {
      label,
      total: readTableCell(row, ["Total", "Count", "Programs", "Value"]) || "-",
      newCount: readTableCell(row, ["New", "New Programs", "New Count"]) || "-",
      newPercent: readTableCell(row, ["New %", "New Percent", "New Percentage", "NewPct"]) || "-"
    });
  });

  return fallback.map((item) => byLabel.get(item.label) || item);
}

function buildActionBullets(input) {
  const recommendations = input.recommendations.slice(0, 5);
  if (recommendations.length) return recommendations;

  const fromProgram = input.programSections.flatMap((section) => section.bullets).slice(0, 5);
  if (fromProgram.length) return fromProgram;

  const fromPublisher = input.publisherSections.flatMap((section) => section.bullets).slice(0, 5);
  if (fromPublisher.length) return fromPublisher;

  return [
    "Confirm top growth and decline publishers with account teams.",
    "Prioritize actions linked directly to measured KPI movement.",
    "Track impact owners and deadlines before the next QBR cycle."
  ];
}

function directionWord(varianceValue) {
  if (varianceValue === null || varianceValue === undefined || Number.isNaN(Number(varianceValue))) return "changed";
  if (Number(varianceValue) > 0) return "increased";
  if (Number(varianceValue) < 0) return "decreased";
  return "was flat";
}

function metricSentence(label, metric, includeDelta = true) {
  if (!metric) return `${label}: data not available.`;
  if (!includeDelta) return `${label}: ${metric.current || "-"}.`;
  const dir = directionWord(metric.varianceValue);
  const variance = cleanDeltaText(metric.variance) || "N/A";
  return `${label} ${dir} ${variance} (${metric.previous || "-"} to ${metric.current || "-"}).`;
}

function getTopDirection(table, directionLabel) {
  const base = tableRows(table, 20);
  if (!base || !base.rows.length) return null;
  const dirIdx = base.columns.findIndex((column) => cleanInlineText(column).toLowerCase() === "direction");
  const pubIdx = base.columns.findIndex((column) => cleanInlineText(column).toLowerCase().includes("publisher"));
  const changeIdx = base.columns.findIndex((column) => cleanInlineText(column).toLowerCase().includes("yoy change"));
  if (dirIdx === -1 || pubIdx === -1) return null;
  const row = base.rows.find((candidate) => cleanInlineText(candidate[dirIdx]).toLowerCase() === cleanInlineText(directionLabel).toLowerCase());
  if (!row) return null;
  const publisher = row[pubIdx] || "Publisher";
  const change = changeIdx > -1 ? row[changeIdx] : "";
  return change ? `${publisher} (${change})` : publisher;
}

function buildKpiAnalysisBullets(input) {
  const m = input.metricMap;
  const conversions = m.conversions || m.sales;
  const conv = m.convrate;
  const ov = m.ordervalue;
  const publisherCommission = m.publcommission || m.publishercommission;
  const digitalWallet = m.digitalwallet;
  const totalEarnings = m.totalearnings || m.totalcommission;
  const cleanPublisherLabel = (value) =>
    cleanInlineText(value || "")
      .replace(/\s*-\s*I$/i, "")
      .replace(/\s*-\s*I(?=\s|$)/gi, "")
      .replace(/\s+/g, " ")
      .trim();
  const sectionLooksLikeKpiAnalysis = (section) =>
    /kpi snapshot|kpi highlights|business implications|confirmed changes|program performance|analysis|insights|signals|implication/i
      .test(cleanInlineText(section?.title || "").toLowerCase());

  const aiCandidatesRaw = [...(input.programSections || []), ...(input.publisherSections || [])]
    .filter((section) =>
      sectionLooksLikeKpiAnalysis(section)
    )
    .flatMap((section) => [...(section.bullets || []), ...(section.paragraphs || [])])
    .map((line) => cleanDeltaText(cleanInlineText(line)))
    .filter((line) => line.length >= 14 && line.length <= 360)
    .filter((line) => !/\bsite\s*id\b/i.test(line))
    .filter((line) => !/^\s*program\s*id\s*\d+/i.test(line))
    .filter((line) => !/^\s*program\s*\d+\b/i.test(line))
    .filter((line) => !/\btotal order value\b[\s:;,-]*.*\byoy change\b/i.test(line))
    .filter((line) => !/\bcurrent sales:\b|\bcurrent ov:\b|\bov yoy change:\b|\bsales yoy %:\b/i.test(line))
    .filter((line) => !/^\d{4}-\d{2}-\d{2}\s+to\s+\d{4}-\d{2}-\d{2}/i.test(line));

  const isHeadingLike = (line) => {
    if (!line) return false;
    if (line.length > 70) return false;
    if (/[.:;!?]$/.test(line)) return false;
    if (/^[+\-]?\d/.test(line)) return false;
    return /^[A-Za-z0-9&()\-/' ]+$/.test(line);
  };

  const aiCandidates = [];
  for (let i = 0; i < aiCandidatesRaw.length; i += 1) {
    const current = aiCandidatesRaw[i];
    const next = aiCandidatesRaw[i + 1];
    if (isHeadingLike(current) && next && !isHeadingLike(next)) {
      aiCandidates.push(`${current}: ${next}`);
      i += 1;
      continue;
    }
    if (isHeadingLike(current)) continue;
    aiCandidates.push(current);
  }

  const looksLikeRawKpiSnapshot = (line) =>
    /^(sales|conversions?|order value|clicks|conv rate|conversion rate|aov|publ(?:isher)? commission(?:\s*\/\s*total commission)?|digital wallet|total earnings?|total commission|cpa|roi)\b.*:/i
      .test(cleanInlineText(line));
  const looksLikeProgramListing = (line) =>
    /^\s*program\s*id\s*\d+/i.test(cleanInlineText(line))
    || /^\s*program\s*\d+\b/i.test(cleanInlineText(line))
    || /\btotal order value\b[\s:;,-]*.*\byoy change\b/i.test(cleanInlineText(line));

  const preferredAi = aiCandidates
    .filter((line) => !looksLikeRawKpiSnapshot(line) && !looksLikeProgramListing(line))
    .map((line) => publisherAnalysisToProgramContext(line))
    .filter((line) => !/^driver not confirmed from available data\.?$/i.test(line))
    .filter((line) => !/^detail not available from current extract\.?$/i.test(line))
    .filter(Boolean);

  const declineRows = (input.tables.topDecliningPublishers?.rows || []).slice(0, 3);
  const declineList = declineRows
    .map((row) => `${cleanPublisherLabel(row["Program Name"] || row.Program || row.Publisher || "Program")} (${cleanInlineText(row["Sales YoY %"] || row["YoY Change"] || "N/A")})`)
    .filter(Boolean)
    .join(", ");

  const bullets = [
    `Conversion Rate Movement: ${metricSentence("Conversion rate", conv)} Conversions ${directionWord(conversions?.varianceValue)} ${conversions?.variance || "N/A"} (${conversions?.difference || "-"}), showing how efficiently programs converted traffic into outcomes.`,
    `Conversion Volume: Total conversions ${directionWord(conversions?.varianceValue)} ${conversions?.variance || "N/A"} (${conversions?.difference || "-"}). ${declineList ? `Largest declining programs were ${declineList}.` : "Largest declining program contribution requires confirmation from mover tables."}`,
    `Order Value: ${metricSentence("Total order value", ov)} Value generation moved alongside conversion volume and earnings performance.`,
    `Publisher Commission: ${metricSentence("Publisher commission", publisherCommission)} Commission remains the core earning component for the publisher.`,
    `Total Earnings: ${metricSentence("Total earnings", totalEarnings)} Digital wallet contribution was ${digitalWallet?.current || "-"} (${digitalWallet?.variance || "N/A"}), so total earnings combine commission and wallet income.`
  ];
  const generated = bullets.map((line) => publisherAnalysisToProgramContext(line)).filter(Boolean);

  const topicTitleByKey = {
    conversion: "Conversion Rate Movement",
    conversionVolume: "Conversion Volume",
    orderValue: "Order Value",
    commission: "Publisher Commission",
    totalEarnings: "Total Earnings"
  };
  const topicOrder = ["conversion", "conversionVolume", "orderValue", "commission", "totalEarnings"];

  const topicKeyFromText = (value) => {
    const text = cleanInlineText(value || "").toLowerCase();
    if (!text) return "";
    if (/conversion volume|conversions|sales volume pressure|sales/.test(text)) return "conversionVolume";
    if (/conversion rate|conv rate/.test(text)) return "conversion";
    if (/order value|total order value/.test(text)) return "orderValue";
    if (/publisher commission|commission/.test(text)) return "commission";
    if (/total earnings?|digital wallet|wallet/.test(text)) return "totalEarnings";
    return "";
  };

  const generatedByTopic = {};
  generated.forEach((line) => {
    const idx = line.indexOf(":");
    if (idx < 0) return;
    const key = topicKeyFromText(line.slice(0, idx));
    const detail = cleanInlineText(line.slice(idx + 1));
    if (key && detail && !generatedByTopic[key]) {
      generatedByTopic[key] = detail;
    }
  });

  const aiByTopic = {};
  const aiExtras = [];
  preferredAi.slice(0, 12).forEach((line) => {
    const text = cleanInlineText(line);
    if (!text) return;
    const idx = text.indexOf(":");
    const hasTitle = idx > 6 && idx < 90;
    const title = hasTitle ? text.slice(0, idx).trim() : text;
    const detail = hasTitle ? text.slice(idx + 1).trim() : text;
    const key = topicKeyFromText(title || text);
    if (key) {
      if (!aiByTopic[key] && detail) aiByTopic[key] = detail;
      return;
    }
    aiExtras.push(text);
  });

  const output = [];
  const seen = new Set();
  const pushUnique = (line) => {
    const cleaned = cleanInlineText(line);
    if (!cleaned) return;
    const fingerprint = cleaned.toLowerCase();
    if (seen.has(fingerprint)) return;
    seen.add(fingerprint);
    output.push(cleaned);
  };

  topicOrder.forEach((topicKey) => {
    const title = topicTitleByKey[topicKey] || "KPI Signal";
    const aiDetail = cleanInlineText(aiByTopic[topicKey] || "");
    const generatedDetail = cleanInlineText(generatedByTopic[topicKey] || "");

    let aiDetailForTopic = aiDetail;
    const detail = generatedDetail || aiDetailForTopic;
    if (detail) pushUnique(`${title}: ${detail}`);
  });

  if (output.length < 5) {
    aiExtras.forEach((line) => pushUnique(line));
  }
  if (output.length < 5) {
    generated.forEach((line) => pushUnique(line));
  }
  return output.slice(0, 5);
}

function buildCostCallout(input) {
  const m = input.metricMap;
  const base = "ROI = Total Order Value \u00F7 Total Commission. A higher ROI indicates greater return per \u00A31 of commission spend.";
  const pub = m.publcommission;
  const total = m.totalcommission;
  if (!pub || !total) return base;

  const pubCurrent = cleanInlineText(pub.current || "");
  const totalCurrent = cleanInlineText(total.current || "");
  if (pubCurrent && totalCurrent && pubCurrent === totalCurrent) {
    return `${base} Publisher Commission and Total Commission are equal in this period - no overrides recorded.`;
  }
  return base;
}

function buildRiskMitigation(riskType, evidence) {
  const combined = `${riskType} ${evidence}`.toLowerCase();
  if (/traffic|click/.test(combined) && /zero|0\s+conversion|no conversion/.test(combined)) {
    return "Pause scale-up until the publisher has tested the tracking link, landing page, and voucher route with a small controlled traffic sample that produces at least one verified conversion.";
  }
  if (/new|brand-new|activated|activation|joined/.test(combined) && /(zero|0|no)\s*(ov|order value|conversion|earning|commission)/.test(combined)) {
    return "For each zero-output new program, confirm tracking is live, add one visible launch placement or voucher, and set a first-week conversion target before adding more exposure.";
  }
  if (/concentration|dependency|dependenc|top\s+\d+/.test(combined)) {
    return "Protect the leading program's placement while adding two similar programs into the next campaign rotation so commission is not dependent on one partner.";
  }
  if (/traffic|click|landing/.test(combined) && /conversion|tracking|leak/.test(combined)) {
    return "Audit the affected program's landing path, voucher validity, and tracking parameters, then retest with one high-intent placement before increasing traffic.";
  }
  if (/decline|decrease|drop|down|negative|loss|fall|fell/.test(combined)) {
    return "Compare the program with its prior top-performing period, restore the best placement or offer that changed, and run a two-week recovery test against the lost metric.";
  }
  if (/commission|cpa|cost/.test(combined)) {
    return "Model commission against conversion value, reduce exposure on low-value traffic, and test a capped incentive where expected order value covers the payout.";
  }
  if (/order value|ov|aov/.test(combined)) {
    return "Prioritise higher-basket placements or bundles with the affected programs, then track order value and conversion mix before adding more low-value volume.";
  }
  if (/conversion|tracking/.test(combined)) {
    return "Check voucher validity, landing-page speed, and tracking handoff, then rerun the placement with the traffic source that historically converted best.";
  }
  return "Turn the finding into one controlled publisher test with a named program, target metric, start date, and success threshold for the next review.";
}

function formatRiskRationale(programName, riskType, evidence) {
  const program = cleanInlineText(programName || "");
  const risk = cleanInlineText(riskType || "Performance variance");
  const proof = cleanInlineText(evidence || "").replace(/[.]+$/g, "");
  const prefix = program ? `${program}: ${risk}` : risk;
  return proof ? `${prefix}. Evidence: ${proof}.` : prefix;
}

function compactRiskRationale(value, maxLen = 150) {
  const text = cleanInlineText(value || "")
    .replace(/\bEvidence:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text || text.length <= maxLen) return text;
  const firstSentence = text.match(/^(.{40,}?[.!?])\s+/);
  if (firstSentence && firstSentence[1].length <= maxLen) return firstSentence[1];
  const clipped = text.slice(0, maxLen);
  const lastSpace = clipped.lastIndexOf(" ");
  return (lastSpace > 80 ? clipped.slice(0, lastSpace) : clipped).trimEnd();
}

function buildRiskTileText(riskType, evidence, programName = "", action = "") {
  const rationale = compactRiskRationale(formatRiskRationale(programName, riskType, evidence), 210);
  const nextStep = cleanInlineText(action || buildRiskMitigation(riskType, evidence));
  return nextStep ? `${rationale} Action: ${nextStep}` : rationale;
}

function formatProgramList(names, limit = 10) {
  const unique = Array.from(new Set(
    (names || [])
      .map((name) => cleanInlineText(name || ""))
      .filter(Boolean)
  ));
  if (!unique.length) return "";
  const visible = unique.slice(0, limit).join(", ");
  const hidden = unique.length - limit;
  return hidden > 0 ? `${visible}, +${hidden} more` : visible;
}

function buildZeroOrderValueActivationRisk(input) {
  const brandNew = input.tables.brandNewPublishers;
  const rows = brandNew && Array.isArray(brandNew.rows) ? brandNew.rows : [];
  if (!rows.length) return null;

  const zeroOvRows = rows.filter((row) => {
    const orderValue = parseNumber(readTableCell(row, ["Order Value", "Current OV", "Current Order Value", "CurrentOrderValue", "CurrentOV"]));
    const conversions = parseNumber(readTableCell(row, ["Conversions", "Current Conversions", "Sales", "Current Sales"]));
    const totalEarnings = parseNumber(readTableCell(row, ["Total Earnings", "Total Earning", "TotalEarnings"]));
    return (orderValue || 0) <= 0 && (conversions || 0) <= 0 && (totalEarnings || 0) <= 0;
  });

  if (!zeroOvRows.length) return null;

  const names = zeroOvRows.map((row) => readTableCell(row, ["Program Name", "Program", "Name"]));
  const namedPrograms = formatProgramList(names);
  const riskType = "New programs activated with no recorded OV";
  const evidence = `${zeroOvRows.length} of ${rows.length} new programs show 0 order value, 0 conversions, and 0 total earnings${namedPrograms ? `: ${namedPrograms}` : ""}`;
  const impact = zeroOvRows.length === rows.length ? "High" : "Medium";

  return [
    buildRiskTileText(riskType, evidence, namedPrograms || "New programs")
  ];
}

function buildStructuredRiskRows(input) {
  const table = input.tables.riskDependenciesTable;
  const rows = table && Array.isArray(table.rows) ? table.rows : [];
  return rows
    .map((row) => {
      const programName = readTableCell(row, ["Program Name", "Program", "Name"]);
      const riskType = readTableCell(row, ["Risk Type", "Risk", "Risk / Dependency", "Issue", "Dependency", "Theme", "Area"]) || "Performance variance";
      const evidence = readTableCell(row, ["Evidence", "Analysis", "Rationale", "Reason", "Impact Detail", "Dependency Detail", "Detail", "Details", "Observation"]);
      const providedMitigation = readTableCell(row, ["Mitigation", "Action", "Recommended Action", "Recommendation", "Next Step", "Next Steps", "Owner Action"]);
      return [
        buildRiskTileText(riskType, evidence, programName, providedMitigation)
      ];
    })
    .filter((row) => row.some(Boolean));
}

function buildRiskRows(input) {
  const structuredRows = buildStructuredRiskRows(input);
  const activationRisk = buildZeroOrderValueActivationRisk(input);
  if (activationRisk) return [...structuredRows.slice(0, 3), activationRisk].slice(0, 4);
  if (structuredRows.length) return structuredRows.slice(0, 4);

  const sourceLines = [];
  const appendLines = (lines, limit = 4) => {
    (lines || []).forEach((line) => {
      if (sourceLines.length >= limit) return;
      const text = publisherAnalysisToProgramContext(line);
      if (!text) return;
      if (sourceLines.some((existing) => existing.toLowerCase() === text.toLowerCase())) return;
      sourceLines.push(text);
    });
  };

  appendLines(input.publisherOverviewObservations);
  if (sourceLines.length < 2) appendLines(buildPublisherOverviewBullets(input));
  if (sourceLines.length < 2) appendLines(buildKpiAnalysisBullets(input));

  const inferRiskLabel = (text) => {
    const lower = cleanInlineText(text).toLowerCase();
    if (/(concentration|top\s+\d+\s+program|dependency|dependenc)/.test(lower)) return "Program concentration risk";
    if (/\bcpa\b|commission|cost per acquisition/.test(lower)) return "Rising CPA trend";
    if (/click|traffic|volume decline/.test(lower)) return "Traffic decline";
    if (/aov|order value|ov /.test(lower)) return "Order value mix volatility";
    if (/\broi\b|return on investment/.test(lower)) return "Return efficiency risk";
    if (/conversion/.test(lower)) return "Conversion quality dependency";
    return "Performance variance risk";
  };

  const inferImpact = (text) => {
    const lower = cleanInlineText(text).toLowerCase();
    if (/(declined|decrease|drop|largest decline|risk|high|pressure)/.test(lower)) return "High";
    if (/(marginal|flat|mixed|moderate|watch)/.test(lower)) return "Medium";
    return "Medium";
  };

  const rows = sourceLines
    .slice(0, 4)
    .map((line) => [
      buildRiskTileText(inferRiskLabel(line), line)
    ]);

  if (rows.length) return rows;

  return [
    [
      buildRiskTileText("Program concentration risk", "Contribution is concentrated in the top programs.")
    ],
    [
      buildRiskTileText("Rising CPA trend", "CPA growth is outpacing sales efficiency.")
    ],
    [
      buildRiskTileText("Traffic decline", "Traffic quality or volume is below the prior benchmark.")
    ]
  ];
}

function buildDeckSpec(input, theme) {
  const slides = [];
  const headline = buildHeadline(input);
  const executiveCardConfig = [
    { key: "conversions", label: "Conversions", iconKey: "sales", icon: "\u2630" },
    { key: "ordervalue", label: "Total Order Value", iconKey: "ordervalue", icon: "\u25A4" },
    { key: "publcommission", label: "Publisher Commission", iconKey: "commission", icon: "\u00A3" },
    { key: "digitalwallet", label: "Digital Wallet", iconKey: "commission", icon: "\u00A3" },
    { key: "totalearnings", label: "Total Earnings", iconKey: "commission", icon: "\u00A3" },
    { key: "convrate", label: "Conversion Rate", iconKey: "convrate", icon: "\u26A1" }
  ];
  const topCards = executiveCardConfig
    .map((cfg) => {
      const card = metricCard(input.metricMap[cfg.key]);
      if (!card) return null;
      const hasIconPath = Boolean(cfg.iconKey && HAS_KPI_ICON[cfg.iconKey]);
      return {
        ...card,
        label: cfg.label,
        icon: cfg.icon,
        iconPath: hasIconPath ? KPI_ICON_PATHS[cfg.iconKey] : ""
      };
    })
    .filter(Boolean);
  const executiveNarrative = buildExecutiveSummaryText(input);
  const reportingSummary = `${input.reportingPeriod} vs ${input.comparisonPeriod}`;

  const volumeRows = buildMetricRows(input.metricMap, [
    ["conversions", "Conversions"],
    ["convrate", "Conversion Rate"],
    ["clicks", "Clicks"],
    ["earningsperclick", "Earnings per Click"],
    ["earningsperconversion", "Earnings per Conversion"],
    ["ordervalue", "Total Order Value"],
    ["publcommission", "Publisher Commission"],
    ["digitalwallet", "Digital Wallet"],
    ["totalearnings", "Total Earnings"],
    ["activeprograms", "Active Programs"]
  ]);

  const brandNew = input.tables.brandNewPublishers;
  const programConnectionStatus = buildProgramConnectionStatusTable(input.tables.programConnectionStatusTable);
  const programConnectionStatusPages = paginateProgramConnectionStatus(
    programConnectionStatus,
    input.programStatusCreatedFromDate
  );
  const programGapAnalysis = buildProgramGapAnalysis(input.tables.programGapAnalysisTable, input);
  const programGapAnalysisImpact = programGapAnalysis?.impact;
  const programGapAnalysisTopPrograms = Array.isArray(programGapAnalysis?.topRevenueRows)
    ? programGapAnalysis.topRevenueRows
    : [];
  const programGapAnalysisReportRows = Array.isArray(programGapAnalysis?.reportRows)
    ? programGapAnalysis.reportRows
    : [];
  const moversCommissionChart = buildMoversCommissionBarChart(input.tables.moversCommissionChart || input.tables.moversCommission, input.locale);
  const competitorAnalysis = buildCompetitorAnalysisTable(input.tables.competitorAnalysisTable);
  const competitorShareChart = buildCompetitorSharePubCommChart(input.tables.competitorSharePubCommChart, input.tables.competitorAnalysisTable);
  const competitorWeeklyChart = buildWeeklyPubCommComboChart(input.tables.competitorWeeklyPubCommChart);
  const topProgramsCompetitorPerformance = buildTopProgramsCompetitorPerformanceTable(input.tables.topProgramsCompetitorPerformanceTable);
  const activationSnapshot = buildProgramActivationSnapshot(input.tables.programActivationSnapshotTable);
  const kpiAnalysisBullets = buildKpiAnalysisBullets(input);
  const publisherOverviewBullets = buildPublisherOverviewBullets(input);
  const segmentPerformanceBlocks = buildSegmentPerformanceBlocks(input);
  const salesGrowthSignals = buildSalesGrowthSignals(input);
  const programBreakdownTable = buildProgramBreakdownTable(input);

  slides.push({
    id: "cover",
    kind: "cover",
    title: titleCaseWords(`${input.client} performance review`),
    subtitle: "",
    headline,
    summary: input.qbrFocusDetail
      ? `${input.qbrFocus}. ${input.qbrFocusDetail}`
      : `A focused year-over-year review of ${input.client} publisher performance, conversion dynamics, earnings and strategic priorities to drive growth and optimise outcomes.`,
    bullets: [`Client: ${input.client}`, `Reporting currency: ${input.currencyCode}`, `Language: ${input.languageName}`],
    kpis: [],
    tables: []
  });

  slides.push({
    id: "reporting-period",
    kind: "reporting-period",
    title: "Reporting Period",
    subtitle: "",
    headline: "",
    summary: "",
    bullets: [
      `${input.reportingPeriod}`
    ],
    kpis: [],
    tables: [],
    callout: ""
  });

  slides.push({
    id: "executive-summary",
    kind: "program-executive-summary",
    title: "Program Performance: Executive Summary",
    headline: "",
    summary: executiveNarrative,
    bullets: [],
    kpis: topCards,
    tables: []
  });

  slides.push({
    id: "kpi-volume-conversion",
    kind: "kpi-table",
    title: "KPI Summary Table: Conversions & Earnings",
    subtitle: "Unified KPI breakdown vs prior year.",
    bullets: [],
    kpis: [],
    tables: [
      {
        title: "KPI Summary",
        columns: ["Metric", "Recent", "Previous", "Difference", "% Variance"],
        rows: volumeRows.length ? volumeRows : [["-", "-", "-", "-", "-"]],
        colAlign: ["left", "center", "center", "center", "center"],
        dense: false
      }
    ],
    footerNote: "Total earnings = Publisher Commission + Digital Wallet. Conversion rate uses the source API's conversion-rate basis."
  });

  slides.push({
    id: "program-activation-snapshot",
    kind: "program-activation-snapshot",
    title: "Program Activation Snapshot",
    subtitle: "",
    bullets: [],
    kpis: [],
    tables: [],
    activationSnapshot,
    summary: "A clean status view for newly joined primary-publisher programs, placed directly after the KPI summary."
  });

  slides.push({
    id: "kpi-cost-roi",
    kind: "program-breakdown",
    title: "Program-Level Analysis: Publisher Commission",
    subtitle: "Per-program view ordered by publisher commission.",
    bullets: [],
    kpis: [],
    tables: [programBreakdownTable]
  });

  slides.push({
    id: "movers-shakers-publisher-commission",
    kind: "movers-bar-chart",
    title: "Movers & Shakers: Publisher Commission",
    subtitle: "Top 10 positive and top 10 negative period-on-period publisher commission changes.",
    bullets: [],
    kpis: [],
    tables: [],
    chart: moversCommissionChart
  });

  slides.push({
    id: "brand-new-publishers",
    kind: "publisher-table",
    title: "Top 10 New Programs",
    subtitle: "Programs joined or first active for the primary publisher in the current period.",
    bullets: [],
    kpis: [],
    tables: [
      buildTopNewProgramsTable(brandNew)
    ],
    callout: "New programs are included where current-period activity exists and no prior-period baseline was found."
  });

  if (programGapAnalysisImpact) {
    slides.push({
      id: "gap-analysis-impact",
      kind: "gap-analysis-impact",
      title: "Growth opportunity in the competitor gap",
      subtitle: "",
      bullets: [],
      kpis: [
        {
          label: "Competitor pub comm opportunity",
          value: programGapAnalysisImpact.totalOpportunityDisplay,
          summary: `${programGapAnalysisImpact.totalPrograms} programs where the primary publisher is not earning`
        },
        {
          label: "Gap programs",
          value: String(programGapAnalysisImpact.totalPrograms),
          summary: "Programs requiring activation, application or recovery"
        },
        {
          label: "Activation opportunities",
          value: String(programGapAnalysisImpact.activationPrograms),
          summary: `${programGapAnalysisImpact.activationValueDisplay} already accepted but inactive`
        }
      ],
      tables: [],
      gapImpact: programGapAnalysisImpact
    });
    if (programGapAnalysisTopPrograms.length) {
      slides.push({
        id: "gap-analysis-top-programs",
        kind: "gap-analysis-top-programs",
        title: "Top 10 competitor-funded gaps",
        subtitle: "Programs receiving publisher commission from competitors while the primary publisher is not earning.",
        bullets: [],
        kpis: [],
        tables: [],
        gapTopPrograms: {
          rows: programGapAnalysisTopPrograms,
          totalRows: programGapAnalysis.rows.length,
          reportRows: programGapAnalysisReportRows.length
        }
      });
    }
  }

  programConnectionStatusPages.forEach((programConnectionStatusPage, pageIndex) => {
    slides.push({
      id: pageIndex === 0 ? "program-connection-status" : `program-connection-status-${pageIndex + 1}`,
      kind: "program-connection-status",
      title: programConnectionStatusPages.length > 1
        ? `Program Connection Status (${pageIndex + 1}/${programConnectionStatusPages.length})`
        : "Program Connection Status",
      subtitle: "",
      bullets: [],
      kpis: [],
      tables: [],
      programConnectionStatus: programConnectionStatusPage
    });
  });

  slides.push({
    id: "competitor-analysis",
    kind: "competitor-analysis",
    title: "Competitor Analysis",
    subtitle: "Anonymous comparison of publisher commission coverage and white-space programmes.",
    bullets: [],
    kpis: [],
    tables: [competitorAnalysis],
    chart: competitorWeeklyChart,
    callout: "Pub Comm of the above is calculated only for programmes where the primary publisher has 0 publisher commission."
  });

  slides.push({
    id: "competitor-share-pub-comm",
    kind: "competitor-share-bar-chart",
    title: "Share Within Competitor Group",
    subtitle: "Publisher commission share for the primary publisher and four comparison publishers.",
    bullets: [],
    kpis: [],
    tables: [],
    chart: competitorShareChart
  });

  slides.push({
    id: "top-programs-competitor-performance",
    kind: "competitor-performance-table",
    title: "Top Programs Competitor Performance",
    subtitle: "Share of publisher commission by program across the primary publisher and four comparison publishers.",
    bullets: [],
    kpis: [],
    tables: [topProgramsCompetitorPerformance]
  });

  slides.push({
    id: "kpi-highlights",
    kind: "insights-blue",
    title: "KPI Highlights & Business Implications",
    subtitle: "What the numbers mean for the business - key signals and context.",
    bullets: kpiAnalysisBullets,
    kpis: [],
    tables: []
  });

  slides.push({
    id: "risks-dependencies",
    kind: "risks-dependencies",
    title: "Risks & Dependencies",
    subtitle: "",
    bullets: [],
    kpis: [],
    tables: [
      {
        title: "Risks & Dependencies",
        columns: ["Risk / rationale"],
        colW: [12.5],
        rows: buildRiskRows(input),
        dense: false
      }
    ]
  });

  slides.push({
    id: "thank-you",
    kind: "thank-you",
    title: `${capitalizeFirstLetter(input.client)} - Thank you.`,
    subtitle: "",
    bullets: [],
    kpis: [],
    tables: []
  });

  if (input.includeAppendix && input.metrics.length) {
    slides.push({
      id: "appendix-program-yoy",
      kind: "appendix",
      title: "Appendix: Program YoY Table",
      bullets: [],
      kpis: [],
      tables: [
        {
          title: "Program YoY Table",
          columns: ["Metric", "Current", "Previous", "Difference", "Variance"],
          rows: input.metrics.map((metric) => [
            metric.label,
            metric.current || "-",
            metric.previous || "-",
            metric.difference || "-",
            metric.variance || "-"
          ]),
          dense: true
        }
      ]
    });
  }

  return {
    metadata: {
      requestId: input.requestId,
      client: input.client,
      deckTitle: input.deckTitle,
      reportingPeriod: input.reportingPeriod,
      comparisonPeriod: input.comparisonPeriod,
      languageCode: input.languageCode,
      languageName: input.languageName,
      locale: input.locale || "en-GB",
      uiLabels: uiLabelsForLanguage(input.languageCode),
      currencyCode: input.currencyCode,
      qbrFocus: input.qbrFocus,
      analysisProgramIds: Array.isArray(input.analysisProgramIds) ? input.analysisProgramIds : [],
      generatedAt: new Date().toISOString()
    },
    theme,
    slides,
    reports: {
      gapAnalysis: programGapAnalysis
        ? {
            title: `${input.client} Gap Analysis Report`,
            client: input.client,
            currencyCode: input.currencyCode,
            locale: input.locale || "en-GB",
            topRows: programGapAnalysisTopPrograms,
            remainingRows: programGapAnalysisReportRows,
            allRows: programGapAnalysis.rows
          }
        : null
    }
  };
}

function toColor(hex) {
  return String(hex || "#000000").replace(/^#/, "");
}

function isBlueKind(kind) {
  return ["cover", "insights-blue", "sales-growth-signals-blue", "recommendations-blue", "segment-performance-blue", "program-activation-snapshot", "thank-you"].includes(kind);
}

function titleRuns(title) {
  const text = cleanInlineText(title);
  if (!text) return [{ text: "Slide Title", options: {} }];
  const phrases = [
    "Growth Publishers",
    "Decline Publishers",
    "Current Performers",
    "Segment Performance",
    "Strategic Recommendations",
    "Risks & Dependencies",
    "Reporting Period",
    "Priority Actions",
    "Performance Overview",
    "Business Implications",
    "Publishers",
    "Overview",
    "Actions",
    "Period"
  ];

  const lower = text.toLowerCase();
  const phrase = phrases.find((candidate) => lower.includes(candidate.toLowerCase()));
  if (!phrase) return [{ text, options: {} }];

  const idx = lower.indexOf(phrase.toLowerCase());
  const before = text.slice(0, idx);
  const middle = text.slice(idx, idx + phrase.length);
  const after = text.slice(idx + phrase.length);

  const runs = [];
  if (before) runs.push({ text: before, options: {} });
  runs.push({ text: middle, options: { color: toColor(DEFAULT_THEME.colors.accent) } });
  if (after) runs.push({ text: after, options: {} });
  return runs;
}

function addDotPattern(slide, x, y, color, transparency = 35) {
  const cols = 9;
  const rows = 7;
  const size = 0.025;
  const gap = 0.11;
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      slide.addShape("ellipse", {
        x: x + col * gap,
        y: y + row * gap,
        w: size,
        h: size,
        line: { color: toColor(color), pt: 0 },
        fill: { color: toColor(color), transparency }
      });
    }
  }
}

function drawPolyline(slide, points, color, lineTransparency = 72, pt = 0.9) {
  for (let i = 0; i < points.length - 1; i += 1) {
    const [x1, y1] = points[i];
    const [x2, y2] = points[i + 1];
    slide.addShape("line", {
      x: x1,
      y: y1,
      w: x2 - x1,
      h: y2 - y1,
      line: { color: toColor(color), pt, transparency: lineTransparency }
    });
  }
}

function addWireframeWatermark(slide, color, lineTransparency = 72) {
  const loops = [
    [[9.55, 0.2], [10.55, 0.0], [11.5, 0.22], [11.95, 1.0], [11.5, 1.76], [10.56, 1.92], [9.78, 1.35], [9.55, 0.2]],
    [[10.72, 0.54], [11.62, 0.7], [12.12, 1.34], [11.86, 2.08], [11.08, 2.28], [10.38, 1.84], [10.25, 1.08], [10.72, 0.54]],
    [[10.06, 1.86], [10.9, 2.04], [11.3, 2.78], [10.96, 3.52], [10.2, 3.7], [9.46, 3.24], [9.32, 2.48], [10.06, 1.86]],
    [[11.36, 2.32], [12.12, 2.58], [12.56, 3.3], [12.28, 4.02], [11.55, 4.28], [10.98, 3.92], [10.88, 3.16], [11.36, 2.32]],
    [[9.86, 3.5], [10.45, 3.78], [10.7, 4.44], [10.28, 5.03], [9.68, 5.12], [9.22, 4.7], [9.2, 4.08], [9.86, 3.5]]
  ];

  loops.forEach((loop) => drawPolyline(slide, loop, color, lineTransparency, 0.9));

  const links = [
    [[10.55, 0.0], [11.62, 0.7]],
    [[10.56, 1.92], [10.9, 2.04]],
    [[11.08, 2.28], [11.36, 2.32]],
    [[10.96, 3.52], [10.7, 4.44]],
    [[11.55, 4.28], [10.7, 4.44]],
    [[9.55, 0.2], [10.06, 1.86]],
    [[11.95, 1.0], [12.12, 2.58]],
    [[9.78, 1.35], [9.2, 4.08]]
  ];

  links.forEach((segment) => drawPolyline(slide, segment, color, lineTransparency + 7, 0.7));
}

function addSlideWatermark(slide, deck, isBlueSlide) {
  if ((isBlueSlide && HAS_TEMPLATE_BLUE_BG) || (!isBlueSlide && HAS_TEMPLATE_LIGHT_BG)) {
    return;
  }
  const watermarkColor = isBlueSlide ? deck.theme.colors.paper : deck.theme.colors.accent;
  addWireframeWatermark(slide, watermarkColor, isBlueSlide ? 70 : 82);
  addDotPattern(slide, 0.08, 6.2, watermarkColor, isBlueSlide ? 42 : 62);
}

function addTemplateBackgroundImage(slide, imagePath) {
  slide.addImage({
    path: imagePath,
    x: 0,
    y: 0,
    w: 13.333,
    h: 7.5
  });
}

function addLightChrome(slide, deck) {
  slide.background = { color: toColor(deck.theme.colors.paper) };
  if (HAS_TEMPLATE_LIGHT_BG) {
    addTemplateBackgroundImage(slide, TEMPLATE_LIGHT_BG_PATH);
  }
}

function addBlueChrome(slide, deck) {
  slide.background = { color: toColor(deck.theme.colors.accent) };
  if (HAS_TEMPLATE_BLUE_BG) {
    addTemplateBackgroundImage(slide, TEMPLATE_BLUE_BG_PATH);
  }
}

function addCyanFifthElementWireframe(slide, box) {
  if (!HAS_TD_FIFTH_ELEMENT_WIREFRAME_CYAN) return;
  slide.addImage({
    path: TD_FIFTH_ELEMENT_WIREFRAME_CYAN_PATH,
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    transparency: box.transparency || 0
  });
}

function addTitle(slide, deck, spec, color, subtitleColor, isBlueSlide = false) {
  const titleText = cleanInlineText(spec.title, "Slide Title");
  const titleRunsData = isBlueSlide ? [{ text: titleText, options: {} }] : titleRuns(titleText);

  slide.addText(titleRunsData, {
    x: 0.7,
    y: 0.58,
    w: 11.8,
    h: 0.62,
    fontFace: deck.theme.fonts.heading,
    fontSize: 28,
    color: toColor(color),
    margin: 0
  });
  if (spec.subtitle) {
    slide.addText(spec.subtitle, {
      x: 0.7,
      y: 1.44,
      w: 12.0,
      h: 0.34,
      fontFace: deck.theme.fonts.body,
      fontSize: 11,
      color: toColor(subtitleColor),
      align: "center",
      margin: 0
    });
  }
  addSlideWatermark(slide, deck, isBlueSlide);
}

function addBullets(slide, deck, bullets, box, color) {
  if (!bullets || !bullets.length) return;
  slide.addText(bullets.map((item) => `\u2022 ${item}`).join("\n"), {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontFace: deck.theme.fonts.body,
    fontSize: 12,
    color: toColor(color || deck.theme.colors.ink),
    breakLine: true,
    margin: 0.06
  });
}

function isTableValueNumeric(value) {
  const text = cleanInlineText(value || "");
  if (!text || text === "-" || /^n\/a$/i.test(text)) return false;
  const normalized = text
    .replace(/[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE]/g, "")
    .replace(/%$/, "")
    .replace(/^\+/, "");
  return /^-?\d+(\.\d+)?$/.test(normalized);
}

function inferTableColumnAlign(table, columnIndex, columnName) {
  if (columnIndex === 0) return "left";
  const explicitAlign = Array.isArray(table.colAlign) && table.colAlign.length === table.columns.length
    ? table.colAlign[columnIndex]
    : null;
  if (explicitAlign === "left" || explicitAlign === "center" || explicitAlign === "right") {
    return explicitAlign;
  }

  const numericHeaderPattern = /sales|clicks?|impressions?|ov|order value|commission|cost|cpa|roi|aov|rate|variance|change|yoy|publishers?|count|rank|id|current|prior|previous|total/i;
  const sampleValues = (table.rows || [])
    .slice(0, 8)
    .map((row) => row[columnIndex]);
  const numericCount = sampleValues.filter(isTableValueNumeric).length;

  if (sampleValues.length && numericCount >= Math.ceil(sampleValues.length * 0.5)) return "right";
  if (numericHeaderPattern.test(cleanInlineText(columnName || ""))) return "right";
  return "left";
}

function addCallout(slide, deck, text, y, darkText = true) {
  if (!text) return;
  slide.addShape("roundRect", {
    x: 0.7,
    y,
    w: 11.95,
    h: 0.92,
    radius: 0.04,
    line: { color: toColor(deck.theme.colors.highlight), pt: 0.5 },
    fill: { color: toColor(deck.theme.colors.highlight), transparency: 12 }
  });
  slide.addText(`\u25AD  ${text}`, {
    x: 0.95,
    y: y + 0.19,
    w: 11.4,
    h: 0.55,
    fontFace: deck.theme.fonts.body,
    fontSize: 10.5,
    color: toColor(darkText ? deck.theme.colors.ink : deck.theme.colors.paper),
    breakLine: true,
    margin: 0
  });
}

function kpiDeltaColor(deck, card) {
  const delta = cleanDeltaText(card && card.delta);
  if (card && card.trend === "up") return toColor(deck.theme.colors.success);
  if (card && card.trend === "down") return toColor(deck.theme.colors.accentAlt);
  if (card && card.trend === "flat") return toColor(deck.theme.colors.warning);
  if (delta.startsWith("+")) return toColor(deck.theme.colors.success);
  if (delta.startsWith("-")) return toColor(deck.theme.colors.accentAlt);
  if (/^(0(?:\.0+)?%?|n\/a|na|-)?$/i.test(delta)) return toColor(deck.theme.colors.warning);
  return toColor(deck.theme.colors.ink);
}

function kpiSummaryParts(card) {
  const summary = cleanInlineText(card.summary || card.value || "-");
  const delta = cleanDeltaText(card.delta || "");
  const finalDeltaMatch = summary.match(/([+-]\s?\d+(?:[.,]\d+)?%|0(?:[.,]0+)?%)\s*$/);
  if (!delta && !finalDeltaMatch) return { summary };

  const matchedDelta = finalDeltaMatch ? finalDeltaMatch[1].replace(/\s+/g, "") : delta;
  let idx = delta ? summary.lastIndexOf(delta) : -1;
  let deltaText = delta;
  if (idx < 0 && finalDeltaMatch) {
    deltaText = finalDeltaMatch[1];
    idx = finalDeltaMatch.index;
  }
  if (idx < 0) return { summary };

  const before = summary.slice(0, idx).trimEnd();
  const after = summary.slice(idx + deltaText.length);
  if (after.trim()) return { summary };
  return { before, delta: matchedDelta };
}

function addKpis(slide, deck, cards, origin, mode = "light") {
  const visible = (cards || []).slice(0, 6);
  if (!visible.length) return;
  const columns = 3;
  const cardW = 3.82;
  const cardH = 1.48;
  const gapX = 0.24;
  const gapY = 0.48;
  visible.forEach((card, index) => {
    const useDiamond = mode === "diamond" && visible.length === 5;
    let x;
    let y;
    if (useDiamond && index >= 3) {
      x = origin.x + ((index - 3) * (cardW + gapX)) + ((cardW + gapX) / 2);
      y = origin.y + cardH + gapY;
    } else {
      const col = index % columns;
      const row = Math.floor(index / columns);
      x = origin.x + col * (cardW + gapX);
      y = origin.y + row * (cardH + gapY);
    }
    slide.addShape("rect", {
      x,
      y,
      w: cardW,
      h: cardH,
      line: { color: toColor(deck.theme.colors.border), pt: 0.7 },
      fill: { color: toColor(deck.theme.colors.paper), transparency: (mode === "blue" || mode === "diamond-blue") ? 6 : 0 }
    });
    slide.addShape("rect", {
      x,
      y,
      w: cardW,
      h: 0.08,
      line: { color: toColor(deck.theme.colors.accent), pt: 0 },
      fill: { color: toColor(deck.theme.colors.accent) }
    });
    slide.addShape("ellipse", {
      x: x + (cardW / 2) - 0.27,
      y: y - 0.27,
      w: 0.54,
      h: 0.54,
      line: { color: toColor(deck.theme.colors.accent), pt: 0 },
      fill: { color: toColor(deck.theme.colors.accent) }
    });
    if (card.iconPath) {
      slide.addImage({
        path: card.iconPath,
        x: x + (cardW / 2) - 0.11,
        y: y - 0.12,
        w: 0.22,
        h: 0.22
      });
    } else {
      slide.addText(card.icon || String(index + 1), {
        x: x + (cardW / 2) - 0.13,
        y: y - 0.13,
        w: 0.26,
        h: 0.24,
        align: "center",
        valign: "mid",
        fontFace: "Segoe UI Symbol",
        fontSize: 12,
        bold: false,
        color: toColor(deck.theme.colors.paper),
        margin: 0
      });
    }
    slide.addText(card.label, {
      x: x + 0.22,
      y: y + 0.50,
      w: cardW - 0.3,
      h: 0.24,
      fontFace: deck.theme.fonts.body,
      fontSize: 10.5,
      color: toColor(deck.theme.colors.ink),
      margin: 0
    });
    const summaryParts = kpiSummaryParts(card);
    const summaryOptions = {
      y: y + 0.86,
      h: 0.52,
      fontFace: deck.theme.fonts.body,
      fontSize: 10.2,
      margin: 0,
      breakLine: true
    };
    if (summaryParts.delta) {
      slide.addText(summaryParts.before, {
        ...summaryOptions,
        x: x + 0.22,
        w: cardW - 1.32,
        color: toColor(deck.theme.colors.ink)
      });
      slide.addText(summaryParts.delta, {
        ...summaryOptions,
        x: x + cardW - 1.06,
        w: 0.88,
        color: kpiDeltaColor(deck, card),
        bold: true
      });
    } else {
      slide.addText(summaryParts.summary, {
        ...summaryOptions,
        x: x + 0.22,
        w: cardW - 0.35,
        color: toColor(deck.theme.colors.ink)
      });
    }
  });
}

function addActivationSnapshotCell(slide, deck, item, box) {
  const total = cleanInlineText(item.total || "-");
  const newCount = cleanInlineText(item.newCount || "-");
  const newPercent = cleanInlineText(item.newPercent || "-");
  const label = cleanInlineText(item.label || "Metric");
  const groupW = Math.min(3.34, box.w - 0.84);
  const groupX = box.x + (box.w - groupW) / 2;
  const detailGap = 0.42;
  const detailW = (groupW - detailGap) / 2;

  slide.addText(total, {
    x: groupX,
    y: box.y + 0.34,
    w: groupW,
    h: 0.62,
    fontFace: deck.theme.fonts.body,
    fontSize: total.length > 3 ? 31 : 38,
    color: toColor("#FFFFFF"),
    align: "center",
    margin: 0,
    fit: "shrink"
  });
  slide.addText(label, {
    x: groupX,
    y: box.y + 1.02,
    w: groupW,
    h: 0.25,
    fontFace: deck.theme.fonts.body,
    fontSize: 13.6,
    color: toColor("#BCEBFF"),
    align: "center",
    margin: 0,
    fit: "shrink"
  });
  [
    { label: "New", value: newCount, x: groupX },
    { label: "New %", value: newPercent, x: groupX + detailW + detailGap }
  ].forEach((detail) => {
    slide.addText(detail.label, {
      x: detail.x,
      y: box.y + 1.58,
      w: detailW,
      h: 0.16,
      fontFace: deck.theme.fonts.body,
      fontSize: 8.2,
      color: toColor(deck.theme.colors.paper),
      align: "center",
      margin: 0
    });
    slide.addText(detail.value, {
      x: detail.x,
      y: box.y + 1.82,
      w: detailW,
      h: 0.26,
      fontFace: deck.theme.fonts.body,
      fontSize: 14.2,
      color: toColor("#BCEBFF"),
      align: "center",
      margin: 0,
      fit: "shrink"
    });
  });
}

function renderProgramActivationSnapshotSlide(slide, deck, spec) {
  addBlueChrome(slide, deck);
  addSlideWatermark(slide, deck, true);

  slide.addText("Activation Status", {
    x: 0.72,
    y: 0.32,
    w: 5.4,
    h: 0.46,
    fontFace: deck.theme.fonts.body,
    fontSize: 28,
    color: toColor("#FFFFFF"),
    margin: 0,
    fit: "shrink"
  });

  const items = spec.activationSnapshot || [];
  const matrix = { x: 0.72, y: 1.02, w: 11.88, h: 5.58 };
  const cellW = matrix.w / 2;
  const cellH = matrix.h / 2;
  const centerX = matrix.x + cellW;
  const centerY = matrix.y + cellH;
  const ruleInsetX = 0.42;
  const ruleGapX = 0.48;
  const ruleInsetY = 0.40;
  const ruleGapY = 0.34;

  [
    { x: centerX, y: matrix.y + ruleInsetY, w: 0, h: cellH - ruleInsetY - ruleGapY },
    { x: centerX, y: centerY + ruleGapY, w: 0, h: cellH - ruleInsetY - ruleGapY },
    { x: matrix.x + ruleInsetX, y: centerY, w: cellW - ruleInsetX - ruleGapX, h: 0 },
    { x: centerX + ruleGapX, y: centerY, w: cellW - ruleInsetX - ruleGapX, h: 0 }
  ].forEach((line) => {
    slide.addShape("line", {
      ...line,
      line: { color: toColor("#FFFFFF"), pt: 1.0, transparency: 0 }
    });
  });

  items.slice(0, 4).forEach((item, index) => {
    addActivationSnapshotCell(slide, deck, item, {
      x: matrix.x + (index % 2) * cellW,
      y: matrix.y + Math.floor(index / 2) * cellH,
      w: cellW,
      h: cellH
    });
  });
}

function formatAxisNumber(value, locale = "en-GB") {
  return Math.round(Number(value) || 0).toLocaleString(locale);
}

function niceAxisMax(value) {
  const max = Math.max(1, Number(value) || 0);
  const magnitude = 10 ** Math.floor(Math.log10(max));
  const normalized = max / magnitude;
  const nice = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return nice * magnitude;
}

function weeklyPubCommSeriesColor(label, index, deck) {
  const normalized = cleanInlineText(label || "").toLowerCase().replace(/\s+/g, " ");
  const accent = deck.theme.colors.accent;
  const red = deck.theme.colors.accentAlt || "#EB5757";
  const colorsByLabel = [
    { pattern: /^(your site|primary publisher|primary site)$/, color: accent },
    { pattern: /^(publisher|competitor)\s*1$|^comp\.?\s*a$/, color: "#F28E2B" },
    { pattern: /^(publisher|competitor)\s*2$|^comp\.?\s*b$/, color: "#9AA3AD" },
    { pattern: /^(publisher|competitor)\s*3$|^comp\.?\s*c$/, color: "#F2C94C" },
    { pattern: /^(publisher|competitor)\s*4$|^comp\.?\s*d$/, color: red }
  ];
  const match = colorsByLabel.find((item) => item.pattern.test(normalized));
  if (match) return match.color;
  return [accent, "#F28E2B", "#9AA3AD", "#F2C94C", red][index % 5];
}

function addWeeklyComboChart(slide, deck, chart, box) {
  if (!chart || !Array.isArray(chart.categories) || !chart.categories.length || !Array.isArray(chart.series)) return;

  const locale = deck.metadata.locale || "en-GB";
  const values = chart.series
    .flatMap((series) => series.values || [])
    .filter((value) => value !== null && value !== undefined && Number.isFinite(Number(value)));
  const axisMax = niceAxisMax(values.length ? Math.max(...values, 0) : 0);
  const chartData = chart.series
    .slice(0, 5)
    .filter((series) => Array.isArray(series.values) && series.values.some((value) => value !== null && value !== undefined && Number.isFinite(Number(value))))
    .map((series) => ({
      name: cleanInlineText(series.label || "Series"),
      labels: chart.categories.map((category) => cleanInlineText(category)),
      values: (series.values || []).map((value) => (
        value === null || value === undefined || !Number.isFinite(Number(value)) ? null : Number(value)
      ))
    }));
  if (!chartData.length) return;
  const palette = chartData.map((series, index) => weeklyPubCommSeriesColor(series.name, index, deck));

  slide.addShape("rect", {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    line: { color: toColor(deck.theme.colors.paper), pt: 0 },
    fill: { color: toColor(deck.theme.colors.paper) }
  });

  slide.addShape("rect", {
    x: box.x + 0.54,
    y: box.y + 0.28,
    w: box.w - 0.66,
    h: box.h - 0.78,
    line: { color: toColor(deck.theme.colors.paper), pt: 0 },
    fill: { color: toColor(deck.theme.colors.paper) }
  });

  slide.addText(cleanInlineText(chart.title || "Publ comm by week"), {
    x: box.x,
    y: box.y,
    w: box.w,
    h: 0.2,
    align: "center",
    fontFace: deck.theme.fonts.body,
    fontSize: 8.5,
    color: toColor(deck.theme.colors.muted),
    margin: 0
  });

  slide.addChart("line", chartData, {
    x: box.x + 0.36,
    y: box.y + 0.28,
    w: box.w - 0.58,
    h: box.h - 0.36,
    chartColors: palette.map((color) => toColor(color)),
    chartArea: {
      fill: { color: toColor(deck.theme.colors.paper), transparency: 0 },
      border: { color: toColor(deck.theme.colors.paper), pt: 0 }
    },
    plotArea: {
      fill: { color: toColor(deck.theme.colors.paper), transparency: 0 },
      border: { color: toColor(deck.theme.colors.paper), pt: 0 }
    },
    displayBlanksAs: "span",
    lineCap: "round",
    lineDataSymbol: "none",
    lineSize: 2.0,
    showLegend: true,
    legendPos: "b",
    legendColor: toColor(deck.theme.colors.muted),
    legendFontFace: deck.theme.fonts.body,
    legendFontSize: 7,
    showTitle: false,
    showValue: false,
    showLabel: false,
    catAxisLabelColor: toColor(deck.theme.colors.muted),
    catAxisLabelFontFace: deck.theme.fonts.body,
    catAxisLabelFontSize: 6,
    catAxisLineColor: toColor("#D9DEE8"),
    catAxisLineSize: 0.45,
    catAxisMajorTickMark: "none",
    valAxisLabelColor: toColor(deck.theme.colors.muted),
    valAxisLabelFontFace: deck.theme.fonts.body,
    valAxisLabelFontSize: 6.2,
    valAxisLabelFormatCode: "#,##0",
    valAxisMinVal: 0,
    valAxisMaxVal: axisMax,
    valAxisMajorUnit: axisMax / 5,
    valAxisLineShow: false,
    valAxisMajorTickMark: "none",
    valGridLine: { color: toColor("#D9DEE8"), size: 0.45, style: "solid", cap: "flat" },
    lang: locale
  });
}

function addMoversBarChart(slide, deck, chart, box) {
  const rows = chart && Array.isArray(chart.rows) ? chart.rows.slice(0, 20) : [];
  const values = rows.map((row) => Math.abs(Number(row.value) || 0));
  const maxAbs = Math.max(...values, 1);
  const labelW = 3.75;
  const valueW = 0.92;
  const plot = {
    x: box.x + labelW,
    y: box.y + 0.32,
    w: box.w - labelW - valueW,
    h: box.h - 0.45
  };
  const baselineX = plot.x + plot.w / 2;
  const halfW = plot.w / 2 - 0.10;
  const rowH = plot.h / Math.max(1, rows.length);
  const barH = Math.min(0.13, Math.max(0.055, rowH * 0.48));

  slide.addText(cleanInlineText(chart.title || "Publisher commission movers"), {
    x: box.x,
    y: box.y,
    w: box.w,
    h: 0.2,
    align: "center",
    fontFace: deck.theme.fonts.body,
    fontSize: 8.8,
    color: toColor(deck.theme.colors.muted),
    margin: 0
  });

  for (let i = -2; i <= 2; i += 1) {
    const x = baselineX + (i / 2) * halfW;
    slide.addShape("line", {
      x,
      y: plot.y,
      w: 0,
      h: plot.h,
      line: {
        color: toColor(i === 0 ? "#7B8190" : "#E0E4EC"),
        pt: i === 0 ? 0.8 : 0.45,
        transparency: i === 0 ? 0 : 12
      }
    });
  }

  rows.forEach((row, index) => {
    const y = plot.y + index * rowH + (rowH - barH) / 2;
    const value = Number(row.value) || 0;
    const barW = Math.max(0.02, (Math.abs(value) / maxAbs) * halfW);
    const positive = value >= 0;
    const barX = positive ? baselineX : baselineX - barW;
    const displayText = cleanInlineText(row.display || String(value));
    const barColor = positive ? deck.theme.colors.accent : deck.theme.colors.accentAlt;

    slide.addText(cleanInlineText(row.label || "-"), {
      x: box.x,
      y: y - 0.035,
      w: labelW - 0.18,
      h: Math.max(0.12, rowH),
      align: "right",
      valign: "mid",
      fontFace: deck.theme.fonts.body,
      fontSize: rows.length > 16 ? 5.8 : 6.4,
      color: toColor(deck.theme.colors.muted),
      margin: 0
    });
    slide.addShape("rect", {
      x: barX,
      y,
      w: barW,
      h: barH,
      line: {
        color: toColor(barColor),
        pt: 0
      },
      fill: {
        color: toColor(barColor),
        transparency: positive ? 4 : 0
      }
    });
    slide.addText(displayText, {
      x: positive ? Math.min(plot.x + plot.w - valueW, baselineX + barW + 0.05) : Math.max(plot.x, baselineX - barW - valueW - 0.05),
      y: y - 0.015,
      w: valueW,
      h: barH + 0.04,
      align: positive ? "left" : "right",
      valign: "mid",
      fontFace: deck.theme.fonts.body,
      fontSize: rows.length > 16 ? 5.7 : 6.2,
      color: toColor(deck.theme.colors.ink),
      margin: 0
    });
  });
}

function addCompetitorShareBarChart(slide, deck, chart, box) {
  const categories = chart && Array.isArray(chart.categories) ? chart.categories.slice(0, 5) : [];
  const series = chart && Array.isArray(chart.series) ? chart.series.slice(0, 2) : [];
  const activeSeries = series.filter((item) => Array.isArray(item.values) && item.values.some((value) => Number(value) > 0));
  const labels = categories.length ? categories : ["Your Site", "Comp. A", "Comp. B", "Comp. C", "Comp. D"];
  const plot = {
    x: box.x + 0.54,
    y: box.y + 0.58,
    w: box.w - 0.78,
    h: box.h - 1.20
  };
  const palette = ["#7CC8DE", deck.theme.colors.accent];
  const groupW = plot.w / Math.max(1, labels.length);
  const maxValue = Math.max(100, ...activeSeries.flatMap((item) => item.values.map((value) => Number(value) || 0)));
  const axisMax = Math.min(100, Math.ceil(maxValue / 10) * 10);
  const zeroY = plot.y + plot.h;

  slide.addText(cleanInlineText(chart?.title || "Share of publisher commission within competitor group"), {
    x: box.x,
    y: box.y,
    w: box.w,
    h: 0.28,
    align: "center",
    fontFace: deck.theme.fonts.body,
    fontSize: 12,
    color: toColor(deck.theme.colors.ink),
    margin: 0
  });

  [0, 25, 50, 75, 100].forEach((tick) => {
    const y = zeroY - (tick / axisMax) * plot.h;
    slide.addShape("line", {
      x: plot.x,
      y,
      w: plot.w,
      h: 0,
      line: { color: toColor(tick === 0 ? "#AEB6C4" : "#DDE3ED"), pt: tick === 0 ? 0.75 : 0.45, transparency: tick === 0 ? 0 : 18 }
    });
  });

  if (!activeSeries.length) {
    slide.addText("Competitor share chart requires publisher commission values for the primary site and comparison group.", {
      x: box.x + 0.6,
      y: box.y + 2.5,
      w: box.w - 1.2,
      h: 0.42,
      align: "center",
      fontFace: deck.theme.fonts.body,
      fontSize: 12,
      color: toColor(deck.theme.colors.muted),
      margin: 0
    });
    return;
  }

  labels.forEach((label, categoryIndex) => {
    const seriesCount = activeSeries.length;
    const barW = Math.min(0.48, groupW / (seriesCount + 1.35));
    const totalBarsW = (barW * seriesCount) + (0.10 * Math.max(0, seriesCount - 1));
    const startX = plot.x + categoryIndex * groupW + (groupW - totalBarsW) / 2;
    activeSeries.forEach((item, seriesIndex) => {
      const value = Math.max(0, Number(item.values[categoryIndex]) || 0);
      const barH = (value / axisMax) * plot.h;
      const x = startX + seriesIndex * (barW + 0.10);
      const y = zeroY - barH;
      slide.addShape("rect", {
        x,
        y,
        w: barW,
        h: Math.max(0.02, barH),
        line: { color: toColor(palette[seriesIndex]), pt: 0 },
        fill: { color: toColor(palette[seriesIndex]), transparency: seriesIndex === 0 ? 8 : 0 }
      });
      slide.addText(`${Math.round(value)}%`, {
        x: x - 0.12,
        y: Math.max(plot.y - 0.02, y - 0.28),
        w: barW + 0.24,
        h: 0.18,
        align: "center",
        fontFace: deck.theme.fonts.body,
        fontSize: 8.5,
        color: toColor(deck.theme.colors.ink),
        margin: 0
      });
    });
    slide.addText(cleanInlineText(label), {
      x: plot.x + categoryIndex * groupW + 0.02,
      y: zeroY + 0.16,
      w: groupW - 0.04,
      h: 0.20,
      align: "center",
      fontFace: deck.theme.fonts.body,
      fontSize: 8.6,
      color: toColor(deck.theme.colors.muted),
      margin: 0,
      fit: "shrink"
    });
  });

  const legendW = 1.78;
  const legendStart = plot.x + Math.max(0, (plot.w - activeSeries.length * legendW) / 2);
  activeSeries.forEach((item, index) => {
    const x = legendStart + index * legendW;
    const y = box.y + box.h - 0.26;
    slide.addShape("rect", {
      x,
      y: y + 0.045,
      w: 0.15,
      h: 0.07,
      line: { color: toColor(palette[index]), pt: 0 },
      fill: { color: toColor(palette[index]) }
    });
    slide.addText(cleanInlineText(item.label), {
      x: x + 0.20,
      y,
      w: legendW - 0.22,
      h: 0.18,
      fontFace: deck.theme.fonts.body,
      fontSize: 7.6,
      color: toColor(deck.theme.colors.muted),
      margin: 0
    });
  });
}

function isDeltaColumn(header) {
  const lower = cleanInlineText(header).toLowerCase();
  return lower.includes("change")
    || lower.includes("variance")
    || lower.includes("yoy")
    || lower.includes("trend")
    || lower.includes("difference")
    || lower.includes("delta")
    || lower.includes("movement");
}

function cellTextColor(table, column, value, deck) {
  const text = cleanInlineText(value);
  const movementColumn = isDeltaColumn(column);
  const compact = text
    .replace(MOJIBAKE_DIRECTION_MARK_REGEX, "")
    .replace(DIRECTION_MARK_REGEX, "")
    .replace(/\s+/g, " ")
    .trim();
  const hasSignedMovement = /^[+]/.test(compact) || /^-/.test(compact) || /\b(?:GBP|EUR|USD|AUD|PLN)\s*[+-]/i.test(compact);
  const hasSignedPercent = /^[+-]\d+(?:[.,]\d+)?%$/.test(compact);

  if (!movementColumn && !hasSignedMovement && !hasSignedPercent) return toColor(deck.theme.colors.ink);
  if (/^\+/.test(compact) || /\b(?:GBP|EUR|USD|AUD|PLN)\s*\+/i.test(compact)) return toColor(deck.theme.colors.success);
  if (/^-/.test(compact) || /\b(?:GBP|EUR|USD|AUD|PLN)\s*-/i.test(compact)) return toColor(deck.theme.colors.accentAlt);
  if (/^(0(?:[.,]0+)?%?|n\/a|na|-)?$/i.test(compact)) return toColor(deck.theme.colors.warning);
  return toColor(deck.theme.colors.ink);
}

function addTable(slide, deck, table, box, mode = "light") {
  const innerX = box.x + 0.12;
  const innerY = box.y + 0.10;
  const innerW = box.w - 0.24;
  const innerH = box.h - 0.20;
  const columnAlignments = table.columns.map(() => "center");

  const header = table.columns.map((column, columnIndex) => ({
    text: column,
    options: {
      bold: true,
      fontFace: deck.theme.fonts.body,
      fontSize: table.dense ? 9 : 10.5,
      color: toColor(deck.theme.colors.ink),
      fill: { color: toColor("#F1F3F7") },
      margin: 0.045,
      align: columnAlignments[columnIndex] || "left",
      valign: "mid"
    }
  }));
  const bodyRows = table.rows.map((row, index) => {
    const firstCell = cleanInlineText(row[0] || "");
    const isSectionRow = /^Top\s+\d+\s+(Up|Down)$/i.test(firstCell);
    const defaultRowFill = isSectionRow
      ? "#E5E8EF"
      : index % 2 === 0
        ? "#F4F5F7"
        : "#ECEDEF";

    return row.map((value, cellIndex) => ({
      text: value,
      options: {
        bold: isSectionRow,
        fontFace: deck.theme.fonts.body,
        fontSize: table.dense ? 9 : (isSectionRow ? 10.5 : 10),
        color: isSectionRow
          ? toColor(deck.theme.colors.ink)
          : cellTextColor(table, table.columns[cellIndex] || "", value, deck),
        fill: {
          color: toColor(cellIndex === table.primaryHighlightColumn ? "#C8F7D2" : defaultRowFill)
        },
        margin: 0.045,
        align: columnAlignments[cellIndex] || "left",
        valign: "mid"
      }
    }));
  });

  let headerH = table.dense ? 0.34 : 0.58;
  let bodyH = table.dense ? 0.30 : 0.58;
  const bodyCount = bodyRows.length;
  const headerWeight = table.dense ? 1.08 : 1.16;
  const fittedBody = innerH / Math.max(1, (bodyCount + headerWeight));
  if (Number.isFinite(fittedBody) && fittedBody > 0) {
    bodyH = Number(Math.min(table.dense ? 0.46 : 0.72, Math.max(table.dense ? 0.22 : 0.30, fittedBody)).toFixed(3));
    headerH = Number((bodyH * headerWeight).toFixed(3));
  }
  const desiredTableH = headerH + (bodyCount * bodyH);
  if (desiredTableH > innerH && desiredTableH > 0) {
    const scale = innerH / desiredTableH;
    headerH = Math.max(0.24, Number((headerH * scale).toFixed(3)));
    bodyH = Math.max(0.22, Number((bodyH * scale).toFixed(3)));
  }

  const effectiveTableH = Number((headerH + (bodyCount * bodyH)).toFixed(3));
  const containerH = Math.min(box.h, Number((effectiveTableH + 0.16).toFixed(3)));
  const rowHeights = [headerH, ...Array.from({ length: bodyCount }, () => bodyH)];
  let colW;
  if (Array.isArray(table.colW) && table.colW.length === table.columns.length) {
    const numeric = table.colW.map((w) => Number(w)).filter((w) => Number.isFinite(w) && w > 0);
    if (numeric.length === table.columns.length) {
      const total = numeric.reduce((sum, w) => sum + w, 0);
      if (total > 0) {
        colW = numeric.map((w) => Number(((w / total) * innerW).toFixed(3)));
      }
    }
  }

  slide.addShape("rect", {
    x: box.x,
    y: box.y,
    w: box.w,
    h: containerH,
    line: { color: toColor("#D6DAE3"), pt: 0.7 },
    fill: { color: toColor(deck.theme.colors.paper) }
  });

  slide.addTable([header, ...bodyRows], {
    x: innerX,
    y: innerY,
    w: innerW,
    h: effectiveTableH,
    colW,
    rowH: rowHeights,
    border: { type: "solid", color: toColor("#E3E6EC"), pt: 0.3 },
    margin: 0.02,
    autoFit: false
  });

  return { containerH, tableH: effectiveTableH };
}

function splitRiskTileText(value) {
  const text = cleanInlineText(value || "");
  const actionMatch = text.match(/\b(?:Publisher action|Action):\s*(.+)$/i);
  const action = actionMatch ? cleanInlineText(actionMatch[1]) : "";
  const analysisText = actionMatch ? cleanInlineText(text.slice(0, actionMatch.index)) : text;
  const compactAnalysis = compactRiskRationale(analysisText, 210);
  const colonIndex = compactAnalysis.indexOf(":");
  const periodIndex = compactAnalysis.indexOf(".");
  const validColon = colonIndex > 4 && colonIndex < 80 ? colonIndex : -1;
  const validPeriod = periodIndex > 4 && periodIndex < 80 ? periodIndex : -1;
  const splitIndex = validColon >= 0 && validPeriod >= 0
    ? Math.min(validColon, validPeriod)
    : Math.max(validColon, validPeriod);
  if (splitIndex < 0) {
    return { heading: "Risk", detail: compactAnalysis, action };
  }
  return {
    heading: cleanInlineText(compactAnalysis.slice(0, splitIndex)),
    detail: cleanInlineText(compactAnalysis.slice(splitIndex + 1)),
    action
  };
}

function renderRisksDependenciesTiles(slide, deck, spec) {
  const rows = spec.tables?.[0]?.rows || [];
  const risks = rows
    .map((row) => cleanInlineText(Array.isArray(row) ? row[0] : row))
    .filter(Boolean)
    .slice(0, 3);
  const items = risks.length ? risks : [
    "Program concentration risk: diversify program mix and reduce reliance on the top contributors."
  ];
  const tiles = [
    { x: 0.78, y: 2.16, w: 3.76, h: 3.82 },
    { x: 4.78, y: 2.16, w: 3.76, h: 3.82 },
    { x: 8.78, y: 2.16, w: 3.76, h: 3.82 }
  ];

  tiles.forEach((tile, index) => {
    const item = items[index] || "";
    const parsed = splitRiskTileText(item);
    slide.addShape("roundRect", {
      x: tile.x,
      y: tile.y,
      w: tile.w,
      h: tile.h,
      radius: 0.08,
      line: { color: toColor("#D7E4FF"), pt: 0 },
      fill: { color: toColor("#D7E4FF") }
    });
    if (!item) return;
    slide.addText(parsed.heading, {
      x: tile.x + 0.32,
      y: tile.y + 0.34,
      w: tile.w - 0.64,
      h: 0.56,
      align: "center",
      fontFace: deck.theme.fonts.heading,
      fontSize: 11.4,
      bold: true,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      breakLine: true,
      fit: "shrink"
    });
    slide.addText("Analysis", {
      x: tile.x + 0.32,
      y: tile.y + 1.12,
      w: tile.w - 0.64,
      h: 0.22,
      fontFace: deck.theme.fonts.heading,
      fontSize: 8.3,
      bold: true,
      color: toColor(deck.theme.colors.accent),
      margin: 0
    });
    slide.addText(parsed.detail || item, {
      x: tile.x + 0.32,
      y: tile.y + 1.40,
      w: tile.w - 0.68,
      h: 0.98,
      fontFace: deck.theme.fonts.body,
      fontSize: 8.1,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      breakLine: true,
      fit: "shrink"
    });
    slide.addText("Action", {
      x: tile.x + 0.32,
      y: tile.y + 2.58,
      w: tile.w - 0.64,
      h: 0.22,
      fontFace: deck.theme.fonts.heading,
      fontSize: 8.3,
      bold: true,
      color: toColor(deck.theme.colors.accent),
      margin: 0
    });
    slide.addText(parsed.action || "Run one controlled publisher test with a target metric and success threshold for the next review.", {
      x: tile.x + 0.32,
      y: tile.y + 2.86,
      w: tile.w - 0.68,
      h: 0.72,
      fontFace: deck.theme.fonts.body,
      fontSize: 7.7,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      breakLine: true,
      fit: "shrink"
    });
  });
}

function truncateDisplayText(value, maxChars) {
  const text = cleanInlineText(value || "");
  if (text.length <= maxChars) return text;
  const clipped = text.slice(0, Math.max(0, maxChars - 3));
  const lastSpace = clipped.lastIndexOf(" ");
  return `${(lastSpace > 12 ? clipped.slice(0, lastSpace) : clipped).trimEnd()}...`;
}

function programConnectionNameFontSize(name, baseFontSize, nameW, rowH) {
  const text = cleanInlineText(name || "");
  if (!text) return baseFontSize;

  const oneLineCapacity = Math.max(10, Math.floor(nameW * 10.5));
  const twoLineCapacity = Math.max(oneLineCapacity, Math.floor(oneLineCapacity * 1.8));
  if (text.length <= oneLineCapacity) return baseFontSize;
  if (text.length <= twoLineCapacity && rowH >= 0.34) return Math.max(4.7, baseFontSize - 0.4);
  if (rowH >= 0.34) return Math.max(4.3, baseFontSize - 0.9);
  return Math.max(4.1, baseFontSize - 1.0);
}

function renderProgramConnectionStatusSlide(slide, deck, spec) {
  addLightChrome(slide, deck);
  addTitle(slide, deck, spec, deck.theme.colors.ink, deck.theme.colors.accent, false);

  const data = spec.programConnectionStatus || {};
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const totalRows = Number(data.totalRows || rows.length) || rows.length;
  const pageIndex = Number(data.pageIndex || 0);
  const pageCount = Math.max(1, Number(data.pageCount || 1) || 1);
  const key = Array.isArray(data.key) ? data.key : CONNECTION_STATUS_META;
  const counts = data.counts || {};
  const cutoffText = cleanInlineText(data.cutoffDate || "");
  const cutoffDate = /^\d{4}-\d{2}-\d{2}$/.test(cutoffText)
    ? new Date(`${cutoffText}T00:00:00Z`)
    : null;
  const cutoffLabel = cutoffDate && !Number.isNaN(cutoffDate.getTime())
    ? formatLongDate(cutoffDate, deck.metadata.locale || "en-GB")
    : cutoffText;
  const countText = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .map(([label, count]) => `${label}: ${count}`)
    .join(" | ");
  const pageText = pageCount > 1 ? ` | Page ${pageIndex + 1}/${pageCount}` : "";
  const cutoffSummary = cutoffLabel ? ` | Created from ${cutoffLabel}` : "";

  slide.addText(`${totalRows} programs reviewed${pageText}${countText ? ` | ${countText}` : ""}${cutoffSummary}`, {
    x: 0.72,
    y: 1.28,
    w: 10.9,
    h: 0.28,
    fontFace: deck.theme.fonts.body,
    fontSize: 8.8,
    color: toColor(deck.theme.colors.muted),
    margin: 0,
    fit: "shrink"
  });

  const keyX = 12.0;
  const keyY = 1.72;
  const keyW = 1.05;
  slide.addShape("roundRect", {
    x: keyX,
    y: keyY,
    w: keyW,
    h: 2.96,
    radius: 0.04,
    rectRadius: 0.04,
    line: { color: toColor("#D6DAE3"), pt: 0.45 },
    fill: { color: toColor("#FFFFFF"), transparency: 5 }
  });
  slide.addText("Key", {
    x: keyX + 0.10,
    y: keyY + 0.12,
    w: keyW - 0.20,
    h: 0.18,
    fontFace: deck.theme.fonts.heading,
    fontSize: 8.2,
    bold: true,
    color: toColor(deck.theme.colors.ink),
    margin: 0
  });
  key.forEach((item, index) => {
    const y = keyY + 0.43 + index * 0.30;
    slide.addShape("ellipse", {
      x: keyX + 0.10,
      y: y + 0.014,
      w: 0.075,
      h: 0.075,
      line: { color: toColor(item.color), pt: 0 },
      fill: { color: toColor(item.color) }
    });
    slide.addText(`${item.id} ${item.shortLabel || item.label}`, {
      x: keyX + 0.22,
      y: y - 0.010,
      w: keyW - 0.28,
      h: 0.14,
      fontFace: deck.theme.fonts.body,
      fontSize: 5.8,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      fit: "shrink"
    });
  });

  const grid = { x: 0.42, y: 1.78, w: 11.40, h: 5.34 };
  const columnCount = programConnectionStatusColumnCount(totalRows);
  const gap = columnCount > 5 ? 0.055 : 0.075;
  const columnW = (grid.w - ((columnCount - 1) * gap)) / columnCount;
  const rowsPerColumn = Math.max(12, Math.ceil(rows.length / columnCount));
  const rowH = grid.h / Math.max(1, rowsPerColumn);
  const compact = rowsPerColumn > 18 || columnCount > 5;
  const fontSize = compact ? 5.0 : (rowH > 0.36 ? 6.4 : 5.8);
  const headerFontSize = compact ? 5.0 : 5.7;
  const textH = Math.max(0.16, rowH - 0.045);

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    const x = grid.x + columnIndex * (columnW + gap);
    slide.addShape("rect", {
      x,
      y: grid.y - 0.24,
      w: columnW,
      h: 0.20,
      line: { color: toColor("#D6DAE3"), pt: 0.25 },
      fill: { color: toColor("#E5E8EF") }
    });
    slide.addText("ID / Program / Status", {
      x: x + 0.06,
      y: grid.y - 0.20,
      w: columnW - 0.12,
      h: 0.12,
      fontFace: deck.theme.fonts.heading,
      fontSize: headerFontSize,
      bold: true,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      fit: "shrink"
    });

    const start = columnIndex * rowsPerColumn;
    const columnRows = rows.slice(start, start + rowsPerColumn);
    columnRows.forEach((row, rowIndex) => {
      const y = grid.y + rowIndex * rowH;
      const fill = rowIndex % 2 === 0 ? "#F7F8FA" : "#ECEFF4";
      const contentY = y + ((rowH - textH) / 2);
      slide.addShape("rect", {
        x,
        y,
        w: columnW,
        h: Math.max(0.10, rowH - 0.014),
        line: { color: toColor("#E3E6EC"), pt: 0.2 },
        fill: { color: toColor(fill) }
      });
      slide.addShape("ellipse", {
        x: x + 0.052,
        y: y + Math.max(0.040, (rowH - 0.082) / 2),
        w: 0.072,
        h: 0.072,
        line: { color: toColor(row.color || "#8A94A6"), pt: 0 },
        fill: { color: toColor(row.color || "#8A94A6") }
      });
      slide.addText(truncateDisplayText(row.programId || "-", 7), {
        x: x + 0.16,
        y: contentY,
        w: compact ? 0.36 : 0.44,
        h: textH,
        fontFace: deck.theme.fonts.mono,
        fontSize,
        bold: true,
        color: toColor(deck.theme.colors.ink),
        margin: 0,
        fit: "shrink"
      });
      const nameX = x + (compact ? 0.56 : 0.66);
      const statusW = compact ? 0.42 : 0.50;
      const nameW = columnW - (nameX - x) - statusW - 0.09;
      slide.addText(cleanInlineText(row.programName || "-"), {
        x: nameX,
        y: contentY,
        w: nameW,
        h: textH,
        fontFace: deck.theme.fonts.body,
        fontSize: programConnectionNameFontSize(row.programName, fontSize, nameW, rowH),
        color: toColor(deck.theme.colors.ink),
        margin: 0,
        fit: "shrink"
      });
      slide.addText(row.shortStatus || row.status || "-", {
        x: x + columnW - statusW - 0.05,
        y: contentY,
        w: statusW,
        h: textH,
        align: "right",
        fontFace: deck.theme.fonts.body,
        fontSize: Math.max(5.1, fontSize - 0.2),
        color: toColor(row.color || deck.theme.colors.muted),
        margin: 0,
        fit: "shrink"
      });
    });
  }
}

function renderGapAnalysisImpactSlide(slide, deck, spec) {
  addLightChrome(slide, deck);
  const impact = spec.gapImpact || {};
  const kpis = Array.isArray(spec.kpis) ? spec.kpis : [];
  const total = kpis[0] || { value: "-", label: "Competitor pub comm opportunity" };
  const programs = kpis[1] || { value: "-", label: "Gap programs" };
  const activation = kpis[2] || { value: "-", label: "Activation opportunities" };

  addCyanFifthElementWireframe(slide, {
    x: 0.30,
    y: 1.02,
    w: 4.65,
    h: 4.65,
    transparency: 22
  });

  slide.addText(cleanInlineText(spec.title || "Growth opportunity in the competitor gap"), {
    x: 5.55,
    y: 0.78,
    w: 6.85,
    h: 0.72,
    fontFace: deck.theme.fonts.heading,
    fontSize: 22,
    color: toColor(deck.theme.colors.ink),
    breakLine: true,
    fit: "shrink",
    margin: 0
  });

  [
    { x: 5.60, metric: programs, detail: cleanInlineText(programs.summary || "Programs requiring action") },
    { x: 9.25, metric: activation, detail: cleanInlineText(activation.summary || "Already accepted but inactive") }
  ].forEach(({ x, metric, detail }) => {
    slide.addText(cleanInlineText(metric.value || "-"), {
      x,
      y: 1.78,
      w: 2.7,
      h: 0.52,
      fontFace: deck.theme.fonts.heading,
      fontSize: 30,
      bold: true,
      color: toColor(deck.theme.colors.ink),
      align: "center",
      margin: 0
    });
    slide.addText(cleanInlineText(metric.label || "-"), {
      x: x - 0.2,
      y: 2.52,
      w: 3.1,
      h: 0.25,
      fontFace: deck.theme.fonts.heading,
      fontSize: 12.5,
      bold: true,
      color: toColor(deck.theme.colors.muted),
      align: "center",
      fit: "shrink",
      margin: 0
    });
    slide.addText(detail, {
      x: x - 0.25,
      y: 2.88,
      w: 3.2,
      h: 0.28,
      fontFace: deck.theme.fonts.body,
      fontSize: 9.8,
      color: toColor(deck.theme.colors.muted),
      align: "center",
      fit: "shrink",
      margin: 0
    });
  });

  slide.addText(cleanInlineText(total.value || "-"), {
    x: 6.45,
    y: 3.58,
    w: 4.4,
    h: 0.68,
    fontFace: deck.theme.fonts.heading,
    fontSize: 36,
    bold: true,
    color: toColor(deck.theme.colors.ink),
    align: "center",
    margin: 0
  });
  slide.addText(cleanInlineText(total.label || "Competitor pub comm opportunity"), {
    x: 6.20,
    y: 4.42,
    w: 4.9,
    h: 0.34,
    fontFace: deck.theme.fonts.heading,
    fontSize: 13.5,
    bold: true,
    color: toColor(deck.theme.colors.muted),
    align: "center",
    fit: "shrink",
    margin: 0
  });
  slide.addText(cleanInlineText(total.summary || "Competitors are earning where the primary publisher is not."), {
    x: 6.32,
    y: 4.86,
    w: 4.68,
    h: 0.42,
    fontFace: deck.theme.fonts.body,
    fontSize: 10.4,
    color: toColor(deck.theme.colors.muted),
    align: "center",
    fit: "shrink",
    margin: 0
  });

  const narrative = `Prioritise accepted and click-led programs first, then build an application pipeline for no-connection gaps. The largest visible gap is ${cleanInlineText(impact.leadingProgram || "the leading program")} (${cleanInlineText(impact.leadingProgramValue || "-")}).`;
  slide.addText(narrative, {
    x: 5.55,
    y: 5.58,
    w: 6.75,
    h: 0.7,
    fontFace: deck.theme.fonts.body,
    fontSize: 11.4,
    color: toColor(deck.theme.colors.muted),
    breakLine: true,
    fit: "shrink",
    margin: 0
  });

}

function gapRegisterStatusColor(status, deck) {
  const lower = cleanInlineText(status || "").toLowerCase();
  if (/^pub comm$/.test(lower)) return deck.theme.colors.accent;
  if (/accepted/.test(lower)) return deck.theme.colors.success;
  if (/click/.test(lower)) return "#7CC8DE";
  if (/no connection|not connected/.test(lower)) return "#8A94A6";
  if (/denied/.test(lower)) return deck.theme.colors.accentAlt;
  if (/ended/.test(lower)) return "#5B6372";
  if (/hold|under consideration|consideration/.test(lower)) return deck.theme.colors.warning;
  return "#8A94A6";
}

function renderGapAnalysisRegisterSlide(slide, deck, spec) {
  addLightChrome(slide, deck);
  addTitle(slide, deck, spec, deck.theme.colors.ink, deck.theme.colors.accent, false);

  const register = spec.gapRegister || {};
  const rows = Array.isArray(register.rows) ? register.rows : [];
  const totalRows = Number(register.totalRows) || rows.length;
  const pageText = register.pageCount > 1 ? ` | Page ${register.pageIndex + 1}/${register.pageCount}` : "";
  slide.addText(`${totalRows} programs reviewed${pageText} | Pub Comm - Specified Sites is competitor-generated publisher commission where the primary publisher is not earning.`, {
    x: 0.72,
    y: 1.28,
    w: 11.6,
    h: 0.28,
    fontFace: deck.theme.fonts.body,
    fontSize: 8.8,
    color: toColor(deck.theme.colors.muted),
    margin: 0,
    fit: "shrink"
  });

  const grid = { x: 0.42, y: 1.78, w: 12.48, h: 5.36 };
  const columnCount = 3;
  const gap = 0.10;
  const columnW = (grid.w - (gap * (columnCount - 1))) / columnCount;
  const rowsPerColumn = Math.max(1, register.rowsPerColumn || Math.ceil(rows.length / columnCount));
  const headerH = 0.23;
  const rowH = Math.min(0.225, (grid.h - headerH - 0.04) / rowsPerColumn);
  const fontSize = rowH < 0.21 ? 4.8 : 5.25;

  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    const x = grid.x + columnIndex * (columnW + gap);
    slide.addShape("rect", {
      x,
      y: grid.y,
      w: columnW,
      h: headerH,
      line: { color: toColor("#D6DAE3"), pt: 0.25 },
      fill: { color: toColor("#E5E8EF") }
    });
    slide.addText("Program / ID / Status / Pub Comm", {
      x: x + 0.06,
      y: grid.y + 0.055,
      w: columnW - 0.12,
      h: 0.12,
      fontFace: deck.theme.fonts.heading,
      fontSize: 5.8,
      bold: true,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      fit: "shrink"
    });

    const columnRows = rows.slice(columnIndex * rowsPerColumn, (columnIndex + 1) * rowsPerColumn);
    columnRows.forEach((row, rowIndex) => {
      const y = grid.y + headerH + 0.035 + rowIndex * rowH;
      const fill = rowIndex % 2 === 0 ? "#F7F8FA" : "#ECEFF4";
      const statusColor = gapRegisterStatusColor(row.primaryStatus, deck);
      slide.addShape("rect", {
        x,
        y,
        w: columnW,
        h: Math.max(0.10, rowH - 0.012),
        line: { color: toColor("#E3E6EC"), pt: 0.16 },
        fill: { color: toColor(fill) }
      });
      slide.addShape("ellipse", {
        x: x + 0.045,
        y: y + Math.max(0.035, (rowH - 0.065) / 2),
        w: 0.060,
        h: 0.060,
        line: { color: toColor(statusColor), pt: 0 },
        fill: { color: toColor(statusColor) }
      });
      slide.addText(cleanInlineText(row.programId || "-"), {
        x: x + 0.13,
        y: y + 0.035,
        w: 0.42,
        h: 0.12,
        fontFace: deck.theme.fonts.mono,
        fontSize,
        bold: true,
        color: toColor(deck.theme.colors.ink),
        margin: 0,
        fit: "shrink"
      });
      slide.addText(compactLabel(row.programName || "-", 28), {
        x: x + 0.58,
        y: y + 0.035,
        w: columnW - 2.18,
        h: 0.12,
        fontFace: deck.theme.fonts.body,
        fontSize,
        color: toColor(deck.theme.colors.ink),
        margin: 0,
        fit: "shrink"
      });
      slide.addText(compactLabel(row.primaryStatus || "-", 13), {
        x: x + columnW - 1.50,
        y: y + 0.035,
        w: 0.80,
        h: 0.12,
        fontFace: deck.theme.fonts.body,
        fontSize: Math.max(4.6, fontSize - 0.1),
        color: toColor(statusColor),
        margin: 0,
        fit: "shrink"
      });
      slide.addText(cleanInlineText(row.registerOpportunityDisplay || row.competitorPubCommDisplay || "-"), {
        x: x + columnW - 0.68,
        y: y + 0.035,
        w: 0.62,
        h: 0.12,
        align: "right",
        fontFace: deck.theme.fonts.body,
        fontSize: Math.max(4.6, fontSize - 0.05),
        color: toColor(deck.theme.colors.ink),
        margin: 0,
        fit: "shrink"
      });
    });
  }
}

function renderGapAnalysisTopProgramsSlide(slide, deck, spec) {
  addLightChrome(slide, deck);
  addTitle(slide, deck, spec, deck.theme.colors.ink, deck.theme.colors.accent, false);

  const payload = spec.gapTopPrograms || {};
  const rows = Array.isArray(payload.rows) ? payload.rows.slice(0, 10) : [];
  const totalRows = Number(payload.totalRows) || rows.length;
  const reportRows = Number(payload.reportRows) || Math.max(0, totalRows - rows.length);
  const maxValue = Math.max(1, ...rows.map((row) => Number(row.competitorPubComm) || 0));

  slide.addText(`${rows.length} highest-value programs shown here. ${reportRows} remaining gap programs are available in the Excel report.`, {
    x: 0.72,
    y: 1.28,
    w: 11.8,
    h: 0.28,
    fontFace: deck.theme.fonts.body,
    fontSize: 9.4,
    color: toColor(deck.theme.colors.muted),
    margin: 0,
    fit: "shrink"
  });

  const grid = { x: 0.58, y: 1.72, w: 12.18, h: 5.38 };
  const headerH = 0.34;
  const rowH = 0.46;
  const cols = [
    { key: "rank", label: "#", x: grid.x, w: 0.42, align: "center" },
    { key: "program", label: "Program / ID", x: grid.x + 0.48, w: 3.78, align: "left" },
    { key: "status", label: "Primary status", x: grid.x + 4.36, w: 1.52, align: "left" },
    { key: "signal", label: "Competitor signal", x: grid.x + 5.98, w: 1.82, align: "left" },
    { key: "bar", label: "Pub Comm - Specified Sites", x: grid.x + 7.88, w: 3.02, align: "left" },
    { key: "value", label: "Value", x: grid.x + 11.00, w: 1.08, align: "right" }
  ];

  slide.addShape("rect", {
    x: grid.x,
    y: grid.y,
    w: grid.w,
    h: headerH,
    line: { color: toColor("#D6DAE3"), pt: 0.25 },
    fill: { color: toColor("#E5E8EF") }
  });
  cols.forEach((col) => {
    slide.addText(col.label, {
      x: col.x + 0.04,
      y: grid.y + 0.10,
      w: col.w - 0.08,
      h: 0.14,
      fontFace: deck.theme.fonts.heading,
      fontSize: 7.6,
      bold: true,
      align: col.align,
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      fit: "shrink"
    });
  });

  rows.forEach((row, index) => {
    const y = grid.y + headerH + index * rowH;
    const fill = index % 2 === 0 ? "#F7F8FA" : "#ECEFF4";
    const statusColor = gapRegisterStatusColor(row.primaryStatus, deck);
    const barW = Math.max(0.08, ((Number(row.competitorPubComm) || 0) / maxValue) * 2.35);
    slide.addShape("rect", {
      x: grid.x,
      y,
      w: grid.w,
      h: rowH - 0.018,
      line: { color: toColor("#E3E6EC"), pt: 0.18 },
      fill: { color: toColor(fill) }
    });
    slide.addText(String(index + 1), {
      x: cols[0].x,
      y: y + 0.15,
      w: cols[0].w,
      h: 0.16,
      fontFace: deck.theme.fonts.heading,
      fontSize: 8,
      bold: true,
      align: "center",
      color: toColor(deck.theme.colors.ink),
      margin: 0
    });
    slide.addText(`${compactLabel(row.programName || "-", 38)}\n${cleanInlineText(row.programId || "-")}`, {
      x: cols[1].x + 0.04,
      y: y + 0.07,
      w: cols[1].w - 0.08,
      h: 0.30,
      fontFace: deck.theme.fonts.body,
      fontSize: 7.2,
      color: toColor(deck.theme.colors.ink),
      breakLine: true,
      fit: "shrink",
      margin: 0
    });
    slide.addShape("ellipse", {
      x: cols[2].x + 0.02,
      y: y + 0.18,
      w: 0.08,
      h: 0.08,
      line: { color: toColor(statusColor), pt: 0 },
      fill: { color: toColor(statusColor) }
    });
    slide.addText(compactLabel(row.primaryStatus || "-", 18), {
      x: cols[2].x + 0.14,
      y: y + 0.15,
      w: cols[2].w - 0.18,
      h: 0.16,
      fontFace: deck.theme.fonts.body,
      fontSize: 7.2,
      color: toColor(statusColor),
      margin: 0,
      fit: "shrink"
    });
    slide.addText(compactLabel(row.competitorSignal || "-", 24), {
      x: cols[3].x + 0.04,
      y: y + 0.15,
      w: cols[3].w - 0.08,
      h: 0.16,
      fontFace: deck.theme.fonts.body,
      fontSize: 7.1,
      color: toColor(deck.theme.colors.muted),
      margin: 0,
      fit: "shrink"
    });
    slide.addShape("rect", {
      x: cols[4].x + 0.04,
      y: y + 0.18,
      w: cols[4].w - 0.20,
      h: 0.08,
      line: { color: toColor("#D9E2F7"), pt: 0 },
      fill: { color: toColor("#D9E2F7") }
    });
    slide.addShape("rect", {
      x: cols[4].x + 0.04,
      y: y + 0.18,
      w: barW,
      h: 0.08,
      line: { color: toColor(deck.theme.colors.accent), pt: 0 },
      fill: { color: toColor(deck.theme.colors.accent) }
    });
    slide.addText(cleanInlineText(row.registerOpportunityDisplay || row.competitorPubCommDisplay || "-"), {
      x: cols[5].x,
      y: y + 0.15,
      w: cols[5].w,
      h: 0.16,
      fontFace: deck.theme.fonts.heading,
      fontSize: 7.8,
      bold: true,
      align: "right",
      color: toColor(deck.theme.colors.ink),
      margin: 0,
      fit: "shrink"
    });
  });
}

function renderSlide(slide, deck, spec, pageNumber) {
  if (spec.kind === "cover") {
    addBlueChrome(slide, deck);
    addSlideWatermark(slide, deck, true);
    const coverTitle = cleanInlineText(spec.title || `${deck.metadata.client} Quarterly Business Review`);
    const match = coverTitle.match(/^(.*?)(business review)$/i);
    const titleRunsData = match
      ? [
          { text: `${match[1]}`, options: { color: toColor(deck.theme.colors.paper) } },
          { text: `${match[2]}`, options: { color: toColor(deck.theme.colors.ink) } }
        ]
      : [{ text: coverTitle, options: { color: toColor(deck.theme.colors.paper) } }];
    slide.addText(titleRunsData, {
      x: 0.68,
      y: 0.68,
      w: 11.55,
      h: 1.22,
      fontFace: deck.theme.fonts.heading,
      fontSize: 29,
      breakLine: true,
      margin: 0
    });
    const coverSummary = cleanInlineText(
      spec.summary
        || "A comprehensive year-over-year analysis of programme performance, publisher dynamics and strategic priorities."
    );
    slide.addText(coverSummary, {
      x: 0.68,
      y: 2.0,
      w: 11.35,
      h: 0.78,
      fontFace: deck.theme.fonts.body,
      fontSize: 12.2,
      color: toColor(deck.theme.colors.paper),
      breakLine: true,
      margin: 0
    });
    if (HAS_TD_LOGO_WHITE) {
      slide.addImage({
        path: TD_LOGO_WHITE_PATH,
        x: 0.62,
        y: 4.22,
        w: 2.15,
        h: 1.74
      });
    }
    addCyanFifthElementWireframe(slide, { x: 8.25, y: 2.78, w: 4.4, h: 4.4 });
    return;
  }

  if (spec.kind === "thank-you") {
    addBlueChrome(slide, deck);
    addSlideWatermark(slide, deck, true);
    addCyanFifthElementWireframe(slide, { x: 5.72, y: 0.08, w: 6.9, h: 6.9 });
    slide.addText(spec.title, {
      x: 0.7,
      y: 1.45,
      w: 10.8,
      h: 0.62,
      fontFace: deck.theme.fonts.heading,
      fontSize: 25,
      color: toColor(deck.theme.colors.paper),
      margin: 0
    });
    slide.addShape("roundRect", {
      x: 0.7,
      y: 2.55,
      w: 2.82,
      h: 1.12,
      radius: 0.06,
      line: { color: toColor(deck.theme.colors.paper), pt: 0.35, transparency: 55 },
      fill: { color: toColor(deck.theme.colors.paper), transparency: 42 }
    });
    slide.addText(uiLabel(deck, "anyQuestions", "Any Questions?"), {
      x: 0.7,
      y: 2.55,
      w: 2.82,
      h: 1.12,
      align: "center",
      valign: "mid",
      fontFace: deck.theme.fonts.heading,
      fontSize: 17.5,
      color: toColor(deck.theme.colors.paper),
      margin: 0
    });
    return;
  }

  if (spec.kind === "program-activation-snapshot") {
    renderProgramActivationSnapshotSlide(slide, deck, spec);
    return;
  }

  if (spec.kind === "program-connection-status") {
    renderProgramConnectionStatusSlide(slide, deck, spec);
    return;
  }

  if (spec.kind === "gap-analysis-impact") {
    renderGapAnalysisImpactSlide(slide, deck, spec);
    return;
  }

  if (spec.kind === "gap-analysis-top-programs") {
    renderGapAnalysisTopProgramsSlide(slide, deck, spec);
    return;
  }

  if (spec.kind === "gap-analysis-register") {
    renderGapAnalysisRegisterSlide(slide, deck, spec);
    return;
  }

  if (isBlueKind(spec.kind)) {
    addBlueChrome(slide, deck);
    addTitle(slide, deck, spec, deck.theme.colors.paper, deck.theme.colors.paper, true);
  } else {
    addLightChrome(slide, deck);
    addTitle(slide, deck, spec, deck.theme.colors.ink, deck.theme.colors.accent, false);
  }

  if (spec.kind === "reporting-period") {
    const locale = deck.metadata.locale || "en-GB";
    const selectedPeriod = cleanInlineText(spec?.bullets?.[0] || deck.metadata.reportingPeriod || "Not specified");
    const selectedPeriodReadable = formatPeriodForSlide(selectedPeriod, locale);
    const comparisonPeriodReadable = formatPeriodForSlide(deck.metadata.comparisonPeriod || "Not specified", locale);
    const selectedPeriodParsed = parseIsoPeriod(selectedPeriod);
    const selectedPeriodEndReadable = selectedPeriodParsed
      ? formatLongDate(selectedPeriodParsed.end, locale)
      : selectedPeriodReadable;

    slide.addText("Current and YoY comparison windows", {
      x: 0.72,
      y: 1.38,
      w: 11.3,
      h: 0.26,
      fontFace: deck.theme.fonts.body,
      fontSize: 10.2,
      color: toColor(deck.theme.colors.accent),
      margin: 0
    });
    [
      {
        heading: uiLabel(deck, "currentPeriod", "Current Period"),
        period: `${uiLabel(deck, "reportingPeriodPrefix", "Reporting Period")}: ${selectedPeriodReadable}`,
        basis: `${uiLabel(deck, "dataAsOfPrefix", "Data as of")}: ${selectedPeriodEndReadable}`
      },
      {
        heading: uiLabel(deck, "comparisonPeriodYoy", "Comparison Period (YoY)"),
        period: `${uiLabel(deck, "comparisonPeriodPrefix", "Comparison Period")}: ${comparisonPeriodReadable}`,
        basis: uiLabel(deck, "basisYoy", "Basis: Year-over-Year (YoY)")
      }
    ].forEach((block, index) => {
      const x = index === 0 ? 0.72 : 6.55;
      slide.addText(block.heading, {
        x,
        y: 1.92,
        w: 5.0,
        h: 0.34,
        fontFace: deck.theme.fonts.heading,
        fontSize: 16,
        color: toColor(deck.theme.colors.ink),
        margin: 0
      });
      slide.addText(block.period, {
        x,
        y: 2.50,
        w: 5.25,
        h: 0.26,
        fontFace: deck.theme.fonts.body,
        fontSize: 9.8,
        color: toColor(deck.theme.colors.muted),
        margin: 0
      });
      slide.addText(block.basis, {
        x,
        y: 2.84,
        w: 5.25,
        h: 0.26,
        fontFace: deck.theme.fonts.body,
        fontSize: 9.8,
        color: toColor(deck.theme.colors.muted),
        margin: 0
      });
    });
    slide.addShape("roundRect", {
      x: 0.72,
      y: 3.58,
      w: 11.95,
      h: 0.88,
      radius: 0.05,
      line: { color: toColor(deck.theme.colors.highlight), pt: 0.5 },
      fill: { color: toColor(deck.theme.colors.highlight), transparency: 10 }
    });
    slide.addText(uiLabel(deck, "allFiguresStatement", "All figures are reported in {currency} unless otherwise stated. YoY variance is calculated as Current Period vs Comparison Period.").replace("{currency}", deck.metadata.currencyCode || ""), {
      x: 0.98,
      y: 3.89,
      w: 11.2,
      h: 0.24,
      fontFace: deck.theme.fonts.heading,
      fontSize: 9.7,
      color: toColor(deck.theme.colors.ink),
      margin: 0
    });
    return;
  }

  if (spec.kind === "risks-dependencies") {
    renderRisksDependenciesTiles(slide, deck, spec);
    return;
  }

  if (spec.kind === "program-executive-summary") {
    addKpis(slide, deck, spec.kpis, { x: 0.7, y: 2.15 }, "diamond");
    if (spec.summary) {
      slide.addText(spec.summary, {
        x: 0.82,
        y: 5.87,
        w: 11.6,
        h: 0.9,
        fontFace: deck.theme.fonts.body,
        fontSize: 11.3,
        color: toColor(deck.theme.colors.ink),
        breakLine: true,
        margin: 0
      });
    }
    return;
  }

  if (spec.kind === "movers-bar-chart") {
    addMoversBarChart(slide, deck, spec.chart, {
      x: 0.58,
      y: 1.42,
      w: 12.18,
      h: 5.75
    });
    return;
  }

  if (spec.kind === "competitor-share-bar-chart") {
    addCompetitorShareBarChart(slide, deck, spec.chart, {
      x: 0.78,
      y: 1.80,
      w: 11.78,
      h: 4.85
    });
    return;
  }

  if (spec.kind === "insights-blue") {
    const insightItems = (spec.bullets || []).slice(0, 5);
    const items = insightItems.length
      ? insightItems
      : [uiLabel(deck, "kpiDriverUnavailable", "Driver not confirmed from available KPI data.")];

    const localizeKpiTitle = (title) => {
      const t = cleanInlineText(title || "").toLowerCase();
      if (/^conversion rate improvement$/.test(t)) return uiLabel(deck, "kpiTitleConversionRateImprovement", "Conversion Rate Improvement");
      if (/^sales volume pressure$/.test(t)) return uiLabel(deck, "kpiTitleSalesVolumePressure", "Sales Volume Pressure");
      if (/^aov growth partially offsetting volume decline$/.test(t)) return uiLabel(deck, "kpiTitleAovGrowthOffset", "AOV Growth Partially Offsetting Volume Decline");
      if (/^rising cpa$/.test(t)) return uiLabel(deck, "kpiTitleRisingCpa", "Rising CPA");
      if (/^(roi trend|trend roi)$/.test(t)) return uiLabel(deck, "kpiTitleRoiTrend", "ROI Trend");
      return cleanInlineText(title || "");
    };

    const inferKpiSignalTitle = (text) => {
      const t = cleanInlineText(text).toLowerCase();
      if (/(conv rate|conversion rate)/.test(t)) return uiLabel(deck, "kpiTitleConversionRateImprovement", "Conversion Rate Improvement");
      if (/(click|sales)/.test(t)) return uiLabel(deck, "kpiTitleSalesVolumePressure", "Sales Volume Pressure");
      if (/(aov|average order value|order value)/.test(t)) return uiLabel(deck, "kpiTitleAovGrowthOffset", "AOV Growth Partially Offsetting Volume Decline");
      if (/\bcpa\b|cost per acquisition|commission/.test(t)) return uiLabel(deck, "kpiTitleRisingCpa", "Rising CPA");
      if (/\broi\b|return on investment/.test(t)) return uiLabel(deck, "kpiTitleRoiTrend", "ROI Trend");
      return uiLabel(deck, "kpiSignalGeneric", "KPI Signal");
    };

    const parsed = items.map((raw) => {
      const text = cleanInlineText(raw);
      const idx = text.indexOf(":");
      if (idx > 8 && idx < 68) {
        return {
          title: localizeKpiTitle(text.slice(0, idx).trim()),
          detail: text.slice(idx + 1).trim() || uiLabel(deck, "kpiDetailUnavailable", "Detail not available from current extract.")
        };
      }
      return {
        title: inferKpiSignalTitle(text),
        detail: text
      };
    });

    const detailChars = parsed.reduce((sum, item) => sum + item.detail.length, 0);
    const titleSize = detailChars > 1250 ? 12.4 : 13.2;
    const detailSize = detailChars > 1250 ? 10.4 : 11.1;

    let y = 1.86;
    parsed.forEach((item) => {
      const detailLength = item.detail.length;
      const blockH = detailLength > 250 ? 1.24 : detailLength > 170 ? 1.06 : 0.92;
      slide.addText(`\u2022 ${item.title}`, {
        x: 0.80,
        y,
        w: 12.0,
        h: 0.30,
        align: "left",
        valign: "top",
        fontFace: deck.theme.fonts.heading,
        fontSize: titleSize,
        color: toColor(deck.theme.colors.paper),
        bold: true,
        margin: 0
      });
      slide.addText(item.detail, {
        x: 1.05,
        y: y + 0.31,
        w: 11.65,
        h: blockH - 0.20,
        align: "left",
        valign: "top",
        fontFace: deck.theme.fonts.body,
        fontSize: detailSize,
        color: toColor(deck.theme.colors.paper),
        breakLine: true,
        margin: 0
      });
      y += blockH + 0.08;
    });
    return;
  }

  if (spec.kind === "sales-growth-signals-blue") {
    const signals = Array.isArray(spec.signals) && spec.signals.length
      ? spec.signals.slice(0, 5)
      : [{ title: "Signal", detail: "No sales growth signal available from the current data extract." }];
    signals.forEach((signal, index) => {
      const y = 1.88 + index * 1.08;
      slide.addShape("roundRect", {
        x: 0.72,
        y: y + 0.06,
        w: 0.34,
        h: 0.34,
        radius: 0.05,
        line: { color: toColor(deck.theme.colors.paper), pt: 0 },
        fill: { color: toColor(deck.theme.colors.paper), transparency: 22 }
      });
      slide.addText(String(index + 1), {
        x: 0.84,
        y: y + 0.14,
        w: 0.10,
        h: 0.18,
        fontFace: deck.theme.fonts.heading,
        fontSize: 11,
        bold: true,
        align: "center",
        color: toColor(deck.theme.colors.ink),
        margin: 0
      });
      slide.addText(cleanInlineText(signal.title || `Signal ${index + 1}`), {
        x: 1.18,
        y,
        w: 11.4,
        h: 0.34,
        fontFace: deck.theme.fonts.heading,
        fontSize: 16,
        color: toColor(deck.theme.colors.paper),
        margin: 0
      });
      slide.addText(cleanInlineText(signal.detail || "Detail not available."), {
        x: 1.18,
        y: y + 0.34,
        w: 11.45,
        h: 0.70,
        fontFace: deck.theme.fonts.body,
        fontSize: 10.8,
        color: toColor(deck.theme.colors.paper),
        breakLine: true,
        margin: 0
      });
    });
    return;
  }

  if (spec.kind === "priority-actions") {
    const actions = (spec.bullets || []).slice(0, 3);
    const cardW = 3.82;
    const y = 2.55;
    [0, 1, 2].forEach((index) => {
      const x = 0.7 + (index * 4.08);
      slide.addShape("roundRect", {
        x,
        y,
        w: cardW,
        h: 3.05,
        radius: 0.04,
        line: { color: toColor(deck.theme.colors.border), pt: 0.7 },
        fill: { color: toColor(deck.theme.colors.paper) }
      });
      slide.addShape("rect", {
        x,
        y,
        w: cardW,
        h: 0.62,
        line: { color: toColor(deck.theme.colors.highlight), pt: 0 },
        fill: { color: toColor(deck.theme.colors.highlight) }
      });
      slide.addText(String(index + 1), {
        x: x + (cardW / 2) - 0.12,
        y: y + 0.12,
        w: 0.24,
        h: 0.26,
        align: "center",
        fontFace: deck.theme.fonts.heading,
        fontSize: 20,
        color: toColor(deck.theme.colors.ink),
        margin: 0
      });
      slide.addText(actions[index] || "Action to be confirmed from available data.", {
        x: x + 0.22,
        y: y + 0.9,
        w: cardW - 0.4,
        h: 2.0,
        fontFace: deck.theme.fonts.body,
        fontSize: 12,
        color: toColor(deck.theme.colors.ink),
        breakLine: true,
        margin: 0
      });
    });
    return;
  }

  if (spec.kind === "publisher-overview") {
    let overviewTableMetrics = null;
    slide.addText(uiLabel(deck, "publisherActivityBySegment", "Publisher Activity by Segment"), {
      x: 0.35,
      y: 2.04,
      w: 5.6,
      h: 0.32,
      fontFace: deck.theme.fonts.body,
      fontSize: 12.5,
      color: toColor(deck.theme.colors.ink),
      margin: 0
    });
    if (spec.tables && spec.tables[0]) {
      overviewTableMetrics = addTable(slide, deck, spec.tables[0], { x: 0.35, y: 2.38, w: 5.55, h: 4.32 });
    }
    const points = (spec.bullets || []).slice(0, 4);
    slide.addText(uiLabel(deck, "keyObservations", "Key Observations"), {
      x: 6.25,
      y: 2.04,
      w: 5.6,
      h: 0.35,
      fontFace: deck.theme.fonts.heading,
      fontSize: 15,
      color: toColor(deck.theme.colors.ink),
      margin: 0
    });
    const notes = points.length ? points : ["Driver not confirmed from available data."];
    const observationRuns = [];
    notes.forEach((item, index) => {
      observationRuns.push({ text: `\u2022 ${item}`, options: { breakLine: true } });
      if (index < notes.length - 1) {
        observationRuns.push({ text: " ", options: { breakLine: true } });
      }
    });
    slide.addText(observationRuns, {
      x: 6.26,
      y: 2.56,
      w: 5.22,
      h: 4.74,
      fontFace: deck.theme.fonts.body,
      fontSize: 11.4,
      color: toColor(deck.theme.colors.ink),
      breakLine: true,
      margin: 0.02,
      valign: "top"
    });
    if (overviewTableMetrics && overviewTableMetrics.containerH < 4.32) {
      slide.addShape("line", {
        x: 0.35,
        y: 2.38 + overviewTableMetrics.containerH + 0.1,
        w: 5.55,
        h: 0,
        line: { color: toColor("#E6EAF2"), pt: 0.6 }
      });
    }
    return;
  }

  if (spec.kind === "segment-performance-blue" || spec.kind === "segment-performance") {
    const blocks = (spec.bullets || []).slice(0, 5);
    const segmentTileRadiusIn = Number((4 / 96).toFixed(4)); // 4px at 96 DPI
    const segmentSignalUnavailable = uiLabel(deck, "segmentSignalUnavailable", "Segment signal not available.");
    const detailedMovementUnavailable = uiLabel(deck, "detailedMovementUnavailable", "Detailed movement not available from this extract.");
    const clampText = (value, maxChars = 9999) => {
      const text = cleanInlineText(value || "");
      if (!text || text.length <= maxChars) return text;
      return `${text.slice(0, maxChars - 1).trimEnd()}…`;
    };
    // Use larger tiles and full-width bottom row so segment analysis does not clip.
    const layout = [
      { x: 0.56, y: 1.54, w: 5.85, h: 1.90 },
      { x: 6.82, y: 1.54, w: 5.85, h: 1.90 },
      { x: 0.56, y: 3.56, w: 5.85, h: 1.90 },
      { x: 6.82, y: 3.56, w: 5.85, h: 1.90 },
      { x: 0.56, y: 5.58, w: 12.11, h: 1.72 }
    ];
    layout.forEach((box, idx) => {
      const isBottomRow = idx === 4;
      const raw = cleanText(blocks[idx] || segmentSignalUnavailable);
      const lines = raw.split(/\r?\n/).map((line) => cleanInlineText(line)).filter(Boolean);
      const heading = clampText(lines[0] || "Segment", isBottomRow ? 160 : 120);
      const detail = clampText(lines.slice(1).join(" ") || detailedMovementUnavailable, isBottomRow ? 620 : 420);
      const headingMatch = heading.match(/^(.+?)\s*-\s*([+-]?\d+(?:[.,]\d+)?%.*)$/i);
      const headingPrefix = headingMatch ? headingMatch[1].trim() : heading;
      const headingSuffix = headingMatch ? headingMatch[2].trim() : "";
      const headingFontSize = isBottomRow ? 12.0 : 12.4;
      const detailFontSize = isBottomRow ? 9.8 : 10.0;

      slide.addShape("roundRect", {
        x: box.x,
        y: box.y,
        w: box.w,
        h: box.h,
        radius: segmentTileRadiusIn,
        rectRadius: segmentTileRadiusIn,
        line: { color: toColor("#AFC4F5"), pt: 0.7 },
        fill: { color: toColor(deck.theme.colors.paper), transparency: 0 }
      });
      const headingRuns = headingSuffix
        ? [
            { text: `${headingPrefix} - `, options: { color: toColor(deck.theme.colors.ink) } },
            { text: headingSuffix, options: { color: toColor(deck.theme.colors.accent) } }
          ]
        : [{ text: headingPrefix, options: { color: toColor(deck.theme.colors.ink) } }];
      slide.addText(headingRuns, {
        x: box.x + 0.22,
        y: box.y + 0.14,
        w: box.w - 0.36,
        h: isBottomRow ? 0.32 : 0.34,
        fontFace: deck.theme.fonts.heading,
        fontSize: headingFontSize,
        margin: 0
      });
      slide.addText(detail, {
        x: box.x + 0.22,
        y: box.y + 0.48,
        w: box.w - 0.36,
        h: box.h - 0.56,
        fontFace: deck.theme.fonts.body,
        fontSize: detailFontSize,
        color: toColor(deck.theme.colors.ink),
        breakLine: true,
        margin: 0
      });
    });
    return;
  }

  if (spec.kind === "new-emerging") {
    const cards = (spec.bullets || []).slice(0, 4);
    const layout = [
      { x: 0.66, y: 1.95 },
      { x: 6.24, y: 1.95 },
      { x: 0.66, y: 4.2 },
      { x: 6.24, y: 4.2 }
    ];
    layout.forEach((box, idx) => {
      slide.addShape("roundRect", {
        x: box.x,
        y: box.y,
        w: 5.45,
        h: 1.9,
        radius: 0.04,
        line: { color: toColor(deck.theme.colors.highlight), pt: 0.8 },
        fill: { color: toColor(deck.theme.colors.highlight), transparency: 10 }
      });
      slide.addText(cards[idx] || "Publisher note not available.", {
        x: box.x + 0.2,
        y: box.y + 0.25,
        w: 5.05,
        h: 1.45,
        fontFace: deck.theme.fonts.body,
        fontSize: 12,
        color: toColor(deck.theme.colors.ink),
        breakLine: true,
        margin: 0
      });
    });
    addCallout(
      slide,
      deck,
      "Emerging publishers are defined as partners with limited active history; full YoY signal is expected in the next cycle.",
      6.25,
      true
    );
    return;
  }

  if (spec.kind === "recommendations-blue") {
    const actions = (spec.bullets || []).slice(0, 5);
    actions.forEach((item, index) => {
      const y = 1.9 + index * 1.15;
      slide.addShape("roundRect", {
        x: 0.54,
        y: y + 0.04,
        w: 0.34,
        h: 0.34,
        radius: 0.04,
        line: { color: toColor(deck.theme.colors.paper), pt: 0.8 },
        fill: { color: toColor(deck.theme.colors.paper), transparency: 18 }
      });
      slide.addText(String(index + 1), {
        x: 0.64,
        y: y + 0.11,
        w: 0.14,
        h: 0.16,
        align: "center",
        fontFace: deck.theme.fonts.heading,
        fontSize: 10,
        bold: true,
        color: toColor(deck.theme.colors.ink),
        margin: 0
      });
      slide.addText(`${index + 1}. ${item}`, {
        x: 1.0,
        y,
        w: 11.7,
        h: 0.88,
        fontFace: deck.theme.fonts.body,
        fontSize: 13,
        color: toColor(deck.theme.colors.paper),
        breakLine: true,
        margin: 0
      });
    });
    return;
  }

  if (spec.kind === "competitor-analysis") {
    if (spec.tables && spec.tables[0]) {
      addTable(slide, deck, spec.tables[0], {
        x: 0.40,
        y: 1.82,
        w: 12.50,
        h: 1.35
      });
    }
    if (spec.chart) {
      addWeeklyComboChart(slide, deck, spec.chart, {
        x: 0.56,
        y: 3.42,
        w: 12.18,
        h: 3.42
      });
    }
    return;
  }

  if (spec.kind === "kpi-table" || spec.kind === "publisher-table" || spec.kind === "competitor-performance-table" || spec.kind === "program-breakdown" || spec.kind === "appendix" || spec.kind === "risks-dependencies") {
    const isProgramTable = spec.kind === "kpi-table" || spec.kind === "program-breakdown";
    const tableY = isProgramTable ? 1.78 : (spec.kind === "competitor-performance-table" ? 1.86 : 1.95);
    const hasFooterNote = spec.kind === "kpi-table" && Boolean(cleanInlineText(spec.footerNote || ""));
    const tableH = isProgramTable ? (hasFooterNote ? 5.05 : 5.55) : (spec.kind === "appendix" ? 5.15 : (spec.kind === "competitor-performance-table" ? 5.05 : 4.85));
    let renderedTable = null;
    if (spec.tables && spec.tables[0]) {
      renderedTable = addTable(slide, deck, spec.tables[0], {
        x: 0.40,
        y: tableY,
        w: 12.50,
        h: tableH
      });
    }
    if (hasFooterNote) {
      const footerY = renderedTable
        ? Math.min(7.05, tableY + renderedTable.containerH + 0.06)
        : 6.95;
      slide.addText(spec.footerNote, {
        x: 0.52,
        y: footerY,
        w: 12.2,
        h: 0.34,
        fontFace: deck.theme.fonts.body,
        fontSize: 10.2,
        color: toColor(deck.theme.colors.ink),
        margin: 0,
        breakLine: true
      });
    }
    return;
  }

  if (spec.kind === "executive-summary" || spec.kind === "program-overview" || spec.kind === "recommendations") {
    if (spec.headline) {
      slide.addText(spec.headline, {
        x: 0.86,
        y: 1.34,
        w: 5.6,
        h: 0.72,
        fontFace: deck.theme.fonts.heading,
        fontSize: 26,
        color: toColor(deck.theme.colors.ink),
        margin: 0
      });
    }
    if (spec.summary) {
      slide.addText(spec.summary, {
        x: 0.88,
        y: spec.headline ? 2.06 : 1.42,
        w: 5.3,
        h: 0.58,
        fontFace: deck.theme.fonts.body,
        fontSize: 12,
        color: toColor(deck.theme.colors.muted),
        margin: 0
      });
    }
    addBullets(slide, deck, spec.bullets, { x: 0.88, y: 2.84, w: 5.3, h: 3.52 });
    if (spec.kpis && spec.kpis.length) {
      addKpis(slide, deck, spec.kpis, { x: 6.82, y: 1.4 });
    }
    return;
  }

  if (spec.kind === "kpi-snapshot") {
    addKpis(slide, deck, spec.kpis, { x: 0.88, y: 1.72 });
    addBullets(slide, deck, spec.bullets, { x: 0.9, y: 4.52, w: 11.3, h: 1.48 });
    return;
  }

  if (spec.tables && spec.tables[0]) {
    addTable(slide, deck, spec.tables[0], {
      x: 0.88,
      y: spec.kind === "appendix" ? 1.42 : 1.5,
      w: 11.52,
      h: spec.kind === "appendix" ? 5.35 : 5.0
    });
  }
}

async function renderDeck(deck) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "QBR PPTX Service";
  pptx.company = deck.theme.companyName;
  pptx.subject = "Quarterly business review";
  pptx.title = deck.metadata.deckTitle;

  deck.slides.forEach((spec, index) => {
    const slide = pptx.addSlide();
    renderSlide(slide, deck, spec, index + 1);
  });

  const output = await pptx.write({ outputType: "nodebuffer" });
  return Buffer.isBuffer(output) ? output : Buffer.from(output);
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function excelColumnName(index) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function excelCellXml(value, rowIndex, colIndex, styleId = 0) {
  const ref = `${excelColumnName(colIndex)}${rowIndex + 1}`;
  const style = styleId ? ` s="${styleId}"` : "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"${style}><v>${value}</v></c>`;
  }
  if (value === null || value === undefined || value === "") return `<c r="${ref}"${style}/>`;
  return `<c r="${ref}" t="inlineStr"${style}><is><t>${xmlEscape(value)}</t></is></c>`;
}

function excelRowXml(row, rowIndex, styleForCell) {
  return `<row r="${rowIndex + 1}">${row.map((cell, colIndex) => (
    excelCellXml(cell, rowIndex, colIndex, styleForCell(rowIndex, colIndex, cell))
  )).join("")}</row>`;
}

function buildGapAnalysisWorkbookBuffer(report) {
  if (!report || !Array.isArray(report.allRows) || !report.allRows.length) return null;
  const remainingRows = Array.isArray(report.remainingRows) ? report.remainingRows : [];
  const topCount = Array.isArray(report.topRows) ? report.topRows.length : 0;
  const client = cleanInlineText(report.client || "Primary Publisher") || "Primary Publisher";
  const title = cleanInlineText(report.title || `${client} Gap Analysis Report`);
  const rows = [
    [`${title} - remaining programs after top ${topCount || 10} shown in PowerPoint`, "", "", "", "", "", ""],
    ["Program Name", "Program ID", client, "Pub Comm - Specified Sites", "Gap Type", "Recommended Action", "Competitor Signal"],
    ...remainingRows.map((row) => [
      row.programName || "-",
      row.programId || "-",
      row.primaryStatus || "-",
      row.registerOpportunityDisplay || row.competitorPubCommDisplay || "-",
      row.gapType || "-",
      row.action || "-",
      row.competitorSignal || "-"
    ]),
    [],
    ["Connection Type", "Description", "", "", "", "", ""],
    ["Pub Comm:", `Programs where ${client} are earning commission`, "", "", "", "", ""],
    ["No Connection:", `Programs where ${client} have never before applied`, "", "", "", "", ""],
    ["Clicks:", `Programs where ${client} are driving clicks on a program but no conversions/commission`, "", "", "", "", ""],
    ["Accepted:", `Programs where ${client} are accepted but driving no clicks / conversions`, "", "", "", "", ""],
    ["Denied:", `Programs where ${client} application was denied`, "", "", "", "", ""],
    ["Ended:", `Programs where ${client} were once accepted but no longer are`, "", "", "", "", ""],
    ["Hold Accepted:", `Programs where ${client} are accepted but temporarily placed on hold`, "", "", "", "", ""]
  ];
  const mergeEndRow = remainingRows.length + 5;
  const sheetData = rows.map((row, rowIndex) => excelRowXml(row, rowIndex, (r) => {
    if (r === 0) return 1;
    if (r === 1 || r === remainingRows.length + 4) return 2;
    return 0;
  })).join("");
  const dimension = `A1:G${rows.length}`;
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>
    <col min="1" max="1" width="44" customWidth="1"/>
    <col min="2" max="2" width="13" customWidth="1"/>
    <col min="3" max="3" width="18" customWidth="1"/>
    <col min="4" max="4" width="22" customWidth="1"/>
    <col min="5" max="5" width="18" customWidth="1"/>
    <col min="6" max="6" width="24" customWidth="1"/>
    <col min="7" max="7" width="26" customWidth="1"/>
  </cols>
  <sheetData>${sheetData}</sheetData>
  <mergeCells count="1"><mergeCell ref="A1:G1"/></mergeCells>
  <autoFilter ref="A2:G${Math.max(2, remainingRows.length + 2)}"/>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Gap Analysis" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", sheetXml);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3"><font><sz val="11"/><name val="Aptos"/></font><font><b/><sz val="14"/><color rgb="FF1F2533"/><name val="Aptos Display"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2F6FF2"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9DEE8"/></left><right style="thin"><color rgb="FFD9DEE8"/></right><top style="thin"><color rgb="FFD9DEE8"/></top><bottom style="thin"><color rgb="FFD9DEE8"/></bottom><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`);
  const now = new Date().toISOString();
  zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xmlEscape(title)}</dc:title><dc:creator>QBR PPTX Service</dc:creator><cp:lastModifiedBy>QBR PPTX Service</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>QBR PPTX Service</Application></Properties>`);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

function safeName(value) {
  return String(value || "qbr_deck")
    .trim()
    .replace(/[\u25B2\u25BC\u25B3\u25BD\u25B4\u25BE]/g, "")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .toLowerCase() || "qbr_deck";
}

async function generatePresentation(payload, options = {}) {
  const normalized = normalizePayload(payload || {});
  const theme = resolveTheme(normalized.themeName, normalized.themeOverrides);
  const deckSpec = buildDeckSpec(normalized, theme);
  const localizedDeckSpec = await localizeDeckSpec(deckSpec, normalized.languageCode);
  const buffer = await renderDeck(localizedDeckSpec);
  const fileName = normalized.outputFileName || `${safeName(localizedDeckSpec.metadata.deckTitle)}_${crypto.randomUUID()}.pptx`;
  const gapReportBuffer = await buildGapAnalysisWorkbookBuffer(localizedDeckSpec.reports?.gapAnalysis);
  const gapReportFileName = gapReportBuffer
    ? fileName.replace(/\.pptx$/i, ".gap-analysis.xlsx")
    : null;

  return { normalized, deckSpec: localizedDeckSpec, buffer, fileName, gapReportBuffer, gapReportFileName };
}

async function saveOutput(result, outputDir) {
  await fs.mkdir(outputDir, { recursive: true });
  const pptxPath = path.join(outputDir, result.fileName);
  await fs.writeFile(pptxPath, result.buffer);
  let gapReportPath = null;
  if (result.gapReportBuffer && result.gapReportFileName) {
    gapReportPath = path.join(outputDir, result.gapReportFileName);
    await fs.writeFile(gapReportPath, result.gapReportBuffer);
  }

  let deckSpecFileName = null;
  if (result.normalized.debug) {
    deckSpecFileName = result.fileName.replace(/\.pptx$/i, ".deck-spec.json");
    await fs.writeFile(path.join(outputDir, deckSpecFileName), JSON.stringify(result.deckSpec, null, 2), "utf8");
  }

  return { pptxPath, deckSpecFileName, gapReportPath };
}

module.exports = {
  generatePresentation,
  saveOutput
};
