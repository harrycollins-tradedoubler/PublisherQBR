const fs = require("node:fs/promises");
const path = require("node:path");

const { createPublisherQbrAgent, redactSensitive } = require("./publisherQbrAgent");

const DEFAULT_PPTX_SERVICE_URL = "http://127.0.0.1:3010";
const DEFAULT_API_KEY = "td-publisher-qbr-local-2026-secret";

function asText(value, fallback = "") {
  if (value === null || value === undefined) return fallback;
  const text = String(value).trim();
  return text || fallback;
}

function parseMaybeJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function extractIncomingPayload(input = {}) {
  const source = input.body && typeof input.body === "object" ? input.body : input;
  const raw = source.incomingText || source.message || source.body?.message || "";
  const prefix = "QBR_REQUEST";
  let payload = {};

  if (source.payload && typeof source.payload === "object" && !Array.isArray(source.payload)) {
    payload = source.payload;
  } else if (source.qbr_payload && typeof source.qbr_payload === "object" && !Array.isArray(source.qbr_payload)) {
    payload = source.qbr_payload;
  } else if (source.body?.payload && typeof source.body.payload === "object" && !Array.isArray(source.body.payload)) {
    payload = source.body.payload;
  } else if (source.body?.qbr_payload && typeof source.body.qbr_payload === "object" && !Array.isArray(source.body.qbr_payload)) {
    payload = source.body.qbr_payload;
  } else if (typeof raw === "string" && raw.startsWith(prefix)) {
    payload = parseMaybeJson(raw.slice(prefix.length).trim()) || {};
  }

  const tdTokens = source.body?.td_tokens
    || source.td_tokens
    || source.qbr_payload?.td_tokens
    || source.body?.qbr_payload?.td_tokens
    || payload.td_tokens
    || {};

  return {
    ...source,
    payload,
    tdTokens
  };
}

function normalizeIdList(input) {
  if (Array.isArray(input)) return input.map((value) => asText(value)).filter(Boolean);
  if (typeof input === "string" && input.trim()) return input.split(",").map((value) => value.trim()).filter(Boolean);
  if (input !== null && input !== undefined && asText(input)) return [asText(input)];
  return [];
}

function extractSourceIdsFromEndpoint(url) {
  if (!url || typeof url !== "string") return [];
  try {
    const parsed = new URL(url);
    return ["sourceID", "sourceId", "sourceid"]
      .map((key) => asText(parsed.searchParams.get(key)))
      .filter(Boolean);
  } catch {
    return [];
  }
}

function toYyyyMmDd(input, fallbackDate) {
  const value = asText(input);
  if (/^\d{8}$/.test(value)) return value;
  const digits = value.replace(/[^0-9]/g, "");
  if (digits.length === 8) return digits;
  const date = fallbackDate instanceof Date ? fallbackDate : new Date(fallbackDate);
  return `${date.getUTCFullYear()}${String(date.getUTCMonth() + 1).padStart(2, "0")}${String(date.getUTCDate()).padStart(2, "0")}`;
}

function parseDateParts(yyyymmdd) {
  const value = asText(yyyymmdd);
  if (!/^\d{8}$/.test(value)) return null;
  const year = Number(value.slice(0, 4));
  const month = Number(value.slice(4, 6));
  const day = Number(value.slice(6, 8));
  if (!year || month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { year, month, day };
}

function formatDateParts(parts) {
  return `${String(parts.year).padStart(4, "0")}${String(parts.month).padStart(2, "0")}${String(parts.day).padStart(2, "0")}`;
}

function firstDayOfNextMonth(parts) {
  return parts.month === 12
    ? { year: parts.year + 1, month: 1, day: 1 }
    : { year: parts.year, month: parts.month + 1, day: 1 };
}

function normalizeDateRangeByInterval(fromDate, toDate, interval) {
  if (interval !== "month") return { fromDate, toDate };
  const fromParts = parseDateParts(fromDate);
  const toParts = parseDateParts(toDate);
  if (!fromParts || !toParts) return { fromDate, toDate };
  const monthFrom = formatDateParts({ year: fromParts.year, month: fromParts.month, day: 1 });
  let monthTo = toParts.day === 1 ? formatDateParts(toParts) : formatDateParts(firstDayOfNextMonth(toParts));
  if (monthTo <= monthFrom) monthTo = formatDateParts(firstDayOfNextMonth(parseDateParts(monthFrom)));
  return { fromDate: monthFrom, toDate: monthTo };
}

function shiftYear(yyyymmdd, years) {
  const value = asText(yyyymmdd);
  return `${Number(value.slice(0, 4)) + years}${value.slice(4, 6)}${value.slice(6, 8)}`;
}

function shiftDays(yyyymmdd, days) {
  const parts = parseDateParts(yyyymmdd);
  if (!parts) return yyyymmdd;
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
  date.setUTCDate(date.getUTCDate() + days);
  return toYyyyMmDd("", date);
}

function formatYmd(yyyymmdd) {
  const value = asText(yyyymmdd);
  return /^\d{8}$/.test(value) ? `${value.slice(0, 4)}-${value.slice(4, 6)}-${value.slice(6, 8)}` : value;
}

function normalizeCompetitorPublishers(rawCompetitors, rawPayload) {
  if (!Array.isArray(rawCompetitors)) return [];
  return rawCompetitors.slice(0, 4).map((competitor, index) => {
    const competitorSourceIds = Array.from(new Set([
      ...normalizeIdList(competitor?.sourceIDs),
      ...normalizeIdList(competitor?.sourceIds),
      ...normalizeIdList(competitor?.sourceID),
      ...normalizeIdList(competitor?.sourceId),
      ...normalizeIdList(competitor?.siteID),
      ...extractSourceIdsFromEndpoint(competitor?.publisherEndpoint || competitor?.publisherEndpointUrl || "")
    ].filter(Boolean)));
    const primarySourceId = competitorSourceIds[0] || "";
    return {
      ...(competitor || {}),
      label: asText(competitor?.label, `Publisher ${index + 1}`),
      clientUsername: asText(competitor?.clientUsername),
      sourceIDs: competitorSourceIds,
      sourceIds: competitorSourceIds,
      sourceID: primarySourceId,
      sourceId: primarySourceId,
      primarySourceId,
      sourceFilterEnabled: Boolean(primarySourceId),
      publisherExportEndpoint: asText(
        competitor?.publisherExportEndpoint || rawPayload.publisherExportEndpoint,
        "https://connect.tradedoubler.com/publisher/report/statistics/export"
      )
    };
  }).filter((competitor) => competitor.clientUsername && competitor.primarySourceId);
}

function normalizeWorkflowInput(input = {}, options = {}) {
  const extracted = extractIncomingPayload(input);
  const rawPayload = extracted.payload || {};
  const today = options.now instanceof Date ? options.now : new Date();
  const fallbackTo = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
  const fallbackFrom = new Date(fallbackTo);
  fallbackFrom.setUTCDate(fallbackFrom.getUTCDate() - 89);

  const endpointSourceIds = extractSourceIdsFromEndpoint(rawPayload.publisherEndpoint || rawPayload.publisherEndpointUrl || "");
  const sourceIds = Array.from(new Set([
    ...normalizeIdList(rawPayload.sourceIDs),
    ...normalizeIdList(rawPayload.sourceIds),
    ...normalizeIdList(rawPayload.sourceID),
    ...normalizeIdList(rawPayload.sourceId),
    ...normalizeIdList(rawPayload.siteID),
    ...endpointSourceIds
  ].filter(Boolean)));
  const primarySourceId = sourceIds[0] || "";

  const reportType = asText(rawPayload.reportType, "program").toLowerCase();
  const requestedIntervalType = asText(rawPayload.intervalType, "day").toLowerCase();
  const validIntervals = new Set(["day", "month", "year", "period"]);
  const intervalType = validIntervals.has(requestedIntervalType) ? requestedIntervalType : "day";
  const range = normalizeDateRangeByInterval(
    toYyyyMmDd(rawPayload.fromDate || rawPayload.startDate || rawPayload.from, fallbackFrom),
    toYyyyMmDd(rawPayload.toDate || rawPayload.endDate || rawPayload.to, fallbackTo),
    intervalType
  );
  const prevFromDate = shiftYear(range.fromDate, -1);
  const prevToDate = shiftYear(range.toDate, -1);
  const programStatusBackdateDays = Math.max(0, Number(rawPayload.programStatusBackdateDays || 90) || 90);
  const languageCode = asText(rawPayload.languageCode, "EN").toUpperCase();
  const languageMap = {
    EN: "English", DE: "German", FR: "French", ES: "Spanish", IT: "Italian",
    NL: "Dutch", PL: "Polish", SV: "Swedish", NO: "Norwegian", DA: "Danish", FI: "Finnish"
  };

  const competitorPublishers = normalizeCompetitorPublishers(
    rawPayload.competitorPublishers || rawPayload.comparisonPublishers || rawPayload.competitors,
    rawPayload
  );

  return {
    ...extracted,
    payload: {
      ...rawPayload,
      td_tokens: rawPayload.td_tokens || extracted.tdTokens,
      analysisLevel: "publisher_program",
      sourceIDs: sourceIds,
      sourceIds,
      sourceID: primarySourceId,
      sourceId: primarySourceId,
      primarySourceId,
      sourceFilterEnabled: rawPayload.sourceFilterEnabled === undefined
        ? Boolean(primarySourceId)
        : String(rawPayload.sourceFilterEnabled).toLowerCase() === "true",
      competitorPublishers,
      competitorPublisherCount: competitorPublishers.length,
      fromDate: range.fromDate,
      toDate: range.toDate,
      currencyCode: asText(rawPayload.currencyCode, "EUR").toUpperCase(),
      languageCode,
      languageName: languageMap[languageCode] || "English",
      reportingPeriod: `${formatYmd(range.fromDate)} to ${formatYmd(range.toDate)}`,
      comparisonPeriod: `${formatYmd(prevFromDate)} to ${formatYmd(prevToDate)}`,
      reportingPeriodLabel: asText(rawPayload.reportingPeriod, "custom"),
      qbrFocus: asText(rawPayload.qbrFocus || rawPayload.focusTheme, "General performance review"),
      qbrFocusDetail: asText(rawPayload.qbrFocusDetail || rawPayload.focusDetail),
      publisherExportEndpoint: asText(rawPayload.publisherExportEndpoint, "https://connect.tradedoubler.com/publisher/report/statistics/export"),
      digitalWalletEndpoint: asText(rawPayload.digitalWalletEndpoint || rawPayload.publisherDigitalWalletEndpoint, "https://connect.tradedoubler.com/publisher/report/payments/digitalwallets"),
      programStatusEndpoint: asText(rawPayload.programStatusEndpoint || rawPayload.publisherProgramsEndpoint, "https://connect.tradedoubler.com/publisher/programs"),
      programStatusCreatedFromDate: formatYmd(rawPayload.programStatusCreatedFromDateRaw || shiftDays(range.fromDate, -programStatusBackdateDays)),
      programStatusBackdateDays,
      digitalWalletCurrentFromDate: range.fromDate,
      digitalWalletCurrentToDate: range.toDate,
      digitalWalletPreviousFromDate: prevFromDate,
      digitalWalletPreviousToDate: prevToDate,
      digitalWalletLimit: Math.max(1, Math.min(100, Number(rawPayload.digitalWalletLimit || 100) || 100)),
      digitalWalletStatus: asText(rawPayload.digitalWalletStatus).toUpperCase(),
      digitalWalletProgramId: asText(rawPayload.digitalWalletProgramId || rawPayload.programId),
      reportType,
      intervalType
    },
    prev: {
      fromDate: prevFromDate,
      toDate: prevToDate
    }
  };
}

function appendParams(baseUrl, params) {
  const url = new URL(baseUrl);
  for (const [key, value] of Object.entries(params)) {
    if (value === undefined || value === null || value === "") continue;
    url.searchParams.set(key, String(value));
  }
  return url;
}

async function parseHttpResponse(response, label) {
  const text = await response.text();
  let body = text;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {}
  if (!response.ok) {
    const detail = body && typeof body === "object" ? body.detail || body.message || body.error : String(body || "");
    throw new Error(`${label} failed: HTTP ${response.status}${detail ? ` ${detail}` : ""}`);
  }
  return body;
}

async function fetchJsonWithRetry(url, init, options = {}) {
  const fetchImpl = options.fetch || globalThis.fetch;
  const retries = Math.max(0, Number(options.retries ?? 2));
  const timeoutMs = Math.max(100, Number(options.timeoutMs ?? 45000));
  const label = options.label || "HTTP request";
  let lastError;

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    try {
      options.logger?.info?.("publisher_qbr_fetch", redactSensitive({ label, url: String(url), attempt, headers: init?.headers || {} }));
      const response = await fetchImpl(url, { ...init, signal: controller.signal });
      return await parseHttpResponse(response, label);
    } catch (error) {
      lastError = error;
      if (attempt >= retries) break;
      options.logger?.warn?.("publisher_qbr_fetch_retry", redactSensitive({ label, url: String(url), attempt, error: error.message }));
    } finally {
      clearTimeout(timeout);
    }
  }
  throw lastError;
}

function rowsFromResponse(response) {
  if (Array.isArray(response)) return response;
  if (!response || typeof response !== "object") return [];
  for (const key of ["items", "data", "rows", "results", "content", "digitalWallets", "wallets"]) {
    if (Array.isArray(response[key])) return response[key];
  }
  return [response];
}

function paginationState(response, currentUrl) {
  if (!response || typeof response !== "object") return null;
  const next = response.next || response.nextPage || response.next_url || response.nextUrl;
  if (typeof next === "string" && next) return new URL(next, currentUrl);
  const total = Number(response.total ?? response.totalCount ?? response.count);
  const limit = Number(response.limit ?? new URL(currentUrl).searchParams.get("limit"));
  const offset = Number(response.offset ?? new URL(currentUrl).searchParams.get("offset") ?? 0);
  if (Number.isFinite(total) && Number.isFinite(limit) && limit > 0 && offset + limit < total) {
    const nextUrl = new URL(currentUrl);
    nextUrl.searchParams.set("offset", String(offset + limit));
    nextUrl.searchParams.set("limit", String(limit));
    return nextUrl;
  }
  return null;
}

async function fetchAllPages(url, init, options = {}) {
  let nextUrl = new URL(url);
  const rows = [];
  const maxPages = Math.max(1, Number(options.maxPages || 25));
  for (let page = 0; page < maxPages && nextUrl; page += 1) {
    const body = await fetchJsonWithRetry(nextUrl, init, options);
    rows.push(...rowsFromResponse(body));
    nextUrl = paginationState(body, nextUrl);
  }
  return rows;
}

function authHeaders(token) {
  return {
    Authorization: `Bearer ${asText(token)}`,
    Accept: "application/json"
  };
}

function buildMetricsUrl(payload, fromDate, toDate, sourceId, endpoint) {
  return appendParams(endpoint || payload.publisherExportEndpoint, {
    reportType: payload.reportType,
    intervalType: payload.intervalType,
    reportCurrencyCode: payload.currencyCode,
    fromDate,
    toDate,
    sourceId: payload.sourceFilterEnabled && sourceId ? sourceId : ""
  });
}

function walletUrl(payload, fromDate, toDate) {
  return appendParams(payload.digitalWalletEndpoint, {
    fromDate,
    toDate,
    limit: payload.digitalWalletLimit,
    offset: 0,
    sortBy: "programId",
    sortOrder: "asc",
    sourceId: payload.sourceFilterEnabled && payload.primarySourceId ? payload.primarySourceId : "",
    programId: payload.digitalWalletProgramId,
    status: payload.digitalWalletStatus
  });
}

function programStatusUrl(payload, sourceId) {
  return appendParams(payload.programStatusEndpoint, { sourceId });
}

function compactRows(rows, limit = 120) {
  return rows.slice(0, limit).map((row) => {
    if (!row || typeof row !== "object" || Array.isArray(row)) return { Value: asText(row, "-") };
    return Object.fromEntries(Object.entries(row).filter(([, value]) => value !== undefined && value !== null));
  });
}

function safeNum(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const normalized = String(value).replace(/[^\d,.-]/g, "").replace(/,/g, "");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function currencySymbol(code) {
  switch (String(code || "").toUpperCase()) {
    case "GBP": return "Â£";
    case "USD": return "$";
    case "AUD": return "A$";
    case "PLN": return "zÅ‚";
    case "SEK":
    case "NOK":
    case "DKK":
      return "kr";
    case "EUR":
    default:
      return "â‚¬";
  }
}

function fmtInt(value) {
  return Math.round(safeNum(value)).toLocaleString("en-GB");
}

function fmtMoney(value, symbol) {
  const n = Math.round(safeNum(value));
  const sign = n < 0 ? "-" : "";
  return `${sign}${symbol}${Math.abs(n).toLocaleString("en-GB")}`;
}

function fmtMoney2(value, symbol) {
  const n = safeNum(value);
  const sign = n < 0 ? "-" : "";
  return `${sign}${symbol}${Math.abs(n).toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function fmtMoneyCompact(value, symbol) {
  const n = safeNum(value);
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(n);
  if (abs >= 1000) return `${sign}${symbol}${(abs / 1000).toFixed(1)}k`;
  return fmtMoney(abs * (n < 0 ? -1 : 1), symbol);
}

function fmtPct(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
  const n = Number(value);
  const sign = n > 0 ? "+" : "";
  return `${sign}${n.toFixed(1)}%`;
}

function formatSharePercentages(values) {
  const nums = values.map((value) => Math.max(0, safeNum(value)));
  const total = nums.reduce((sum, value) => sum + value, 0);
  if (total <= 0) return nums.map(() => "0%");
  const exact = nums.map((value) => (value / total) * 100);
  const rounded = exact.map((value) => Math.floor(value));
  let remaining = 100 - rounded.reduce((sum, value) => sum + value, 0);
  exact
    .map((value, index) => ({
      index,
      remainder: value - rounded[index],
      value: nums[index]
    }))
    .sort((a, b) => (b.remainder - a.remainder) || (b.value - a.value) || (a.index - b.index))
    .forEach((entry) => {
      if (remaining <= 0) return;
      rounded[entry.index] += 1;
      remaining -= 1;
    });
  return rounded.map((value) => `${value}%`);
}

function yoyPct(current, previous) {
  const prev = safeNum(previous);
  if (!prev) return null;
  return ((safeNum(current) - prev) / prev) * 100;
}

function pickFirst(row, keys) {
  if (!row || typeof row !== "object") return "";
  for (const key of keys) {
    if (row[key] !== undefined && row[key] !== null && asText(row[key])) return row[key];
  }
  const lowerMap = Object.fromEntries(Object.keys(row).map((key) => [key.toLowerCase(), key]));
  for (const key of keys) {
    const match = lowerMap[String(key).toLowerCase()];
    if (match && row[match] !== undefined && row[match] !== null && asText(row[match])) return row[match];
  }
  return "";
}

function programId(row) {
  return asText(pickFirst(row, ["programId", "programID", "Program ID", "campaignId", "offerId", "id"]));
}

function programName(row, fallbackId = "") {
  return asText(
    pickFirst(row, ["programName", "Program Name", "programDisplayName", "campaignName", "offerName", "name"]),
    fallbackId ? `Program ${fallbackId}` : "Unknown Program"
  );
}

function programKey(row) {
  const id = programId(row);
  return id || programName(row).toLowerCase();
}

function rowClicks(row) {
  return safeNum(pickFirst(row, ["clicks", "Clicks"]));
}

function rowConversions(row) {
  return safeNum(pickFirst(row, ["conversions", "sales", "iSales", "Sales", "Current Sales"]));
}

function rowOrderValue(row) {
  return safeNum(pickFirst(row, ["orderValue", "salesOrderValue", "iSalesOrderValue", "eventOrderValue", "Order Value", "Current OV"]));
}

function rowCommission(row) {
  return safeNum(pickFirst(row, [
    "publisherCommission",
    "publisher_commission",
    "pubCommission",
    "publisherEarnings",
    "salesCommission",
    "iSalesCommission",
    "commission",
    "commissionAmount",
    "commissionValue",
    "Publisher Commission"
  ]));
}

function parseDateOnly(value) {
  const text = asText(value);
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  match = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (match) return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : new Date(Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate()));
}

function isoDate(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function isoWeekStartFromYearWeek(year, week) {
  const fourthJan = new Date(Date.UTC(year, 0, 4));
  const fourthJanDay = fourthJan.getUTCDay() || 7;
  const weekOneMonday = new Date(fourthJan);
  weekOneMonday.setUTCDate(fourthJan.getUTCDate() - fourthJanDay + 1);
  weekOneMonday.setUTCDate(weekOneMonday.getUTCDate() + ((week - 1) * 7));
  return weekOneMonday;
}

function rowWeekLabel(row) {
  const raw = pickFirst(row, [
    "week",
    "Week",
    "weekStart",
    "weekStartDate",
    "week_start",
    "periodStart",
    "periodStartDate",
    "date",
    "Date",
    "period",
    "Period",
    "transactionDate",
    "eventDate"
  ]);
  const text = asText(raw);
  const isoWeek = text.match(/^(\d{4})-?W(\d{1,2})$/i);
  if (isoWeek) return isoDate(isoWeekStartFromYearWeek(Number(isoWeek[1]), Number(isoWeek[2])));
  const date = parseDateOnly(text);
  if (!date) return "";
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() - day + 1);
  return isoDate(date);
}

function weeklyPublisherCommissionChart(primaryRows = [], competitorRows = [], fallbackTotals = []) {
  const seriesKeys = ["Primary", "Comp. A", "Comp. B", "Comp. C", "Comp. D"];
  const weekMap = new Map();
  const ensure = (week) => {
    if (!weekMap.has(week)) {
      weekMap.set(week, {
        Week: week,
        Primary: 0,
        "Comp. A": 0,
        "Comp. B": 0,
        "Comp. C": 0,
        "Comp. D": 0
      });
    }
    return weekMap.get(week);
  };
  const addRows = (rows, key) => {
    for (const row of rows || []) {
      const week = rowWeekLabel(row);
      if (!week) continue;
      ensure(week)[key] += rowCommission(row);
    }
  };

  addRows(primaryRows, "Primary");
  competitorRows.slice(0, 4).forEach((rows, index) => addRows(rows, seriesKeys[index + 1]));

  const weeklyRows = Array.from(weekMap.values())
    .sort((a, b) => a.Week.localeCompare(b.Week))
    .slice(-13)
    .map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, key === "Week" ? value : Math.round(safeNum(value) * 100) / 100])));

  if (weeklyRows.length) return weeklyRows;
  return [{
    Week: "Current Period",
    Primary: Math.round(safeNum(fallbackTotals[0]?.publisherCommission) * 100) / 100,
    "Comp. A": Math.round(safeNum(fallbackTotals[1]?.publisherCommission) * 100) / 100,
    "Comp. B": Math.round(safeNum(fallbackTotals[2]?.publisherCommission) * 100) / 100,
    "Comp. C": Math.round(safeNum(fallbackTotals[3]?.publisherCommission) * 100) / 100,
    "Comp. D": Math.round(safeNum(fallbackTotals[4]?.publisherCommission) * 100) / 100
  }];
}

function rowWallet(row) {
  return safeNum(pickFirst(row, [
    "digitalWallet",
    "digitalWallets",
    "digital_wallet",
    "digitalWalletAmount",
    "digitalWalletCommission",
    "walletAmount",
    "walletCommission",
    "amount",
    "publisherValue",
    "publisherAmount",
    "paymentAmount",
    "total",
    "totalAmount"
  ]));
}

function aggregatePrograms(metricRows = [], walletRows = []) {
  const programs = new Map();
  const ensure = (row) => {
    const key = programKey(row);
    if (!key) return null;
    if (!programs.has(key)) {
      const id = programId(row);
      programs.set(key, {
        key,
        programId: id || "-",
        programName: programName(row, id),
        clicks: 0,
        conversions: 0,
        orderValue: 0,
        publisherCommission: 0,
        digitalWallet: 0
      });
    }
    return programs.get(key);
  };

  for (const row of metricRows || []) {
    const program = ensure(row);
    if (!program) continue;
    program.clicks += rowClicks(row);
    program.conversions += rowConversions(row);
    program.orderValue += rowOrderValue(row);
    program.publisherCommission += rowCommission(row);
  }

  for (const row of walletRows || []) {
    const program = ensure(row);
    if (!program) continue;
    program.digitalWallet += rowWallet(row);
  }

  return Array.from(programs.values()).map((program) => ({
    ...program,
    totalEarnings: program.publisherCommission + program.digitalWallet
  }));
}

function totals(programs = []) {
  return programs.reduce((out, program) => {
    out.clicks += program.clicks;
    out.conversions += program.conversions;
    out.orderValue += program.orderValue;
    out.publisherCommission += program.publisherCommission;
    out.digitalWallet += program.digitalWallet;
    out.totalEarnings += program.totalEarnings;
    if (program.clicks || program.conversions || program.orderValue || program.totalEarnings) out.activePrograms += 1;
    if (program.publisherCommission > 0) out.programsWithCommission += 1;
    return out;
  }, {
    clicks: 0,
    conversions: 0,
    orderValue: 0,
    publisherCommission: 0,
    digitalWallet: 0,
    totalEarnings: 0,
    activePrograms: 0,
    programsWithCommission: 0
  });
}

function conversionRate(total) {
  return total.clicks ? (total.conversions / total.clicks) * 100 : 0;
}

function earningPerClick(total) {
  return total.clicks ? total.totalEarnings / total.clicks : 0;
}

function earningPerConversion(total) {
  return total.conversions ? total.totalEarnings / total.conversions : 0;
}

function metricRow(period, total, symbol) {
  return {
    Period: period,
    Conversions: fmtInt(total.conversions),
    "Conversion Rate": `${conversionRate(total).toFixed(2)}%`,
    "Conv Rate": `${conversionRate(total).toFixed(2)}%`,
    Clicks: fmtInt(total.clicks),
    "Earnings per Click": fmtMoney2(earningPerClick(total), symbol),
    "Earnings per Conversion": fmtMoney2(earningPerConversion(total), symbol),
    "Order Value": fmtMoney(total.orderValue, symbol),
    "Publisher Commission": fmtMoney(total.publisherCommission, symbol),
    "Digital Wallet": fmtMoney(total.digitalWallet, symbol),
    "Total Earnings": fmtMoney(total.totalEarnings, symbol),
    "Active Programs": fmtInt(total.activePrograms),
    "Programs w/ Commission": fmtInt(total.programsWithCommission)
  };
}

function metricVarianceRow(current, previous) {
  return {
    Period: "% Variance",
    Conversions: fmtPct(yoyPct(current.conversions, previous.conversions)),
    "Conversion Rate": fmtPct(yoyPct(conversionRate(current), conversionRate(previous))),
    "Conv Rate": fmtPct(yoyPct(conversionRate(current), conversionRate(previous))),
    Clicks: fmtPct(yoyPct(current.clicks, previous.clicks)),
    "Earnings per Click": fmtPct(yoyPct(earningPerClick(current), earningPerClick(previous))),
    "Earnings per Conversion": fmtPct(yoyPct(earningPerConversion(current), earningPerConversion(previous))),
    "Order Value": fmtPct(yoyPct(current.orderValue, previous.orderValue)),
    "Publisher Commission": fmtPct(yoyPct(current.publisherCommission, previous.publisherCommission)),
    "Digital Wallet": fmtPct(yoyPct(current.digitalWallet, previous.digitalWallet)),
    "Total Earnings": fmtPct(yoyPct(current.totalEarnings, previous.totalEarnings)),
    "Active Programs": fmtPct(yoyPct(current.activePrograms, previous.activePrograms)),
    "Programs w/ Commission": fmtPct(yoyPct(current.programsWithCommission, previous.programsWithCommission))
  };
}

function diffTotals(current, previous) {
  const out = {};
  for (const key of Object.keys(current)) out[key] = safeNum(current[key]) - safeNum(previous[key]);
  return out;
}

function varianceTotals(current, previous) {
  const out = {};
  for (const key of Object.keys(current)) out[key] = yoyPct(current[key], previous[key]);
  return out;
}

function programDisplayRow(program, previous, symbol) {
  return {
    "Program ID": program.programId || "-",
    "Program Name": program.programName || "-",
    "Publisher Commission": fmtMoney(program.publisherCommission, symbol),
    "Digital Wallet": fmtMoney(program.digitalWallet, symbol),
    "Total Earnings": fmtMoney(program.totalEarnings, symbol),
    Conversions: fmtInt(program.conversions),
    "Order Value": fmtMoney(program.orderValue, symbol),
    "Publisher Commission YoY %": fmtPct(yoyPct(program.publisherCommission, previous?.publisherCommission)),
    "Earnings YoY %": fmtPct(yoyPct(program.totalEarnings, previous?.totalEarnings)),
    "Conversions YoY %": fmtPct(yoyPct(program.conversions, previous?.conversions))
  };
}

function statusLabel(row) {
  const statusId = asText(pickFirst(row, ["statusId", "statusID", "Status ID"]));
  const fallback = asText(pickFirst(row, ["connectionStatus", "statusName", "status", "Status Name"]));
  const labels = {
    0: "Not Connected",
    1: "Under Consideration",
    2: "Hold Under Consideration",
    3: "Accepted",
    4: "Ended",
    5: "Denied",
    6: "Hold Accepted",
    9: "Ending"
  };
  return labels[statusId] || fallback || "Unknown";
}

function isoDateOnly(value) {
  const text = asText(value);
  if (!text) return "";
  const isoMatch = text.match(/\d{4}-\d{2}-\d{2}/);
  if (isoMatch) return isoMatch[0];
  const compact = text.replace(/[^0-9]/g, "");
  if (/^\d{8}$/.test(compact)) return `${compact.slice(0, 4)}-${compact.slice(4, 6)}-${compact.slice(6, 8)}`;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString().slice(0, 10);
}

function filterProgramStatusRowsByCutoff(rows = [], cutoffDate) {
  const cutoff = isoDateOnly(cutoffDate);
  if (!cutoff) return rows || [];
  return (rows || []).filter((row) => {
    const createdDate = isoDateOnly(pickFirst(row, ["createdDate", "Created Date", "created", "createdAt"]));
    return Boolean(createdDate && createdDate >= cutoff);
  });
}

function statusKeyFromParts(id, name) {
  const programIdValue = asText(id);
  if (programIdValue && programIdValue !== "-") return programIdValue;
  return asText(name).toLowerCase();
}

function normalizeGapStatus(status) {
  const text = asText(status);
  if (!text) return "No Connection";
  if (/^(not connected|not connected yet|none|unknown)$/i.test(text)) return "No Connection";
  if (/^hold uc$/i.test(text)) return "Hold Under Consideration";
  if (/^hold accepted$/i.test(text)) return "Hold Accepted";
  return text.replace(/^under consideration$/i, "Under Consideration");
}

function buildProgramStatusMap(rows = [], cutoffDate) {
  const map = new Map();
  for (const row of filterProgramStatusRowsByCutoff(rows, cutoffDate)) {
    const id = programId(row) || asText(pickFirst(row, ["id"]));
    const name = programName(row, id);
    const key = statusKeyFromParts(id, name);
    if (!key) continue;
    map.set(key, normalizeGapStatus(statusLabel(row)));
  }
  return map;
}

function deriveGapStatus(program, statusMap) {
  if (program && safeNum(program.publisherCommission) > 0) return "Pub Comm";
  if (program && safeNum(program.clicks) > 0) return "Clicks";
  const key = program ? statusKeyFromParts(program.programId, program.programName) : "";
  return normalizeGapStatus((key && statusMap.get(key)) || "No Connection");
}

function classifyGap(primaryStatus) {
  const normalized = normalizeGapStatus(primaryStatus);
  if (normalized === "Accepted") return { type: "Activation", action: "Activate" };
  if (normalized === "Clicks") return { type: "Click Leakage", action: "Fix tracking / conversion" };
  if (normalized === "No Connection") return { type: "Application", action: "Apply" };
  if (/^(Denied|Ended|Ending|Hold Accepted|On Hold)$/i.test(normalized)) return { type: "Recovery", action: "Recover" };
  if (/under consideration/i.test(normalized)) return { type: "Pending", action: "Monitor" };
  return { type: "Review", action: "Review" };
}

function programStatusTable(rows = [], cutoffDate) {
  const statusOrder = {
    Accepted: 1,
    "Hold Accepted": 2,
    "Under Consideration": 3,
    "Hold Under Consideration": 4,
    "Not Connected": 5,
    Ending: 6,
    Ended: 7,
    Denied: 8
  };
  return filterProgramStatusRowsByCutoff(rows, cutoffDate).map((row) => {
    const id = programId(row) || asText(pickFirst(row, ["id"]));
    return {
      _createdDate: isoDateOnly(pickFirst(row, ["createdDate", "Created Date", "created", "createdAt"])),
      "Program ID": id || "-",
      "Program Name": programName(row, id),
      "Connection Status": statusLabel(row)
    };
  })
    .sort((a, b) => {
      const byStatus = (statusOrder[a["Connection Status"]] || 99) - (statusOrder[b["Connection Status"]] || 99);
      if (byStatus) return byStatus;
      const byDate = String(b._createdDate || "").localeCompare(String(a._createdDate || ""));
      if (byDate) return byDate;
      return String(a["Program Name"] || "").localeCompare(String(b["Program Name"] || ""));
    })
    .map(({ _createdDate, ...row }) => row);
}

function makeTables(payload, fetched) {
  const symbol = currencySymbol(payload.currencyCode);
  const currentPrograms = aggregatePrograms(fetched.publisherCurrent, fetched.walletCurrent);
  const previousPrograms = aggregatePrograms(fetched.publisherPrevious, fetched.walletPrevious);
  const previousByKey = new Map(previousPrograms.map((program) => [program.key, program]));
  const currentTotal = totals(currentPrograms);
  const previousTotal = totals(previousPrograms);
  const difference = diffTotals(currentTotal, previousTotal);
  const variance = varianceTotals(currentTotal, previousTotal);
  const programLevel = currentPrograms
    .slice()
    .sort((a, b) => (b.publisherCommission - a.publisherCommission) || (b.totalEarnings - a.totalEarnings) || (b.orderValue - a.orderValue))
    .map((program) => programDisplayRow(program, previousByKey.get(program.key), symbol));
  const movers = currentPrograms
    .map((program) => {
      const previous = previousByKey.get(program.key) || {};
      const diff = program.publisherCommission - safeNum(previous.publisherCommission);
      return {
        Program: program.programName || "-",
        "Program Name": program.programName || "-",
        "Program ID": program.programId || "-",
        "Chart Label": `${program.programName || "-"} (${fmtPct(yoyPct(program.publisherCommission, previous.publisherCommission))})`,
        "Current Commission": fmtMoney(program.publisherCommission, symbol),
        "YoY Change": fmtMoney(diff, symbol),
        "YoY %": fmtPct(yoyPct(program.publisherCommission, previous.publisherCommission)),
        Direction: diff > 0 ? "Up" : diff < 0 ? "Down" : "Neutral",
        "Publisher Commission Change": fmtMoney(diff, symbol),
        "Publisher Commission Change Value": Math.round(diff * 100) / 100,
        "Current Publisher Commission": fmtMoney(program.publisherCommission, symbol),
        "Previous Publisher Commission": fmtMoney(previous.publisherCommission, symbol),
        "Publisher Commission YoY %": fmtPct(yoyPct(program.publisherCommission, previous.publisherCommission)),
        _diff: diff
      };
    })
    .filter((row) => row._diff !== 0);
  const moverRows = [
    ...movers.filter((row) => row._diff > 0).sort((a, b) => b._diff - a._diff).slice(0, 10),
    ...movers.filter((row) => row._diff < 0).sort((a, b) => a._diff - b._diff).slice(0, 10)
  ].map(({ _diff, ...row }) => row);
  const brandNewPrograms = currentPrograms
    .filter((program) => !previousByKey.has(program.key) || safeNum(previousByKey.get(program.key).totalEarnings) <= 0)
    .sort((a, b) => b.totalEarnings - a.totalEarnings)
    .map((program) => ({
      "Program ID": program.programId || "-",
      "Program Name": program.programName || "-",
      Conversions: fmtInt(program.conversions),
      "Order Value": fmtMoney(program.orderValue, symbol),
      "Publisher Commission": fmtMoney(program.publisherCommission, symbol),
      "Digital Wallet": fmtMoney(program.digitalWallet, symbol),
      "Total Earnings": fmtMoney(program.totalEarnings, symbol)
    }));
  const statusCutoff = payload.programStatusCreatedFromDate || payload.programStatusCreatedFromDateRaw;
  const statusRows = programStatusTable(fetched.publisherProgramStatus, statusCutoff);
  const competitorGroups = (fetched.competitors || []).map((item, index) => {
    const programs = aggregatePrograms(item.current || [], []);
    const previousProgramsForCompetitor = aggregatePrograms(item.previous || [], []);
    const total = totals(programs);
    const previousTotalForCompetitor = totals(previousProgramsForCompetitor);
    return {
      label: item.label || `Publisher ${index + 1}`,
      programs,
      previousPrograms: previousProgramsForCompetitor,
      total,
      previousTotal: previousTotalForCompetitor,
      statusRows: item.status || [],
      byKey: new Map(programs.map((program) => [program.key, program]))
    };
  });
  const primaryShareTotal = currentTotal.publisherCommission + competitorGroups.reduce((sum, group) => sum + group.total.publisherCommission, 0);
  const previousShareTotal = previousTotal.publisherCommission + competitorGroups.reduce((sum, group) => sum + group.previousTotal.publisherCommission, 0);
  const shareRow = (label, total, priorTotal) => ({
    Publisher: label,
    "Publisher Commission PP": fmtMoney(priorTotal.publisherCommission, symbol),
    "Publisher Commission": fmtMoney(total.publisherCommission, symbol),
    "Share PP": previousShareTotal ? `${((priorTotal.publisherCommission / previousShareTotal) * 100).toFixed(1)}%` : "0.0%",
    Share: primaryShareTotal ? `${((total.publisherCommission / primaryShareTotal) * 100).toFixed(1)}%` : "0.0%"
  });
  const competitorShare = [
    shareRow("Primary", currentTotal, previousTotal),
    ...competitorGroups.map((group) => shareRow(group.label, group.total, group.previousTotal))
  ];
  const competitorSharePubCommChart = (() => {
    const labels = ["Your Site", "Comp. A", "Comp. B", "Comp. C", "Comp. D"];
    const currentValues = [currentTotal.publisherCommission, ...competitorGroups.slice(0, 4).map((group) => group.total.publisherCommission)];
    const previousValues = [previousTotal.publisherCommission, ...competitorGroups.slice(0, 4).map((group) => group.previousTotal.publisherCommission)];
    while (currentValues.length < labels.length) currentValues.push(0);
    while (previousValues.length < labels.length) previousValues.push(0);
    const makeRow = (metric, values) => {
      const row = { "Competitor Group Summary": metric };
      labels.forEach((label, index) => {
        row[label] = fmtMoney(values[index] || 0, symbol);
      });
      return row;
    };
    return [
      makeRow("Publisher Commission PP", previousValues),
      makeRow("Publisher Commission", currentValues)
    ];
  })();
  const competitorWeeklyPubComm = weeklyPublisherCommissionChart(
    fetched.publisherCurrent,
    (fetched.competitors || []).map((item) => item.current || []),
    [currentTotal, ...competitorGroups.map((group) => group.total)]
  );
  const allKeys = Array.from(new Set([
    ...currentPrograms.map((program) => program.key),
    ...competitorGroups.flatMap((group) => group.programs.map((program) => program.key))
  ]));
  const programNameByKey = new Map(currentPrograms.map((program) => [program.key, program]));
  competitorGroups.forEach((group) => group.programs.forEach((program) => {
    if (!programNameByKey.has(program.key)) programNameByKey.set(program.key, program);
  }));
  const topCompetitorPerformance = allKeys
    .map((key) => {
      const ref = programNameByKey.get(key) || {};
      const primary = currentPrograms.find((program) => program.key === key);
      const values = [
        safeNum(primary?.publisherCommission),
        ...competitorGroups.slice(0, 4).map((group) => safeNum(group.byKey.get(key)?.publisherCommission))
      ];
      while (values.length < 5) values.push(0);
      const shares = formatSharePercentages(values);
      return {
        _primaryCommission: safeNum(primary?.publisherCommission),
        _totalCommission: values.reduce((sum, value) => sum + value, 0),
        "Program Name": ref.programName || "-",
        [asText(payload.clientName || payload.programName || payload.clientUsername, "Primary Publisher")]: shares[0],
        "Comp. A": shares[1],
        "Comp. B": shares[2],
        "Comp. C": shares[3],
        "Comp. D": shares[4]
      };
    })
    .filter((row) => row._totalCommission > 0)
    .sort((a, b) => (b._primaryCommission - a._primaryCommission) || (b._totalCommission - a._totalCommission) || String(a["Program Name"]).localeCompare(String(b["Program Name"])))
    .map(({ _primaryCommission, _totalCommission, ...row }) => row);
  const primaryStatusByKey = buildProgramStatusMap(fetched.publisherProgramStatus, statusCutoff);
  const competitorStatusByKey = competitorGroups.slice(0, 4).map((group) => buildProgramStatusMap(group.statusRows || [], statusCutoff));
  const primaryByKey = new Map(currentPrograms.map((program) => [program.key, program]));
  const competitorByKey = competitorGroups.slice(0, 4).map((group) => new Map(group.programs.map((program) => [program.key, program])));
  const programsByKey = new Map();
  const rememberProgram = (key, id, name) => {
    const resolvedKey = asText(key) || statusKeyFromParts(id, name);
    if (!resolvedKey || programsByKey.has(resolvedKey)) return;
    programsByKey.set(resolvedKey, {
      key: resolvedKey,
      programId: asText(id, "-"),
      programName: asText(name || id || resolvedKey, resolvedKey)
    });
  };
  currentPrograms.forEach((program) => rememberProgram(program.key, program.programId, program.programName));
  competitorGroups.slice(0, 4).forEach((group) => group.programs.forEach((program) => rememberProgram(program.key, program.programId, program.programName)));
  filterProgramStatusRowsByCutoff(fetched.publisherProgramStatus, statusCutoff).forEach((row) => {
    const id = programId(row) || asText(pickFirst(row, ["id"]));
    rememberProgram(statusKeyFromParts(id, programName(row, id)), id, programName(row, id));
  });
  competitorGroups.slice(0, 4).forEach((group) => filterProgramStatusRowsByCutoff(group.statusRows || [], statusCutoff).forEach((row) => {
    const id = programId(row) || asText(pickFirst(row, ["id"]));
    rememberProgram(statusKeyFromParts(id, programName(row, id)), id, programName(row, id));
  }));
  const gapRows = Array.from(programsByKey.values())
    .map((programRef) => {
      const primaryProgram = primaryByKey.get(programRef.key);
      const primaryStatus = deriveGapStatus(primaryProgram || programRef, primaryStatusByKey);
      const competitorStatuses = competitorByKey.map((map, index) => deriveGapStatus(map.get(programRef.key), competitorStatusByKey[index]));
      const competitorValue = competitorByKey.reduce((sum, map) => sum + safeNum(map.get(programRef.key)?.publisherCommission), 0);
      const gap = classifyGap(primaryStatus);
      const competitorCount = competitorStatuses.filter((status) => status === "Pub Comm").length;
      return {
        "Program Name": programRef.programName,
        "Program ID": programRef.programId || "-",
        Primary: primaryStatus,
        "Primary status": primaryStatus,
        "Primary Status": primaryStatus,
        "Comp. A": competitorStatuses[0] || "No Connection",
        "Comp. B": competitorStatuses[1] || "No Connection",
        "Comp. C": competitorStatuses[2] || "No Connection",
        "Comp. D": competitorStatuses[3] || "No Connection",
        "Competitor signal": `${competitorCount}/4 competitors earning`,
        "Competitor Signal": `${competitorCount}/4 competitors earning`,
        "Competitor Pub Comm": fmtMoney(competitorValue, symbol),
        "Gap Type": gap.type,
        "Recommended Action": gap.action,
        _competitorValue: competitorValue
      };
    })
    .filter((row) => row["Primary status"] !== "Pub Comm")
    .filter((row) => row._competitorValue > 0 || row["Primary status"] === "Clicks" || row["Primary status"] !== "No Connection")
    .sort((a, b) => {
      if (b._competitorValue !== a._competitorValue) return b._competitorValue - a._competitorValue;
      const statusPriority = { Accepted: 1, Clicks: 2, "No Connection": 3, Denied: 4, Ended: 5, "Hold Accepted": 6 };
      return (statusPriority[a["Primary status"]] || 99) - (statusPriority[b["Primary status"]] || 99)
        || String(a["Program Name"]).localeCompare(String(b["Program Name"]));
    });
  const gaps = gapRows.map((row, index) => ({
    Rank: String(index + 1),
    Program: row["Program Name"],
    "Program ID": row["Program ID"],
    "Program Name": row["Program Name"],
    "Program / ID": `${row["Program Name"]} / ${row["Program ID"]}`,
    "Primary status": row["Primary status"],
    "Primary Status": row["Primary Status"],
    "Competitor signal": row["Competitor signal"],
    "Competitor Signal": row["Competitor Signal"],
    "Comp. A": row["Comp. A"],
    "Comp. B": row["Comp. B"],
    "Comp. C": row["Comp. C"],
    "Comp. D": row["Comp. D"],
    "Pub Comm - Specified Sites Value": row["Competitor Pub Comm"],
    "Pub Comm - Specified Sites": row["Competitor Pub Comm"],
    Value: row["Competitor Pub Comm"],
    "Competitor Pub Comm": row["Competitor Pub Comm"],
    "Gap Type": row["Gap Type"],
    "Recommended Action": row["Recommended Action"],
    _competitorValue: row._competitorValue
  }));
  const gapSummaryByType = Array.from(gapRows.reduce((map, row) => {
    const type = row["Gap Type"] || "Review";
    const existing = map.get(type) || { count: 0, commission: 0 };
    existing.count += 1;
    existing.commission += safeNum(row["Competitor Pub Comm"]);
    map.set(type, existing);
    return map;
  }, new Map()).entries()).map(([gapType, value]) => ({
    "Gap Type": gapType,
    Programs: fmtInt(value.count),
    "Competitor Pub Comm": fmtMoney(value.commission, symbol)
  }));
  const gapSummaryByTypeMap = new Map(gapSummaryByType.map((row) => [row["Gap Type"], row]));
  const activationStatusCount = statusRows.filter((row) => row["Connection Status"] === "Accepted").length;
  const activationGapCount = safeNum(gapSummaryByTypeMap.get("Activation")?.Programs);
  const gapProgramsCount = gapRows.length + Math.max(0, activationStatusCount - activationGapCount);
  const activationValue = safeNum(gapSummaryByTypeMap.get("Activation")?.["Competitor Pub Comm"]);
  const competitorOpportunityValue = gapRows.reduce((sum, row) => sum + safeNum(row["Competitor Pub Comm"]), 0);
  const programGapAnalysisSummaryCards = [
    {
      Metric: "Gap programs",
      Value: fmtInt(gapProgramsCount),
      Detail: "Programs requiring activation, application or recovery."
    },
    {
      Metric: "Activation opportunities",
      Value: fmtInt(activationStatusCount || activationGapCount),
      Detail: `${fmtMoney(activationValue, symbol)} already accepted but inactive`
    },
    {
      Metric: "Competitor pub comm opportunity",
      Value: fmtMoneyCompact(competitorOpportunityValue, symbol),
      Detail: `${fmtInt(gapProgramsCount)} programs where the primary publisher is not earning.`
    },
    {
      Metric: "Largest visible gap",
      Value: gaps[0]?.Value || fmtMoney(0, symbol),
      Detail: gaps[0]?.Program ? `Largest visible gap is ${gaps[0].Program}.` : "No visible competitor-funded gap."
    }
  ];
  const riskDependencies = programLevel.slice(0, 5).map((row) => ({
    Issue: "Commission concentration",
    Analysis: `${row["Program Name"]} contributes ${row["Publisher Commission"]} publisher commission in the current period.`,
    Action: "Monitor concentration in the next QBR cycle."
  }));
  const summaryRows = [
    ["Conversions", currentTotal.conversions, previousTotal.conversions, false],
    ["Total Order Value", currentTotal.orderValue, previousTotal.orderValue, true],
    ["Publisher Commission", currentTotal.publisherCommission, previousTotal.publisherCommission, true],
    ["Digital Wallet", currentTotal.digitalWallet, previousTotal.digitalWallet, true],
    ["Total Earnings", currentTotal.totalEarnings, previousTotal.totalEarnings, true],
    ["Conversion Rate", conversionRate(currentTotal), conversionRate(previousTotal), false]
  ].map(([metric, current, previous, money]) => ({
    Metric: metric,
    Recent: money ? fmtMoney(current, symbol) : (metric === "Conversion Rate" ? `${current.toFixed(2)}%` : fmtInt(current)),
    Previous: money ? fmtMoney(previous, symbol) : (metric === "Conversion Rate" ? `${previous.toFixed(2)}%` : fmtInt(previous)),
    Difference: money ? fmtMoney(safeNum(current) - safeNum(previous), symbol) : (metric === "Conversion Rate" ? `${(safeNum(current) - safeNum(previous)).toFixed(2)}pp` : fmtInt(safeNum(current) - safeNum(previous))),
    "% Variance": fmtPct(yoyPct(current, previous))
  }));

  return {
    publisherPerformanceSummaryTable: summaryRows,
    kpiSummaryTable: [
      metricRow("Recent", currentTotal, symbol),
      metricRow("Previous", previousTotal, symbol),
      metricRow("Difference", difference, symbol),
      metricVarianceRow(currentTotal, previousTotal)
    ],
    kpiVarianceColorHintsTable: [
      { Metric: "Conversions", Tone: variance.conversions > 0 ? "positive" : variance.conversions < 0 ? "negative" : "neutral" },
      { Metric: "Publisher Commission", Tone: variance.publisherCommission > 0 ? "positive" : variance.publisherCommission < 0 ? "negative" : "neutral" },
      { Metric: "Digital Wallet", Tone: variance.digitalWallet > 0 ? "positive" : variance.digitalWallet < 0 ? "negative" : "neutral" },
      { Metric: "Total Earnings", Tone: variance.totalEarnings > 0 ? "positive" : variance.totalEarnings < 0 ? "negative" : "neutral" }
    ],
    programActivationSnapshotTable: [
      { Metric: "Joined programs", Total: fmtInt(currentPrograms.length), New: fmtInt(brandNewPrograms.length), "New %": currentPrograms.length ? `${((brandNewPrograms.length / currentPrograms.length) * 100).toFixed(1)}%` : "0.0%" },
      { Metric: "With clicks", Total: fmtInt(currentPrograms.filter((program) => program.clicks > 0).length), New: fmtInt(brandNewPrograms.filter((row) => safeNum(row.Conversions) > 0 || safeNum(row["Order Value"]) > 0).length), "New %": currentTotal.activePrograms ? `${((brandNewPrograms.length / currentTotal.activePrograms) * 100).toFixed(1)}%` : "0.0%" },
      { Metric: "Pub commission", Total: fmtInt(currentTotal.programsWithCommission), New: fmtInt(brandNewPrograms.filter((row) => safeNum(row["Publisher Commission"]) > 0).length), "New %": currentTotal.programsWithCommission ? `${((brandNewPrograms.filter((row) => safeNum(row["Publisher Commission"]) > 0).length / currentTotal.programsWithCommission) * 100).toFixed(1)}%` : "0.0%" },
      { Metric: "Inactive", Total: fmtInt(Math.max(currentPrograms.length - currentTotal.activePrograms, 0)), New: "0", "New %": "0.0%" }
    ],
    programLevelBreakdown: programLevel.slice(0, 12),
    allProgramsApiScope: programLevel.slice(0, 120),
    moversShakersCommissionChart: moverRows,
    brandNewProgramsTable: (brandNewPrograms.length ? brandNewPrograms : programLevel).slice(0, 10),
    programConnectionStatusTable: statusRows,
    programGapAnalysisTable: gaps.map(({ _competitorValue, ...row }) => row),
    programGapAnalysisSummaryTable: programGapAnalysisSummaryCards,
    programGapAnalysisByTypeTable: gapSummaryByType,
    competitorAnalysisTable: competitorShare,
    competitorWeeklyPubCommChart: competitorWeeklyPubComm,
    competitorSharePubCommChart: competitorSharePubCommChart,
    topProgramsCompetitorPerformanceTable: topCompetitorPerformance.slice(0, 10),
    kpiHighlightsTable: summaryRows.map((row) => ({
      KPI: row.Metric,
      Recent: row.Recent,
      Previous: row.Previous,
      "YoY %": row["% Variance"],
      Highlight: `${row.Metric}: ${row.Recent} vs ${row.Previous} (${row["% Variance"]}).`
    })),
    kpiHighlightNarrativeTable: [],
    riskDependenciesTable: riskDependencies,
    reportingPeriodTable: [
      { Window: "Current Period", Period: payload.reportingPeriod, "Data as of": payload.reportingPeriod.split(" to ").pop() || payload.reportingPeriod },
      { Window: "Comparison Period (YoY)", Period: payload.comparisonPeriod, Basis: "Year-over-Year (YoY)" }
    ]
  };
}

function buildDataForAI(payload, tables, diagnostics) {
  return JSON.stringify({
    payload: {
      clientUsername: payload.clientUsername,
      reportingPeriod: payload.reportingPeriod,
      comparisonPeriod: payload.comparisonPeriod,
      languageName: payload.languageName,
      currencyCode: payload.currencyCode,
      qbrFocus: payload.qbrFocus
    },
    tables,
    diagnostics
  }, null, 2);
}

function normalizePublisherPptxUrl(payload) {
  const candidate = asText(
    payload.publisherPptxApiUrl
    || payload.publisher_pptx_api_url
    || payload.publisherPresentonApiUrl
    || payload.publisher_presenton_api_url
    || payload.presentonApiUrl
    || payload.presenton_api_url
    || process.env.QBR_PPTX_SERVICE_URL
    || process.env.PUBLISHER_QBR_PPTX_SERVICE_URL
    || DEFAULT_PPTX_SERVICE_URL
  );
  if (/:3011(?:\D|$)/.test(candidate)) return DEFAULT_PPTX_SERVICE_URL;
  return candidate.replace(/\/+$/, "");
}

function buildFinalPptxPayload(payload, publisherAnalysis, tables) {
  const client = asText(payload.clientName || payload.programName || payload.clientUsername, "Publisher QBR");
  const slideBlueprint = [
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
  ];
  const publisherPptxApiUrl = normalizePublisherPptxUrl(payload);
  const slideTableBindings = {
    reporting_period: "reportingPeriodTable",
    kpi_tiles: "publisherPerformanceSummaryTable",
    kpi_summary_table: "kpiSummaryTable + kpiVarianceColorHintsTable",
    program_activation_snapshot: "programActivationSnapshotTable",
    program_level_analysis: "programLevelBreakdown",
    movers_commission_chart: "moversShakersCommissionChart",
    brand_new_programs: "brandNewProgramsTable",
    growth_opportunity_gap: "programGapAnalysisSummaryTable",
    top_competitor_funded_gaps: "programGapAnalysisTable",
    program_connection_status_1: "programConnectionStatusTable",
    program_connection_status_2: "programConnectionStatusTable",
    competitor_analysis: "competitorAnalysisTable + competitorWeeklyPubCommChart",
    share_within_competitor_group: "competitorSharePubCommChart",
    top_programs_competitor_performance: "topProgramsCompetitorPerformanceTable",
    kpi_highlights: "kpiHighlightsTable + kpiHighlightNarrativeTable",
    risks_dependencies: "riskDependenciesTable"
  };
  const requestedSlides = Number(payload.n_slides ?? payload.targetSlides ?? payload.target_slides);
  const targetSlides = Number.isFinite(requestedSlides) && requestedSlides > 0
    ? Math.min(Math.floor(requestedSlides), slideBlueprint.length)
    : slideBlueprint.length;

  return {
    client,
    deckTitle: `QBR - ${client}`,
    targetSlides,
    templateSlideCount: slideBlueprint.length,
    slideBlueprint: slideBlueprint.slice(0, targetSlides),
    slideTableBindings,
    reportingPeriod: payload.reportingPeriod,
    comparisonPeriod: payload.comparisonPeriod,
    programStatusCreatedFromDate: payload.programStatusCreatedFromDate,
    qbrFocus: payload.qbrFocus,
    qbrFocusDetail: payload.qbrFocusDetail,
    analysisLevel: "publisher_program",
    programOutput: "",
    programYoYTable: tables.kpiSummaryTable || [],
    programScopeTable: tables.programLevelBreakdown || [],
    allProgramsApiScopeTable: tables.allProgramsApiScope || [],
    publisherAnalysis: asText(publisherAnalysis).slice(0, 6000),
    kpiHighlightNarrative: [],
    publisherInsights: [],
    publisherTables: tables,
    languageCode: payload.languageCode,
    languageName: payload.languageName,
    currencyCode: payload.currencyCode,
    templateFamily: "publisher_qbr",
    forbiddenTemplateFamily: "advertiser_qbr",
    presentonTemplateId: /advertiser/i.test(asText(payload.publisherPresentonTemplateId || payload.publisherTemplateId)) ? "" : asText(payload.publisherPresentonTemplateId || payload.publisherTemplateId),
    presentonAdditionalInstructions: [
      "Publisher QBR template only. Use the Reward Gateway-style publisher service layout, headings, tables, and tone.",
      "Ignore generic templateId, presentonTemplateId, gammaThemeId, and non-publisher template fields from the incoming request.",
      "The publisher deck must use the publisher PPTX service. Reject service URLs for port 3011.",
      `Slide 1 cover title must be "${client} performance review"; never use "affiliate program" wording.`,
      `Write all narrative text in ${payload.languageName}.`
    ].join("\n"),
    presentonExportAs: asText(payload.presentonExportAs || payload.presenton_export_as || payload.gammaExportAs || payload.gamma_export_as, "pptx"),
    publisherPptxApiUrl,
    presentonApiUrl: publisherPptxApiUrl
  };
}

function projectN8nResponse(d = {}) {
  return {
    success: Boolean(
      d.pptx_url
      || d.path
      || d.presentation_url
      || d.presentation_id
      || d.presentationId
    ) && !d.error,
    provider: d.provider || "presenton",
    message: d.error ? "Presenton generation failed" : (d.message || "Presentation generated successfully"),
    generation_status: d.status || null,
    generation_id: d.presentation_id || d.presentationId || null,
    presentation_id: d.presentation_id || d.presentationId || null,
    presentation_url: d.presentation_url || null,
    edit_url: d.edit_path || d.editPath || null,
    pptx_url: d.pptx_url || d.path || d.download_url || d.file_url || null,
    gap_analysis_report_url: d.gap_analysis_report_url || null,
    gap_analysis_report_file_name: d.gap_analysis_report_file_name || null,
    file_name: d.file_name || null,
    theme: d.template || d.theme || null,
    slide_count: d.slide_count || null,
    error: d.error || null
  };
}

async function fetchWorkflowData(normalized, options) {
  const payload = normalized.payload;
  const primaryToken = payload.td_tokens?.impersonate_access_token || normalized.tdTokens?.impersonate_access_token;
  const requestOptions = {
    fetch: options.fetch,
    logger: options.logger,
    retries: options.retries,
    timeoutMs: options.timeoutMs
  };

  const publisherCurrent = await fetchAllPages(buildMetricsUrl(payload, payload.fromDate, payload.toDate, payload.primarySourceId), {
    method: "GET",
    headers: authHeaders(primaryToken)
  }, { ...requestOptions, label: "Publisher current statistics" });
  const publisherPrevious = await fetchAllPages(buildMetricsUrl(payload, normalized.prev.fromDate, normalized.prev.toDate, payload.primarySourceId), {
    method: "GET",
    headers: authHeaders(primaryToken)
  }, { ...requestOptions, label: "Publisher previous statistics" });
  const walletCurrent = await fetchAllPages(walletUrl(payload, payload.digitalWalletCurrentFromDate, payload.digitalWalletCurrentToDate), {
    method: "GET",
    headers: authHeaders(primaryToken)
  }, { ...requestOptions, label: "Publisher current digital wallets" });
  const walletPrevious = await fetchAllPages(walletUrl(payload, payload.digitalWalletPreviousFromDate, payload.digitalWalletPreviousToDate), {
    method: "GET",
    headers: authHeaders(primaryToken)
  }, { ...requestOptions, label: "Publisher previous digital wallets" });
  const publisherProgramStatus = await fetchAllPages(programStatusUrl(payload, payload.primarySourceId), {
    method: "GET",
    headers: authHeaders(primaryToken)
  }, { ...requestOptions, label: "Publisher program status" });

  const competitors = [];
  for (const competitor of payload.competitorPublishers || []) {
    const competitorToken = competitor.td_tokens?.impersonate_access_token || competitor.tdTokens?.impersonate_access_token;
    const current = await fetchAllPages(buildMetricsUrl(payload, payload.fromDate, payload.toDate, competitor.primarySourceId, competitor.publisherExportEndpoint), {
      method: "GET",
      headers: authHeaders(competitorToken)
    }, { ...requestOptions, label: `${competitor.label} current statistics` });
    const previous = await fetchAllPages(buildMetricsUrl(payload, normalized.prev.fromDate, normalized.prev.toDate, competitor.primarySourceId, competitor.publisherExportEndpoint), {
      method: "GET",
      headers: authHeaders(competitorToken)
    }, { ...requestOptions, label: `${competitor.label} previous statistics` });
    const status = await fetchAllPages(programStatusUrl(payload, competitor.primarySourceId), {
      method: "GET",
      headers: authHeaders(competitorToken)
    }, { ...requestOptions, label: `${competitor.label} program status` });
    competitors.push({ label: competitor.label, current, previous, status });
  }

  return {
    publisherCurrent,
    publisherPrevious,
    walletCurrent,
    walletPrevious,
    publisherProgramStatus,
    competitors
  };
}

async function assertPublisherPptxService(baseUrl, options = {}) {
  const fetchImpl = options.fetch || globalThis.fetch;
  const healthUrl = `${String(baseUrl).replace(/\/+$/, "")}/health`;
  const response = await fetchImpl(healthUrl, {
    method: "GET",
    headers: { Accept: "application/json" }
  });
  const body = await parseHttpResponse(response, "Publisher PPTX service health check");
  if (!body || body.service !== "publisher-qbr-pptx-service") {
    throw new Error(`Publisher PPTX service health check failed: expected publisher-qbr-pptx-service at ${healthUrl}, got ${body?.service || "unknown service"}. Run qbr-pptx-service\\restart-server.ps1 and qbr-pptx-service\\assert-publisher-port.ps1.`);
  }
  return body;
}

async function postPptxPayload(payload, options) {
  const fetchImpl = options.fetch || globalThis.fetch;
  const apiKey = process.env.PUBLISHER_QBR_API_KEY || process.env.QBR_PPTX_API_KEY || process.env.API_KEY || DEFAULT_API_KEY;
  const baseUrl = normalizePublisherPptxUrl(payload);
  await assertPublisherPptxService(baseUrl, options);
  const url = `${baseUrl}/generate`;
  const response = await fetchImpl(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      "x-api-key": apiKey
    },
    body: JSON.stringify(payload)
  });
  return parseHttpResponse(response, "Publisher PPTX generation");
}

function debugDirectory(options = {}) {
  return options.debugDir || process.env.PUBLISHER_QBR_DEBUG_DIR || "";
}

async function writeDebugArtifact(name, data, options = {}) {
  const dir = debugDirectory(options);
  if (!dir) return null;
  await fs.mkdir(dir, { recursive: true });
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filePath = path.join(dir, `${timestamp}-${name}.json`);
  await fs.writeFile(filePath, JSON.stringify(redactSensitive(data), null, 2));
  return filePath;
}

async function runPublisherQbrWorkflow(input = {}, options = {}) {
  const logger = options.logger || console;
  const normalized = normalizeWorkflowInput(input, options);
  logger.info?.("publisher_qbr_workflow_start", redactSensitive({
    clientUsername: normalized.payload.clientUsername,
    sourceID: normalized.payload.primarySourceId,
    competitorPublisherCount: normalized.payload.competitorPublisherCount,
    td_tokens: normalized.payload.td_tokens
  }));

  const fetched = await fetchWorkflowData(normalized, options);
  const tables = makeTables(normalized.payload, fetched);
  const diagnostics = {
    currentRows: fetched.publisherCurrent.length,
    previousRows: fetched.publisherPrevious.length,
    competitorPublishers: fetched.competitors.length,
    reportingPeriod: normalized.payload.reportingPeriod,
    comparisonPeriod: normalized.payload.comparisonPeriod,
    currencyCode: normalized.payload.currencyCode
  };
  const dataForAI = buildDataForAI(normalized.payload, tables, diagnostics);
  const agent = options.agent || createPublisherQbrAgent({
    modelClient: options.modelClient,
    tools: options.tools || {},
    maxIterations: options.maxIterations,
    logger,
    fetch: options.fetch
  });
  const agentOutput = await agent.run({ dataForAI, payload: normalized.payload, tables, diagnostics });
  const pptxPayload = buildFinalPptxPayload(normalized.payload, agentOutput.markdown, tables);
  await writeDebugArtifact("publisher-qbr-run", {
    diagnostics,
    tableRowCounts: Object.fromEntries(Object.entries(tables).map(([key, value]) => [key, Array.isArray(value) ? value.length : null])),
    normalizedPayload: normalized.payload,
    finalPptxPayload: pptxPayload,
    agentOutput: {
      deterministic: Boolean(agentOutput.deterministic),
      fallbackReason: agentOutput.fallbackReason || null,
      markdownPreview: asText(agentOutput.markdown).slice(0, 1000)
    }
  }, options);
  const pptxResponse = await postPptxPayload(pptxPayload, options);
  return projectN8nResponse(pptxResponse);
}

module.exports = {
  DEFAULT_PPTX_SERVICE_URL,
  extractIncomingPayload,
  normalizeWorkflowInput,
  buildMetricsUrl,
  fetchJsonWithRetry,
  fetchAllPages,
  buildFinalPptxPayload,
  assertPublisherPptxService,
  writeDebugArtifact,
  projectN8nResponse,
  runPublisherQbrWorkflow
};



