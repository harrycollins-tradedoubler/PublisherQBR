const DEFAULTS = {
  oauthUrl: "https://connect.tradedoubler.com/uaa/oauth/token",
  impersonateUrl: "https://connect.tradedoubler.com/uaa/admin/impersonate?username=",
  oauthBasic: "dGRjb25uZWN0X3B1Ymxpc2hlcjoxMjM0NTY=",
  qbrWebhookUrl: "http://127.0.0.1:3020/webhook-local/publisher-qbr-v5-competitor-weekly-chart-20260505"
};

const STORAGE_KEYS = {
  connectionConfig: "publisherQbrConnectionConfig"
};

const form = document.getElementById("publisherForm");
const saveConnectionButton = document.getElementById("saveConnection");
const submitButton = document.getElementById("submitRequest");
const adminUsernameInput = document.getElementById("adminUsername");
const adminPasswordInput = document.getElementById("adminPassword");
const qbrWebhookUrlInput = document.getElementById("qbrWebhookUrl");
const oauthUrlInput = document.getElementById("oauthUrl");
const impersonateUrlInput = document.getElementById("impersonateUrl");
const oauthBasicInput = document.getElementById("oauthBasic");
const accessTokenInput = document.getElementById("accessToken");
const clientUsernameInput = document.getElementById("clientUsername");
const sourceIDInput = document.getElementById("sourceID");
const impersonatePrimaryButton = document.getElementById("impersonatePrimary");
const clearSessionButton = document.getElementById("clearSession");
const sessionStatus = document.getElementById("sessionStatus");
const reportingPeriodInput = document.getElementById("reportingPeriod");
const startDateInput = document.getElementById("startDate");
const endDateInput = document.getElementById("endDate");
const summaryList = document.getElementById("summaryList");
const statusOutput = document.getElementById("statusOutput");
const resultLink = document.getElementById("resultLink");

let impersonatedPrimaryUsername = "";

function normalizeBaseUrl(value, fallback = "") {
  const trimmed = String(value || "").trim();
  if (!trimmed) return fallback;

  const withProtocol = /^[a-z][a-z\d+\-.]*:\/\//i.test(trimmed) ? trimmed : `http://${trimmed}`;
  return withProtocol.replace(/\/+$/, "");
}

function normalizeDate(date) {
  return String(date || "").replaceAll("-", "");
}

function daysBetween(start, end) {
  const startDate = new Date(`${start}T00:00:00`);
  const endDate = new Date(`${end}T00:00:00`);
  return Math.floor((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));
}

function formatDateForInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function startOfQuarter(date) {
  const month = date.getMonth();
  return new Date(date.getFullYear(), Math.floor(month / 3) * 3, 1);
}

function endOfQuarter(date) {
  const start = startOfQuarter(date);
  return new Date(start.getFullYear(), start.getMonth() + 3, 0);
}

function calculateRange(period) {
  const today = new Date();
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (period === "last_30_days") {
    const start = new Date(end);
    start.setDate(start.getDate() - 29);
    return { start, end };
  }

  if (period === "last_90_days") {
    const start = new Date(end);
    start.setDate(start.getDate() - 89);
    return { start, end };
  }

  if (period === "this_quarter") {
    return { start: startOfQuarter(end), end };
  }

  if (period === "last_quarter") {
    const thisQuarterStart = startOfQuarter(end);
    const lastQuarterEnd = new Date(thisQuarterStart.getFullYear(), thisQuarterStart.getMonth(), 0);
    const lastQuarterStart = startOfQuarter(lastQuarterEnd);
    return { start: lastQuarterStart, end: endOfQuarter(lastQuarterStart) };
  }

  return null;
}

function applyPresetRange() {
  const range = calculateRange(reportingPeriodInput.value);
  const isCustom = reportingPeriodInput.value === "custom";
  startDateInput.disabled = !isCustom;
  endDateInput.disabled = !isCustom;

  if (!isCustom && range) {
    startDateInput.value = formatDateForInput(range.start);
    endDateInput.value = formatDateForInput(range.end);
  }
}

function getStorageValue(key) {
  return new Promise((resolve) => {
    chrome.storage.local.get([key], (items) => resolve(items[key]));
  });
}

function setStorageValue(key, value) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ [key]: value }, resolve);
  });
}

function getConnectionConfig() {
  return {
    adminUsername: adminUsernameInput.value.trim(),
    adminPassword: adminPasswordInput.value,
    oauthUrl: oauthUrlInput.value.trim() || DEFAULTS.oauthUrl,
    impersonateUrl: impersonateUrlInput.value.trim() || DEFAULTS.impersonateUrl,
    oauthBasic: oauthBasicInput.value.trim() || DEFAULTS.oauthBasic,
    qbrWebhookUrl: normalizeBaseUrl(qbrWebhookUrlInput.value.trim(), DEFAULTS.qbrWebhookUrl)
  };
}

async function saveConnectionConfig() {
  const cfg = getConnectionConfig();
  await setStorageValue(STORAGE_KEYS.connectionConfig, {
    adminUsername: cfg.adminUsername,
    oauthUrl: cfg.oauthUrl,
    impersonateUrl: cfg.impersonateUrl,
    oauthBasic: cfg.oauthBasic,
    qbrWebhookUrl: cfg.qbrWebhookUrl
  });
  return cfg;
}

function applyConnectionConfig(cfg = {}) {
  adminUsernameInput.value = cfg.adminUsername || "";
  adminPasswordInput.value = "";
  oauthUrlInput.value = cfg.oauthUrl || DEFAULTS.oauthUrl;
  impersonateUrlInput.value = cfg.impersonateUrl || DEFAULTS.impersonateUrl;
  oauthBasicInput.value = cfg.oauthBasic || DEFAULTS.oauthBasic;
  qbrWebhookUrlInput.value = normalizeBaseUrl(
    cfg.qbrWebhookUrl || cfg.webhookUrl || cfg.qbrApiBaseUrl || cfg.apiBaseUrl,
    DEFAULTS.qbrWebhookUrl
  );
}

function sendRuntimeMessage(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (result) => {
      if (chrome.runtime.lastError) {
        const rawMessage = chrome.runtime.lastError.message || "Chrome runtime message failed.";
        const reloadHint = /message port closed|receiving end does not exist/i.test(rawMessage)
          ? " Reload the unpacked extension in chrome://extensions, then reopen the extension page."
          : "";
        reject(new Error(`${rawMessage}${reloadHint}`));
        return;
      }
      resolve(result);
    });
  });
}

async function sendExtensionRequest(message) {
  const result = await sendRuntimeMessage({
    ...message,
    cfg: getConnectionConfig()
  });

  if (!result || !result.ok) {
    throw new Error(result?.error || "The extension service worker did not return a response.");
  }

  return result.data;
}

function redact(value) {
  if (Array.isArray(value)) return value.map(redact);
  if (!value || typeof value !== "object") return value;

  return Object.entries(value).reduce((next, [key, item]) => {
    if (/token|authorization|password|secret/i.test(key)) {
      next[key] = item ? "[redacted]" : item;
      return next;
    }
    next[key] = redact(item);
    return next;
  }, {});
}

function writeStatus(label, value) {
  statusOutput.textContent = `${label}\n${JSON.stringify(redact(value), null, 2)}`;
}

function isRuntimeReloadError(error) {
  return /message port closed|receiving end does not exist|service worker/i.test(error?.message || "");
}

function clearResultLink() {
  resultLink.replaceChildren();
}

function appendResultLink(label, href, downloadName = "") {
  if (!href) return;

  const link = document.createElement("a");
  link.href = href;
  link.target = "_blank";
  link.rel = "noreferrer";
  if (downloadName) link.download = downloadName;
  link.textContent = label;
  resultLink.append(link);
}

function collectComparisonPublishers(data) {
  const publishers = [];
  for (let i = 1; i <= 4; i += 1) {
    const clientUsername = String(data.get(`competitor${i}Username`) || "").trim();
    const sourceID = String(data.get(`competitor${i}SourceID`) || "").trim();
    if (!clientUsername && !sourceID) continue;
    publishers.push({
      label: `Publisher ${i}`,
      clientUsername,
      sourceID,
      siteID: sourceID
    });
  }
  return publishers;
}

function getFormPayload() {
  const data = new FormData(form);
  const sourceID = String(data.get("sourceID") || "").trim();
  const comparisonPublishers = collectComparisonPublishers(data);

  return {
    type: "PUBLISHER_QBR_REQUEST",
    analysisLevel: "publisher",
    clientUsername: String(data.get("clientUsername") || "").trim(),
    sourceID,
    siteID: sourceID,
    comparisonPublishers,
    competitors: comparisonPublishers,
    languageCode: String(data.get("languageCode") || "").trim(),
    currencyCode: String(data.get("currencyCode") || "").trim(),
    reportingPeriod: String(data.get("reportingPeriod") || "").trim(),
    startDate: String(data.get("startDate") || "").trim(),
    endDate: String(data.get("endDate") || "").trim(),
    fromDate: normalizeDate(data.get("startDate")),
    toDate: normalizeDate(data.get("endDate")),
    tdSession: {
      mode: "extension_publisher_impersonation",
      tokensIncluded: false
    },
    requestedFrom: "publisher-qbr-extension"
  };
}

function validateRequest(payload) {
  if (!payload.clientUsername) return "Primary username is required.";
  if (!payload.sourceID) return "Primary site ID / sourceID is required.";
  if (!payload.startDate || !payload.endDate) return "Select a valid reporting period.";
  if (payload.startDate > payload.endDate) return "Start date must be before end date.";
  if (daysBetween(payload.startDate, payload.endDate) > 366) return "Date ranges are limited to 366 days.";

  for (const publisher of payload.comparisonPublishers) {
    if (!publisher.clientUsername || !publisher.sourceID) {
      return `${publisher.label} needs both username and site ID, or leave the row blank.`;
    }
  }

  return "";
}

function findFirstValue(data, keys) {
  if (!data || typeof data !== "object") return "";

  for (const key of keys) {
    if (data[key]) return data[key];
  }

  for (const value of Object.values(data)) {
    const nested = findFirstValue(value, keys);
    if (nested) return nested;
  }

  return "";
}

function showWebhookResultLinks(response) {
  clearResultLink();
  const data = response?.data || response;
  const resultUrl = findFirstValue(data, ["reportUrl", "resultUrl", "downloadUrl", "download_url", "pptx_url", "file_url", "url"]);
  const fileName = findFirstValue(data, ["file_name", "filename"]) || "publisher-qbr-report.pptx";
  if (resultUrl) appendResultLink("Open result", resultUrl, fileName);
}

function updateSummary() {
  const payload = getFormPayload();
  const items = [
    ["Primary Username", payload.clientUsername || "Not set"],
    ["Primary Site ID", payload.sourceID || "Not set"],
    ["Impersonated", impersonatedPrimaryUsername || "Not connected"],
    ["Comparisons", payload.comparisonPublishers.length ? String(payload.comparisonPublishers.length) : "None"],
    ["Comparison Sites", payload.comparisonPublishers.length ? payload.comparisonPublishers.map((item) => `${item.clientUsername} / ${item.sourceID}`).join(", ") : "None"],
    ["Language / Currency", `${payload.languageCode || "-"} / ${payload.currencyCode || "-"}`],
    ["Date Range", payload.startDate && payload.endDate ? `${payload.startDate} to ${payload.endDate}` : "Not set"],
    ["Webhook", getConnectionConfig().qbrWebhookUrl]
  ];

  summaryList.replaceChildren(
    ...items.map(([term, description]) => {
      const row = document.createElement("div");
      const dt = document.createElement("dt");
      const dd = document.createElement("dd");
      dt.textContent = term;
      dd.textContent = description;
      row.append(dt, dd);
      return row;
    })
  );
}

async function saveConnection() {
  clearResultLink();
  saveConnectionButton.disabled = true;
  writeStatus("Saving TD connection settings...", {
    qbrWebhookUrl: qbrWebhookUrlInput.value,
    oauthUrl: oauthUrlInput.value
  });

  try {
    const cfg = await saveConnectionConfig();
    let response = { status: "saved locally" };
    let workerWarning = "";

    try {
      response = await sendExtensionRequest({ type: "SAVE_CONFIG" });
    } catch (error) {
      if (!isRuntimeReloadError(error)) throw error;
      workerWarning = error.message;
    }

    writeStatus("TD connection settings saved", {
      ...response,
      qbrWebhookUrl: cfg.qbrWebhookUrl,
      passwordStored: false,
      serviceWorkerReady: !workerWarning,
      warning: workerWarning || undefined
    });
    return true;
  } catch (error) {
    writeStatus("TD connection settings failed", { error: error.message });
    return false;
  } finally {
    saveConnectionButton.disabled = false;
    updateSummary();
  }
}

async function impersonatePrimary() {
  const clientUsername = clientUsernameInput.value.trim();
  if (!clientUsername) {
    writeStatus("Impersonation blocked", { error: "Primary username is required." });
    return false;
  }

  impersonatePrimaryButton.disabled = true;
  sessionStatus.textContent = "Impersonating primary publisher...";

  try {
    await saveConnectionConfig();
    const response = await sendExtensionRequest({
      type: "IMPERSONATE_PUBLISHER",
      username: clientUsername,
      bearerToken: accessTokenInput.value.trim()
    });

    impersonatedPrimaryUsername = clientUsername;
    sessionStatus.textContent = `Impersonated ${clientUsername}.`;
    updateSummary();
    writeStatus("Primary publisher impersonated via extension", {
      username: clientUsername,
      data: response,
      tokenStoredInExtension: true
    });
    return true;
  } catch (error) {
    impersonatedPrimaryUsername = "";
    sessionStatus.textContent = "Primary publisher impersonation failed.";
    updateSummary();
    writeStatus("Primary publisher impersonation failed", {
      error: error.message,
      hint: "Check admin credentials, OAuth Basic, and the publisher username. You can also provide an admin bearer token override."
    });
    return false;
  } finally {
    impersonatePrimaryButton.disabled = false;
  }
}

async function clearSessionFields() {
  accessTokenInput.value = "";
  adminPasswordInput.value = "";
  impersonatedPrimaryUsername = "";
  sessionStatus.textContent = "No publisher impersonated.";
  clearResultLink();
  updateSummary();

  try {
    await sendRuntimeMessage({ type: "CLEAR_STATE" });
    writeStatus("TD session cleared", { extensionTokenFieldCleared: true, extensionSessionCleared: true });
  } catch (error) {
    writeStatus("Extension fields cleared", {
      extensionTokenFieldCleared: true,
      extensionSessionClearError: error.message
    });
  }
}

async function submitRequest() {
  clearResultLink();

  if (!form.reportValidity()) {
    writeStatus("Form validation failed", { message: "Complete required fields before submitting." });
    return;
  }

  const payload = getFormPayload();
  const validationError = validateRequest(payload);
  if (validationError) {
    writeStatus("Form validation failed", { error: validationError });
    return;
  }

  submitButton.disabled = true;

  try {
    await saveConnectionConfig();
    writeStatus("Submitting publisher request via extension service worker...", payload);
    const response = await sendExtensionRequest({
      type: "SUBMIT_PUBLISHER_QBR",
      payload,
      bearerToken: accessTokenInput.value.trim()
    });
    writeStatus("Publisher request response", response);
    showWebhookResultLinks(response);
  } catch (error) {
    writeStatus("Publisher request failed", { error: error.message });
  } finally {
    submitButton.disabled = false;
  }
}

async function init() {
  const storedConfig = await getStorageValue(STORAGE_KEYS.connectionConfig);
  const legacyBaseUrl = await getStorageValue("publisherQbrApiBaseUrl");
  applyConnectionConfig({
    ...DEFAULTS,
    ...(storedConfig || {}),
    qbrWebhookUrl: storedConfig?.qbrWebhookUrl || storedConfig?.qbrApiBaseUrl || legacyBaseUrl || DEFAULTS.qbrWebhookUrl
  });
  applyPresetRange();
  updateSummary();

  try {
    await sendExtensionRequest({ type: "PING" });
  } catch (error) {
    if (isRuntimeReloadError(error)) {
      writeStatus("Extension background needs reload", {
        savedSettingsAvailable: true,
        action: "Reload the unpacked Publisher QBR extension in chrome://extensions, then reopen this page.",
        detail: error.message
      });
    } else {
      writeStatus("Extension background check failed", { error: error.message });
    }
  }
}

form.addEventListener("input", updateSummary);
form.addEventListener("change", updateSummary);
reportingPeriodInput.addEventListener("change", () => {
  applyPresetRange();
  updateSummary();
});

for (const input of [adminUsernameInput, qbrWebhookUrlInput, oauthUrlInput, impersonateUrlInput, oauthBasicInput]) {
  input.addEventListener("blur", () => {
    saveConnectionConfig().catch((error) => writeStatus("Connection config save failed", { error: error.message }));
  });
}

saveConnectionButton.addEventListener("click", saveConnection);
impersonatePrimaryButton.addEventListener("click", impersonatePrimary);
clearSessionButton.addEventListener("click", clearSessionFields);
submitButton.addEventListener("click", submitRequest);

init().catch((error) => writeStatus("Extension initialization failed", { error: error.message }));

