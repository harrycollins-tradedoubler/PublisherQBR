const form = document.getElementById("qbrForm");
const submitBtn = document.getElementById("submitBtn");
const responseOutput = document.getElementById("responseOutput");
const summaryPanel = document.getElementById("summaryPanel");
const apiBaseUrlInput = document.getElementById("apiBaseUrl");

const reportingPeriodInput = document.getElementById("reportingPeriod");
const startDateInput = document.getElementById("startDate");
const endDateInput = document.getElementById("endDate");
const DEFAULT_API_BASE_URL = "http://localhost:3000";
const STORAGE_KEY = "publisherQbrApiBaseUrl";

function normalizeApiBaseUrl(value) {
  return String(value || DEFAULT_API_BASE_URL).trim().replace(/\/+$/, "");
}

function getApiBaseUrl() {
  return normalizeApiBaseUrl(apiBaseUrlInput?.value || DEFAULT_API_BASE_URL);
}

function apiUrl(path) {
  return `${getApiBaseUrl()}${path}`;
}

async function requestApiJson(path, options = {}) {
  const url = apiUrl(path);
  const method = options.method || "GET";
  const headers = options.headers || {};
  const body = options.body;

  if (globalThis.chrome?.runtime?.sendMessage) {
    const messageResponse = await chrome.runtime.sendMessage({
      type: "QBR_API_REQUEST",
      url,
      method,
      headers,
      body,
    });

    if (!messageResponse?.ok) {
      throw new Error(messageResponse?.error || "Failed to fetch");
    }

    return messageResponse.response;
  }

  const response = await fetch(url, { method, headers, body });
  return {
    ok: response.ok,
    status: response.status,
    statusText: response.statusText,
    data: await response.json(),
  };
}

async function restoreExtensionSettings() {
  if (!apiBaseUrlInput) return;
  apiBaseUrlInput.value = DEFAULT_API_BASE_URL;

  if (!globalThis.chrome?.storage?.local) return;
  const data = await chrome.storage.local.get([STORAGE_KEY]);
  apiBaseUrlInput.value = normalizeApiBaseUrl(data[STORAGE_KEY]);
}

async function saveExtensionSettings() {
  if (!apiBaseUrlInput || !globalThis.chrome?.storage?.local) return;
  await chrome.storage.local.set({ [STORAGE_KEY]: getApiBaseUrl() });
}

function formatDateForInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function startOfQuarter(date) {
  const month = date.getMonth();
  const quarterStartMonth = Math.floor(month / 3) * 3;
  return new Date(date.getFullYear(), quarterStartMonth, 1);
}

function endOfQuarter(date) {
  const start = startOfQuarter(date);
  return new Date(start.getFullYear(), start.getMonth() + 3, 0);
}

function calculateRange(period) {
  const today = new Date();
  const end = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  switch (period) {
    case "last_30_days": {
      const start = new Date(end);
      start.setDate(start.getDate() - 29);
      return { start, end };
    }
    case "last_90_days": {
      const start = new Date(end);
      start.setDate(start.getDate() - 89);
      return { start, end };
    }
    case "this_quarter": {
      return { start: startOfQuarter(end), end };
    }
    case "last_quarter": {
      const thisQuarterStart = startOfQuarter(end);
      const lastQuarterEnd = new Date(thisQuarterStart.getFullYear(), thisQuarterStart.getMonth(), 0);
      const lastQuarterStart = startOfQuarter(lastQuarterEnd);
      return { start: lastQuarterStart, end: endOfQuarter(lastQuarterStart) };
    }
    default:
      return null;
  }
}

function applyPresetRange() {
  const period = reportingPeriodInput.value;
  const range = calculateRange(period);
  const isCustom = period === "custom";

  startDateInput.disabled = !isCustom;
  endDateInput.disabled = !isCustom;

  if (!isCustom && range) {
    startDateInput.value = formatDateForInput(range.start);
    endDateInput.value = formatDateForInput(range.end);
  }
}

function buildSummary() {
  const data = new FormData(form);
  const clientUsername = String(data.get("clientUsername") || "").trim() || "Not set";
  const sourceID = String(data.get("sourceID") || "").trim() || "Not set (optional)";
  const languageCode = String(data.get("languageCode") || "").trim() || "-";
  const currencyCode = String(data.get("currencyCode") || "").trim() || "-";
  const reportingPeriod = String(data.get("reportingPeriod") || "").trim() || "-";
  const startDate = String(data.get("startDate") || "").trim() || "-";
  const endDate = String(data.get("endDate") || "").trim() || "-";
  const competitors = collectCompetitors(data);
  const competitorSummary = competitors.length
    ? competitors
        .map((competitor) => `<div><strong>${competitor.label}:</strong> ${competitor.clientUsername} / ${competitor.sourceID}</div>`)
        .join("")
    : '<div><strong>Competitors:</strong> Not set (optional)</div>';

  summaryPanel.innerHTML = `
    <div><strong>Primary Username:</strong> ${clientUsername}</div>
    <div><strong>Primary sourceID:</strong> ${sourceID}</div>
    ${competitorSummary}
    <div><strong>Language / Currency:</strong> ${languageCode} / ${currencyCode}</div>
    <div><strong>Reporting Period:</strong> ${reportingPeriod}</div>
    <div><strong>Date Range:</strong> ${startDate} to ${endDate}</div>
  `;
}

function collectCompetitors(formData) {
  const competitors = [];
  for (let i = 1; i <= 4; i += 1) {
    const clientUsername = String(formData.get(`competitor${i}Username`) || "").trim();
    const sourceID = String(formData.get(`competitor${i}SourceID`) || "").trim();
    if (!clientUsername && !sourceID) continue;
    competitors.push({
      label: `Publisher ${i}`,
      clientUsername,
      sourceID,
    });
  }
  return competitors;
}

function setLoading(loading) {
  submitBtn.disabled = loading;
  submitBtn.textContent = loading ? "Generating..." : "Send QBR Request";
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function buildDownloadArtifacts(summary) {
  const executionId = String(summary?.executionId || "").trim();
  if (!executionId) return [];

  const result = summary?.result || {};
  const baseDownloadUrl = apiUrl(`/api/publisher-qbr/download/${encodeURIComponent(executionId)}`);
  const artifacts = [];

  if (result.pptx_url) {
    artifacts.push({
      key: "pptx",
      label: "PowerPoint",
      fileName: result.file_name || "publisher-qbr-report.pptx",
      downloadUrl: `${baseDownloadUrl}?artifact=pptx`,
      directUrl: result.pptx_url,
      primary: true,
    });
  }

  if (result.gap_analysis_report_url) {
    artifacts.push({
      key: "gap-analysis",
      label: "Gap Analysis Excel",
      fileName: result.gap_analysis_report_file_name || "publisher-qbr-gap-analysis.xlsx",
      downloadUrl: `${baseDownloadUrl}?artifact=gap-analysis`,
      directUrl: result.gap_analysis_report_url,
      primary: false,
    });
  }

  return artifacts;
}

function triggerDownload(downloadUrl, fileName) {
  const link = document.createElement("a");
  link.href = downloadUrl;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  link.remove();
}

async function queueDownloads(artifacts) {
  for (let index = 0; index < artifacts.length; index += 1) {
    if (index > 0) {
      await sleep(350);
    }
    triggerDownload(artifacts[index].downloadUrl, artifacts[index].fileName);
  }
}

async function pollQbrStatus(executionId) {
  const maxAttempts = 180;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    await sleep(attempt === 1 ? 1200 : 5000);

    const response = await requestApiJson(`/api/publisher-qbr/status/${encodeURIComponent(executionId)}`);
    const json = response.data;

    if (!response.ok || !json.ok) {
      throw new Error(json.error || "Unable to check QBR generation status.");
    }

    responseOutput.textContent = JSON.stringify(
      {
        message: json.finished ? "QBR generation finished." : "QBR generation running...",
        executionId,
        status: json.status,
        lastNodeExecuted: json.lastNodeExecuted,
        result: json.result,
        error: json.error,
      },
      null,
      2
    );

    if (json.finished) {
      if (json.status === "success" && json.result) {
        return json;
      }
      throw new Error(json.error?.message || `QBR generation ended with status: ${json.status}`);
    }
  }

  throw new Error("QBR generation is still running. Check n8n executions for the latest status.");
}

function startPptxDownload(summary) {
  const executionId = summary?.executionId;
  if (!executionId) return;

  const artifacts = buildDownloadArtifacts(summary);
  const safeExecutionId = escapeHtml(executionId);
  const artifactMarkup = artifacts
    .map((artifact) => {
      const safeLabel = escapeHtml(artifact.label);
      const safeFileName = escapeHtml(artifact.fileName);
      const safeDownloadUrl = escapeHtml(artifact.downloadUrl);
      const safeDirectUrl = escapeHtml(artifact.directUrl || "");
      const actionClass = artifact.primary ? "download-link" : "secondary-link";
      return `
        <a class="${actionClass}" href="${safeDownloadUrl}" download="${safeFileName}">Download ${safeLabel}</a>
        ${artifact.directUrl ? `<a class="secondary-link" href="${safeDirectUrl}" target="_blank" rel="noopener">Open ${safeLabel} Source</a>` : ""}
      `;
    })
    .join("");
  const fileSummary = artifacts.length
    ? artifacts.map((artifact) => escapeHtml(artifact.fileName)).join("<br />")
    : "No downloadable files were returned.";

  responseOutput.innerHTML = `
    <div>QBR generation finished. Downloads should start automatically.</div>
    <div><strong>Execution:</strong> ${safeExecutionId}</div>
    <div><strong>Files:</strong><br />${fileSummary}</div>
    <div class="download-actions">
      ${artifactMarkup}
    </div>
  `;

  if (artifacts.length) {
    queueDownloads(artifacts);
  }
}

reportingPeriodInput.addEventListener("change", () => {
  applyPresetRange();
  buildSummary();
});

form.addEventListener("input", () => {
  buildSummary();
});

apiBaseUrlInput?.addEventListener("change", () => {
  saveExtensionSettings().catch(() => {});
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const formData = new FormData(form);
  const sourceID = String(formData.get("sourceID") || "").trim();
  const competitors = collectCompetitors(formData);
  const payload = {
    clientUsername: String(formData.get("clientUsername") || "").trim(),
    languageCode: String(formData.get("languageCode") || "").trim(),
    currencyCode: String(formData.get("currencyCode") || "").trim(),
    reportingPeriod: String(formData.get("reportingPeriod") || "").trim(),
    startDate: String(formData.get("startDate") || "").trim(),
    endDate: String(formData.get("endDate") || "").trim(),
  };

  if (sourceID) {
    payload.sourceID = sourceID;
  }
  if (competitors.length) {
    payload.competitors = competitors.map(({ clientUsername, sourceID }) => ({ clientUsername, sourceID }));
  }

  setLoading(true);
  responseOutput.textContent = "Submitting publisher QBR request...";

  try {
    await saveExtensionSettings();

    const response = await requestApiJson("/api/publisher-qbr", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const json = response.data;

    if (!response.ok) {
      throw new Error(json.error || "Request failed.");
    }

    if (json.accepted && json.executionId) {
      responseOutput.textContent = JSON.stringify(
        {
          message: "QBR generation started.",
          executionId: json.executionId,
          workflowId: json.n8n?.workflowId,
          webhookUrl: json.n8n?.webhookUrl,
        },
        null,
        2
      );
      const summary = await pollQbrStatus(json.executionId);
      startPptxDownload(summary);
      return;
    }

    responseOutput.textContent = JSON.stringify(json, null, 2);
  } catch (error) {
    responseOutput.innerHTML = `<span class="error-text">Error: ${error.message}</span>`;
  } finally {
    setLoading(false);
  }
});

restoreExtensionSettings()
  .catch(() => {})
  .finally(() => {
    applyPresetRange();
    buildSummary();
  });
