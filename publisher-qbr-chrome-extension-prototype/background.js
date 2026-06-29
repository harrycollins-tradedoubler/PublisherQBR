chrome.action.onClicked.addListener(() => {
  chrome.tabs.create({ url: chrome.runtime.getURL("app.html") });
});

const state = {
  adminToken: null,
  publisherTokens: new Map()
};

function normalizeBaseUrl(value, fallback) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return fallback;
  const withProtocol = /^[a-z][a-z\d+\-.]*:\/\//i.test(trimmed) ? trimmed : `http://${trimmed}`;
  return withProtocol.replace(/\/+$/, "");
}

function normalizeCfg(cfg = {}) {
  return {
    adminUsername: String(cfg.adminUsername || cfg.username || "").trim(),
    adminPassword: String(cfg.adminPassword || cfg.password || ""),
    oauthUrl: String(cfg.oauthUrl || "https://connect.tradedoubler.com/uaa/oauth/token").trim(),
    impersonateUrl: String(cfg.impersonateUrl || "https://connect.tradedoubler.com/uaa/admin/impersonate?username=").trim(),
    oauthBasic: String(cfg.oauthBasic || "dGRjb25uZWN0X3B1Ymxpc2hlcjoxMjM0NTY=").trim(),
    qbrWebhookUrl: normalizeBaseUrl(
      cfg.qbrWebhookUrl || cfg.webhookUrl || cfg.qbrApiBaseUrl || cfg.apiBaseUrl,
      "http://127.0.0.1:3020/webhook-local/publisher-qbr-v5-competitor-weekly-chart-20260505"
    )
  };
}

function toFormBody(values) {
  return Object.entries(values)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value == null ? "" : value)}`)
    .join("&");
}

function timeoutMsFor(label) {
  return /qbr|webhook|publisher/i.test(label) ? 180000 : 45000;
}

async function readJsonResponse(response, label) {
  const text = await response.text();
  let data = text;
  try {
    data = text ? JSON.parse(text) : null;
  } catch (_error) {
    data = text;
  }

  if (!response.ok) {
    const detail = data && typeof data === "object"
      ? data.detail || data.message || data.error
      : String(data || "").slice(0, 1500);
    throw new Error(`${label} failed: HTTP ${response.status}${detail ? ` ${detail}` : ""}`);
  }

  return data;
}

async function fetchJson(url, init, label) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMsFor(label));

  try {
    const response = await fetch(url, { ...init, signal: controller.signal });
    return await readJsonResponse(response, label);
  } catch (error) {
    if (error && error.name === "AbortError") {
      throw new Error(`${label} timed out while calling ${url}`);
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

async function oauthPasswordToken(cfg) {
  if (!cfg.adminUsername || !cfg.adminPassword) {
    throw new Error("Admin username/password are required unless an admin bearer token override is provided.");
  }
  if (!cfg.oauthUrl || !cfg.oauthBasic) {
    throw new Error("OAuth URL and OAuth Basic are required.");
  }

  const data = await fetchJson(
    cfg.oauthUrl,
    {
      method: "POST",
      headers: {
        "Accept": "application/json",
        "Authorization": `Basic ${cfg.oauthBasic.replace(/^Basic\s+/i, "")}`,
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: toFormBody({
        grant_type: "password",
        username: cfg.adminUsername,
        password: cfg.adminPassword
      })
    },
    "OAuth token"
  );

  const token = data && data.access_token;
  if (!token) throw new Error("No access_token in OAuth response.");
  return String(token);
}

async function ensureAdminToken(cfg, bearerToken) {
  const override = String(bearerToken || "").trim();
  if (override) {
    state.adminToken = override;
    return state.adminToken;
  }

  if (state.adminToken) return state.adminToken;
  state.adminToken = await oauthPasswordToken(cfg);
  return state.adminToken;
}

async function impersonatePublisher(cfg, username, bearerToken) {
  const targetUser = String(username || "").trim();
  if (!targetUser) throw new Error("Publisher username is required.");

  const cachedToken = state.publisherTokens.get(targetUser);
  if (cachedToken) {
    return {
      username: targetUser,
      tokenStoredInExtension: true,
      reusedToken: true
    };
  }

  const adminToken = await ensureAdminToken(cfg, bearerToken);
  const data = await fetchJson(
    `${cfg.impersonateUrl}${encodeURIComponent(targetUser)}`,
    {
      method: "GET",
      headers: {
        "Accept": "application/json",
        "Authorization": `Bearer ${adminToken}`
      }
    },
    "Publisher impersonation"
  );

  const token = data && data.access_token;
  if (!token) throw new Error("No access_token in impersonation response.");

  state.publisherTokens.set(targetUser, String(token));
  return {
    username: targetUser,
    tokenStoredInExtension: true,
    reusedToken: false
  };
}

function tokensForPublisher(username) {
  const impersonateToken = state.publisherTokens.get(String(username || "").trim());
  if (!state.adminToken || !impersonateToken) return null;

  return {
    user_access_token: state.adminToken,
    impersonate_access_token: impersonateToken
  };
}

function publisherSessionsForPayload(payload) {
  const sessions = [];
  const primaryTokens = tokensForPublisher(payload.clientUsername);
  if (primaryTokens) {
    sessions.push({
      role: "primary",
      clientUsername: payload.clientUsername,
      sourceID: payload.sourceID,
      siteID: payload.siteID || payload.sourceID,
      td_tokens: primaryTokens
    });
  }

  const comparisons = Array.isArray(payload.comparisonPublishers) ? payload.comparisonPublishers : [];
  for (const publisher of comparisons) {
    const tokens = tokensForPublisher(publisher.clientUsername);
    if (!tokens) continue;
    sessions.push({
      role: "comparison",
      label: publisher.label,
      clientUsername: publisher.clientUsername,
      sourceID: publisher.sourceID,
      siteID: publisher.siteID || publisher.sourceID,
      td_tokens: tokens
    });
  }

  return sessions;
}

function publisherWithTokenSession(publisher) {
  const tokens = tokensForPublisher(publisher.clientUsername);
  const hasTokens = Boolean(tokens);

  return {
    ...publisher,
    tokenSessionAvailable: hasTokens,
    ...(hasTokens
      ? {
          td_tokens: tokens,
          tdTokens: tokens,
          user_access_token: tokens.user_access_token,
          impersonate_access_token: tokens.impersonate_access_token
        }
      : {})
  };
}

function tokensByPublisherForPayload(payload) {
  const entries = [];
  const primaryTokens = tokensForPublisher(payload.clientUsername);
  if (payload.clientUsername && primaryTokens) {
    entries.push([payload.clientUsername, primaryTokens]);
  }

  const comparisons = Array.isArray(payload.comparisonPublishers) ? payload.comparisonPublishers : [];
  for (const publisher of comparisons) {
    const username = String(publisher.clientUsername || "").trim();
    const tokens = tokensForPublisher(username);
    if (username && tokens) entries.push([username, tokens]);
  }

  return Object.fromEntries(entries);
}

async function buildPublisherQbrPayload(cfg, payload, bearerToken) {
  const clientUsername = String(payload && payload.clientUsername || "").trim();
  if (!clientUsername) throw new Error("clientUsername is required.");

  await impersonatePublisher(cfg, clientUsername, bearerToken);

  const comparisons = Array.isArray(payload.comparisonPublishers) ? payload.comparisonPublishers : [];
  for (const publisher of comparisons) {
    if (publisher.clientUsername) {
      await impersonatePublisher(cfg, publisher.clientUsername, bearerToken);
    }
  }

  const primaryTokens = tokensForPublisher(clientUsername);
  if (!primaryTokens) throw new Error("No TD tokens available for the primary publisher.");

  const publisherSessions = publisherSessionsForPayload(payload);
  const comparisonsWithTokens = comparisons.map(publisherWithTokenSession);
  const requestPayload = {
    ...payload,
    td_tokens: primaryTokens,
    publisherSessions,
    tdTokensByPublisher: tokensByPublisherForPayload(payload),
    comparisonPublishers: comparisonsWithTokens,
    competitorPublishers: comparisonsWithTokens,
    competitors: comparisonsWithTokens,
    tdSession: {
      ...(payload.tdSession || {}),
      mode: "extension_publisher_impersonation",
      tokensIncluded: true,
      tokenStorage: "extension_service_worker_session"
    }
  };

  return {
    requestPayload,
    tdSession: {
      mode: "extension_publisher_impersonation",
      clientUsername,
      publisherSessionCount: publisherSessions.length,
      tokenStoredInExtension: true
    }
  };
}

async function submitPublisherQbr(cfg, payload, bearerToken) {
  const prepared = await buildPublisherQbrPayload(cfg, payload, bearerToken);
  const data = await fetchJson(
    cfg.qbrWebhookUrl,
    {
      method: "POST",
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        message: `QBR_REQUEST ${JSON.stringify(prepared.requestPayload)}`,
        thread_id: crypto.randomUUID(),
        payload: prepared.requestPayload,
        qbr_payload: prepared.requestPayload,
        td_tokens: prepared.requestPayload.td_tokens,
        publisherSessions: prepared.requestPayload.publisherSessions
      })
    },
    "Publisher QBR local runner"
  );

  return {
    ...data,
    tdSession: prepared.tdSession
  };
}

async function proxyApiRequest(msg) {
  const response = await fetch(msg.url, {
    method: msg.method || "GET",
    headers: msg.headers || {},
    body: msg.body
  });
  const text = await response.text();
  let data = text;

  try {
    data = text ? JSON.parse(text) : null;
  } catch (_error) {
    data = text;
  }

  return {
    ok: true,
    response: {
      ok: response.ok,
      status: response.status,
      statusText: response.statusText,
      data
    }
  };
}

async function handleMessage(msg) {
  if (!msg || !msg.type) {
    return { ok: false, error: "Missing message type." };
  }

  if (msg.type === "QBR_API_REQUEST") {
    return proxyApiRequest(msg);
  }

  if (msg.type === "PING") {
    return {
      ok: true,
      data: {
        status: "ready",
        serviceWorker: "publisher-qbr",
        version: "0.2.1"
      }
    };
  }

  if (msg.type === "CLEAR_STATE") {
    state.adminToken = null;
    state.publisherTokens.clear();
    return { ok: true, data: { status: "cleared" } };
  }

  const cfg = normalizeCfg(msg.cfg || {});
  const bearerToken = String(msg.bearerToken || "").trim();

  if (msg.type === "SAVE_CONFIG") {
    return { ok: true, data: { status: "saved", qbrWebhookUrl: cfg.qbrWebhookUrl } };
  }

  if (msg.type === "IMPERSONATE_PUBLISHER") {
    const data = await impersonatePublisher(cfg, msg.username, bearerToken);
    return { ok: true, data };
  }

  if (msg.type === "SUBMIT_PUBLISHER_QBR") {
    const data = await submitPublisherQbr(cfg, msg.payload, bearerToken);
    return { ok: true, data };
  }

  return { ok: false, error: `Unknown message type: ${msg.type}` };
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  handleMessage(msg)
    .then((response) => sendResponse(response))
    .catch((error) => {
      sendResponse({
        ok: false,
        error: error && error.message ? error.message : String(error)
      });
    });

  return true;
});

