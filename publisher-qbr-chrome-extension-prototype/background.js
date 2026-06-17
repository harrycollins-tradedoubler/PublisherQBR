chrome.action.onClicked.addListener(async () => {
  await chrome.tabs.create({ url: chrome.runtime.getURL("app.html") });
});

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (!msg || msg.type !== "QBR_API_REQUEST") {
    return false;
  }

  (async () => {
    try {
      const response = await fetch(msg.url, {
        method: msg.method || "GET",
        headers: msg.headers || {},
        body: msg.body,
      });
      const text = await response.text();
      let data = text;

      try {
        data = text ? JSON.parse(text) : null;
      } catch (_) {
        data = text;
      }

      sendResponse({
        ok: true,
        response: {
          ok: response.ok,
          status: response.status,
          statusText: response.statusText,
          data,
        },
      });
    } catch (error) {
      sendResponse({
        ok: false,
        error: error && error.message ? error.message : String(error),
      });
    }
  })();

  return true;
});
