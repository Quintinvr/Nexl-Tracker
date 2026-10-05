/*
 * Nexl Check Bridge — background service worker.
 *
 * Fetches a fixed list of READ-ONLY Nexl Controller pages using the browser's existing Nexl login,
 * and hands the HTML back to the Nexl Check Excel add-in. It never sends POST requests, never
 * submits forms, and refuses any URL that isn't on the allow-list below.
 */
const NEXL = "https://controller.nexl.online";
const ADDIN_ORIGIN = "https://quintinvr.github.io";

const ALLOWED = [
  /^\/php\/ajax\/statusscreen\/get\.status\.screen\.template\.php\?status_type=(active|completed)&p_region=(all|pe|durban|ct)&p_tab_selected=[1-4]$/,
  /^\/php\/ajax\/instruction_containers\/get\.containers\.fresh\.php\?p_id=\d{1,8}(\.\d{1,2})?$/,
  /^\/php\/ajax\/drivertracking\/get\.driver\.table\.php$/,
  /^\/php\/ajax\/search\/get\.search\.php\?p_filter=container&p_word=[A-Za-z0-9]{4,20}$/,
];

// Every one of these endpoints returns a <table> when the session is valid.
const looksValid = (text) => typeof text === "string" && /<table[\s>]/i.test(text);

async function directFetch(path) {
  const r = await fetch(NEXL + path, { method: "GET", credentials: "include", cache: "no-store", redirect: "follow" });
  return { status: r.status, url: r.url, text: await r.text() };
}

// Fallback: run the same GET inside an open Nexl tab (same-origin, so the login cookie is always sent).
async function viaNexlTab(path) {
  const tabs = await chrome.tabs.query({ url: NEXL + "/*" });
  if (!tabs.length) return null;
  const [res] = await chrome.scripting.executeScript({
    target: { tabId: tabs[0].id },
    func: (u) => fetch(u, { method: "GET", credentials: "include", cache: "no-store" }).then((r) => r.text()),
    args: [path],
  });
  return res && typeof res.result === "string" ? { status: 200, text: res.result } : null;
}

async function nexlGet(path) {
  if (typeof path !== "string" || !ALLOWED.some((re) => re.test(path))) {
    return { ok: false, error: "NOT_ALLOWED", detail: "Blocked request: " + String(path).slice(0, 120) };
  }
  try {
    const d = await directFetch(path);
    if (looksValid(d.text)) return { ok: true, text: d.text };
  } catch (e) {
    /* fall through to tab fallback */
  }
  try {
    const t = await viaNexlTab(path);
    if (t && looksValid(t.text)) return { ok: true, text: t.text };
  } catch (e) {
    /* ignore */
  }
  return { ok: false, error: "NOT_LOGGED_IN", detail: "Log in to Nexl Controller in this browser, then sync again." };
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  // Only answer the Nexl Check add-in frame (or this extension's own popup).
  const fromAddin = sender.origin === ADDIN_ORIGIN || (sender.url || "").startsWith(ADDIN_ORIGIN + "/");
  const fromSelf = sender.id === chrome.runtime.id && !sender.tab;
  if (!fromAddin && !fromSelf) return false;

  if (msg && msg.type === "ping") {
    sendResponse({ ok: true, version: chrome.runtime.getManifest().version });
    return false;
  }
  if (msg && msg.type === "fetch") {
    nexlGet(msg.path).then(sendResponse);
    return true; // async response
  }
  sendResponse({ ok: false, error: "UNKNOWN_REQUEST" });
  return false;
});
