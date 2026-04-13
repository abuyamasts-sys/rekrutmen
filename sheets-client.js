(() => {
  const ENDPOINT = () => (window.AIRTIS_SHEETS_ENDPOINT || "").toString().trim();
  const SECRET = () => (window.AIRTIS_SHEETS_SECRET || "").toString().trim();
  const QUEUE_KEY = "AIRTIS_SHEET_QUEUE";

  function nowIso() {
    return new Date().toISOString();
  }

  function safeParseJson(text) {
    try { return JSON.parse(text); } catch { return null; }
  }

  function getQueue() {
    try {
      const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]");
      return Array.isArray(q) ? q : [];
    } catch {
      return [];
    }
  }

  function setQueue(queue) {
    try { localStorage.setItem(QUEUE_KEY, JSON.stringify(queue)); } catch {}
  }

  async function postJson(url, body) {
    // Apps Script Web App does not reliably support CORS for browser fetch.
    // Use sendBeacon when available (fire-and-forget), otherwise no-cors fetch.
    const textBody = JSON.stringify(body);

    if (navigator.sendBeacon) {
      const ok = navigator.sendBeacon(url, new Blob([textBody], { type: "text/plain;charset=utf-8" }));
      if (!ok) throw new Error("sendBeacon failed");
      return { ok: true, beacon: true };
    }

    // no-cors: browser won't expose response, but request will be delivered.
    await fetch(url, { method: "POST", mode: "no-cors", body: textBody });
    return { ok: true, nocors: true };
  }

  function jsonpGet(params) {
    const endpoint = ENDPOINT();
    if (!endpoint) return Promise.resolve({ ok: false, error: "missing_endpoint" });

    const cbName = "__airtis_cb_" + Math.random().toString(36).slice(2);
    const secret = SECRET();

    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      const qs = new URLSearchParams();
      Object.entries(params || {}).forEach(([k, v]) => {
        if (v == null) return;
        const s = String(v).trim();
        if (!s) return;
        qs.set(k, s);
      });
      if (secret) qs.set("secret", secret);
      qs.set("cb", cbName);

      const url = endpoint + (endpoint.includes("?") ? "&" : "?") + qs.toString();

      window[cbName] = (payload) => {
        try { delete window[cbName]; } catch {}
        script.remove();
        resolve(payload);
      };

      script.onerror = () => {
        try { delete window[cbName]; } catch {}
        script.remove();
        reject(new Error("jsonp_failed"));
      };

      script.src = url;
      document.head.appendChild(script);
    });
  }

  async function validateToken(token) {
    const t = (token || "").toString().trim();
    if (!t) return { ok: true, valid: false, reason: "missing_token" };

    try {
      const res = await jsonpGet({ kind: "validate_token", token: t });
      if (!res || !res.ok) return { ok: false, valid: false, error: res?.error || "invalid_response" };
      return res;
    } catch (e) {
      return { ok: false, valid: false, error: e?.message || String(e) };
    }
  }

  async function flushQueue() {
    const url = ENDPOINT();
    if (!url) return { ok: false, skipped: true, reason: "missing_endpoint" };

    const queue = getQueue();
    if (!queue.length) return { ok: true, flushed: 0 };

    // With beacon/no-cors we can't confirm delivery; treat as best-effort.
    // Keep queue small: after attempting, clear it.
    for (const item of queue) {
      try { await postJson(url, item); } catch {}
    }
    setQueue([]);
    return { ok: true, flushed: queue.length, remaining: 0 };
  }

  async function send(kind, payload) {
    const url = ENDPOINT();
    const body = {
      secret: SECRET() || undefined,
      kind,
      at: nowIso(),
      tz: Intl.DateTimeFormat().resolvedOptions().timeZone || "",
      userAgent: navigator.userAgent,
      payload
    };

    // Always queue first (safer for intermittent network)
    const queue = getQueue();
    queue.push(body);
    setQueue(queue);

    // Best-effort flush
    try { await flushQueue(); } catch {}
    return { queued: true };
  }

  window.AirtisSheets = {
    send,
    flushQueue,
    jsonpGet,
    validateToken
  };
})();
