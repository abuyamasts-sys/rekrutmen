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
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const text = await res.text();
    const json = safeParseJson(text);
    if (!res.ok) {
      const msg = (json && (json.error || json.message)) ? (json.error || json.message) : text;
      throw new Error(`Sheets HTTP ${res.status}: ${msg}`);
    }
    return json || { ok: true, raw: text };
  }

  async function flushQueue() {
    const url = ENDPOINT();
    if (!url) return { ok: false, skipped: true, reason: "missing_endpoint" };

    const queue = getQueue();
    if (!queue.length) return { ok: true, flushed: 0 };

    const remaining = [];
    let flushed = 0;
    for (const item of queue) {
      try {
        await postJson(url, item);
        flushed++;
      } catch (e) {
        remaining.push(item);
      }
    }
    setQueue(remaining);
    return { ok: remaining.length === 0, flushed, remaining: remaining.length };
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
    flushQueue
  };
})();

