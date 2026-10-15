// I01 Outlook feasibility probe: core logic (I01-T02..T05, T09, T11).
// Read-only by design: it never calls send/save/set/display*/move/delete APIs.
// Privacy: it records only metadata (field presence, lengths, counts, requirement sets), never
// subject or body text, addresses, or attachment names. The account fingerprint is a truncated
// SHA-256 used only to tell accounts apart.
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.I01Probe = api;
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const PROBE_VERSION = "0.1.0";
  const MAILBOX_SETS = Array.from({ length: 16 }, (_, i) => "1." + (i + 1));

  function classifyEnvironment(Office) {
    if (!Office || !Office.context) return { ok: false, code: "office_unavailable", guidance: "Office.js did not initialize. Open this page from the Outlook add-in, not a normal browser tab." };
    if (!Office.context.mailbox) return { ok: false, code: "mailbox_unavailable", guidance: "This host is not Outlook, or the mailbox API is unavailable (possible tenant restriction)." };
    if (!Office.context.mailbox.item) return { ok: false, code: "no_item", guidance: "No mail item is selected. Select a synthetic test message and reopen the pane." };
    return { ok: true, code: "ready", guidance: "" };
  }

  function requirementSets(Office) {
    const supported = [];
    const req = Office && Office.context && Office.context.requirements;
    for (const v of MAILBOX_SETS) {
      let ok = false;
      try { ok = !!(req && req.isSetSupported("Mailbox", v)); } catch (e) { ok = false; }
      if (ok) supported.push(v);
    }
    return { supported: supported, highest: supported.length ? supported[supported.length - 1] : null };
  }

  function diagnostics(Office) {
    const d = (Office && Office.context && Office.context.diagnostics) || {};
    return { host: d.host || null, platform: d.platform || null, version: d.version || null };
  }

  // Read mode exposes subject as a string; compose mode exposes an object with getAsync.
  function detectMode(item) {
    if (!item) return "unknown";
    if (typeof item.subject === "string") return "read";
    if (item.subject && typeof item.subject.getAsync === "function") return "compose";
    return "unknown";
  }

  function present(v) {
    if (v === undefined) return "unknown";   // property not exposed in this mode/host: never inferred
    if (v === null || v === "") return "absent";
    return "present";
  }

  function callAsync(fn, args, timeoutMs) {
    return new Promise(function (resolve) {
      let done = false;
      const timer = setTimeout(function () { if (!done) { done = true; resolve({ status: "timeout" }); } }, timeoutMs);
      try {
        fn.apply(null, args.concat([function (res) {
          if (done) return;
          done = true;
          clearTimeout(timer);
          if (res && res.status === "succeeded") resolve({ status: "succeeded", value: res.value });
          else resolve({ status: "failed", error: (res && res.error && (res.error.code || res.error.name)) || "unknown" });
        }]));
      } catch (e) {
        done = true;
        clearTimeout(timer);
        resolve({ status: "threw", error: e && e.name ? e.name : "error" });
      }
    });
  }

  async function bodyInfo(Office, item, timeoutMs) {
    const out = {};
    if (!item.body || typeof item.body.getAsync !== "function") return { html: { status: "unsupported" }, text: { status: "unsupported" } };
    for (const kind of ["Html", "Text"]) {
      const coercion = Office.CoercionType ? Office.CoercionType[kind] : kind.toLowerCase();
      const r = await callAsync(item.body.getAsync.bind(item.body), [coercion], timeoutMs);
      out[kind.toLowerCase()] = r.status === "succeeded"
        ? { status: "succeeded", length: String(r.value || "").length, empty: !String(r.value || "").trim() }
        : { status: r.status, error: r.error || null };
    }
    return out;
  }

  function attachmentInfo(list) {
    if (!Array.isArray(list)) return { status: "unknown" };
    return {
      status: "present",
      count: list.length,
      inline: list.filter(function (a) { return a && a.isInline; }).length,
      types: list.map(function (a) { return (a && a.attachmentType) || "unknown"; }),
      sizes: list.map(function (a) { return a && typeof a.size === "number" ? a.size : null; }),
    };
  }

  async function itemSummary(Office, item, opts) {
    const timeoutMs = (opts && opts.timeoutMs) || 10000;
    const mode = detectMode(item);
    const s = { mode: mode, itemType: item.itemType || "unknown" };
    if (mode === "read") {
      s.subject = { state: present(item.subject), length: (item.subject || "").length };
      s.from = present(item.from);
      s.to = Array.isArray(item.to) ? { state: "present", count: item.to.length } : { state: present(item.to) };
      s.cc = Array.isArray(item.cc) ? { state: "present", count: item.cc.length } : { state: present(item.cc) };
      s.dateTimeCreated = present(item.dateTimeCreated);
      s.internetMessageId = present(item.internetMessageId);
      s.conversationId = present(item.conversationId);
      s.itemId = present(item.itemId);
      s.attachments = attachmentInfo(item.attachments);
    } else if (mode === "compose") {
      const subj = await callAsync(item.subject.getAsync.bind(item.subject), [], timeoutMs);
      s.subject = subj.status === "succeeded" ? { state: present(subj.value), length: String(subj.value || "").length } : { state: "unknown", status: subj.status };
      s.from = "unknown";
      s.internetMessageId = "unknown";
      s.conversationId = present(item.conversationId);
      s.itemId = "unknown";
      if (typeof item.getAttachmentsAsync === "function") {
        const a = await callAsync(item.getAttachmentsAsync.bind(item), [], timeoutMs);
        s.attachments = a.status === "succeeded" ? attachmentInfo(a.value) : { status: a.status };
      } else {
        s.attachments = { status: "unsupported" };
      }
    }
    s.body = await bodyInfo(Office, item, timeoutMs);
    return s;
  }

  async function fingerprint(text, subtle) {
    if (!text) return null;
    const c = subtle || (typeof crypto !== "undefined" && crypto.subtle);
    if (!c) return "unavailable";
    const data = new TextEncoder().encode(String(text).trim().toLowerCase());
    const buf = await c.digest("SHA-256", data);
    return Array.from(new Uint8Array(buf)).slice(0, 4).map(function (b) { return b.toString(16).padStart(2, "0"); }).join("");
  }

  function maskAddress(addr) {
    if (!addr || addr.indexOf("@") < 1) return null;
    const parts = addr.split("@");
    return parts[0][0] + "***@" + parts[1];
  }

  async function accountInfo(Office, subtle) {
    const p = (Office.context.mailbox && Office.context.mailbox.userProfile) || {};
    return {
      accountType: p.accountType || "unknown",   // Mailbox 1.6: enterprise | office365 | outlookCom | gmail
      fingerprint: await fingerprint(p.emailAddress, subtle),
      maskedLabel: maskAddress(p.emailAddress),  // displayed in the pane only; excluded from reports
    };
  }

  // Tracks ItemChanged events so an in-flight capture can be invalidated (I03-T05/T13 precursor).
  function createItemTracker() {
    let generation = 0;
    return {
      onItemChanged: function () { generation += 1; },
      begin: function () { const g = generation; return { isStale: function () { return g !== generation; } }; },
      generation: function () { return generation; },
    };
  }

  function registerItemChanged(Office, tracker) {
    return new Promise(function (resolve) {
      const mb = Office.context.mailbox;
      if (!mb || typeof mb.addHandlerAsync !== "function" || !Office.EventType || !Office.EventType.ItemChanged) {
        resolve({ status: "unsupported" });
        return;
      }
      try {
        mb.addHandlerAsync(Office.EventType.ItemChanged, tracker.onItemChanged, function (r) {
          resolve({ status: r && r.status === "succeeded" ? "registered" : "failed" });
        });
      } catch (e) {
        resolve({ status: "threw" });
      }
    });
  }

  async function loopbackTest(fetchFn, targets, token, timeoutMs) {
    const results = [];
    for (const t of targets) {
      const ctl = typeof AbortController !== "undefined" ? new AbortController() : null;
      const timer = setTimeout(function () { if (ctl) ctl.abort(); }, timeoutMs || 5000);
      const opts = { method: "GET", cache: "no-store", signal: ctl ? ctl.signal : undefined };
      if (t.authed) opts.headers = { "X-Probe-Token": token || "" };
      const started = Date.now();
      try {
        const r = await fetchFn(t.url, opts);
        results.push({ id: t.id, url: t.url, outcome: r.ok ? "ok" : "http_" + r.status, ms: Date.now() - started });
      } catch (e) {
        results.push({ id: t.id, url: t.url, outcome: e && e.name === "AbortError" ? "timeout" : "error", error: e && e.name ? e.name + ": " + (e.message || "") : String(e), ms: Date.now() - started });
      } finally {
        clearTimeout(timer);
      }
    }
    return results;
  }

  // httpPort: plain-HTTP probe API on 127.0.0.1; httpsPort: HTTPS probe server on localhost (dev certificate).
  function loopbackTargets(httpPort, httpsPort) {
    const p = httpPort || 47802;
    const s = httpsPort || 47801;
    return [
      { id: "http-127-ping", url: "http://127.0.0.1:" + p + "/probe/ping?case=outlook-http" },
      { id: "https-localhost-ping", url: "https://localhost:" + s + "/probe/ping?case=outlook-https" },
      { id: "http-127-authed", url: "http://127.0.0.1:" + p + "/probe/authed?case=outlook-authed", authed: true },
      { id: "https-localhost-authed", url: "https://localhost:" + s + "/probe/authed?case=outlook-https-authed", authed: true },
    ];
  }

  // The report that may leave the pane: metadata only. maskedLabel is removed on purpose.
  function buildReport(parts) {
    const acct = Object.assign({}, parts.account || {});
    delete acct.maskedLabel;
    return {
      probe: "I01-outlook", version: PROBE_VERSION, at: new Date().toISOString(),
      environment: parts.environment, diagnostics: parts.diagnostics, requirementSets: parts.requirementSets,
      account: acct, item: parts.item, itemChanged: parts.itemChanged, loopback: parts.loopback || [],
      snapshotExport: parts.snapshotExport || null,
    };
  }

  // Snapshot-export probe (ADR-0006 path 1). Builds a schema-v1-like Markdown snapshot of a SYNTHETIC
  // test message so the real export path (size, encoding, download/clipboard) can be measured.
  // The pane only calls this after the user ticks "this is a synthetic test message".
  function buildTestSnapshot(o) {
    const capturedAt = o.capturedAt || new Date();
    const expiresAt = new Date(capturedAt.getTime() + (o.ttlMinutes || 30) * 60000);
    const esc = function (v) { return JSON.stringify(v === undefined ? null : v); };
    const body = String(o.bodyText || "").split(">>>").join("> > >");  // keep the untrusted block delimiter unambiguous
    return [
      "---",
      "schema_version: \"1.0-probe\"",
      "capture_id: " + esc(o.captureId || "i01-probe"),
      "source: " + esc(o.source || "unknown"),
      "account_label: " + esc(o.accountLabel || null),
      "subject: " + esc(o.subject === undefined ? null : o.subject),
      "captured_at: " + esc(capturedAt.toISOString()),
      "expires_at: " + esc(expiresAt.toISOString()),
      "coverage: " + esc(o.coverage || "unknown"),
      "warnings: " + JSON.stringify(o.warnings || ["probe_snapshot_not_production"]),
      "untrusted_content: true",
      "---",
      "",
      "<<<UNTRUSTED_EMAIL_CONTENT",
      body,
      ">>>",
      "",
    ].join("\n");
  }

  function utf8Bytes(text) { return new TextEncoder().encode(text).length; }

  const FORBIDDEN_CALLS = ["sendAsync", "saveAsync", "setAsync", "displayReplyForm", "displayReplyAllForm",
    "displayNewMessageForm", "moveAsync", "deleteAsync", "removeAttachmentAsync", "addFileAttachmentAsync"];

  return {
    PROBE_VERSION, MAILBOX_SETS, FORBIDDEN_CALLS, classifyEnvironment, requirementSets, diagnostics, detectMode,
    present, callAsync, bodyInfo, attachmentInfo, itemSummary, fingerprint, maskAddress, accountInfo,
    createItemTracker, registerItemChanged, loopbackTest, loopbackTargets, buildReport, buildTestSnapshot, utf8Bytes,
  };
});
