// I01 Outlook probe task pane glue. All logic lives in probe-core.js (unit-tested).
/* global Office, I01Probe */
"use strict";

(function () {
  const state = { parts: {}, tracker: I01Probe.createItemTracker() };
  const $ = (id) => document.getElementById(id);
  const show = () => { $("out").textContent = JSON.stringify(I01Probe.buildReport(state.parts), null, 2); };

  async function runProbe() {
    const env = I01Probe.classifyEnvironment(typeof Office !== "undefined" ? Office : null);
    state.parts.environment = env;
    if (!env.ok) { show(); return; }
    state.parts.diagnostics = I01Probe.diagnostics(Office);
    state.parts.requirementSets = I01Probe.requirementSets(Office);
    const acct = await I01Probe.accountInfo(Office);
    $("account").textContent = (acct.maskedLabel || "unknown") + " (" + acct.accountType + ", fp " + acct.fingerprint + ")";
    state.parts.account = acct;
    const ticket = state.tracker.begin();
    const item = await I01Probe.itemSummary(Office, Office.context.mailbox.item, { timeoutMs: 10000 });
    item.staleDuringRead = ticket.isStale();
    item.itemChangedGeneration = state.tracker.generation();
    state.parts.item = item;
    show();
  }

  async function runLoopback() {
    const q = new URLSearchParams(location.search);
    const targets = I01Probe.loopbackTargets(Number(q.get("probePort") || 47802), Number(q.get("probeHttpsPort") || 47801));
    state.parts.loopback = await I01Probe.loopbackTest(fetch.bind(window), targets, $("token").value, 6000);
    state.parts.paneOrigin = location.origin;
    show();
  }

  async function copyReport() {
    const text = $("out").textContent;
    try { await navigator.clipboard.writeText(text); $("copy").textContent = "Copied"; }
    catch (e) { const r = document.createRange(); r.selectNodeContents($("out")); const s = getSelection(); s.removeAllRanges(); s.addRange(r); $("copy").textContent = "Select-all done: press Ctrl+C"; }
  }

  // Tests the explicit-file fallback (ADR-0006): can the pane hand the user a file?
  function downloadReport() {
    try {
      const blob = new Blob([$("out").textContent], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "i01-outlook-probe-report.json";
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
      $("download").textContent = "Download triggered (check Downloads)";
    } catch (e) { $("download").textContent = "Download failed: " + e.name; }
  }

  // Builds a full Markdown snapshot of the current (synthetic) message. Content never enters the report;
  // only the outcome and byte size do.
  async function testSnapshot(mode) {
    const acct = state.parts.account || {};
    const prepared = await I01Probe.prepareSnapshot(Office, state.tracker, {
      syntheticConfirmed: $("synthetic").checked, timeoutMs: 10000, parser: new DOMParser(), accountLabel: acct.maskedLabel,
      source: acct.accountType === "outlookCom" ? "outlook_personal" : (acct.accountType === "office365" || acct.accountType === "enterprise" ? "outlook365" : "unknown"),
    });
    const result = { mode: mode, ok: prepared.ok };
    if (!prepared.ok) {
      result.outcome = "not_exported";
      result.error = prepared.error;
      result.detail = prepared.detail;
    } else {
      result.bytes = prepared.bytes;
      result.bodySource = prepared.bodySource;
      Object.assign(result, await I01Probe.deliverSnapshot(prepared.markdown, mode,
        { document: document, URL: URL, Blob: Blob, clipboard: navigator.clipboard, setTimeout: setTimeout }));
    }
    state.parts.snapshotExport = result;  // outcome and size only; never the snapshot text
    show();
  }

  async function sendReport() {
    const port = Number(new URLSearchParams(location.search).get("probePort") || 47802);
    const report = I01Probe.buildReport(state.parts);
    try {
      const r = await fetch("http://127.0.0.1:" + port + "/probe/report", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ case: "outlook-report", outcome: "report", detail: JSON.stringify(report).slice(0, 380) }),
      });
      $("send").textContent = "Sent (HTTP " + r.status + ")";
    } catch (e) { $("send").textContent = "Send failed: " + e.name; }
  }

  Office.onReady(async function () {
    state.parts.itemChanged = await I01Probe.registerItemChanged(Office, state.tracker);
    $("run").onclick = runProbe;
    $("loopback").onclick = runLoopback;
    $("copy").onclick = copyReport;
    $("download").onclick = downloadReport;
    $("synthetic").onchange = () => { $("snapDownload").disabled = $("snapCopy").disabled = !$("synthetic").checked; };
    $("snapDownload").onclick = () => testSnapshot("download");
    $("snapCopy").onclick = () => testSnapshot("clipboard");
    $("send").onclick = sendReport;
    runProbe();
  });
})();
