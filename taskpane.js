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
    const item = Office.context.mailbox.item;
    const isRead = I01Probe.detectMode(item) === "read";
    const subject = isRead ? item.subject : (await I01Probe.callAsync(item.subject.getAsync.bind(item.subject), [], 10000)).value;
    const body = await I01Probe.callAsync(item.body.getAsync.bind(item.body), [Office.CoercionType.Text], 10000);
    const acct = state.parts.account || {};
    const md = I01Probe.buildTestSnapshot({ subject: subject, bodyText: body.value || "", accountLabel: acct.maskedLabel,
      source: acct.accountType === "outlookCom" ? "outlook_personal" : (acct.accountType === "unknown" ? "unknown" : "outlook365"),
      coverage: "unknown" });
    const result = { mode: mode, bytes: I01Probe.utf8Bytes(md), bodyStatus: body.status, outcome: "pending" };
    try {
      if (mode === "download") {
        const a = document.createElement("a");
        a.href = URL.createObjectURL(new Blob([md], { type: "text/markdown;charset=utf-8" }));
        a.download = "i01-test-snapshot.md";
        document.body.appendChild(a); a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
        result.outcome = "download_triggered_verify_file";
      } else {
        await navigator.clipboard.writeText(md);
        result.outcome = "clipboard_ok";
      }
    } catch (e) { result.outcome = "failed: " + e.name; }
    state.parts.snapshotExport = result;
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
