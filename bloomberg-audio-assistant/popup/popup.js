const STATE_KEY = "bcaState";
const elements = {};

document.addEventListener("DOMContentLoaded", () => {
  [
    "cms-status",
    "data-status",
    "scan",
    "start",
    "resume",
    "stop",
    "repair-index",
    "export-log",
    "run-status",
    "current-item",
    "count-stops",
    "count-tasks",
    "count-ready",
    "count-complete",
    "count-skipped",
    "count-failed",
    "log",
  ].forEach((id) => {
    elements[id] = document.getElementById(id);
  });

  elements.scan.addEventListener("click", () => sendToActiveTab({ type: "bca:scan" }));
  elements.start.addEventListener("click", () => sendToActiveTab({ type: "bca:start" }));
  elements.resume.addEventListener("click", () => sendToActiveTab({ type: "bca:resume" }));
  elements.stop.addEventListener("click", () => sendToActiveTab({ type: "bca:stop" }));
  elements["repair-index"].addEventListener("click", () => sendToActiveTab({ type: "bca:repairIndex" }));
  elements["export-log"].addEventListener("click", exportLog);

  refresh();
  setInterval(refresh, 1000);
});

function storageGet(keys) {
  return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
}

function queryActiveTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => resolve(tabs[0] || null));
  });
}

function runtimeMessage(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(response);
    });
  });
}

function isBloombergCmsTab(tab) {
  return Boolean(tab && tab.url && tab.url.startsWith("https://cms.bloombergconnects.org/"));
}

function setDot(element, text, className) {
  element.textContent = text;
  element.className = `dot-status ${className || ""}`.trim();
}

async function refresh() {
  const [tab, data] = await Promise.all([queryActiveTab(), storageGet(STATE_KEY)]);

  if (isBloombergCmsTab(tab)) {
    setDot(elements["cms-status"], "Bloomberg CMS tab is active", "success");
    setButtonsDisabled(false);
  } else {
    setDot(elements["cms-status"], "Open Bloomberg Connects CMS", "error");
    setButtonsDisabled(true);
    elements["export-log"].disabled = false;
  }

  const state = data[STATE_KEY] || {};
  const transcriptCount = state.transcriptCount || 0;
  setDot(
    elements["data-status"],
    transcriptCount ? `${transcriptCount} transcript sections loaded` : "Transcript data not loaded yet",
    transcriptCount ? "success" : "warning"
  );
  renderState(state);
}

function renderState(state) {
  const safeState = state || {};
  const counts = safeState.counts || {};
  const status = safeState.status || "idle";
  const phase = safeState.phase || "idle";
  const index = Number.isInteger(safeState.index) ? safeState.index : 0;
  const workTotal = Array.isArray(safeState.workItems) ? safeState.workItems.length : 0;

  elements["run-status"].textContent =
    status === "idle"
      ? "Idle."
      : `${status} (${phase}) - ${Math.min(index, workTotal)} of ${workTotal} tasks processed.`;

  if (safeState.current && safeState.current.label) {
    elements["current-item"].textContent = safeState.current.label;
  } else {
    elements["current-item"].textContent = "";
  }

  elements["count-stops"].textContent = counts.eligibleStops || 0;
  elements["count-tasks"].textContent = counts.totalLanguageTasks || 0;
  elements["count-ready"].textContent = counts.ready || 0;
  elements["count-complete"].textContent = counts.processed || counts.alreadyComplete || 0;
  elements["count-skipped"].textContent = counts.skipped || 0;
  elements["count-failed"].textContent = counts.failed || counts.missingTranscript || 0;

  elements.stop.disabled = status !== "running";
  elements.resume.disabled = !["paused", "stopped"].includes(status);

  const logEntries = Array.isArray(safeState.log) ? safeState.log.slice(-20).reverse() : [];
  elements.log.textContent = "";
  if (!logEntries.length) {
    const item = document.createElement("li");
    item.textContent = "No activity yet.";
    elements.log.appendChild(item);
    return;
  }

  logEntries.forEach((entry) => {
    const item = document.createElement("li");
    if (entry.itemLabel) {
      const strong = document.createElement("strong");
      strong.textContent = `${entry.itemLabel}: `;
      item.appendChild(strong);
    }
    item.appendChild(document.createTextNode(entry.message || ""));
    elements.log.appendChild(item);
  });
}

function setButtonsDisabled(disabled) {
  ["scan", "start", "resume", "stop", "repair-index"].forEach((id) => {
    elements[id].disabled = disabled;
  });
}

async function sendToActiveTab(payload) {
  setButtonsDisabled(true);
  try {
    const response = await runtimeMessage({ type: "bca:sendToActiveTab", payload });
    if (!response || !response.ok) {
      throw new Error(response && response.error ? response.error : "Extension command failed.");
    }
    renderState(response.state);
  } catch (error) {
    elements["run-status"].textContent = error.message;
  } finally {
    setButtonsDisabled(false);
    refresh();
  }
}

async function exportLog() {
  elements["export-log"].disabled = true;
  try {
    const response = await runtimeMessage({ type: "bca:exportLog" });
    if (!response || !response.ok) {
      throw new Error(response && response.error ? response.error : "Could not export log.");
    }
  } catch (error) {
    elements["run-status"].textContent = error.message;
  } finally {
    elements["export-log"].disabled = false;
  }
}
