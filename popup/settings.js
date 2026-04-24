const STATE_KEY = "espAltTextAssistantState";
const DIRECTORY_MARKER = "/cp/files/directory/6";

const elements = {};

document.addEventListener("DOMContentLoaded", () => {
  [
    "cms-status",
    "chatgpt-status",
    "open-chatgpt",
    "scan",
    "start",
    "stop",
    "run-status",
    "current-item",
    "count-images",
    "count-skip",
    "count-shorten",
    "count-generate",
    "count-saved",
    "count-failed",
    "log",
  ].forEach((id) => {
    elements[id] = document.getElementById(id);
  });

  elements.scan.addEventListener("click", () => sendToActiveEeTab("esp:scanDirectory"));
  elements.start.addEventListener("click", () => sendToActiveEeTab("esp:startRun"));
  elements.stop.addEventListener("click", () => sendToActiveEeTab("esp:stopRun"));
  elements["open-chatgpt"].addEventListener("click", openChatGPT);

  refresh();
  setInterval(refresh, 1000);
});

function storageGet(keys) {
  return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
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

function queryActiveTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      resolve(tabs[0] || null);
    });
  });
}

function queryEasternStateTabs() {
  return new Promise((resolve) => {
    chrome.tabs.query({ url: "https://easternstate.org/*" }, (tabs) => {
      resolve(Array.isArray(tabs) ? tabs : []);
    });
  });
}

function sendTabMessage(tabId, message) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      resolve(response);
    });
  });
}

function isEasternStateTab(tab) {
  return Boolean(tab && tab.url && tab.url.startsWith("https://easternstate.org/"));
}

function isDirectoryTab(tab) {
  return Boolean(
    isEasternStateTab(tab) && decodeURIComponent(tab.url).includes(DIRECTORY_MARKER)
  );
}

function setDot(element, text, className) {
  element.textContent = text;
  element.className = `dot-status ${className || ""}`.trim();
}

async function refresh() {
  const [tab, data, chatgpt] = await Promise.all([
    queryActiveTab(),
    storageGet(STATE_KEY),
    runtimeMessage({ type: "esp:getChatGPTStatus" }).catch((error) => ({ ok: false, error: error.message })),
  ]);

  if (isDirectoryTab(tab)) {
    setDot(elements["cms-status"], "General directory is active", "success");
    elements.scan.disabled = false;
    elements.start.disabled = false;
  } else if (isEasternStateTab(tab)) {
    setDot(elements["cms-status"], "Open the General file directory", "warning");
    elements.scan.disabled = true;
    elements.start.disabled = true;
  } else {
    setDot(elements["cms-status"], "Open Eastern State CMS", "error");
    elements.scan.disabled = true;
    elements.start.disabled = true;
  }

  if (chatgpt.ok && chatgpt.available) {
    setDot(elements["chatgpt-status"], "ChatGPT tab is open", "success");
  } else {
    setDot(elements["chatgpt-status"], "ChatGPT tab is needed", "warning");
  }

  renderState(data[STATE_KEY]);
}

function renderState(state) {
  const safeState = state || {};
  const counts = safeState.counts || {};
  const status = safeState.status || "idle";
  const phase = safeState.phase || "idle";
  const workTotal = Array.isArray(safeState.workItems) ? safeState.workItems.length : 0;
  const index = Number.isInteger(safeState.index) ? safeState.index : 0;

  elements["run-status"].textContent =
    status === "idle"
      ? "Idle."
      : `${status} (${phase}) - ${Math.min(index, workTotal)} of ${workTotal} change candidates processed.`;

  if (safeState.current && safeState.current.fileName) {
    elements["current-item"].textContent = `${safeState.current.fileName} - ${safeState.current.action || phase}`;
  } else if (workTotal && status === "scanned") {
    elements["current-item"].textContent = `${workTotal} files need shortening or generated alt text.`;
  } else {
    elements["current-item"].textContent = "";
  }

  elements["count-images"].textContent = counts.imageRows || 0;
  elements["count-skip"].textContent = counts.skip || 0;
  elements["count-shorten"].textContent = counts.shorten || 0;
  elements["count-generate"].textContent = counts.generate || 0;
  elements["count-saved"].textContent = counts.saved || 0;
  elements["count-failed"].textContent = counts.failed || 0;

  elements.stop.disabled = status !== "running";

  const logEntries = Array.isArray(safeState.log) ? safeState.log.slice(-20).reverse() : [];
  elements.log.innerHTML = "";

  if (!logEntries.length) {
    const item = document.createElement("li");
    item.textContent = "No activity yet.";
    elements.log.appendChild(item);
    return;
  }

  logEntries.forEach((entry) => {
    const item = document.createElement("li");
    const file = entry.fileName ? `<strong>${escapeHtml(entry.fileName)}</strong>: ` : "";
    item.innerHTML = `${file}${escapeHtml(entry.message || "")}`;
    elements.log.appendChild(item);
  });
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

async function sendToActiveEeTab(type) {
  const activeTab = await queryActiveTab();
  const requiresDirectory = type !== "esp:stopRun";
  let targetTab = activeTab;

  if (!requiresDirectory && !isEasternStateTab(targetTab)) {
    const easternStateTabs = await queryEasternStateTabs();
    const fallbackEasternTab = easternStateTabs
      .slice()
      .sort((a, b) => {
        const activeDelta = Number(Boolean(b && b.active)) - Number(Boolean(a && a.active));
        if (activeDelta) return activeDelta;
        return (b && b.lastAccessed ? b.lastAccessed : 0) - (a && a.lastAccessed ? a.lastAccessed : 0);
      })[0] || null;
    targetTab = easternStateTabs.find((tab) => isDirectoryTab(tab)) || fallbackEasternTab;
  }

  const isAllowedTab = requiresDirectory ? isDirectoryTab(targetTab) : isEasternStateTab(targetTab);

  if (!isAllowedTab) {
    setDot(
      elements["cms-status"],
      requiresDirectory ? "Open the General file directory" : "Open an Eastern State CMS tab",
      "error"
    );
    return;
  }

  setButtonsDisabled(true);

  try {
    const response = await sendTabMessage(targetTab.id, { type });
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

function setButtonsDisabled(disabled) {
  elements.scan.disabled = disabled;
  elements.start.disabled = disabled;
  elements.stop.disabled = disabled;
  elements["open-chatgpt"].disabled = disabled;
}

async function openChatGPT() {
  elements["open-chatgpt"].disabled = true;
  try {
    await runtimeMessage({ type: "esp:openChatGPT" });
  } finally {
    elements["open-chatgpt"].disabled = false;
    refresh();
  }
}
