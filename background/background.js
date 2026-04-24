const CHATGPT_URL = "https://chatgpt.com/";
const CHATGPT_URL_PATTERNS = ["https://chatgpt.com/*", "https://chat.openai.com/*"];
const CHATGPT_URL_PREFIXES = ["https://chatgpt.com", "https://chat.openai.com"];
const MAX_IMAGE_BYTES = 12 * 1024 * 1024;
const CHATGPT_CONTENT_SCRIPT = "content-scripts/chatgpt.js";
const RECEIVER_RETRY_ATTEMPTS = 8;
const RECEIVER_RETRY_DELAY_MS = 500;
const PREFERRED_CHATGPT_TAB_KEY = "espPreferredChatGPTTabId";

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

function executeScript(tabId, files) {
  return new Promise((resolve, reject) => {
    chrome.scripting.executeScript({ target: { tabId }, files }, () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve();
    });
  });
}

function executeScriptFunction(tabId, func, args = []) {
  return new Promise((resolve, reject) => {
    chrome.scripting.executeScript({ target: { tabId }, func, args }, (injectionResults) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      if (!Array.isArray(injectionResults) || !injectionResults.length) {
        resolve(undefined);
        return;
      }

      resolve(injectionResults[0].result);
    });
  });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getTab(tabId) {
  return new Promise((resolve, reject) => {
    chrome.tabs.get(tabId, (tab) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(tab);
    });
  });
}

function queryTabs(queryInfo) {
  return new Promise((resolve) => {
    chrome.tabs.query(queryInfo, resolve);
  });
}

function createTab(createProperties) {
  return new Promise((resolve) => {
    chrome.tabs.create(createProperties, resolve);
  });
}

function storageGet(keys) {
  return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
}

function storageSet(values) {
  return new Promise((resolve) => chrome.storage.local.set(values, resolve));
}

async function getPreferredChatGPTTabId() {
  const data = await storageGet(PREFERRED_CHATGPT_TAB_KEY);
  return Number(data[PREFERRED_CHATGPT_TAB_KEY]) || null;
}

async function setPreferredChatGPTTabId(tabId) {
  if (!tabId) return;
  await storageSet({ [PREFERRED_CHATGPT_TAB_KEY]: tabId });
}

async function findChatGPTTab() {
  const tabs = await listChatGPTTabs();
  return tabs[0] || null;
}

function isChatGPTUrl(url) {
  if (!url) return false;
  return CHATGPT_URL_PREFIXES.some((prefix) => url.startsWith(prefix));
}

async function listChatGPTTabs() {
  const tabMap = new Map();

  for (const pattern of CHATGPT_URL_PATTERNS) {
    const tabs = await queryTabs({ url: pattern });
    tabs.forEach((tab) => {
      if (tab && tab.id && isChatGPTUrl(tab.url || "")) {
        tabMap.set(tab.id, tab);
      }
    });
  }

  const preferredTabId = await getPreferredChatGPTTabId();
  return Array.from(tabMap.values()).sort((a, b) => {
    const preferredA = a.id === preferredTabId ? 1 : 0;
    const preferredB = b.id === preferredTabId ? 1 : 0;
    if (preferredA !== preferredB) return preferredB - preferredA;

    const scoreA = (a.active ? 8 : 0) + (!a.discarded ? 4 : 0) + (a.status === "complete" ? 2 : 0);
    const scoreB = (b.active ? 8 : 0) + (!b.discarded ? 4 : 0) + (b.status === "complete" ? 2 : 0);
    if (scoreA !== scoreB) return scoreB - scoreA;
    return (b.lastAccessed || 0) - (a.lastAccessed || 0);
  });
}

async function waitForTabReady(tabId, timeoutMs = 20000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const tab = await getTab(tabId);
    if (tab && tab.status === "complete" && isChatGPTUrl(tab.url || "")) {
      return tab;
    }
    await delay(250);
  }

  throw new Error("Timed out waiting for ChatGPT tab to load.");
}

async function ensureReceiverReady(tabId) {
  let lastError = new Error("Could not establish connection.");
  for (let attempt = 0; attempt < RECEIVER_RETRY_ATTEMPTS; attempt += 1) {
    try {
      await executeScript(tabId, [CHATGPT_CONTENT_SCRIPT]);
      const isReady = await executeScriptFunction(tabId, () => Boolean(globalThis.__ESP_CHATGPT_READY__));
      if (!isReady) {
        throw new Error("ChatGPT helper script is not ready.");
      }
      return;
    } catch (error) {
      lastError = error;
      await delay(RECEIVER_RETRY_DELAY_MS);
    }
  }

  throw lastError;
}

async function openFreshChatGPTTab() {
  const newTab = await createTab({ url: CHATGPT_URL, active: false });
  if (!newTab || !newTab.id) {
    throw new Error("Could not open a new ChatGPT tab.");
  }

  await waitForTabReady(newTab.id);
  return newTab;
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";

  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }

  return btoa(binary);
}

async function fetchImageAsDataUrl(imageUrl) {
  if (!imageUrl) {
    throw new Error("No preview image URL was found on the edit page.");
  }

  const response = await fetch(imageUrl, { credentials: "include" });
  if (!response.ok) {
    throw new Error(`Unable to fetch image preview (${response.status}).`);
  }

  const contentLength = Number(response.headers.get("content-length") || "0");
  if (contentLength > MAX_IMAGE_BYTES) {
    throw new Error("Image preview is too large to send to ChatGPT.");
  }

  const contentType = response.headers.get("content-type") || "image/jpeg";
  if (!contentType.startsWith("image/")) {
    throw new Error(`Preview URL did not return an image (${contentType}).`);
  }

  const buffer = await response.arrayBuffer();
  if (buffer.byteLength > MAX_IMAGE_BYTES) {
    throw new Error("Image preview is too large to send to ChatGPT.");
  }

  return {
    dataUrl: `data:${contentType};base64,${arrayBufferToBase64(buffer)}`,
    mimeType: contentType,
  };
}

async function handleAltTextRequest(payload) {
  const chatgptTabs = await listChatGPTTabs();
  if (!chatgptTabs.length) {
    try {
      chatgptTabs.push(await openFreshChatGPTTab());
    } catch (error) {
      return {
        ok: false,
        error: `Open ChatGPT in another tab and make sure you are logged in. (${error.message})`,
      };
    }
  }

  let imagePayload = null;
  if (payload.action === "generate") {
    try {
      imagePayload = await fetchImageAsDataUrl(payload.imageUrl);
    } catch (error) {
      return {
        ok: false,
        error: error.message,
      };
    }
  }

  let lastError = new Error("Could not establish connection.");
  for (const tab of chatgptTabs) {
    if (!tab.id) {
      continue;
    }

    try {
      await waitForTabReady(tab.id);
      await ensureReceiverReady(tab.id);
      const response = await executeScriptFunction(
        tab.id,
        async (requestPayload) => {
          try {
            if (typeof globalThis.__ESP_CHATGPT_REQUEST_ALT_TEXT__ !== "function") {
              return {
                ok: false,
                error: "ChatGPT helper not available in tab.",
              };
            }

            return await globalThis.__ESP_CHATGPT_REQUEST_ALT_TEXT__(requestPayload);
          } catch (error) {
            return {
              ok: false,
              error: error?.message || String(error ?? "Unknown ChatGPT helper error."),
            };
          }
        },
        [{
          ...payload,
          imageDataUrl: imagePayload ? imagePayload.dataUrl : "",
          imageMimeType: imagePayload ? imagePayload.mimeType : "",
        }]
      );

      if (!response || !response.ok) {
        lastError = new Error(response && response.error ? response.error : "ChatGPT did not return alt text.");
        continue;
      }

      await setPreferredChatGPTTabId(tab.id);
      return response;
    } catch (error) {
      lastError = error;
    }
  }

  {
    return {
      ok: false,
      error: `Unable to communicate with the ChatGPT tab: ${lastError.message}`,
    };
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") {
    sendResponse({ ok: false, error: "Unknown message." });
    return false;
  }

  if (message.type === "esp:getChatGPTStatus") {
    findChatGPTTab()
      .then((tab) => {
        sendResponse({
          ok: true,
          available: Boolean(tab),
          tabId: tab ? tab.id : null,
        });
      })
      .catch((error) => {
        sendResponse({ ok: false, error: error.message });
      });
    return true;
  }

  if (message.type === "esp:openChatGPT") {
    findChatGPTTab()
      .then((tab) => {
        if (tab && tab.id) {
          setPreferredChatGPTTabId(tab.id).catch(() => {});
          chrome.tabs.update(tab.id, { active: true });
          if (tab.windowId) {
            chrome.windows.update(tab.windowId, { focused: true });
          }
          sendResponse({ ok: true, tabId: tab.id });
          return null;
        }

        return createTab({ url: CHATGPT_URL }).then((newTab) => {
          if (newTab && newTab.id) {
            setPreferredChatGPTTabId(newTab.id).catch(() => {});
          }
          sendResponse({ ok: true, tabId: newTab.id });
        });
      })
      .catch((error) => {
        sendResponse({ ok: false, error: error.message });
      });
    return true;
  }

  if (message.type === "esp:requestAltText") {
    handleAltTextRequest(message.payload || {})
      .then(sendResponse)
      .catch((error) => {
        sendResponse({ ok: false, error: error.message });
      });
    return true;
  }

  sendResponse({ ok: false, error: `Unhandled message type: ${message.type}` });
  return false;
});
