const STATE_KEY = "bcaState";
const ENABLE_DEBUGGER_INPUT = false;
const DEBUGGER_INPUT_DISABLED_MESSAGE = "Debugger-backed input is disabled for stability.";

function queryActiveTab() {
  return new Promise((resolve) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => resolve(tabs[0] || null));
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

async function focusTabAndWindow(tab) {
  if (!tab || !tab.id) throw new Error("Could not identify the CMS tab to focus.");
  if (tab.windowId) {
    await new Promise((resolve, reject) => {
      chrome.windows.update(tab.windowId, { focused: true }, () => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }
        resolve();
      });
    });
  }
  await new Promise((resolve, reject) => {
    chrome.tabs.update(tab.id, { active: true }, () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve();
    });
  });
}

function storageGet(keys) {
  return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
}

function downloadJson(filename, value) {
  const json = JSON.stringify(value, null, 2);
  const url = `data:application/json;charset=utf-8,${encodeURIComponent(json)}`;
  chrome.downloads.download(
    {
      url,
      filename,
      saveAs: true,
      conflictAction: "uniquify",
    },
    () => {}
  );
}

function debuggerAttach(target) {
  return new Promise((resolve, reject) => {
    chrome.debugger.attach(target, "1.3", () => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve();
    });
  });
}

function debuggerDetach(target) {
  return new Promise((resolve) => {
    chrome.debugger.detach(target, () => resolve());
  });
}

function debuggerSendCommand(target, method, params) {
  return new Promise((resolve, reject) => {
    chrome.debugger.sendCommand(target, method, params, (result) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      resolve(result);
    });
  });
}

async function trustedTabClick(tabId, x, y) {
  const target = { tabId };
  const params = {
    x,
    y,
    button: "left",
    clickCount: 1,
  };
  await debuggerAttach(target);
  try {
    await debuggerSendCommand(target, "Input.dispatchMouseEvent", { ...params, type: "mousePressed" });
    await debuggerSendCommand(target, "Input.dispatchMouseEvent", { ...params, type: "mouseReleased" });
  } finally {
    await debuggerDetach(target);
  }
}

async function trustedTabType(tabId, text) {
  const target = { tabId };
  await debuggerAttach(target);
  try {
    await debuggerSendCommand(target, "Input.insertText", { text: String(text || "") });
  } finally {
    await debuggerDetach(target);
  }
}

async function trustedTabInsertText(tabId, text) {
  const target = { tabId };
  await debuggerAttach(target);
  try {
    await debuggerSendCommand(target, "Input.insertText", { text: String(text || "") });
  } finally {
    await debuggerDetach(target);
  }
}

async function trustedTabPressKey(tabId, key) {
  const target = { tabId };
  const keyCodeMap = {
    ArrowDown: 40,
    Enter: 13,
    Home: 36,
    Space: 32,
  };
  const codeMap = {
    ArrowDown: "ArrowDown",
    Enter: "Enter",
    Home: "Home",
    Space: "Space",
  };
  const keyValueMap = {
    Space: " ",
  };
  const windowsVirtualKeyCode = keyCodeMap[key] || 0;
  const code = codeMap[key] || key;
  const keyValue = keyValueMap[key] || key;
  await debuggerAttach(target);
  try {
    await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
      type: "rawKeyDown",
      key: keyValue,
      code,
      windowsVirtualKeyCode,
      nativeVirtualKeyCode: windowsVirtualKeyCode,
    });
    await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
      type: "keyUp",
      key: keyValue,
      code,
      windowsVirtualKeyCode,
      nativeVirtualKeyCode: windowsVirtualKeyCode,
    });
  } finally {
    await debuggerDetach(target);
  }
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function trustedTabLanguagePickerFlow(tabId, x, y, text) {
  const target = { tabId };
  const clickParams = {
    x,
    y,
    button: "left",
    clickCount: 1,
  };
  await debuggerAttach(target);
  try {
    if (Number.isFinite(x) && Number.isFinite(y)) {
      await debuggerSendCommand(target, "Input.dispatchMouseEvent", { ...clickParams, type: "mousePressed" });
      await debuggerSendCommand(target, "Input.dispatchMouseEvent", { ...clickParams, type: "mouseReleased" });
      await wait(120);
    }
    for (const char of String(text || "")) {
      const upper = char.toUpperCase();
      const code = /^[a-z]$/i.test(char) ? `Key${upper}` : char === " " ? "Space" : "";
      await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
        type: "rawKeyDown",
        key: char,
        code,
        text: char,
        unmodifiedText: char,
      });
      await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
        type: "char",
        key: char,
        code,
        text: char,
        unmodifiedText: char,
      });
      await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
        type: "keyUp",
        key: char,
        code,
      });
    }
    await wait(500);
    await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
      type: "rawKeyDown",
      key: "Enter",
      code: "Enter",
      windowsVirtualKeyCode: 13,
      nativeVirtualKeyCode: 13,
    });
    await debuggerSendCommand(target, "Input.dispatchKeyEvent", {
      type: "keyUp",
      key: "Enter",
      code: "Enter",
      windowsVirtualKeyCode: 13,
      nativeVirtualKeyCode: 13,
    });
  } finally {
    await debuggerDetach(target);
  }
}

async function mainWorldSelectAddTranslationLanguage(tabId, languageLabel, query) {
  const [result] = await chrome.scripting.executeScript({
    target: { tabId },
    world: "MAIN",
    args: [languageLabel, query],
    func: (targetLanguageLabel, targetQuery) => {
      const normalizeSpaces = (value) => String(value || "").replace(/\s+/g, " ").trim();
      const normalizeComparable = (value) =>
        normalizeSpaces(value)
          .toLowerCase()
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "")
          .replace(/[^a-z0-9]+/g, " ")
          .trim();
      const readValue = (element) => {
        try {
          return element && "value" in element ? element.value || "" : "";
        } catch (_) {
          return "";
        }
      };
      const controlText = (element) =>
        normalizeSpaces(element && (readValue(element) || element.textContent || element.getAttribute?.("aria-label") || element.getAttribute?.("placeholder") || ""));
      const visibleElements = (selector, root = document) =>
        Array.from(root.querySelectorAll(selector)).filter((element) => {
          const rect = element.getBoundingClientRect();
          const style = window.getComputedStyle(element);
          return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
        });
      const getReactFiber = (node) => {
        if (!(node instanceof Element)) return null;
        const key = Object.keys(node).find((entry) => entry.startsWith("__reactFiber$") || entry.startsWith("__reactInternalInstance$"));
        return key ? node[key] : null;
      };
      const flattenOptions = (options) => {
        const result = [];
        const visit = (entry) => {
          if (!entry) return;
          if (Array.isArray(entry)) {
            entry.forEach(visit);
            return;
          }
          if (Array.isArray(entry.options)) {
            entry.options.forEach(visit);
            return;
          }
          if (typeof entry === "object") result.push(entry);
        };
        visit(options);
        return result;
      };
      const optionLabel = (option) => normalizeSpaces(option && (option.label || option.name || option.title || readValue(option) || ""));
      const findReactSelectComponent = (node) => {
        let fiber = getReactFiber(node);
        while (fiber) {
          const props = fiber.memoizedProps || fiber.pendingProps || {};
          const selectProps = props.selectProps || {};
          if (Array.isArray(props.options) && typeof props.onChange === "function") {
            return { fiber, props, source: "props" };
          }
          if (Array.isArray(selectProps.options) && typeof selectProps.onChange === "function") {
            return { fiber, props: selectProps, source: "selectProps" };
          }
          fiber = fiber.return;
        }
        return null;
      };
      const collectReactSelectComponents = (root, input, control) => {
        const anchors = [
          input,
          control,
          document.activeElement,
          ...Array.from(
            root.querySelectorAll(
              "input[id^='react-select-'][id$='-input'], input.docent-select__input, input[class*='select__input'], [class*='docent-select'], [class*='select'], [class*='Select'], [class*='control'], [class*='Control'], [role='combobox'], [aria-haspopup='listbox']"
            )
          ),
          ...Array.from(root.querySelectorAll("*")).slice(0, 500),
        ].filter((element) => element instanceof Element);
        const seen = new Set();
        const components = [];
        for (const anchor of anchors) {
          const component = findReactSelectComponent(anchor);
          if (!component || !component.props || seen.has(component.props)) continue;
          seen.add(component.props);
          components.push(component);
        }
        return components;
      };
      const languageComparableCandidates = (languageLabel, query) => {
        const target = normalizeComparable(languageLabel);
        const queryComparable = normalizeComparable(query);
        const values = new Set([target, queryComparable].filter(Boolean));
        if (target.includes("portuguese") || queryComparable.includes("portuguese")) values.add("portugues");
        if (target.includes("spanish") || queryComparable.includes("spanish")) values.add("espanol");
        if (target.includes("french") || queryComparable.includes("french")) values.add("francais");
        if (target.includes("german") || queryComparable.includes("german")) values.add("deutsch");
        if (target.includes("italian") || queryComparable.includes("italian")) values.add("italiano");
        return Array.from(values);
      };
      const findMatchingOption = (options, languageLabel, query) => {
        const candidates = languageComparableCandidates(languageLabel, query);
        return (
          options.find((option) => candidates.includes(normalizeComparable(optionLabel(option)))) ||
          options.find((option) => candidates.some((candidate) => normalizeComparable(optionLabel(option)).includes(candidate))) ||
          options.find((option) => candidates.some((candidate) => candidate.includes(normalizeComparable(optionLabel(option))))) ||
          null
        );
      };
      const dispatchPointClick = (element) => {
        if (!(element instanceof Element)) return false;
        element.scrollIntoView?.({ block: "center", inline: "center" });
        const rect = element.getBoundingClientRect();
        if (!rect.width || !rect.height) return false;
        const x = rect.left + rect.width / 2;
        const y = rect.top + rect.height / 2;
        const hit = document.elementFromPoint(x, y) || element;
        const init = { bubbles: true, cancelable: true, composed: true, view: window, clientX: x, clientY: y, button: 0 };
        hit.dispatchEvent(new PointerEvent("pointerdown", { ...init, pointerId: 1, pointerType: "mouse", isPrimary: true }));
        hit.dispatchEvent(new MouseEvent("mousedown", init));
        hit.dispatchEvent(new PointerEvent("pointerup", { ...init, pointerId: 1, pointerType: "mouse", isPrimary: true }));
        hit.dispatchEvent(new MouseEvent("mouseup", init));
        hit.dispatchEvent(new MouseEvent("click", init));
        return true;
      };
      const setNativeValue = (input, value) => {
        if (!input) return false;
        const descriptor = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(input), "value");
        const previous = readValue(input);
        if (descriptor && descriptor.set) descriptor.set.call(input, value);
        else input.value = value;
        if (input._valueTracker && typeof input._valueTracker.setValue === "function") {
          input._valueTracker.setValue(previous);
        }
        return true;
      };
      const sendKey = (target, key, code = key, extra = {}) => {
        const init = { bubbles: true, cancelable: true, composed: true, key, code, ...extra };
        target.dispatchEvent(new KeyboardEvent("keydown", init));
        target.dispatchEvent(new KeyboardEvent("keypress", init));
        target.dispatchEvent(new KeyboardEvent("keyup", init));
      };
      const typeQuery = (input, text) => {
        if (!(input instanceof Element)) return false;
        input.focus?.();
        setNativeValue(input, "");
        input.dispatchEvent(new Event("input", { bubbles: true, cancelable: true, composed: true }));
        let composed = "";
        for (const char of text) {
          composed += char;
          sendKey(input, char, /^[a-z]$/i.test(char) ? `Key${char.toUpperCase()}` : "", { keyCode: char.charCodeAt(0), which: char.charCodeAt(0) });
          setNativeValue(input, composed);
          input.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, composed: true, data: char, inputType: "insertText" }));
          input.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, composed: true, data: char, inputType: "insertText" }));
        }
        return true;
      };

      const dialog =
        visibleElements('[role="dialog"], [class*="modal"], [class*="Dialog"], div')
          .filter((element) => {
            const text = normalizeComparable(element.textContent);
            return text.includes("add translation") && text.includes("choose a language");
          })
          .sort((a, b) => a.textContent.length - b.textContent.length)[0] || document;
      const input =
        dialog.querySelector("input[id^='react-select-'][id$='-input'], input.docent-select__input, input[class*='select__input'], input[role='combobox']") ||
        document.activeElement;
      const control =
        input?.closest?.("[class*='select__control'], [class*='Select__control'], [class*='control'], [class*='Control']") ||
        dialog.querySelector("[class*='select__control'], [class*='Select__control'], [class*='control'], [class*='Control']");
      const addButton = () =>
        visibleElements("button, [role='button']", dialog).find((button) => normalizeComparable(controlText(button)) === "add");
      const isReady = () => {
        const button = addButton();
        return Boolean(button && !button.disabled && button.getAttribute("aria-disabled") !== "true");
      };

      const anchors = collectReactSelectComponents(dialog, input, control);
      const componentDebug = [];
      for (const component of anchors) {
        const options = flattenOptions(component.props.options);
        componentDebug.push(`${component.source}:${options.length}:${options.slice(0, 8).map(optionLabel).join("|")}`);
        const option = findMatchingOption(options, targetLanguageLabel, targetQuery);
        if (option) {
          const handler = component.props.onChange;
          const context = component.fiber.stateNode || component.fiber.memoizedProps;
          if (handler && context) {
            handler.call(context, option, { action: "select-option", option, name: component.props.name });
          } else if (handler) {
            handler(option, { action: "select-option", option, name: component.props.name });
          }
          return { ok: true, method: "main-world-react-onChange", debug: `selector-v45-add-translation-confirm; ${componentDebug.join("; ")}` };
        }
      }

      if (control) {
        dispatchPointClick(control);
      }
      if (input instanceof Element) {
        input.focus?.();
        dispatchPointClick(input);
        sendKey(input, "ArrowDown", "ArrowDown", { keyCode: 40, which: 40 });
        typeQuery(input, targetQuery);
        const option =
          visibleElements("[role='option'], [role='menuitem'], li, div", document)
            .filter((element) => {
              const text = normalizeComparable(controlText(element));
              return text.includes(normalizeComparable(targetLanguageLabel)) || text.includes(normalizeComparable(targetQuery));
            })
            .sort((a, b) => controlText(a).length - controlText(b).length)[0] || null;
        if (option) {
          dispatchPointClick(option);
        return { ok: true, method: "main-world-option-click-dispatched", debug: `selector-v45-add-translation-confirm; ${controlText(option)}` };
        }
        sendKey(input, "Enter", "Enter", { keyCode: 13, which: 13 });
        return { ok: true, method: "main-world-type-enter-dispatched", debug: `selector-v45-add-translation-confirm; value=${readValue(input)}; ready=${isReady()}` };
      }

      return {
        ok: false,
        method: "none",
        debug: `selector-v45-add-translation-confirm; input=${input instanceof Element ? `${input.tagName.toLowerCase()}#${input.id || ""}.${String(input.className || "").split(/\s+/).slice(0, 2).join(".")} value=${readValue(input)} expanded=${input.getAttribute("aria-expanded") || ""}` : "none"}; control=${control ? controlText(control).slice(0, 80) : "none"}; ready=${isReady()}; components=${componentDebug.join("; ") || "none"}`,
      };
    },
  });
  return result && result.result ? result.result : { ok: false, debug: "No main-world script result." };
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") {
    return false;
  }

  if (message.type === "bca:sendToActiveTab") {
    queryActiveTab()
      .then((tab) => {
        if (!tab || !tab.id || !String(tab.url || "").startsWith("https://cms.bloombergconnects.org/")) {
          throw new Error("Open a Bloomberg Connects CMS tab first.");
        }
        return sendTabMessage(tab.id, message.payload);
      })
      .then((response) => sendResponse(response))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:focusCmsTab") {
    Promise.resolve()
      .then(() => {
        const tab = sender.tab;
        if (!tab || !tab.id || !String(tab.url || "").startsWith("https://cms.bloombergconnects.org/")) {
          throw new Error("Could not identify this Bloomberg Connects CMS tab.");
        }
        return focusTabAndWindow(tab);
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:exportLog") {
    storageGet(STATE_KEY)
      .then((data) => {
        const stamp = new Date().toISOString().replace(/[:.]/g, "-");
        downloadJson(`bloomberg-audio-assistant-log-${stamp}.json`, data[STATE_KEY] || {});
        sendResponse({ ok: true });
      })
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:trustedClick") {
    if (!ENABLE_DEBUGGER_INPUT) {
      sendResponse({ ok: false, error: DEBUGGER_INPUT_DISABLED_MESSAGE });
      return false;
    }
    Promise.resolve()
      .then(() => {
        const tabId = sender.tab && sender.tab.id;
        if (!tabId) throw new Error("Could not identify the CMS tab for trusted click.");
        const x = Number(message.x);
        const y = Number(message.y);
        if (!Number.isFinite(x) || !Number.isFinite(y)) throw new Error("Invalid trusted click coordinates.");
        return trustedTabClick(tabId, x, y);
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:trustedType") {
    if (!ENABLE_DEBUGGER_INPUT) {
      sendResponse({ ok: false, error: DEBUGGER_INPUT_DISABLED_MESSAGE });
      return false;
    }
    Promise.resolve()
      .then(() => {
        const tabId = sender.tab && sender.tab.id;
        if (!tabId) throw new Error("Could not identify the CMS tab for trusted typing.");
        const text = String(message.text || "");
        if (!text) throw new Error("Missing trusted typing text.");
        return trustedTabType(tabId, text);
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:trustedInsertText") {
    if (!ENABLE_DEBUGGER_INPUT) {
      sendResponse({ ok: false, error: DEBUGGER_INPUT_DISABLED_MESSAGE });
      return false;
    }
    Promise.resolve()
      .then(() => {
        const tabId = sender.tab && sender.tab.id;
        if (!tabId) throw new Error("Could not identify the CMS tab for trusted insertText.");
        const text = String(message.text || "");
        if (!text) throw new Error("Missing trusted insertText text.");
        return trustedTabInsertText(tabId, text);
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:trustedPressKey") {
    if (!ENABLE_DEBUGGER_INPUT) {
      sendResponse({ ok: false, error: DEBUGGER_INPUT_DISABLED_MESSAGE });
      return false;
    }
    Promise.resolve()
      .then(() => {
        const tabId = sender.tab && sender.tab.id;
        if (!tabId) throw new Error("Could not identify the CMS tab for trusted key press.");
        const key = String(message.key || "");
        if (!key) throw new Error("Missing key for trusted key press.");
        return trustedTabPressKey(tabId, key);
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:trustedLanguagePickerFlow") {
    if (!ENABLE_DEBUGGER_INPUT) {
      sendResponse({ ok: false, error: DEBUGGER_INPUT_DISABLED_MESSAGE });
      return false;
    }
    Promise.resolve()
      .then(() => {
        const tabId = sender.tab && sender.tab.id;
        if (!tabId) throw new Error("Could not identify the CMS tab for language picker flow.");
        const x = Number(message.x);
        const y = Number(message.y);
        const text = String(message.text || "");
        if (!text) throw new Error("Missing language picker text.");
        return trustedTabLanguagePickerFlow(tabId, x, y, text);
      })
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:mainWorldSelectAddTranslationLanguage") {
    Promise.resolve()
      .then(() => {
        const tabId = sender.tab && sender.tab.id;
        if (!tabId) throw new Error("Could not identify the CMS tab for main-world selection.");
        const languageLabel = String(message.languageLabel || "");
        const query = String(message.query || languageLabel.split("(")[0].trim() || languageLabel);
        if (!languageLabel) throw new Error("Missing language label for main-world selection.");
        return mainWorldSelectAddTranslationLanguage(tabId, languageLabel, query);
      })
      .then((response) => sendResponse({ ok: Boolean(response && response.ok), response }))
      .catch((error) => sendResponse({ ok: false, response: { ok: false, debug: error.message } }));
    return true;
  }

  return false;
});
