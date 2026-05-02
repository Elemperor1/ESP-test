(function () {
  const BRIDGE_VERSION = "selector-v45-add-translation-confirm";
  const REQUEST_EVENT = `bca:select-add-translation-language:${BRIDGE_VERSION}`;
  const RESULT_EVENT = `bca:select-add-translation-language-result:${BRIDGE_VERSION}`;
  if (window.__bcaReactSelectBridgeVersion === BRIDGE_VERSION) return;
  window.__bcaReactSelectBridgeVersion = BRIDGE_VERSION;
  window.__bcaReactSelectBridgeInstalled = true;

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
      if (Array.isArray(props.options) && typeof props.onChange === "function") return { props, source: "props" };
      if (Array.isArray(selectProps.options) && typeof selectProps.onChange === "function") return { props: selectProps, source: "selectProps" };
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
    if (window.PointerEvent) {
      hit.dispatchEvent(new PointerEvent("pointerdown", { ...init, pointerId: 1, pointerType: "mouse", isPrimary: true }));
    }
    hit.dispatchEvent(new MouseEvent("mousedown", init));
    if (window.PointerEvent) {
      hit.dispatchEvent(new PointerEvent("pointerup", { ...init, pointerId: 1, pointerType: "mouse", isPrimary: true }));
    }
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
    if (input._valueTracker && typeof input._valueTracker.setValue === "function") input._valueTracker.setValue(previous);
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
  const findAddTranslationDialog = () =>
    visibleElements('[role="dialog"], [class*="modal"], [class*="Dialog"], div')
      .filter((element) => {
        const text = normalizeComparable(element.textContent);
        return text.includes("add translation") && text.includes("choose a language");
      })
      .sort((a, b) => a.textContent.length - b.textContent.length)[0] || document;

  function selectAddTranslationLanguage(languageLabel, query) {
    const dialog = findAddTranslationDialog();
    if (!(dialog instanceof Element)) {
      return { ok: false, method: "page-bridge-no-dialog", debug: `version=${BRIDGE_VERSION}; dialog not found` };
    }

    const inputCandidate = dialog.querySelector("input[id^='react-select-'][id$='-input'], input.docent-select__input, input[class*='select__input'], input[role='combobox']");
    const input = inputCandidate instanceof Element ? inputCandidate : (document.activeElement instanceof Element ? document.activeElement : null);
    const control =
      (input && input.closest?.("[class*='select__control'], [class*='Select__control'], [class*='control'], [class*='Control']")) ||
      dialog.querySelector("[class*='select__control'], [class*='Select__control'], [class*='control'], [class*='Control']");
    const addButton = () => visibleElements("button, [role='button']", dialog).find((button) => normalizeComparable(controlText(button)) === "add");
    const isReady = () => {
      const button = addButton();
      return Boolean(button && !button.disabled && button.getAttribute("aria-disabled") !== "true");
    };

    const components = collectReactSelectComponents(dialog, input, control);
    const componentDebug = [];
    for (const component of components) {
      if (!component || !component.props || !Array.isArray(component.props.options)) continue;
      const options = flattenOptions(component.props.options);
      componentDebug.push(`${component.source}:${options.length}:${options.slice(0, 8).map(optionLabel).join("|")}`);
      const option = findMatchingOption(options, languageLabel, query);
      if (option) {
        try {
          const handler = component.props.onChange;
          const context = component.fiber.stateNode || component.fiber.memoizedProps;
          if (handler && context) {
            handler.call(context, option, { action: "select-option", option, name: component.props.name });
          } else if (handler) {
            handler(option, { action: "select-option", option, name: component.props.name });
          }
          return { ok: true, method: "page-bridge-react-onChange", debug: `version=${BRIDGE_VERSION}; ${componentDebug.join("; ")}` };
        } catch (e) {
          componentDebug.push(`onChange-error: ${e.message}`);
        }
      }
    }

    if (control instanceof Element) dispatchPointClick(control);
    if (input instanceof Element) {
      input.focus?.();
      dispatchPointClick(input);
      sendKey(input, "ArrowDown", "ArrowDown", { keyCode: 40, which: 40 });
      typeQuery(input, query);
      const option =
        visibleElements("[role='option'], [role='menuitem'], li, div", document)
          .filter((element) => {
            if (!(element instanceof Element)) return false;
            const text = normalizeComparable(controlText(element));
            return text.includes(normalizeComparable(languageLabel)) || text.includes(normalizeComparable(query));
          })
          .sort((a, b) => controlText(a).length - controlText(b).length)[0] || null;
      if (option instanceof Element) {
        dispatchPointClick(option);
        return { ok: true, method: "page-bridge-option-click", debug: `version=${BRIDGE_VERSION}; ${controlText(option)}` };
      }
      sendKey(input, "Enter", "Enter", { keyCode: 13, which: 13 });
      return { ok: true, method: "page-bridge-type-enter", debug: `version=${BRIDGE_VERSION}; value=${readValue(input)}; ready=${isReady()}` };
    }

    return {
      ok: false,
      method: "page-bridge-none",
      debug: `version=${BRIDGE_VERSION}; input=${input instanceof Element ? `${input.tagName.toLowerCase()}#${input.id || ""}.${String(input.className || "").split(/\s+/).slice(0, 2).join(".")} value=${readValue(input)} expanded=${input.getAttribute("aria-expanded") || ""}` : "none"}; control=${control instanceof Element ? controlText(control).slice(0, 80) : "none"}; ready=${isReady()}; components=${componentDebug.join("; ") || "none"}`,
    };
  }

  window.addEventListener(REQUEST_EVENT, (event) => {
    const requestId = event.detail && event.detail.requestId;
    const languageLabel = event.detail && event.detail.languageLabel;
    const query = event.detail && event.detail.query;
    let result;
    try {
      result = selectAddTranslationLanguage(languageLabel, query);
    } catch (error) {
      result = { ok: false, method: "page-bridge-error", debug: error && error.message ? error.message : String(error) };
    }
    window.dispatchEvent(new CustomEvent(RESULT_EVENT, { detail: { requestId, result } }));
  });
})();
