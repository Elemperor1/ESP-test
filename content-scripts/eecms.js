const STATE_KEY = "espAltTextAssistantState";
const GENERAL_DIRECTORY_MARKER = "/cp/files/directory/6";
const MAX_ALT_LENGTH = 149;
const SAVE_CHECK_DELAY_MS = 2800;
const ACTION_ONLY_TEXT_RE = /^(?:\uF141|\u22EF|\.{3}|actions?)$/i;
const CMS_MARKER = "/eecms.php";
const PANEL_ROOT_ID = "esp-alt-text-assistant-panel-root";
const IGNORED_CELL_SELECTOR = [
  "button",
  "input",
  "select",
  "textarea",
  "script",
  "style",
  "svg",
  "path",
  "[role='button']",
  "[aria-haspopup]",
  "[data-dropdown-toggle]",
  "[hidden]",
  "[aria-hidden='true']",
  ".hidden",
  ".sr-only",
  ".visually-hidden",
  "[class*='dropdown']",
  "[class*='toolbar']",
  "[class*='actions']",
].join(",");
const PANEL_STYLE = `
  :host {
    all: initial;
  }

  .esp-panel {
    position: fixed;
    right: 16px;
    bottom: 16px;
    width: 320px;
    max-height: calc(100vh - 32px);
    overflow: hidden;
    background: rgba(20, 25, 38, 0.96);
    color: #f5f7fb;
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-radius: 10px;
    box-shadow: 0 18px 40px rgba(0, 0, 0, 0.28);
    font: 13px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    letter-spacing: 0;
    z-index: 2147483647;
  }

  .esp-panel.is-collapsed .esp-body {
    display: none;
  }

  .esp-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    padding: 12px 14px;
    background: rgba(255, 255, 255, 0.04);
    border-bottom: 1px solid rgba(255, 255, 255, 0.08);
  }

  .esp-title {
    font-size: 13px;
    font-weight: 700;
  }

  .esp-phase {
    font-size: 11px;
    opacity: 0.75;
  }

  .esp-collapse {
    border: 0;
    border-radius: 6px;
    padding: 4px 8px;
    background: rgba(255, 255, 255, 0.08);
    color: inherit;
    cursor: pointer;
    font: inherit;
  }

  .esp-body {
    display: grid;
    gap: 12px;
    padding: 12px 14px 14px;
  }

  .esp-status-list {
    display: grid;
    gap: 6px;
  }

  .esp-status {
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 0;
    color: rgba(245, 247, 251, 0.9);
  }

  .esp-status::before {
    content: "";
    width: 8px;
    height: 8px;
    border-radius: 999px;
    flex: 0 0 auto;
    background: #76819a;
  }

  .esp-status.success::before {
    background: #55d38a;
  }

  .esp-status.warning::before {
    background: #f0b24d;
  }

  .esp-status.error::before {
    background: #ef6b73;
  }

  .esp-actions,
  .esp-secondary-actions {
    display: grid;
    gap: 8px;
  }

  .esp-actions {
    grid-template-columns: repeat(3, minmax(0, 1fr));
  }

  .esp-button {
    border: 0;
    border-radius: 8px;
    padding: 9px 10px;
    background: #5f6cf5;
    color: #fff;
    cursor: pointer;
    font: inherit;
    font-weight: 600;
  }

  .esp-button.secondary {
    background: rgba(255, 255, 255, 0.1);
  }

  .esp-button.danger {
    background: #cc5965;
  }

  .esp-button:disabled {
    cursor: not-allowed;
    opacity: 0.45;
  }

  .esp-run-status,
  .esp-current {
    color: rgba(245, 247, 251, 0.92);
  }

  .esp-current {
    font-size: 12px;
    opacity: 0.86;
  }

  .esp-counts {
    display: grid;
    grid-template-columns: repeat(3, minmax(0, 1fr));
    gap: 8px;
  }

  .esp-count {
    padding: 9px 10px;
    border-radius: 8px;
    background: rgba(255, 255, 255, 0.05);
  }

  .esp-count-label {
    display: block;
    font-size: 10px;
    text-transform: uppercase;
    opacity: 0.72;
  }

  .esp-count-value {
    display: block;
    margin-top: 2px;
    font-size: 16px;
    font-weight: 700;
  }

  .esp-log {
    display: grid;
    gap: 6px;
    max-height: 188px;
    overflow: auto;
    margin: 0;
    padding-left: 18px;
  }

  .esp-log li {
    color: rgba(245, 247, 251, 0.86);
  }

  .esp-log strong {
    color: #fff;
  }
`;

let panelElements = null;
let panelRefreshTimer = null;
let panelStorageListenerRegistered = false;

function storageGet(keys) {
  return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
}

function storageSet(values) {
  return new Promise((resolve) => chrome.storage.local.set(values, resolve));
}

function sendRuntimeMessage(message) {
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

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeSpaces(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function shouldIgnoreCellElement(element) {
  if (!(element instanceof Element)) return false;

  return element.matches(IGNORED_CELL_SELECTOR);
}

function isElementHidden(element) {
  if (!(element instanceof Element)) return false;

  const style = window.getComputedStyle(element);
  return style.display === "none" || style.visibility === "hidden";
}

function isActionCell(cell) {
  if (!(cell instanceof Element)) return false;

  const text = normalizeSpaces(cell.innerText || cell.textContent || "");
  if (ACTION_ONLY_TEXT_RE.test(text)) {
    return true;
  }

  return Boolean(
    cell.matches(IGNORED_CELL_SELECTOR) ||
      cell.querySelector(
        [
          "button",
          "[role='button']",
          "[aria-haspopup]",
          "[data-dropdown-toggle]",
          "[class*='dropdown']",
          "[class*='toolbar']",
          "[class*='actions']",
        ].join(",")
      )
  );
}

function resolveDescriptionCell(cells, headerMap) {
  const descriptionCell = cells[headerMap.description];
  if (!descriptionCell) return null;

  // ExpressionEngine can omit the blank Description td altogether, which shifts
  // the action menu cell into the description slot for that row.
  if (headerMap.description === cells.length - 1 && isActionCell(descriptionCell)) {
    return null;
  }

  return descriptionCell;
}

function getMeaningfulCellText(cell) {
  if (!cell) return "";
  if (shouldIgnoreCellElement(cell)) return "";

  if (isElementHidden(cell)) {
    return "";
  }

  const parts = [];
  const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT, {
    acceptNode(node) {
      const value = normalizeSpaces(node.nodeValue);
      if (!value) return NodeFilter.FILTER_REJECT;

      let current = node.parentElement;
      while (current) {
        if (shouldIgnoreCellElement(current)) {
          return NodeFilter.FILTER_REJECT;
        }

        if (isElementHidden(current)) {
          return NodeFilter.FILTER_REJECT;
        }

        if (current === cell) break;
        current = current.parentElement;
      }

      return NodeFilter.FILTER_ACCEPT;
    },
  });

  while (walker.nextNode()) {
    parts.push(normalizeSpaces(walker.currentNode.nodeValue));
  }

  const text = normalizeSpaces(parts.join(" "));
  return ACTION_ONLY_TEXT_RE.test(text) ? "" : text;
}

function isGeneralDirectoryPage() {
  return decodeURIComponent(window.location.href).includes(GENERAL_DIRECTORY_MARKER);
}

function isCmsPage() {
  return decodeURIComponent(window.location.href).includes(CMS_MARKER);
}

function isFileEditPage() {
  const url = decodeURIComponent(window.location.href);
  return url.includes("/cp/files/file/view/") || url.includes("/cp/files/file/edit/");
}

function absoluteUrl(url) {
  try {
    return new URL(url, window.location.href).href;
  } catch (error) {
    return "";
  }
}

function comparableUrl(url) {
  try {
    const parsed = new URL(url, window.location.href);
    const pathname = parsed.pathname.replace(/\/+$/, "");
    return `${parsed.origin}${pathname}${parsed.search}`;
  } catch (error) {
    return "";
  }
}

function isExactEditUrlMatch(editUrl) {
  const expected = comparableUrl(editUrl);
  const current = comparableUrl(window.location.href);
  return Boolean(expected) && expected === current;
}

function defaultState() {
  return {
    status: "idle",
    phase: "idle",
    listUrl: "",
    index: 0,
    items: [],
    workItems: [],
    current: null,
    counts: {
      totalRows: 0,
      imageRows: 0,
      skip: 0,
      shorten: 0,
      generate: 0,
      unsupported: 0,
      saved: 0,
      failed: 0,
      alreadyCompliant: 0,
    },
    log: [],
    saveConfirmed: false,
    updatedAt: new Date().toISOString(),
  };
}

async function readState() {
  const data = await storageGet(STATE_KEY);
  return data[STATE_KEY] || defaultState();
}

async function writeState(state) {
  const nextState = {
    ...state,
    updatedAt: new Date().toISOString(),
    log: (state.log || []).slice(-120),
  };

  await storageSet({ [STATE_KEY]: nextState });
  return nextState;
}

function pushLog(state, level, message, item) {
  state.log = state.log || [];
  state.log.push({
    level,
    message,
    fileName: item && item.fileName ? item.fileName : "",
    at: new Date().toISOString(),
  });
}

function getCells(row) {
  return Array.from(row.children).filter((child) => child.matches("td, th"));
}

function buildHeaderMap(table) {
  const headers = Array.from(table.querySelectorAll("thead th"));
  const cells = headers.length
    ? headers
    : getCells(Array.from(table.querySelectorAll("tr")).find((row) => row.querySelector("th")) || document.createElement("tr"));
  const map = {};

  cells.forEach((cell, index) => {
    const text = normalizeSpaces(cell.textContent).toLowerCase();
    if (text.includes("title")) map.title = index;
    if (text.includes("file name")) map.fileName = index;
    if (text.includes("file type")) map.fileType = index;
    if (text.includes("description")) map.description = index;
  });

  return {
    title: map.title ?? 2,
    fileName: map.fileName ?? 3,
    fileType: map.fileType ?? 4,
    description: map.description ?? 7,
  };
}

function isImageRow(fileType, fileName) {
  if (normalizeSpaces(fileType).toLowerCase() === "image") return true;
  return /\.(?:avif|gif|jpe?g|png|svg|webp)$/i.test(fileName || "");
}

function classifyDescription(description) {
  const cleaned = normalizeSpaces(description);
  if (!cleaned) return "generate";
  if (cleaned.length < 150) return "skip";
  return "shorten";
}

function findEditLink(row, cells, headerMap) {
  const titleCell = cells[headerMap.title] || row;
  const titleLink = titleCell.querySelector('a[href*="/cp/files/file/"], a[href*="cp/files/file"]');
  if (titleLink && titleLink.href) return titleLink.href;

  const rowLink = row.querySelector('a[href*="/cp/files/file/"], a[href*="cp/files/file"]');
  return rowLink && rowLink.href ? rowLink.href : "";
}

function scanDirectoryRows() {
  if (!isGeneralDirectoryPage()) {
    throw new Error("Open the Eastern State General file directory before scanning.");
  }

  const items = [];
  let totalRows = 0;
  let unsupported = 0;

  document.querySelectorAll("table").forEach((table) => {
    const headerMap = buildHeaderMap(table);
    const bodyRows = Array.from(table.querySelectorAll("tbody tr"));
    const rows = bodyRows.length ? bodyRows : Array.from(table.querySelectorAll("tr")).slice(1);

    rows.forEach((row) => {
      const cells = getCells(row);
      if (cells.length < 5 || row.querySelector("th")) return;

      totalRows += 1;
      const fileName = getMeaningfulCellText(cells[headerMap.fileName]);
      const fileType = getMeaningfulCellText(cells[headerMap.fileType]);
      const descriptionCell = resolveDescriptionCell(cells, headerMap);
      const description = getMeaningfulCellText(descriptionCell);
      const title = getMeaningfulCellText(cells[headerMap.title]);
      const editUrl = findEditLink(row, cells, headerMap);

      if (!fileName || !editUrl || !isImageRow(fileType, fileName)) {
        unsupported += 1;
        return;
      }

      const action = classifyDescription(description);
      items.push({
        id: `${items.length + 1}-${fileName}`,
        rowNumber: totalRows,
        title,
        fileName,
        fileType,
        description,
        descriptionLength: description.length,
        action,
        editUrl: absoluteUrl(editUrl),
        status: action === "skip" ? "skipped" : "pending",
      });
    });
  });

  const counts = {
    totalRows,
    imageRows: items.length,
    skip: items.filter((item) => item.action === "skip").length,
    shorten: items.filter((item) => item.action === "shorten").length,
    generate: items.filter((item) => item.action === "generate").length,
    unsupported,
    saved: 0,
    failed: 0,
    alreadyCompliant: 0,
  };

  return {
    items,
    workItems: items.filter((item) => item.action !== "skip"),
    counts,
  };
}

async function scanAndStore() {
  const scan = scanDirectoryRows();
  const state = defaultState();
  state.status = "scanned";
  state.phase = "ready";
  state.listUrl = window.location.href;
  state.items = scan.items;
  state.workItems = scan.workItems;
  state.counts = scan.counts;
  pushLog(state, "info", `Scanned ${scan.items.length} image rows. ${scan.workItems.length} need changes.`);
  await writeState(state);
  return state;
}

function findFieldByLabel(labelText) {
  const normalizedLabel = labelText.toLowerCase();
  const labels = Array.from(document.querySelectorAll("label"));

  for (const label of labels) {
    const text = normalizeSpaces(label.textContent).toLowerCase();
    if (!text || !text.includes(normalizedLabel)) continue;

    const forId = label.getAttribute("for");
    if (forId) {
      const direct = document.getElementById(forId);
      if (direct) return direct;
    }

    const container =
      label.closest(".field-control, .fieldset-field, .field-instruct, .form-standard, .field-group, div") ||
      label.parentElement;
    const field = container && container.querySelector("textarea, input[type='text'], input:not([type])");
    if (field) return field;
  }

  const fallbackSelectors =
    labelText.toLowerCase() === "description"
      ? ["textarea[name*='description' i]", "textarea[id*='description' i]"]
      : [`input[name*='${labelText}' i]`, `input[id*='${labelText}' i]`];

  for (const selector of fallbackSelectors) {
    const field = document.querySelector(selector);
    if (field) return field;
  }

  return null;
}

function findDescriptionField() {
  return findFieldByLabel("Description") || document.querySelector("textarea");
}

function findTitleField() {
  return findFieldByLabel("Title") || document.querySelector("input[type='text']");
}

function getFieldValue(field) {
  return field ? normalizeSpaces(field.value || field.textContent || "") : "";
}

function setFieldValue(field, value) {
  field.focus();
  const prototype = Object.getPrototypeOf(field);
  const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");

  if (descriptor && descriptor.set) {
    descriptor.set.call(field, value);
  } else {
    field.value = value;
  }

  field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
  field.dispatchEvent(new Event("change", { bubbles: true }));
  field.dispatchEvent(new Event("blur", { bubbles: true }));
}

function looksLikeImageUrl(url) {
  return /\.(?:avif|gif|jpe?g|png|svg|webp)(?:[?#]|$)/i.test(url || "");
}

function normalizeIdentifier(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function tokenizeIdentifier(value) {
  return normalizeIdentifier(value)
    .split(/\s+/)
    .filter((token) => token.length >= 3);
}

function isLikelyUiImageSource(source) {
  const text = String(source || "").toLowerCase();
  return /(avatar|profile|gravatar|user(?:pic|image)?|placeholder|default[-_]?user|logo|icon)/.test(text);
}

function isLikelyUiImageElement(image) {
  if (!(image instanceof Element)) return false;
  if (image.closest("header, nav, [role='navigation'], [class*='toolbar'], [class*='topbar'], [class*='menu']")) {
    return true;
  }

  const context = [
    image.className || "",
    image.id || "",
    image.getAttribute("aria-label") || "",
    image.getAttribute("alt") || "",
    image.closest("[class]") ? image.closest("[class]").className || "" : "",
  ]
    .join(" ")
    .toLowerCase();

  return /(avatar|profile|user|account|logo|icon|badge)/.test(context);
}

function scoreImageCandidate(source, width, height, fileName, title) {
  if (!source) return Number.NEGATIVE_INFINITY;
  if (source.startsWith("data:")) return Number.NEGATIVE_INFINITY;

  const lowerSource = source.toLowerCase();
  const normalizedSource = normalizeIdentifier(decodeURIComponent(source));
  let score = Math.min((width || 0) * (height || 0), 4000000) / 4000;

  if (isLikelyUiImageSource(source)) {
    score -= 800;
  }

  if (/\/(?:uploads?|files?)\//.test(lowerSource)) {
    score += 180;
  }

  if (/\/cp\//.test(lowerSource) || /\/assets?\//.test(lowerSource)) {
    score -= 120;
  }

  const fileTokens = tokenizeIdentifier(fileName);
  const titleTokens = tokenizeIdentifier(title).slice(0, 4);
  const matchingFileTokens = fileTokens.filter((token) => normalizedSource.includes(token)).length;
  const matchingTitleTokens = titleTokens.filter((token) => normalizedSource.includes(token)).length;

  score += matchingFileTokens * 120;
  score += matchingTitleTokens * 40;

  if (fileName && lowerSource.includes(String(fileName).toLowerCase())) {
    score += 220;
  }

  return score;
}

function extractBackgroundImageUrl(value) {
  const match = String(value || "").match(/url\((['"]?)(.+?)\1\)/i);
  return match ? absoluteUrl(match[2]) : "";
}

function buildGeneralUploadUrl(fileName) {
  const cleaned = String(fileName || "").trim().replace(/^\/+/, "");
  if (!cleaned) return "";
  const encoded = cleaned
    .split("/")
    .filter(Boolean)
    .map((part) => encodeURIComponent(part))
    .join("/");
  return absoluteUrl(`/uploads/general/${encoded}`);
}

function scoreImageLinkUrl(url, fileName, title) {
  if (!url) return Number.NEGATIVE_INFINITY;
  const lower = url.toLowerCase();
  const normalized = normalizeIdentifier(decodeURIComponent(url));
  let score = 0;

  if (looksLikeImageUrl(url)) score += 200;
  if (/\/uploads?\//.test(lower)) score += 160;
  if (/download/.test(lower)) score += 70;
  if (/avatar|profile|logo|icon|placeholder/.test(lower)) score -= 500;

  const fileTokens = tokenizeIdentifier(fileName);
  const titleTokens = tokenizeIdentifier(title).slice(0, 4);
  score += fileTokens.filter((token) => normalized.includes(token)).length * 120;
  score += titleTokens.filter((token) => normalized.includes(token)).length * 40;

  if (fileName && lower.includes(String(fileName).toLowerCase())) {
    score += 260;
  }

  return score;
}

function findDirectImageLinkUrl(fileName, title) {
  const links = Array.from(document.querySelectorAll("a[href]"))
    .map((link) => {
      const href = absoluteUrl(link.href);
      const text = normalizeSpaces(link.textContent || link.getAttribute("aria-label") || "").toLowerCase();
      const context = normalizeSpaces(link.className || "").toLowerCase();

      // Favor explicit download/preview controls and direct image hrefs.
      if (!href) return null;
      if (!looksLikeImageUrl(href) && !text.includes("download") && !text.includes("view")) return null;
      if (context.includes("avatar") || context.includes("profile")) return null;

      const score = scoreImageLinkUrl(href, fileName, title);
      return { href, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);

  return links.length && links[0].score > 0 ? links[0].href : "";
}

function findPreviewImageUrl(fileName, title) {
  const images = Array.from(document.querySelectorAll("img[src]"));
  const candidates = images
    .map((image) => {
      const source = absoluteUrl(image.currentSrc || image.src || "");
      const width = image.naturalWidth || image.width || 0;
      const height = image.naturalHeight || image.height || 0;
      const uiPenalty = isLikelyUiImageElement(image) ? 600 : 0;
      const score = scoreImageCandidate(source, width, height, fileName, title) - uiPenalty;
      return { source, width, height, score };
    })
    .filter((candidate) => {
      if (!candidate.source) return false;
      if (candidate.width < 120 || candidate.height < 120) return false;
      return candidate.score > 0;
    })
    .sort((a, b) => b.score - a.score);

  const bestImageCandidate = candidates[0];
  if (bestImageCandidate) {
    return bestImageCandidate.source;
  }

  const legacyCandidate = images.filter((image) => {
    const source = image.currentSrc || image.src || "";
    const width = image.naturalWidth || image.width || 0;
    const height = image.naturalHeight || image.height || 0;
    const alt = `${image.alt || ""} ${image.className || ""}`.toLowerCase();

    if (!source || source.startsWith("data:")) return false;
    if (alt.includes("logo") || alt.includes("avatar")) return false;
    if (width < 120 || height < 120) return false;
    return true;
  });

  const image = legacyCandidate[0] || images.find((candidate) => looksLikeImageUrl(candidate.src));
  if (image) {
    return absoluteUrl(image.currentSrc || image.src);
  }

  const backgroundCandidates = Array.from(
    document.querySelectorAll("[style*='background-image'], [class*='preview'], [class*='thumbnail'], [class*='image']")
  ).filter((element) => {
    if (!(element instanceof HTMLElement)) return false;
    const rect = element.getBoundingClientRect();
    if (rect.width < 80 || rect.height < 80) return false;

    const style = window.getComputedStyle(element);
    return Boolean(extractBackgroundImageUrl(style.backgroundImage || element.style.backgroundImage));
  });

  const backgroundUrl = backgroundCandidates
    .map((element) => {
      const style = window.getComputedStyle(element);
      const source = extractBackgroundImageUrl(style.backgroundImage || element.style.backgroundImage);
      const rect = element.getBoundingClientRect();
      const score = scoreImageCandidate(source, rect.width, rect.height, fileName, title);
      return { source, score };
    })
    .filter((candidate) => candidate.source && candidate.score > 0)
    .sort((a, b) => b.score - a.score)
    .map((candidate) => candidate.source)[0];

  if (backgroundUrl) {
    return backgroundUrl;
  }

  const previewLink = Array.from(document.querySelectorAll("a[href]"))
    .map((link) => {
      const href = absoluteUrl(link.href);
      const text = normalizeSpaces(link.textContent || link.getAttribute("aria-label") || "").toLowerCase();
      if (!looksLikeImageUrl(href) || text.includes("logo") || text.includes("avatar")) return null;
      const score = scoreImageCandidate(href, 300, 300, fileName, title);
      return { href, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score)[0];

  return previewLink ? absoluteUrl(previewLink.href) : "";
}

function cleanAltText(rawText) {
  let text = normalizeSpaces(rawText)
    .replace(/^(?:alt text|description|answer)\s*:\s*/i, "")
    .trim();

  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    text = text.slice(1, -1).trim();
  }

  return normalizeSpaces(text);
}

function isValidAltText(text) {
  const cleaned = cleanAltText(text);
  return cleaned.length > 0 && cleaned.length <= MAX_ALT_LENGTH;
}

async function requestAltTextWithRetries(payload, state, item) {
  let previousOutput = "";

  for (let attempt = 1; attempt <= 3; attempt += 1) {
    pushLog(state, "info", `Requesting ChatGPT ${payload.action} attempt ${attempt}.`, item);
    await writeState(state);

    const response = await sendRuntimeMessage({
      type: "esp:requestAltText",
      payload: {
        ...payload,
        maxLength: MAX_ALT_LENGTH,
        attempt,
        previousOutput,
      },
    });

    if (!response || !response.ok) {
      throw new Error(response && response.error ? response.error : "ChatGPT request failed.");
    }

    const candidate = cleanAltText(response.altText || "");
    if (isValidAltText(candidate)) {
      return candidate;
    }

    previousOutput = candidate || "(blank response)";
    pushLog(state, "warn", `ChatGPT returned ${candidate.length} characters; retrying.`, item);
    await writeState(state);
  }

  throw new Error("ChatGPT did not return valid alt text under 150 characters.");
}

function findSaveButton() {
  const buttons = Array.from(document.querySelectorAll("button, input[type='submit'], a.btn"));
  return buttons.find((button) => {
    const text = normalizeSpaces(button.textContent || button.value || "").toLowerCase();
    return text === "save" || text.startsWith("save ");
  }) || null;
}

function findErrorNotice() {
  return document.querySelector(".alert-error, .app-notice--error, .notice--error, .alert--error");
}

function findSuccessNotice() {
  return document.querySelector(".alert-success, .app-notice--success, .notice--success, .alert--success");
}

async function failAndPause(message, item) {
  const state = await readState();
  state.status = "paused";
  state.phase = "failed";
  state.current = {
    itemId: item ? item.id : "",
    fileName: item ? item.fileName : "",
    error: message,
  };
  state.counts.failed += 1;
  if (item) {
    const storedItem = state.workItems.find((candidate) => candidate.id === item.id);
    if (storedItem) {
      storedItem.status = "manual";
      storedItem.error = message;
    }
  }
  pushLog(state, "error", message, item);
  await writeState(state);
  alert(`Alt text assistant paused: ${message}`);
}

async function markCurrentSavedAndContinue() {
  let state = await readState();
  const item = state.workItems[state.index];
  if (!item) {
    await completeRun(state);
    return;
  }

  item.status = "saved";
  item.generatedDescription = state.current ? state.current.expectedDescription : "";
  state.counts.saved += 1;
  state.index += 1;
  state.phase = "ready";
  state.current = null;
  pushLog(state, "success", "Saved description.", item);
  state = await writeState(state);

  setTimeout(() => {
    continueRun().catch((error) => failAndPause(error.message, item));
  }, 500);
}

async function completeRun(state) {
  state.status = "complete";
  state.phase = "complete";
  state.current = null;
  pushLog(state, "success", "Run complete.");
  await writeState(state);

  if (state.listUrl && !isGeneralDirectoryPage()) {
    window.location.href = state.listUrl;
  }
}

async function checkSavingState(state) {
  const item = state.workItems[state.index];
  if (!item || !state.current || !state.current.expectedDescription) {
    await completeRun(state);
    return;
  }

  if (isGeneralDirectoryPage()) {
    await markCurrentSavedAndContinue();
    return;
  }

  const descriptionField = findDescriptionField();
  const expectedDescription = normalizeSpaces(state.current.expectedDescription);
  const fieldMatchesExpected = Boolean(descriptionField && getFieldValue(descriptionField) === expectedDescription);
  const successNotice = findSuccessNotice();
  if (fieldMatchesExpected && successNotice) {
    await markCurrentSavedAndContinue();
    return;
  }

  const errorNotice = findErrorNotice();
  if (errorNotice) {
    await failAndPause(normalizeSpaces(errorNotice.textContent) || "ExpressionEngine reported a save error.", item);
    return;
  }

  const checkCount = (state.current.saveCheckCount || 0) + 1;
  state.current.saveCheckCount = checkCount;
  await writeState(state);

  if (checkCount >= 5) {
    await failAndPause("Could not confirm that ExpressionEngine saved the description.", item);
    return;
  }

  setTimeout(() => {
    readState().then(checkSavingState);
  }, SAVE_CHECK_DELAY_MS);
}

async function saveDescriptionAndResume(item, altText) {
  const saveButton = findSaveButton();
  if (!saveButton) {
    throw new Error("Could not find the ExpressionEngine Save button.");
  }

  let state = await readState();
  state.phase = "saving";
  state.current = {
    itemId: item.id,
    fileName: item.fileName,
    expectedDescription: altText,
    saveCheckCount: 0,
  };
  pushLog(state, "info", "Saving description.", item);
  await writeState(state);

  saveButton.click();

  setTimeout(() => {
    readState().then(checkSavingState);
  }, SAVE_CHECK_DELAY_MS);
}

async function processCurrentEditPage(state) {
  const item = state.workItems[state.index];
  if (!item) {
    await completeRun(state);
    return;
  }

  const descriptionField = findDescriptionField();
  if (!descriptionField) {
    await failAndPause("Could not find the Description field on the edit page.", item);
    return;
  }

  const currentDescription = getFieldValue(descriptionField);
  const currentAction = classifyDescription(currentDescription);
  if (currentAction === "skip") {
    item.status = "already-compliant";
    state.counts.alreadyCompliant += 1;
    state.index += 1;
    state.phase = "ready";
    state.current = null;
    pushLog(state, "info", "Already compliant on edit page; skipped.", item);
    await writeState(state);
    await continueRun();
    return;
  }

  const titleField = findTitleField();
  const title = getFieldValue(titleField) || item.title;
  const action = currentAction === "generate" ? "generate" : "shorten";
  let imageUrl = "";

  if (action === "generate") {
    const discoveredPreviewUrl = findPreviewImageUrl(item.fileName, title);
    const discoveredLinkUrl = !discoveredPreviewUrl ? findDirectImageLinkUrl(item.fileName, title) : "";
    const deterministicUrl = buildGeneralUploadUrl(item.fileName);
    imageUrl = discoveredPreviewUrl || discoveredLinkUrl || deterministicUrl || "";
  }

  if (action === "generate" && !imageUrl) {
    await failAndPause("Missing alt text, but no usable image URL was found on the page or from the file path.", item);
    return;
  }

  state.phase = "processing";
  state.current = {
    itemId: item.id,
    fileName: item.fileName,
    action,
  };
  await writeState(state);

  try {
    if (action === "generate") {
      pushLog(state, "info", `Using image URL: ${imageUrl}`, item);
      await writeState(state);
    }

    const altText = await requestAltTextWithRetries(
      {
        action,
        existingText: currentDescription || item.description,
        title,
        fileName: item.fileName,
        imageUrl,
      },
      state,
      item
    );

    setFieldValue(descriptionField, altText);
    await delay(250);
    await saveDescriptionAndResume(item, altText);
  } catch (error) {
    await failAndPause(error.message, item);
  }
}

async function continueRun() {
  const state = await readState();
  if (state.status !== "running") return;

  if (state.phase === "saving") {
    await checkSavingState(state);
    return;
  }

  if (state.index >= state.workItems.length) {
    await completeRun(state);
    return;
  }

  const item = state.workItems[state.index];
  if (!item) {
    await completeRun(state);
    return;
  }

  if (isFileEditPage() && isExactEditUrlMatch(item.editUrl)) {
    await processCurrentEditPage(state);
    return;
  }

  window.location.href = item.editUrl;
}

async function startRun() {
  let state = await readState();

  if (!state.workItems || !state.workItems.length) {
    state = await scanAndStore();
  }

  if (!state.workItems.length) {
    state.status = "complete";
    state.phase = "complete";
    pushLog(state, "success", "No files need changes.");
    await writeState(state);
    return state;
  }

  const confirmed = window.confirm(
    `This will edit and save public CMS file descriptions for ${state.workItems.length} images in the General directory. Continue?`
  );

  if (!confirmed) {
    state.status = "scanned";
    state.phase = "ready";
    state.saveConfirmed = false;
    pushLog(state, "info", "Save pass cancelled before any CMS changes.");
    await writeState(state);
    return state;
  }

  state.status = "running";
  state.phase = "ready";
  state.index = state.index || 0;
  state.saveConfirmed = true;
  pushLog(state, "info", "Started save-as-you-go processing.");
  await writeState(state);

  setTimeout(() => {
    continueRun().catch((error) => failAndPause(error.message, state.workItems[state.index]));
  }, 250);

  return state;
}

async function stopRun() {
  const state = await readState();
  state.status = "stopped";
  state.phase = "stopped";
  pushLog(state, "warn", "Run stopped by user.");
  await writeState(state);
  return state;
}

function getPanelRoot() {
  return document.getElementById(PANEL_ROOT_ID);
}

function createPanelElement(tagName, className, textContent) {
  const element = document.createElement(tagName);
  if (className) {
    element.className = className;
  }
  if (typeof textContent === "string") {
    element.textContent = textContent;
  }
  return element;
}

function ensureInjectedPanel() {
  if (!isCmsPage()) {
    return null;
  }

  const existingRoot = getPanelRoot();
  if (existingRoot && panelElements) {
    return panelElements;
  }

  const host = document.createElement("div");
  host.id = PANEL_ROOT_ID;
  const shadowRoot = host.attachShadow({ mode: "open" });

  const style = document.createElement("style");
  style.textContent = PANEL_STYLE;
  shadowRoot.appendChild(style);

  const panel = createPanelElement("section", "esp-panel");
  const header = createPanelElement("header", "esp-header");
  const titleWrap = createPanelElement("div");
  const title = createPanelElement("div", "esp-title", "Alt Text Assistant");
  const phase = createPanelElement("div", "esp-phase", "Idle");
  titleWrap.append(title, phase);
  const collapseButton = createPanelElement("button", "esp-collapse", "Hide");
  collapseButton.type = "button";
  header.append(titleWrap, collapseButton);

  const body = createPanelElement("div", "esp-body");
  const statusList = createPanelElement("div", "esp-status-list");
  const cmsStatus = createPanelElement("div", "esp-status", "Checking CMS page");
  const chatgptStatus = createPanelElement("div", "esp-status", "Checking ChatGPT");
  statusList.append(cmsStatus, chatgptStatus);

  const actionRow = createPanelElement("div", "esp-actions");
  const scanButton = createPanelElement("button", "esp-button", "Scan");
  scanButton.type = "button";
  const startButton = createPanelElement("button", "esp-button", "Start Saving");
  startButton.type = "button";
  const stopButton = createPanelElement("button", "esp-button danger", "Stop");
  stopButton.type = "button";
  actionRow.append(scanButton, startButton, stopButton);

  const secondaryActionRow = createPanelElement("div", "esp-secondary-actions");
  const openChatGPTButton = createPanelElement("button", "esp-button secondary", "Open ChatGPT");
  openChatGPTButton.type = "button";
  secondaryActionRow.append(openChatGPTButton);

  const runStatus = createPanelElement("div", "esp-run-status", "Idle.");
  const currentItem = createPanelElement("div", "esp-current");

  const counts = createPanelElement("div", "esp-counts");
  const countKeys = [
    ["Images", "images"],
    ["Skip", "skip"],
    ["Shorten", "shorten"],
    ["Generate", "generate"],
    ["Saved", "saved"],
    ["Failed", "failed"],
  ];
  const countMap = {};
  countKeys.forEach(([label, key]) => {
    const countCard = createPanelElement("div", "esp-count");
    const countLabel = createPanelElement("span", "esp-count-label", label);
    const countValue = createPanelElement("span", "esp-count-value", "0");
    countCard.append(countLabel, countValue);
    counts.appendChild(countCard);
    countMap[key] = countValue;
  });

  const logList = createPanelElement("ol", "esp-log");
  const emptyLog = createPanelElement("li", "", "No activity yet.");
  logList.appendChild(emptyLog);

  body.append(statusList, actionRow, secondaryActionRow, runStatus, currentItem, counts, logList);
  panel.append(header, body);
  shadowRoot.appendChild(panel);
  document.documentElement.appendChild(host);

  panelElements = {
    host,
    panel,
    body,
    phase,
    cmsStatus,
    chatgptStatus,
    scanButton,
    startButton,
    stopButton,
    openChatGPTButton,
    runStatus,
    currentItem,
    countMap,
    logList,
    collapseButton,
  };

  collapseButton.addEventListener("click", () => {
    const isCollapsed = panel.classList.toggle("is-collapsed");
    collapseButton.textContent = isCollapsed ? "Show" : "Hide";
  });

  scanButton.addEventListener("click", async () => {
    setPanelButtonsBusy(true);
    try {
      await scanAndStore();
    } catch (error) {
      window.alert(error.message);
    } finally {
      setPanelButtonsBusy(false);
      refreshInjectedPanel();
    }
  });

  startButton.addEventListener("click", async () => {
    setPanelButtonsBusy(true);
    try {
      await startRun();
    } catch (error) {
      window.alert(error.message);
    } finally {
      setPanelButtonsBusy(false);
      refreshInjectedPanel();
    }
  });

  stopButton.addEventListener("click", async () => {
    setPanelButtonsBusy(true);
    try {
      await stopRun();
    } catch (error) {
      window.alert(error.message);
    } finally {
      setPanelButtonsBusy(false);
      refreshInjectedPanel();
    }
  });

  openChatGPTButton.addEventListener("click", async () => {
    setPanelButtonsBusy(true);
    try {
      await sendRuntimeMessage({ type: "esp:openChatGPT" });
    } catch (error) {
      window.alert(error.message);
    } finally {
      setPanelButtonsBusy(false);
      refreshInjectedPanel();
    }
  });

  if (!panelStorageListenerRegistered && chrome.storage && chrome.storage.onChanged) {
    chrome.storage.onChanged.addListener((changes, areaName) => {
      if (areaName === "local" && changes[STATE_KEY]) {
        refreshInjectedPanel();
      }
    });
    panelStorageListenerRegistered = true;
  }

  if (!panelRefreshTimer) {
    panelRefreshTimer = window.setInterval(() => {
      refreshInjectedPanel();
    }, 1000);
  }

  return panelElements;
}

function setStatusAppearance(element, text, className) {
  if (!element) return;

  element.textContent = text;
  element.className = `esp-status ${className || ""}`.trim();
}

function setPanelButtonsBusy(disabled) {
  if (!panelElements) return;

  panelElements.scanButton.disabled = disabled;
  panelElements.startButton.disabled = disabled;
  panelElements.stopButton.disabled = disabled;
  panelElements.openChatGPTButton.disabled = disabled;
}

function renderPanelLog(logList, entries) {
  logList.textContent = "";

  if (!entries.length) {
    logList.appendChild(createPanelElement("li", "", "No activity yet."));
    return;
  }

  entries.forEach((entry) => {
    const item = createPanelElement("li");
    if (entry.fileName) {
      const strong = createPanelElement("strong", "", entry.fileName);
      item.appendChild(strong);
      item.appendChild(document.createTextNode(": "));
    }
    item.appendChild(document.createTextNode(entry.message || ""));
    logList.appendChild(item);
  });
}

async function refreshInjectedPanel() {
  const elements = ensureInjectedPanel();
  if (!elements) {
    return;
  }

  const [state, chatgpt] = await Promise.all([
    readState().catch(() => defaultState()),
    sendRuntimeMessage({ type: "esp:getChatGPTStatus" }).catch((error) => ({
      ok: false,
      error: error.message,
      available: false,
    })),
  ]);

  if (isGeneralDirectoryPage()) {
    setStatusAppearance(elements.cmsStatus, "General directory is active", "success");
  } else if (isFileEditPage()) {
    setStatusAppearance(elements.cmsStatus, "Edit page active", "warning");
  } else {
    setStatusAppearance(elements.cmsStatus, "Open the General directory", "error");
  }

  if (chatgpt.ok && chatgpt.available) {
    setStatusAppearance(elements.chatgptStatus, "ChatGPT tab is open", "success");
  } else {
    setStatusAppearance(elements.chatgptStatus, "ChatGPT tab is needed", "warning");
  }

  const counts = state.counts || {};
  const workTotal = Array.isArray(state.workItems) ? state.workItems.length : 0;
  const index = Number.isInteger(state.index) ? state.index : 0;
  const status = state.status || "idle";
  const phase = state.phase || "idle";

  elements.phase.textContent = `${status} / ${phase}`;
  elements.runStatus.textContent =
    status === "idle"
      ? "Idle."
      : `${status} (${phase}) - ${Math.min(index, workTotal)} of ${workTotal} change candidates processed.`;

  if (state.current && state.current.fileName) {
    elements.currentItem.textContent = `${state.current.fileName} - ${state.current.action || phase}`;
  } else if (workTotal && status === "scanned") {
    elements.currentItem.textContent = `${workTotal} files need shortening or generated alt text.`;
  } else {
    elements.currentItem.textContent = "";
  }

  elements.countMap.images.textContent = String(counts.imageRows || 0);
  elements.countMap.skip.textContent = String(counts.skip || 0);
  elements.countMap.shorten.textContent = String(counts.shorten || 0);
  elements.countMap.generate.textContent = String(counts.generate || 0);
  elements.countMap.saved.textContent = String(counts.saved || 0);
  elements.countMap.failed.textContent = String(counts.failed || 0);

  elements.scanButton.disabled = !isGeneralDirectoryPage();
  elements.startButton.disabled = !isGeneralDirectoryPage();
  elements.stopButton.disabled = status !== "running";

  const logEntries = Array.isArray(state.log) ? state.log.slice(-8).reverse() : [];
  renderPanelLog(elements.logList, logEntries);
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") {
    return false;
  }

  if (message.type === "esp:scanDirectory") {
    scanAndStore()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "esp:startRun") {
    startRun()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "esp:stopRun") {
    stopRun()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  return false;
});

setTimeout(() => {
  readState()
    .then((state) => {
      if (state.status === "running") {
        return continueRun();
      }
      return null;
    })
    .catch((error) => {
      console.error("Eastern State Alt Text Assistant resume error:", error);
    });
}, 800);

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => {
    ensureInjectedPanel();
    refreshInjectedPanel();
  });
} else {
  ensureInjectedPanel();
  refreshInjectedPanel();
}
