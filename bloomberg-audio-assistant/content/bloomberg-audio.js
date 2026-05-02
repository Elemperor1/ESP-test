const BCA_STATE_KEY = "bcaState";
const BCA_PANEL_ROOT_ID = "bca-panel-root";
const BCA_DB_NAME = "bloomberg-audio-assistant";
const BCA_AUDIO_STORE = "audio";
const BCA_SETTINGS_STORE = "settings";
const BCA_AUDIO_FOLDER_HANDLE_KEY = "audioFolderHandle";
const BCA_AUDIO_FOLDER_META_KEY = "bcaAudioFolderMeta";
const BCA_AUDIO_DOWNLOAD_KEY = "bcaAudioDownloadRun";
const BCA_SAVE_CHECK_DELAY_MS = 2000;
const BCA_UPLOAD_TIMEOUT_MS = 180000;
const BCA_ADD_TRANSLATION_RETRY_ATTEMPTS = 3;
const BCA_ADD_TRANSLATION_RETRY_BACKOFF_MS = 600;
const BCA_EXTENSION_VERSION = "1.0.45";
const BCA_BUILD_LABEL = "selector-v45-add-translation-confirm";
const BCA_PAGE_BRIDGE_EVENT_SUFFIX = "selector-v45-add-translation-confirm";
const BCA_PAGE_BRIDGE_REQUEST_EVENT = `bca:select-add-translation-language:${BCA_PAGE_BRIDGE_EVENT_SUFFIX}`;
const BCA_PAGE_BRIDGE_RESULT_EVENT = `bca:select-add-translation-language-result:${BCA_PAGE_BRIDGE_EVENT_SUFFIX}`;
const BCA_ENABLE_DEBUGGER_INPUT = false;
let lastMainWorldLanguageSelectionDebug = "";
let lastDownloadAudioClickDebug = "";
let pageWorldBridgeInstalledVersion = "";
let bcaAudioFolderHandle = null;
let bcaAudioFolderMeta = null;
const BCA_PANEL_STYLE = `
  :host { all: initial; }
  .bca-panel {
    position: fixed;
    right: 16px;
    bottom: 16px;
    width: 340px;
    max-height: calc(100vh - 32px);
    overflow: hidden;
    background: rgba(18, 24, 32, 0.96);
    color: #f7fafc;
    border: 1px solid rgba(255, 255, 255, 0.14);
    border-radius: 10px;
    box-shadow: 0 18px 44px rgba(0, 0, 0, 0.3);
    font: 13px/1.45 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    letter-spacing: 0;
    z-index: 2147483647;
  }
  .bca-panel.is-collapsed .bca-body { display: none; }
  .bca-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    padding: 12px 14px;
    background: rgba(255, 255, 255, 0.05);
    border-bottom: 1px solid rgba(255, 255, 255, 0.1);
  }
  .bca-title { font-size: 13px; font-weight: 800; }
  .bca-phase { font-size: 11px; opacity: 0.74; }
  .bca-build { font-size: 10px; opacity: 0.58; }
  .bca-collapse {
    border: 0;
    border-radius: 6px;
    padding: 4px 8px;
    background: rgba(255, 255, 255, 0.09);
    color: inherit;
    cursor: pointer;
    font: inherit;
  }
  .bca-body { display: grid; gap: 12px; padding: 12px 14px 14px; }
  .bca-status-list { display: grid; gap: 6px; }
  .bca-status {
    display: flex;
    align-items: center;
    gap: 8px;
    min-width: 0;
    color: rgba(247, 250, 252, 0.9);
  }
  .bca-status::before {
    content: "";
    width: 8px;
    height: 8px;
    border-radius: 999px;
    flex: 0 0 auto;
    background: #8b98a8;
  }
  .bca-status.success::before { background: #47c77a; }
  .bca-status.warning::before { background: #f0b24d; }
  .bca-status.error::before { background: #ef6b73; }
  .bca-actions { display: grid; grid-template-columns: 1.35fr 0.8fr 0.8fr 0.7fr; gap: 8px; }
  .bca-folder-row { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
  .bca-folder-status { grid-column: 1 / -1; }
  .bca-button {
    border: 0;
    border-radius: 8px;
    padding: 9px 8px;
    background: #4f7df3;
    color: #fff;
    cursor: pointer;
    font: inherit;
    font-weight: 700;
  }
  .bca-button.danger { background: #cc5965; }
  .bca-button.secondary { background: rgba(255, 255, 255, 0.12); }
  .bca-button:disabled { cursor: not-allowed; opacity: 0.45; }
  .bca-run-status, .bca-current { color: rgba(247, 250, 252, 0.92); }
  .bca-current { font-size: 12px; opacity: 0.86; overflow-wrap: anywhere; }
  .bca-counts { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
  .bca-count { padding: 9px 10px; border-radius: 8px; background: rgba(255, 255, 255, 0.06); }
  .bca-count-label { display: block; font-size: 10px; text-transform: uppercase; opacity: 0.72; }
  .bca-count-value { display: block; margin-top: 2px; font-size: 16px; font-weight: 800; }
  .bca-log { display: grid; gap: 6px; max-height: 188px; overflow: auto; margin: 0; padding-left: 18px; }
  .bca-log li { color: rgba(247, 250, 252, 0.86); }
  .bca-log strong { color: #fff; }
`;

let bcaConfig = null;
let bcaTranscripts = null;
let bcaPanelElements = null;
let bcaPanelRefreshTimer = null;
let bcaStorageListenerRegistered = false;
let bcaContinueInFlight = false;
let bcaAudioDownloadInFlight = false;
let bcaLastAutoContinueAt = 0;
let bcaLastFocusRequestAt = 0;

function storageGet(keys) {
  return new Promise((resolve) => chrome.storage.local.get(keys, resolve));
}

function storageSet(values) {
  return new Promise((resolve) => chrome.storage.local.set(values, resolve));
}

function storageRemove(keys) {
  return new Promise((resolve) => chrome.storage.local.remove(keys, resolve));
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function assertExtensionBuildCurrent() {
  const manifest = chrome.runtime && typeof chrome.runtime.getManifest === "function" ? chrome.runtime.getManifest() : null;
  if (!manifest || !manifest.version) return;
  if (manifest.version !== BCA_EXTENSION_VERSION) {
    throw new Error(
      `Extension build mismatch. Content script is ${BCA_EXTENSION_VERSION} (${BCA_BUILD_LABEL}) but manifest is ${manifest.version}. Reload the unpacked extension and this CMS tab before continuing.`
    );
  }
}

async function ensureCmsWindowFocused() {
  const now = Date.now();
  if (now - bcaLastFocusRequestAt < 1200 && document.visibilityState === "visible") return true;
  bcaLastFocusRequestAt = now;

  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "bca:focusCmsTab" }, async (response) => {
      if (chrome.runtime.lastError) {
        resolve(false);
        return;
      }
      if (response && response.ok) {
        await delay(250);
        resolve(true);
        return;
      }
      resolve(false);
    });
  });
}

function normalizeSpaces(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeKey(value) {
  return normalizeSpaces(value).toLowerCase();
}

function normalizeComparable(value) {
  return normalizeSpaces(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function compactComparable(value) {
  return normalizeComparable(value).replace(/\s+/g, "");
}

function spaceLetterNumberRuns(value) {
  return normalizeSpaces(value)
    .replace(/([A-Za-z])(\d)/g, "$1 $2")
    .replace(/(\d)([A-Za-z])/g, "$1 $2");
}

function comparableTextMatches(actual, expected) {
  const actualComparable = normalizeComparable(actual);
  const expectedComparable = normalizeComparable(expected);
  if (!actualComparable || !expectedComparable) return false;
  if (
    actualComparable === expectedComparable ||
    actualComparable.includes(expectedComparable) ||
    expectedComparable.includes(actualComparable)
  ) {
    return true;
  }

  const actualCompact = compactComparable(actual);
  const expectedCompact = compactComparable(expected);
  return (
    actualCompact === expectedCompact ||
    (expectedCompact.length >= 8 && actualCompact.includes(expectedCompact)) ||
    (actualCompact.length >= 8 && expectedCompact.includes(actualCompact))
  );
}

function titleCase(value) {
  return normalizeSpaces(value).replace(/\p{L}[\p{L}''-]*/gu, (word) => {
    if (word === word.toUpperCase() || word === word.toLowerCase()) {
      return word.charAt(0).toLocaleUpperCase() + word.slice(1).toLocaleLowerCase();
    }
    return word;
  });
}

function parseStopNumber(...values) {
  for (const value of values) {
    const text = String(value || "");
    const match = text.match(/(?:^|[^\d])0?(\d{1,3})(?:[.\s_-]|$)/);
    if (match) return Number(match[1]);
  }
  return null;
}

function cleanStopTitle(title, stopNumber) {
  let value = normalizeSpaces(title)
    .replace(/\.(wav|mp3|m4a|aac)$/i, "")
    .replace(/\bNA\b$/i, "")
    .replace(/\bUNAPPROVED\b/gi, "")
    .trim();

  if (stopNumber) {
    value = value.replace(new RegExp(`^0*${stopNumber}\\s*[.:-]?\\s*`, "i"), "");
  }

  if (value.includes(":")) {
    const parts = value.split(":");
    value = parts[parts.length - 1];
  }

  return titleCase(value || title);
}

function titleFromFileName(fileName, stopNumber) {
  return cleanStopTitle(fileName || "", stopNumber);
}

function chooseCatalogTitle({ fileName, includedIn, transcript, fallbackTitle, stopNumber }) {
  if (fallbackTitle) return normalizeSpaces(fallbackTitle);
  if (includedIn) return titleCase(includedIn);
  if (transcript && transcript.title) return transcript.title;
  return cleanStopTitle(fileName || "", stopNumber);
}

function activeTargetLanguages(config) {
  return (config.targetLanguages || []).filter((language) => language && language.enabled !== false);
}

function workItemLanguageIsEnabled(item, config) {
  return activeTargetLanguages(config).some(
    (language) => language.key === item.languageKey || language.cmsLabel === item.languageLabel
  );
}

function firstTranscriptParagraph(transcript) {
  for (const entry of sanitizedTranscriptEntries(transcript)) {
    for (const paragraph of Array.isArray(entry.paragraphs) ? entry.paragraphs : []) {
      const value = normalizeSpaces(paragraph);
      if (value) return value;
    }
  }
  return "";
}

function stripTitleDecorations(title) {
  return normalizeSpaces(title)
    .replace(/^["""''']+/, "")
    .replace(/["""''']+$/, "")
    .replace(/(?:\.{3}|…)+$/, "")
    .trim();
}

function takeWords(text, count) {
  const words = normalizeSpaces(text).split(/\s+/).filter(Boolean);
  if (!words.length) return "";
  if (words.length <= count) return words.join(" ");
  return `${words.slice(0, count).join(" ")}...`;
}

function translatedTitleFromCatalogTitle(sourceTitle, transcript, languageKey = "") {
  const source = normalizeSpaces(sourceTitle);
  if (!source) return titleCase(transcript && transcript.title ? transcript.title : "");

  const paragraph = firstTranscriptParagraph(transcript);
  if (!paragraph) return titleCase(transcript && transcript.title ? transcript.title : source);

  const sourceCore = stripTitleDecorations(source);
  const sourceWordCount = sourceCore.split(/\s+/).filter(Boolean).length;
  const isQuoted = /^["""]/.test(source);
  const isExcerpt = /(?:\.{3}|…)/.test(source) || isQuoted;
  let translated;
  if (isExcerpt) {
    translated = takeWords(paragraph, Math.max(4, sourceWordCount || 8));
  } else {
    const normalizedSource = normalizeComparable(source);
    const topic = normalizeSpaces(transcript && transcript.title ? transcript.title : "");
    if (topic && normalizedSource.startsWith("learn about") && normalizedSource.includes("at eastern state")) {
      const key = normalizeComparable(languageKey);
      if (key === "portuguese") translated = `Saiba mais sobre ${topic} na Eastern State`;
      else if (key === "german") translated = `Erfahren Sie mehr uber ${topic} im Eastern State`;
      else if (key === "french") translated = `En savoir plus sur ${topic} a Eastern State`;
      else if (key === "italian") translated = `Scopri di piu su ${topic} a Eastern State`;
      else if (key === "spanish") translated = `Aprende sobre ${topic} en Eastern State`;
    }
    translated = translated || titleCase(transcript.title || paragraph);
  }
  return isQuoted && translated ? `"${translated.replace(/^["""]+|["""]+$/g, "")}"` : translated;
}

function extractItemId(url) {
  const match = String(url || "").match(/\/catalog\/audios\/(\d+)/);
  return match ? match[1] : "";
}

function isBloombergAudioCatalogPage() {
  const url = new URL(window.location.href);
  return url.origin === "https://cms.bloombergconnects.org" && url.pathname.replace(/\/+$/, "") === "/catalog/audios";
}

function isBloombergAudioEditPage() {
  const url = new URL(window.location.href);
  return url.origin === "https://cms.bloombergconnects.org" && /\/catalog\/audios\/\d+/.test(url.pathname);
}

function getChromeUrl(path) {
  return chrome.runtime.getURL(path);
}

async function fetchJsonIfExists(path) {
  const response = await fetch(getChromeUrl(path));
  if (!response.ok) {
    throw new Error(`Could not load ${path}.`);
  }
  return response.json();
}

async function loadConfig() {
  if (bcaConfig) return bcaConfig;
  try {
    bcaConfig = await fetchJsonIfExists("config.json");
  } catch (error) {
    bcaConfig = await fetchJsonIfExists("config.default.json");
  }
  return bcaConfig;
}

async function loadTranscripts() {
  if (bcaTranscripts) return bcaTranscripts;
  bcaTranscripts = await fetchJsonIfExists("data/transcripts.json");
  return bcaTranscripts;
}

function defaultState() {
  return {
    status: "idle",
    phase: "idle",
    index: 0,
    items: [],
    workItems: [],
    completedTasks: {},
    audioItemIndex: {},
    routeRepairAttempts: {},
    pendingOpenItem: null,
    current: null,
    transcriptCount: 0,
    counts: {
      totalRows: 0,
      eligibleStops: 0,
      excludedStops: 0,
      totalLanguageTasks: 0,
      ready: 0,
      skipped: 0,
      alreadyComplete: 0,
      processed: 0,
      failed: 0,
      missingTranscript: 0,
    },
    log: [],
    updatedAt: new Date().toISOString(),
  };
}

async function readState() {
  const data = await storageGet(BCA_STATE_KEY);
  return data[BCA_STATE_KEY] || defaultState();
}

async function writeState(state) {
  const nextState = {
    ...state,
    updatedAt: new Date().toISOString(),
    log: (state.log || []).slice(-250),
  };
  await storageSet({ [BCA_STATE_KEY]: nextState });
  return nextState;
}

function itemLabel(item) {
  if (!item) return "";
  const stop = item.stopNumber ? `${item.stopNumber} ` : "";
  const lang = item.languageLabel ? ` / ${item.languageLabel}` : "";
  return `${stop}${item.title || item.fileName || item.itemId}${lang}`.trim();
}

function pushLog(state, level, message, item, status) {
  state.log = state.log || [];
  state.log.push({
    level,
    status: status || level,
    message,
    itemId: item ? item.itemId || "" : "",
    itemLabel: itemLabel(item),
    at: new Date().toISOString(),
  });
}

function getCells(row) {
  return Array.from(row.children).filter((child) => child.matches("td, th"));
}

function getCatalogRows() {
  const tableRows = [];
  document.querySelectorAll("table").forEach((table) => {
    const headerMap = buildHeaderMap(table);
    const bodyRows = Array.from(table.querySelectorAll("tbody tr"));
    const rows = bodyRows.length ? bodyRows : Array.from(table.querySelectorAll("tr")).slice(1);
    rows.forEach((row) => {
      const cells = getCells(row);
      if (cells.length >= 2 && !row.querySelector("th")) {
        tableRows.push({ row, cells, headerMap });
      }
    });
  });

  if (tableRows.length) return tableRows;

  const roleRows = Array.from(document.querySelectorAll('[role="row"]')).slice(1);
  return roleRows
    .map((row) => {
      const cells = Array.from(row.querySelectorAll('[role="cell"], [role="gridcell"]'));
      return {
        row,
        cells,
        headerMap: {
          title: 1,
          fileName: 2,
          languages: 4,
        },
      };
    })
    .filter((entry) => entry.cells.length >= 2);
}

function catalogRowDataFromEntry({ row, cells, headerMap }) {
  const title = meaningfulText(cells[headerMap.title]);
  const fileName = meaningfulText(cells[headerMap.fileName]);
  const includedIn = meaningfulText(cells[headerMap.includedIn]);
  const languagesText = meaningfulText(cells[headerMap.languages]);
  const editUrl = findEditUrl(row, cells, headerMap);
  const itemId = extractItemId(editUrl);
  const stopNumber = parseStopNumber(fileName);
  return {
    row,
    cells,
    headerMap,
    title,
    fileName,
    includedIn,
    languagesText,
    editUrl,
    itemId,
    stopNumber,
    rowKey: buildRowKey(stopNumber, fileName, includedIn || title),
  };
}

function findCatalogScrollContainer() {
  const candidates = Array.from(document.querySelectorAll("main, section, div, [role='grid'], [role='table']"))
    .map((element) => {
      const style = window.getComputedStyle(element);
      const overflow = `${style.overflowY} ${style.overflow}`;
      const scrollable = element.scrollHeight - element.clientHeight > 24;
      const hasRows = Boolean(element.querySelector("table tr, [role='row']"));
      const rect = element.getBoundingClientRect();
      return { element, scrollable, hasRows, area: rect.width * rect.height, overflow };
    })
    .filter(({ scrollable, hasRows, overflow }) => scrollable && hasRows && !overflow.includes("hidden"))
    .sort((a, b) => b.element.scrollHeight - b.element.clientHeight - (a.element.scrollHeight - a.element.clientHeight));

  if (candidates.length) return candidates[0].element;
  return document.scrollingElement || document.documentElement;
}

async function collectCatalogRowDataAcrossScroll() {
  const container = findCatalogScrollContainer();
  const originalScrollTop = container.scrollTop;
  const rowMap = new Map();
  let stagnantPasses = 0;
  let lastScrollTop = -1;

  const collectVisible = () => {
    let added = 0;
    getCatalogRows().forEach((entry) => {
      const data = catalogRowDataFromEntry(entry);
      if (!data.fileName && !data.title) return;
      const key = data.rowKey || `${data.stopNumber || ""}:${normalizeComparable(data.fileName)}:${normalizeComparable(data.title)}`;
      if (!rowMap.has(key)) {
        rowMap.set(key, data);
        added += 1;
      }
    });
    return added;
  };

  container.scrollTop = 0;
  await delay(250);

  for (let pass = 0; pass < 80; pass += 1) {
    const added = collectVisible();
    const maxScrollTop = Math.max(0, container.scrollHeight - container.clientHeight);
    if (container.scrollTop >= maxScrollTop - 4) break;

    const nextScrollTop = Math.min(maxScrollTop, container.scrollTop + Math.max(280, Math.floor(container.clientHeight * 0.85)));
    if (nextScrollTop === lastScrollTop || added === 0) stagnantPasses += 1;
    else stagnantPasses = 0;
    if (stagnantPasses >= 8) break;

    lastScrollTop = container.scrollTop;
    container.scrollTop = nextScrollTop;
    await delay(220);
  }

  collectVisible();
  container.scrollTop = originalScrollTop;
  await delay(120);
  return Array.from(rowMap.values()).sort((a, b) => (a.stopNumber || 9999) - (b.stopNumber || 9999));
}

function firstObjectValue(object, keys) {
  if (!object || typeof object !== "object") return "";
  for (const key of keys) {
    if (typeof object[key] === "string" || typeof object[key] === "number") return String(object[key]);
  }
  return "";
}

function collectAudioRecordsFromObject(value, records, seen = new WeakSet()) {
  if (!value || typeof value !== "object") return;
  if (seen.has(value)) return;
  seen.add(value);

  if (Array.isArray(value)) {
    value.forEach((entry) => collectAudioRecordsFromObject(entry, records, seen));
    return;
  }

  const fileName = firstObjectValue(value, ["fileName", "filename", "file_name", "originalFileName", "original_filename", "name"]);
  const itemId = firstObjectValue(value, ["id", "itemId", "audioId", "audio_id"]);
  const stopNumber = parseStopNumber(fileName);
  const looksLikeAudioFile = /\.(wav|mp3|m4a|aac)$/i.test(fileName);
  const numericItemId = /^\d{4,}$/.test(itemId) ? itemId : "";
  if (looksLikeAudioFile && stopNumber && numericItemId) {
    const title = firstObjectValue(value, ["title", "displayTitle", "name"]);
    const includedIn = firstObjectValue(value, ["includedIn", "included_in", "tourStopTitle", "stopTitle"]);
    records.push({
      title,
      fileName,
      includedIn,
      languagesText: "",
      itemId: numericItemId,
      editUrl: `https://cms.bloombergconnects.org/catalog/audios/${numericItemId}`,
      stopNumber,
      rowKey: buildRowKey(stopNumber, fileName, includedIn || title),
    });
  }

  Object.values(value).forEach((entry) => collectAudioRecordsFromObject(entry, records, seen));
}

function collectAudioRecordsFromBrowserStorage() {
  const records = [];
  const readStorage = (storage) => {
    for (let index = 0; index < storage.length; index += 1) {
      const key = storage.key(index);
      const raw = key ? storage.getItem(key) : "";
      if (!raw || raw.length > 5_000_000) continue;
      if (!/audio|catalog|apollo|query|redux|persist|bloomberg|docent/i.test(`${key} ${raw.slice(0, 500)}`)) continue;
      try {
        collectAudioRecordsFromObject(JSON.parse(raw), records);
      } catch (_error) {
        // Non-JSON storage values are expected in the CMS app.
      }
    }
  };
  readStorage(window.localStorage);
  readStorage(window.sessionStorage);
  return records;
}

function buildHeaderMap(table) {
  const headers = Array.from(table.querySelectorAll("thead th"));
  const cells = headers.length
    ? headers
    : getCells(Array.from(table.querySelectorAll("tr")).find((row) => row.querySelector("th")) || document.createElement("tr"));
  const map = {};

  cells.forEach((cell, index) => {
    const text = normalizeKey(cell.textContent);
    if (text.includes("title")) map.title = index;
    if (text.includes("file name")) map.fileName = index;
    if (text.includes("included in")) map.includedIn = index;
    if (text.includes("languages")) map.languages = index;
    if (text.includes("updated")) map.updated = index;
  });

  return {
    title: map.title ?? 1,
    fileName: map.fileName ?? 2,
    includedIn: map.includedIn ?? 3,
    languages: map.languages ?? 4,
  };
}

function meaningfulText(element) {
  if (!element) return "";
  const clone = element.cloneNode(true);
  clone.querySelectorAll("button, svg, path, script, style, [aria-hidden='true']").forEach((node) => node.remove());
  return normalizeSpaces(clone.textContent || "");
}

function findEditUrl(row, cells, headerMap) {
  const titleCell = cells[headerMap.title] || row;
  const link =
    titleCell.querySelector('a[href*="/catalog/audios/"]') ||
    row.querySelector('a[href*="/catalog/audios/"]');
  if (link && link.href) return link.href;

  const candidates = [row, ...Array.from(row.querySelectorAll("*"))];
  for (const element of candidates) {
    for (const attribute of Array.from(element.attributes || [])) {
      const value = String(attribute.value || "");
      const pathMatch = value.match(/\/catalog\/audios\/\d+/);
      if (pathMatch) return new URL(pathMatch[0], window.location.origin).href;
    }
  }
  return "";
}

function buildRowKey(stopNumber, fileName, title) {
  return [stopNumber || "", normalizeComparable(fileName), normalizeComparable(title)].join("|");
}

function buildAudioEditUrl(item, config) {
  if (item && item.editUrl) return item.editUrl;
  if (!item || !item.itemId) return "";
  const baseCatalog = ((config && config.cmsCatalogUrl) || "https://cms.bloombergconnects.org/catalog/audios").replace(/\/+$/, "");
  return `${baseCatalog}/${item.itemId}`;
}

function audioItemAliases(item) {
  if (!item) return [];
  const aliases = new Set();
  const stop = item.stopNumber || parseStopNumber(item.fileName, item.title, item.includedIn);
  const add = (prefix, value) => {
    const text = compactComparable(value);
    if (stop && text) aliases.add(`${prefix}:${stop}:${text}`);
  };
  if (item.itemId) aliases.add(`id:${item.itemId}`);
  if (item.rowKey) aliases.add(`row:${item.rowKey}`);
  add("file", item.fileName);
  add("file", spaceLetterNumberRuns(item.fileName || ""));
  add("title", item.title);
  add("title", item.rawTitle);
  add("title", item.sourceTitle);
  add("included", item.includedIn);
  return Array.from(aliases);
}

function audioIndexRecordFromCatalogRow(rowData, config) {
  const editUrl = rowData.editUrl || "";
  const itemId = rowData.itemId || extractItemId(editUrl);
  const stopNumber = rowData.stopNumber || parseStopNumber(rowData.fileName, rowData.includedIn, rowData.title);
  const rowKey = rowData.rowKey || buildRowKey(stopNumber, rowData.fileName, rowData.includedIn || rowData.title);
  return {
    itemId,
    editUrl: itemId ? buildAudioEditUrl({ itemId, editUrl }, config) : editUrl,
    stopNumber,
    fileName: rowData.fileName || "",
    includedIn: rowData.includedIn || "",
    sourceTitle: rowData.title || "",
    title: rowData.title || "",
    rowKey,
    rowText: normalizeSpaces([rowData.title, rowData.fileName, rowData.includedIn, rowData.languagesText].filter(Boolean).join(" | ")),
    lastSeenAt: new Date().toISOString(),
  };
}

function mergeAudioItemIndex(existingIndex, rowRecords, config) {
  const nextIndex = { ...(existingIndex || {}) };
  for (const rowData of rowRecords || []) {
    const record = audioIndexRecordFromCatalogRow(rowData, config);
    if (!record.stopNumber && !record.itemId && !record.rowKey) continue;
    for (const alias of audioItemAliases(record)) {
      const existing = nextIndex[alias] || {};
      nextIndex[alias] = {
        ...existing,
        ...record,
        itemId: record.itemId || existing.itemId || "",
        editUrl: record.editUrl || existing.editUrl || "",
      };
    }
  }
  return nextIndex;
}

function uniqueAudioIndexRecords(index) {
  const seen = new Set();
  const records = [];
  for (const record of Object.values(index || {})) {
    const key = record.itemId || record.rowKey || `${record.stopNumber}:${normalizeComparable(record.fileName)}:${normalizeComparable(record.includedIn || record.title)}`;
    if (!key || seen.has(key)) continue;
    seen.add(key);
    records.push(record);
  }
  return records;
}

function audioIndexMatchScore(item, record) {
  if (!item || !record) return 0;
  if (item.itemId && record.itemId && item.itemId === record.itemId) return 1000;
  if (item.rowKey && record.rowKey && item.rowKey === record.rowKey) return 900;
  if (item.stopNumber && record.stopNumber && item.stopNumber !== record.stopNumber) return 0;

  let score = item.stopNumber && record.stopNumber && item.stopNumber === record.stopNumber ? 120 : 0;
  if (item.fileName && record.fileName && comparableTextMatches(record.fileName, item.fileName)) score += 180;
  if (item.includedIn && record.includedIn && comparableTextMatches(record.includedIn, item.includedIn)) score += 80;
  if (item.title && (record.title || record.sourceTitle) && comparableTextMatches(record.title || record.sourceTitle, item.title)) score += 60;
  if (item.rawTitle && record.sourceTitle && comparableTextMatches(record.sourceTitle, item.rawTitle)) score += 60;
  return score;
}

function resolveAudioItemFromIndex(state, item, config) {
  const index = state && state.audioItemIndex ? state.audioItemIndex : {};
  for (const alias of audioItemAliases(item)) {
    const record = index[alias];
    if (record && (record.itemId || record.editUrl)) {
      return applyAudioIndexRecordToItem(item, record, config);
    }
  }

  const scored = uniqueAudioIndexRecords(index)
    .map((record) => ({ record, score: audioIndexMatchScore(item, record) }))
    .filter(({ record, score }) => score >= 220 && (record.itemId || record.editUrl))
    .sort((a, b) => b.score - a.score);
  if (!scored.length) return null;
  if (scored.length > 1 && scored[0].score === scored[1].score && scored[0].record.itemId !== scored[1].record.itemId) return null;
  return applyAudioIndexRecordToItem(item, scored[0].record, config);
}

function applyAudioIndexRecordToItem(item, record, config) {
  const itemId = record.itemId || extractItemId(record.editUrl);
  const editUrl = buildAudioEditUrl({ itemId, editUrl: record.editUrl }, config);
  const audioCacheKey = itemId ? `${itemId}:${(config && config.sourceLanguageCode) || "en-US"}` : item.audioCacheKey || "";
  return {
    ...item,
    itemId: itemId || item.itemId || "",
    editUrl: editUrl || item.editUrl || "",
    audioCacheKey,
    rowKey: item.rowKey || record.rowKey || "",
    fileName: item.fileName || record.fileName || "",
    includedIn: item.includedIn || record.includedIn || "",
    sourceTitle: item.sourceTitle || record.sourceTitle || "",
  };
}

function sameAudioItem(candidate, item) {
  if (!candidate || !item) return false;
  if (candidate.itemId && item.itemId && candidate.itemId === item.itemId) return true;
  if (candidate.rowKey && item.rowKey && candidate.rowKey === item.rowKey) return true;
  if (candidate.stopNumber && item.stopNumber && candidate.stopNumber !== item.stopNumber) return false;
  if (candidate.fileName && item.fileName && comparableTextMatches(candidate.fileName, item.fileName)) return true;
  return Boolean(candidate.stopNumber && item.stopNumber && candidate.stopNumber === item.stopNumber);
}

function applyResolvedAudioRouteToState(state, item, resolved, config) {
  if (!state || !resolved || (!resolved.itemId && !resolved.editUrl)) return state;
  const hydrated = applyAudioIndexRecordToItem(item, resolved, config);
  const update = (candidate) => (sameAudioItem(candidate, item) ? { ...candidate, ...hydrated, languageKey: candidate.languageKey, languageLabel: candidate.languageLabel } : candidate);
  state.items = (state.items || []).map(update);
  state.workItems = (state.workItems || []).map(update);
  return state;
}

function routeResolutionKey(item) {
  return [item && item.stopNumber, normalizeComparable(item && item.fileName), item && item.rowKey, item && item.languageKey].join("|");
}

function parseExistingLanguages(text, languages) {
  const normalized = normalizeComparable(text);
  return languages.filter((language) => normalized.includes(normalizeComparable(language.cmsLabel))).map((language) => language.cmsLabel);
}

function isExcludedStop(stopNumber, text, config) {
  if (!stopNumber) return true;
  if ((config.excludeStops || []).includes(stopNumber)) return true;
  const allowlist = Array.isArray(config.stopAllowlist) ? config.stopAllowlist.filter(Boolean) : [];
  if (allowlist.length && !allowlist.includes(stopNumber)) return true;
  const normalized = normalizeComparable(text);
  return (config.excludeTitlePatterns || []).some((pattern) => normalized.includes(normalizeComparable(pattern)));
}

function transcriptFor(transcripts, languageKey, stopNumber) {
  return transcripts && transcripts.languages && transcripts.languages[languageKey]
    ? transcripts.languages[languageKey].stops[String(stopNumber)] || null
    : null;
}

async function scanCatalog() {
  assertExtensionBuildCurrent();
  const [config, transcripts] = await Promise.all([loadConfig(), loadTranscripts()]);
  const targetLanguages = activeTargetLanguages(config);
  if (!isBloombergAudioCatalogPage()) {
    throw new Error("Open the Bloomberg Connects Audios catalog before scanning.");
  }

  const previous = await readState().catch(() => null);
  const previousCompletedTasks = (previous && previous.completedTasks) || {};
  const previousAudioItemIndex = (previous && previous.audioItemIndex) || {};

  const state = defaultState();
  state.completedTasks = { ...previousCompletedTasks };
  state.audioItemIndex = { ...previousAudioItemIndex };
  state.status = "scanned";
  state.phase = "dry-run";
  state.transcriptCount = Object.values(transcripts.languages || {}).reduce(
    (sum, language) => sum + Object.keys(language.stops || {}).length,
    0
  );

  const items = [];
  const workItems = [];
  let totalRows = 0;
  let excludedStops = 0;
  let missingTranscript = 0;
  let alreadyComplete = 0;

  await clearCatalogSearch();
  const catalogRows = await collectCatalogRowDataAcrossScroll();
  const storageRows = collectAudioRecordsFromBrowserStorage();
  state.audioItemIndex = mergeAudioItemIndex(state.audioItemIndex, [...catalogRows, ...storageRows], config);

  catalogRows.forEach((rowData) => {
    totalRows += 1;

    const { title, fileName, includedIn, languagesText, editUrl, itemId, stopNumber, rowKey } = rowData;
    const sourceText = `${fileName} ${includedIn}`;

    if (isExcludedStop(stopNumber, sourceText, config)) {
      excludedStops += 1;
      return;
    }

    const existingLanguages = parseExistingLanguages(languagesText, targetLanguages);
    const representativeTranscript = targetLanguages
      .map((language) => transcriptFor(transcripts, language.key, stopNumber))
      .find(Boolean);
    const displayTitle = chooseCatalogTitle({
      fileName,
      includedIn,
      transcript: representativeTranscript,
      fallbackTitle: title,
      stopNumber,
    });
    const item = resolveAudioItemFromIndex(
      state,
      {
        itemId,
        stopNumber,
        title: displayTitle,
        sourceTitle: title,
        rawTitle: title,
        includedIn,
        fileName,
        editUrl,
        rowKey: rowKey || buildRowKey(stopNumber, fileName, includedIn || displayTitle),
        existingLanguages,
      },
      config
    ) || {
      itemId,
      stopNumber,
      title: displayTitle,
      sourceTitle: title,
      rawTitle: title,
      includedIn,
      fileName,
      editUrl,
      rowKey: rowKey || buildRowKey(stopNumber, fileName, includedIn || displayTitle),
      existingLanguages,
    };
    items.push(item);

    for (const language of targetLanguages) {
      const transcript = transcriptFor(transcripts, language.key, stopNumber);
      const sourceCatalogTitle = normalizeSpaces(title || "");
      const translatedTitle = transcript ? translatedTitleFromCatalogTitle(sourceCatalogTitle || displayTitle, transcript, language.key) : "";
      const task = {
        ...item,
        languageKey: language.key,
        languageLabel: language.cmsLabel,
        accessibilityLabel: language.accessibilityLabel || language.cmsLabel,
        transcriptTitle: transcript ? transcript.title : "",
        translatedTitle,
        hasTranscript: Boolean(transcript),
        audioCacheKey: item.itemId ? `${item.itemId}:${config.sourceLanguageCode || "en-US"}` : "",
      };

      if (Array.isArray(state.completedTasks[item.rowKey]) && state.completedTasks[item.rowKey].includes(language.key)) {
        pushLog(state, "info", "Previously marked complete; queued for edit-page verification.", task, "verifyComplete");
      }

      if (!transcript) {
        missingTranscript += 1;
        pushLog(state, "error", "Missing transcript section.", task, "missingTranscript");
        continue;
      }

      workItems.push({ ...task, status: "pending" });
    }
  });

  state.items = items;
  state.workItems = workItems;
  state.counts = {
    totalRows,
    eligibleStops: items.length,
    excludedStops,
    totalLanguageTasks: workItems.length + alreadyComplete + missingTranscript,
    ready: workItems.length,
    skipped: excludedStops,
    alreadyComplete,
    processed: 0,
    failed: 0,
    missingTranscript,
  };
  pushLog(state, "info", `Dry run found ${workItems.length} ready language tasks across ${items.length} eligible stops.`);
  return writeState(state);
}

function visibleCatalogRowSamples(limit = 8) {
  return getCatalogRows()
    .slice(0, limit)
    .map((entry) => {
      const data = catalogRowDataFromEntry(entry);
      return normalizeSpaces([data.fileName, data.includedIn, data.title].filter(Boolean).join(" | "));
    })
    .filter(Boolean);
}

function unresolvedRouteItems(state) {
  const config = bcaConfig || {};
  const uniqueItems = uniqueAudioItems([...(state.items || []), ...(state.workItems || [])]);
  return uniqueItems.filter((item) => {
    if (item.itemId || item.editUrl) return false;
    return !resolveAudioItemFromIndex(state, item, config);
  });
}

async function repairAudioItemIndex({ state = null, silent = false } = {}) {
  assertExtensionBuildCurrent();
  const config = await loadConfig();
  const workingState = state || (await readState());

  if (!isBloombergAudioCatalogPage()) {
    window.location.href = config.cmsCatalogUrl;
    pushLog(workingState, "warn", "Repair Item Index opened the Audios catalog. Run repair again after the catalog loads.");
    return writeState(workingState);
  }

  await clearCatalogSearch();
  const rows = await collectCatalogRowDataAcrossScroll();
  const storageRows = collectAudioRecordsFromBrowserStorage();
  workingState.audioItemIndex = mergeAudioItemIndex(workingState.audioItemIndex || {}, [...rows, ...storageRows], config);

  const hydrate = (item) => {
    const resolved = resolveAudioItemFromIndex(workingState, item, config);
    return resolved || item;
  };
  workingState.items = (workingState.items || []).map(hydrate);
  workingState.workItems = (workingState.workItems || []).map(hydrate);

  const unresolved = unresolvedRouteItems(workingState);
  workingState.routeRepairAttempts = {};
  const routeItemCount = uniqueAudioItems([...(workingState.items || []), ...(workingState.workItems || [])]).length;
  const resolvedCount = Math.max(0, routeItemCount - unresolved.length);
  if (!silent) {
    if (unresolved.length) {
      pushLog(
        workingState,
        "warn",
        `Repair Item Index scanned ${rows.length} rows and resolved ${resolvedCount} stops. Unresolved: ${unresolved
          .slice(0, 6)
          .map((item) => item.fileName || item.title)
          .join("; ")}.`
      );
    } else {
      pushLog(workingState, "success", `Repair Item Index scanned ${rows.length} rows and all queued stops have route data.`);
    }
  }
  return writeState(workingState);
}

function openAudioDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(BCA_DB_NAME, 2);
    request.onerror = () => reject(request.error || new Error("Could not open audio cache."));
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(BCA_AUDIO_STORE)) {
        db.createObjectStore(BCA_AUDIO_STORE, { keyPath: "key" });
      }
      if (!db.objectStoreNames.contains(BCA_SETTINGS_STORE)) {
        db.createObjectStore(BCA_SETTINGS_STORE, { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
  });
}

async function audioCacheGet(key) {
  const db = await openAudioDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(BCA_AUDIO_STORE, "readonly");
    const request = tx.objectStore(BCA_AUDIO_STORE).get(key);
    request.onerror = () => reject(request.error || new Error("Could not read audio cache."));
    request.onsuccess = () => resolve(request.result || null);
    tx.oncomplete = () => db.close();
  });
}

async function audioCachePut(record) {
  const db = await openAudioDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(BCA_AUDIO_STORE, "readwrite");
    const request = tx.objectStore(BCA_AUDIO_STORE).put(record);
    request.onerror = () => reject(request.error || new Error("Could not write audio cache."));
    request.onsuccess = () => resolve(record);
    tx.oncomplete = () => db.close();
  });
}

async function dbSettingGet(key) {
  const db = await openAudioDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(BCA_SETTINGS_STORE, "readonly");
    const request = tx.objectStore(BCA_SETTINGS_STORE).get(key);
    request.onerror = () => reject(request.error || new Error("Could not read extension setting."));
    request.onsuccess = () => resolve(request.result || null);
    tx.oncomplete = () => db.close();
  });
}

async function dbSettingPut(record) {
  const db = await openAudioDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(BCA_SETTINGS_STORE, "readwrite");
    const request = tx.objectStore(BCA_SETTINGS_STORE).put(record);
    request.onerror = () => reject(request.error || new Error("Could not write extension setting."));
    request.onsuccess = () => resolve(record);
    tx.oncomplete = () => db.close();
  });
}

async function ensureAudioFolderPermission(handle, requestAccess = false, mode = "read") {
  if (!handle || typeof handle.queryPermission !== "function") return false;
  const options = { mode };
  let permission = await handle.queryPermission(options);
  if (permission !== "granted" && requestAccess && typeof handle.requestPermission === "function") {
    permission = await handle.requestPermission(options);
  }
  return permission === "granted";
}

async function getAudioFolderHandle(requestAccess = false, mode = "read") {
  if (bcaAudioFolderHandle && (await ensureAudioFolderPermission(bcaAudioFolderHandle, requestAccess, mode))) {
    return bcaAudioFolderHandle;
  }

  const stored = await dbSettingGet(BCA_AUDIO_FOLDER_HANDLE_KEY).catch(() => null);
  const handle = stored && stored.handle;
  if (handle && (await ensureAudioFolderPermission(handle, requestAccess, mode))) {
    bcaAudioFolderHandle = handle;
    bcaAudioFolderMeta = {
      name: stored.name || handle.name || "Selected folder",
      selectedAt: stored.selectedAt || "",
    };
    return handle;
  }

  return null;
}

async function countAudioFilesInFolder(directoryHandle, limit = 500) {
  let count = 0;
  for await (const entry of walkAudioFolder(directoryHandle)) {
    count += 1;
    if (count >= limit) return count;
  }
  return count;
}

async function selectLocalAudioFolder() {
  if (typeof window.showDirectoryPicker !== "function") {
    throw new Error("This Chrome profile does not expose folder picking to the CMS page.");
  }

  const handle = await window.showDirectoryPicker({ mode: "readwrite" });
  const granted = await ensureAudioFolderPermission(handle, true, "readwrite");
  if (!granted) throw new Error("Chrome did not grant read access to the selected audio folder.");

  const selectedAt = new Date().toISOString();
  bcaAudioFolderHandle = handle;
  bcaAudioFolderMeta = { name: handle.name || "Selected folder", selectedAt };
  await dbSettingPut({ key: BCA_AUDIO_FOLDER_HANDLE_KEY, handle, name: bcaAudioFolderMeta.name, selectedAt });
  await storageSet({ [BCA_AUDIO_FOLDER_META_KEY]: bcaAudioFolderMeta });

  const audioCount = await countAudioFilesInFolder(handle);
  const state = await readState();
  pushLog(state, "success", `Selected local audio folder: ${bcaAudioFolderMeta.name} (${audioCount} audio files found).`);
  await writeState(state);
  return bcaAudioFolderMeta;
}

async function* walkAudioFolder(directoryHandle, depth = 0) {
  if (!directoryHandle || depth > 3 || typeof directoryHandle.values !== "function") return;
  for await (const entry of directoryHandle.values()) {
    if (entry.kind === "file" && isAudioLikeFileName(entry.name || "")) {
      yield { fileHandle: entry, directoryHandle };
    } else if (entry.kind === "directory") {
      yield* walkAudioFolder(entry, depth + 1);
    }
  }
}

function isAudioLikeFileName(fileName) {
  const name = String(fileName || "");
  if (/\.(crdownload|download|tmp|part)$/i.test(name)) return false;
  return /\.(wav|mp3|m4a|aac|ogg|flac)$/i.test(name) || normalizeComparable(name).includes("sourceurl");
}

function localAudioTitleKey(fileName, stopNumber) {
  return normalizeComparable(cleanStopTitle(sanitizeAudioFileName(fileName || ""), stopNumber));
}

function localAudioMatchScore(fileName, item) {
  if (!fileName || !item || parseStopNumber(fileName) !== item.stopNumber) return 0;
  const localTitle = localAudioTitleKey(fileName, item.stopNumber);
  if (!localTitle) return 0;

  const targetTitles = [item.fileName, item.title, item.includedIn, item.rawTitle]
    .map((value) => localAudioTitleKey(value, item.stopNumber))
    .filter(Boolean);
  if (targetTitles.some((target) => target === localTitle)) return 100;
  if (targetTitles.some((target) => target && (target.includes(localTitle) || localTitle.includes(target)))) return 92;

  const targetTokens = new Set(
    targetTitles
      .join(" ")
      .split(" ")
      .filter((token) => token.length >= 4 && !["audio", "tour", "stop"].includes(token))
  );
  if (!targetTokens.size) return 0;
  const localTokens = new Set(localTitle.split(" ").filter((token) => token.length >= 4));
  const overlap = Array.from(targetTokens).filter((token) => localTokens.has(token)).length;
  const ratio = overlap / targetTokens.size;
  return ratio >= 0.75 ? 80 : 0;
}

async function findLocalAudioEntryForItem(item) {
  const handle = await getAudioFolderHandle(false);
  if (!handle || !item.stopNumber) return null;

  const candidates = [];
  for await (const { fileHandle, directoryHandle } of walkAudioFolder(handle)) {
    const score = localAudioMatchScore(fileHandle.name, item);
    if (score > 0) candidates.push({ score, fileHandle, directoryHandle });
  }

  candidates.sort((a, b) => b.score - a.score || a.fileHandle.name.localeCompare(b.fileHandle.name));
  const best = candidates[0];
  if (!best || best.score < 80) return null;
  return { fileHandle: best.fileHandle, directoryHandle: best.directoryHandle, file: await best.fileHandle.getFile(), matchScore: best.score };
}

async function buildLocalAudioRecord(item) {
  const entry = await findLocalAudioEntryForItem(item);
  const file = entry && entry.file;
  if (!file || !file.size) return null;
  return {
    key: item.audioCacheKey,
    itemId: item.itemId,
    stopNumber: item.stopNumber,
    fileName: file.name,
    contentType: file.type || "audio/wav",
    size: file.size,
    blob: file,
    source: "localFolder",
    downloadedAt: new Date().toISOString(),
  };
}

function audioExtensionFor(blob, fileName) {
  const extension = /\.([a-z0-9]+)$/i.exec(fileName || "")?.[1] || "";
  if (/^(wav|mp3|m4a|aac|ogg|flac)$/i.test(extension)) return extension.toLowerCase();
  if (blob && blob.type) {
    if (blob.type.includes("mpeg")) return "mp3";
    if (blob.type.includes("mp4") || blob.type.includes("m4a")) return "m4a";
    if (blob.type.includes("aac")) return "aac";
    if (blob.type.includes("ogg")) return "ogg";
    if (blob.type.includes("flac")) return "flac";
  }
  return "wav";
}

function sanitizeAudioFileName(fileName) {
  return normalizeSpaces(
    String(fileName || "")
      .replace(/easternStatePenitentiary_?UNAPPROVED_?/gi, "")
      .replace(/easternStatePenitentiary_?UNNAPROVED_?/gi, "")
      .replace(/easternStatePenitentiary_?/gi, "")
      .replace(/_?sourceUrl\b/gi, "")
      .replace(/[\\/:*?"<>|]/g, " ")
      .replace(/\s+/g, " ")
  ).trim();
}

function cleanLocalAudioFileName(item, blob) {
  const extension = audioExtensionFor(blob, item.fileName);
  const stop = String(item.stopNumber).padStart(3, "0");
  return sanitizeAudioFileName(`${stop} ${cleanStopTitle(item.title || item.fileName, item.stopNumber)}.${extension}`);
}

async function writeBlobToAudioFolder(directoryHandle, fileName, blob) {
  const safeName = sanitizeAudioFileName(fileName);
  const fileHandle = await directoryHandle.getFileHandle(safeName, { create: true });
  const writable = await fileHandle.createWritable();
  await writable.write(blob);
  await writable.close();
  return safeName;
}

async function writeResponseToAudioFolder(directoryHandle, fileName, response) {
  const safeName = sanitizeAudioFileName(fileName);
  const fileHandle = await directoryHandle.getFileHandle(safeName, { create: true });
  const writable = await fileHandle.createWritable();
  try {
    if (response.body && typeof response.body.pipeTo === "function") {
      await response.body.pipeTo(writable);
    } else {
      await writable.write(await response.blob());
      await writable.close();
    }
  } catch (error) {
    try {
      await writable.abort();
    } catch {}
    throw error;
  }
  return safeName;
}

async function hasFileInDirectory(directoryHandle, fileName) {
  try {
    await directoryHandle.getFileHandle(fileName, { create: false });
    return true;
  } catch {
    return false;
  }
}

async function normalizeLocalAudioEntry(entry, item, state) {
  if (!entry || !entry.file || !entry.directoryHandle || !entry.fileHandle) return false;
  const cleanName = cleanLocalAudioFileName(item, entry.file);
  if (entry.file.name === cleanName) return false;

  if (await hasFileInDirectory(entry.directoryHandle, cleanName)) {
    pushLog(state, "warn", `Could not rename ${entry.file.name}; ${cleanName} already exists.`, item, "skipped");
    return false;
  }

  await writeBlobToAudioFolder(entry.directoryHandle, cleanName, entry.file);
  await entry.directoryHandle.removeEntry(entry.fileHandle.name);
  pushLog(state, "success", `Renamed local audio ${entry.file.name} to ${cleanName}.`, item);
  return true;
}

async function snapshotAudioFolder(directoryHandle) {
  const files = new Map();
  for await (const entry of walkAudioFolder(directoryHandle)) {
    const file = await entry.fileHandle.getFile();
    files.set(entry.fileHandle.name, {
      ...entry,
      file,
      signature: `${file.size}:${file.lastModified}`,
    });
  }
  return files;
}

function findNewOrChangedAudioEntry(before, after) {
  const changed = [];
  for (const [name, entry] of after) {
    const previous = before.get(name);
    if (!previous || previous.signature !== entry.signature) changed.push(entry);
  }
  changed.sort((a, b) => (b.file.lastModified || 0) - (a.file.lastModified || 0));
  return changed[0] || null;
}

async function waitForDownloadedAudioFile(directoryHandle, before, item, timeoutMs = 45000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const after = await snapshotAudioFolder(directoryHandle);
    const stopMatch = Array.from(after.values())
      .filter((entry) => parseStopNumber(entry.fileHandle.name) === item.stopNumber)
      .sort((a, b) => (b.file.lastModified || 0) - (a.file.lastModified || 0))[0];
    if (stopMatch) return stopMatch;

    const changed = findNewOrChangedAudioEntry(before, after);
    if (changed) return changed;
    await delay(1000);
  }
  return null;
}

function uniqueAudioItems(items) {
  const seen = new Set();
  const result = [];
  for (const item of items || []) {
    const key = item.itemId || item.rowKey || `${item.stopNumber}:${item.fileName}`;
    if (!item.stopNumber || seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }
  return result;
}

async function waitForLocation(test, timeoutMs = 12000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    if (test()) return true;
    await delay(200);
  }
  return false;
}

async function resolveAudioItemIdForDownload(item, state) {
  if (item.itemId) return item;

  if (!isBloombergAudioCatalogPage()) {
    window.location.href = (await loadConfig()).cmsCatalogUrl;
    await waitForLocation(isBloombergAudioCatalogPage, 12000);
  }

  pushLog(state, "info", `Opening row to resolve CMS id for ${item.fileName || item.title}.`, item);
  await writeState(state);

  const opened = await openCatalogRow(item);
  if (!opened) throw new Error(`Could not open catalog row for ${item.fileName || item.title}`);

  const reachedEditPage = await waitForLocation(isBloombergAudioEditPage, 12000);
  if (!reachedEditPage) throw new Error(`Could not open edit page for ${item.fileName || item.title}`);

  const itemId = extractItemId(window.location.href);
  if (!itemId) throw new Error(`Could not resolve CMS id for ${item.fileName || item.title}`);

  const config = await loadConfig();
  const resolved = {
    ...item,
    itemId,
    audioCacheKey: `${itemId}:${config.sourceLanguageCode || "en-US"}`,
    editUrl: window.location.href,
  };

  state.items = (state.items || []).map((candidate) =>
    candidate.rowKey && item.rowKey && candidate.rowKey === item.rowKey ? { ...candidate, itemId, editUrl: resolved.editUrl } : candidate
  );
  state.workItems = (state.workItems || []).map((candidate) =>
    candidate.rowKey && item.rowKey && candidate.rowKey === item.rowKey
      ? { ...candidate, itemId, audioCacheKey: resolved.audioCacheKey, editUrl: resolved.editUrl }
      : candidate
  );
  await writeState(state);
  return resolved;
}

async function downloadMissingAudioToFolder() {
  if (!isBloombergAudioCatalogPage()) {
    throw new Error("Open the Bloomberg Connects Audios catalog before downloading audio.");
  }

  const folderHandle = await getAudioFolderHandle(true, "readwrite");
  if (!folderHandle) {
    throw new Error("Choose a local audio folder before downloading missing audio.");
  }

  const confirmed = window.confirm(
    "This will download missing CMS source audio into the selected folder and rename matching local audio files to clean stop-title filenames. Continue?"
  );
  if (!confirmed) return null;
  await storageRemove(BCA_AUDIO_DOWNLOAD_KEY);

  const scannedState = await scanCatalog();
  const state = await readState();
  const items = uniqueAudioItems(scannedState.items);

  pushLog(state, "info", `Checking ${items.length} CMS audio stops against the selected local folder.`);
  state.phase = "audio-download";
  await writeState(state);
  await storageSet({
    [BCA_AUDIO_DOWNLOAD_KEY]: {
      status: "running",
      index: 0,
      items,
      counts: { downloaded: 0, renamed: 0, skipped: 0, failed: 0 },
      updatedAt: new Date().toISOString(),
    },
  });

  setTimeout(() => continueAudioDownloadRun().catch((error) => failAudioDownloadRun(error)), 250);
  return state;
}

async function readAudioDownloadRun() {
  const data = await storageGet(BCA_AUDIO_DOWNLOAD_KEY);
  return data[BCA_AUDIO_DOWNLOAD_KEY] || null;
}

async function writeAudioDownloadRun(run) {
  run.updatedAt = new Date().toISOString();
  await storageSet({ [BCA_AUDIO_DOWNLOAD_KEY]: run });
  return run;
}

async function failAudioDownloadRun(error, item = null) {
  const run = await readAudioDownloadRun();
  if (run) {
    run.status = "paused";
    run.error = error.message;
    run.counts = run.counts || {};
    run.counts.failed = (run.counts.failed || 0) + 1;
    await writeAudioDownloadRun(run);
  }
  const state = await readState();
  state.phase = "audio-download-paused";
  pushLog(state, "error", `Audio download paused: ${error.message}`, item, "missingAudio");
  await writeState(state);
  alert(`Bloomberg Audio Assistant (${BCA_BUILD_LABEL}) paused: ${error.message}`);
}

async function finishAudioDownloadRun(run) {
  const counts = run.counts || {};
  const state = await readState();
  state.phase = "dry-run";
  pushLog(
    state,
    "success",
    `Audio folder ready: ${counts.downloaded || 0} downloaded, ${counts.renamed || 0} renamed, ${counts.skipped || 0} already present, ${counts.failed || 0} failed.`
  );
  await writeState(state);
  await storageRemove(BCA_AUDIO_DOWNLOAD_KEY);
}

async function markAudioDownloadItemDone(run, status, continueNow = true) {
  run.counts = run.counts || {};
  run.counts[status] = (run.counts[status] || 0) + 1;
  run.index += 1;
  await writeAudioDownloadRun(run);
  if (continueNow) setTimeout(() => continueAudioDownloadRun().catch((error) => failAudioDownloadRun(error)), 250);
}

async function continueAudioDownloadRun() {
  if (bcaAudioDownloadInFlight) return;
  bcaAudioDownloadInFlight = true;
  try {
  const run = await readAudioDownloadRun();
  if (!run || run.status !== "running") return;
  const items = run.items || [];
  if (run.index >= items.length) {
    await finishAudioDownloadRun(run);
    return;
  }

  const item = items[run.index];
  const config = await loadConfig();
  const folderHandle = await getAudioFolderHandle(true, "readwrite");
  if (!folderHandle) throw new Error("Choose a local audio folder before downloading missing audio.");

  const state = await readState();
  state.phase = "audio-download";
  state.current = { itemId: item.itemId || "", label: itemLabel(item) };
  await writeState(state);

  const existing = await findLocalAudioEntryForItem(item);
  if (existing) {
    const renamed = await normalizeLocalAudioEntry(existing, item, state);
    await writeState(state);
    await markAudioDownloadItemDone(run, renamed ? "renamed" : "skipped");
    return;
  }

  if (!item.itemId && !isBloombergAudioEditPage()) {
    if (!isBloombergAudioCatalogPage()) {
      window.location.href = config.cmsCatalogUrl;
      return;
    }
    pushLog(state, "info", `Opening row to resolve CMS id for ${item.fileName || item.title}.`, item);
    await writeState(state);
    const opened = await openCatalogRow(item);
    if (!opened) throw new Error(`Could not open catalog row for ${item.fileName || item.title}`);
    setTimeout(() => continueAudioDownloadRun().catch((error) => failAudioDownloadRun(error, item)), 900);
    return;
  }

  const itemId = item.itemId || extractItemId(window.location.href);
  if (!itemId) throw new Error(`Could not resolve CMS id for ${item.fileName || item.title}`);
  const resolvedItem = {
    ...item,
    itemId,
    audioCacheKey: `${itemId}:${config.sourceLanguageCode || "en-US"}`,
    editUrl: isBloombergAudioEditPage() ? window.location.href : item.editUrl,
  };
  if (!isBloombergAudioEditPage()) throw new Error(`Could not open edit page for ${item.fileName || item.title}`);

  pushLog(state, "info", `Fetching CMS source audio directly for ${resolvedItem.fileName || resolvedItem.title}.`, resolvedItem);
  await writeState(state);
  const response = await fetchAudioResponseFromUrl(buildDownloadUrl(resolvedItem, config), `CMS source audio for ${resolvedItem.fileName || resolvedItem.title || resolvedItem.itemId}`);
  const finalName = cleanLocalAudioFileName(resolvedItem, { type: response.headers.get("content-type") || "" });
  await writeResponseToAudioFolder(folderHandle, finalName, response);
  pushLog(state, "success", `Downloaded ${finalName} into selected folder.`, resolvedItem);
  await writeState(state);
  run.items[run.index] = resolvedItem;
  await markAudioDownloadItemDone(run, "downloaded", isBloombergAudioCatalogPage());

  if (!isBloombergAudioCatalogPage()) {
    window.location.href = config.cmsCatalogUrl;
  }
  } finally {
    bcaAudioDownloadInFlight = false;
  }
}

function buildDownloadUrl(item, config) {
  if (!item.itemId) {
    throw statusError("missingAudio", "Missing CMS audio item id; open the CMS row again and resume.");
  }
  const sourceLanguageCode = encodeURIComponent(config.sourceLanguageCode || "en-US");
  return `https://cms.bloombergconnects.org/api/download/audio/${encodeURIComponent(item.itemId)}/${sourceLanguageCode}/sourceUrl`;
}

function cleanAudioFileName(item, blob) {
  const extension = audioExtensionFor(blob, item.fileName);
  const stop = String(item.stopNumber).padStart(3, "0");
  return `${stop} ${cleanStopTitle(item.title || item.fileName, item.stopNumber)}.${extension}`;
}

async function fetchAudioResponseFromUrl(url, context = "CMS audio") {
  if (!url) throw new Error(`Missing URL for ${context}.`);

  let response;
  try {
    response = await fetch(url, { credentials: "include", redirect: "follow" });
  } catch (error) {
    throw new Error(`${context} could not be fetched: ${error.message}. url=${url.slice(0, 180)}`);
  }

  if (!response.ok) {
    throw new Error(`${context} returned HTTP ${response.status}. url=${url.slice(0, 180)}`);
  }

  const contentType = response.headers.get("content-type") || "";
  if (/application\/json|text\//i.test(contentType)) {
    const text = await response.text();
    const nestedUrl = extractAudioUrlFromTextResponse(text);
    if (nestedUrl) return fetchAudioResponseFromUrl(nestedUrl, `${context} redirected source URL`);
    throw new Error(`${context} returned ${contentType || "text"} instead of audio. Preview=${normalizeSpaces(text).slice(0, 220)}`);
  }

  const contentLength = Number(response.headers.get("content-length") || 0);
  if (contentLength > 0 && contentLength < 1024) {
    throw new Error(`${context} returned a file that is too small to be usable (${contentLength} bytes). url=${url.slice(0, 180)}`);
  }
  return response;
}

async function fetchAudioBlobFromUrl(url, context = "CMS audio") {
  const response = await fetchAudioResponseFromUrl(url, context);

  const blob = await response.blob();
  if (!blob || !blob.size) {
    throw new Error(`${context} returned an empty file. url=${url.slice(0, 180)}`);
  }
  if (blob.size < 1024) {
    throw new Error(`${context} returned a file that is too small to be usable (${blob.size} bytes). url=${url.slice(0, 180)}`);
  }
  return blob;
}

function extractAudioUrlFromTextResponse(text) {
  const trimmed = String(text || "").trim();
  if (/^https?:\/\//i.test(trimmed)) return trimmed;
  try {
    const parsed = JSON.parse(trimmed);
    const candidates = [
      parsed.sourceUrl,
      parsed.sourceURL,
      parsed.downloadUrl,
      parsed.downloadURL,
      parsed.url,
      parsed.data && parsed.data.sourceUrl,
      parsed.data && parsed.data.downloadUrl,
      parsed.data && parsed.data.url,
    ].filter(Boolean);
    return candidates.find((candidate) => /^https?:\/\//i.test(String(candidate))) || "";
  } catch {
    const match = trimmed.match(/https?:\/\/[^\s"'<>]+/i);
    return match ? match[0] : "";
  }
}

async function fetchCmsAudioBlob(item, config) {
  return fetchAudioBlobFromUrl(buildDownloadUrl(item, config), `CMS source audio for ${item.fileName || item.title || item.itemId}`);
}

async function ensureCachedAudio(item, state) {
  const config = await loadConfig();
  if (!item.audioCacheKey && item.itemId) {
    item.audioCacheKey = `${item.itemId}:${config.sourceLanguageCode || "en-US"}`;
  }
  if (!item.audioCacheKey) {
    throw statusError("missingAudio", "Missing CMS audio cache key; open the CMS row again and resume.");
  }

  const localRecord = await buildLocalAudioRecord(item);
  if (localRecord && localRecord.blob && localRecord.blob.size > 0) {
    pushLog(state, "success", `Using local audio file from selected folder: ${localRecord.fileName}.`, item);
    await writeState(state);
    return localRecord;
  }

  throw statusError(
    "missingAudio",
    "No matching local audio file found. Choose the local audio folder, run Download Missing Audio, then resume. Save mode no longer downloads audio into the browser cache."
  );
}

function statusError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function visibleElements(selector, root = document) {
  return Array.from(root.querySelectorAll(selector)).filter((element) => {
    if (!(element instanceof Element)) return false;
    const rect = element.getBoundingClientRect();
    const style = window.getComputedStyle(element);
    return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
  });
}

function elementArea(element) {
  const rect = element.getBoundingClientRect();
  return rect.width * rect.height;
}

function directText(element) {
  return normalizeSpaces(
    Array.from(element.childNodes)
      .filter((node) => node.nodeType === Node.TEXT_NODE)
      .map((node) => node.textContent || "")
      .join(" ")
  );
}

function findPanelFromHeading(headingText) {
  const target = normalizeComparable(headingText);
  const heading = visibleElements("h1, h2, h3, h4, div, span", document).find((element) => {
    const own = normalizeComparable(directText(element));
    const all = normalizeComparable(controlText(element));
    return own === target || all === target;
  });
  if (!heading) return null;

  const viewportArea = Math.max(window.innerWidth * window.innerHeight, 1);
  const candidates = [];
  let ancestor = heading.parentElement;
  while (ancestor && ancestor !== document.body && ancestor !== document.documentElement) {
    const text = normalizeComparable(ancestor.textContent);
    const hasHeading = text.includes(normalizeComparable(headingText));
    const hasAction = text.includes("cancel") || text.includes("add") || text.includes("save");
    const area = elementArea(ancestor);
    const inputCount = ancestor.querySelectorAll("input, textarea, [role='combobox'], [role='textbox'], select").length;
    const largeEnough = area > viewportArea * 0.08;
    const notFullscreen = area < viewportArea * 0.98;
    if (hasHeading && hasAction && largeEnough && notFullscreen && inputCount > 0) {
      candidates.push(ancestor);
    }
    ancestor = ancestor.parentElement;
  }

  return candidates.sort((a, b) => elementArea(a) - elementArea(b))[0] || null;
}

function findStructuredPanel(kind) {
  const target = normalizeComparable(kind);
  const viewportArea = Math.max(window.innerWidth * window.innerHeight, 1);
  const candidates = visibleElements('[role="dialog"], [class*="modal"], [class*="drawer"], [class*="Dialog"], form, section, article, div', document)
    .filter((element) => element !== document.body && element !== document.documentElement)
    .map((element) => {
      const text = normalizeComparable(element.textContent);
      const area = elementArea(element);
      return { element, text, area };
    })
    .filter(({ text, area }) => {
      if (!text.includes(target)) return false;
      if (area < viewportArea * 0.04 || area > viewportArea * 0.96) return false;
      if (target.includes("edit audio")) {
        return text.includes("cancel") && text.includes("save") && text.includes("audio") && text.includes("title") && text.includes("transcript");
      }
      if (target.includes("add translation")) {
        return text.includes("choose a language") && text.includes("cancel") && text.includes("add");
      }
      return true;
    })
    .sort((a, b) => a.area - b.area);

  return candidates[0] ? candidates[0].element : null;
}

function findEditAudioPanelFromHeading() {
  return findPanelFromHeading("Edit Audio");
}

function findAddTranslationPanelFromHeading() {
  return findPanelFromHeading("Add Translation");
}

function findEditAudioRoot() {
  return findStructuredPanel("Edit Audio") || findEditAudioPanelFromHeading() || findModalRoot();
}

function findModalRoot() {
  const structuredAddTranslation = findStructuredPanel("Add Translation");
  if (structuredAddTranslation) return structuredAddTranslation;
  const structuredEditAudio = findStructuredPanel("Edit Audio");
  if (structuredEditAudio) return structuredEditAudio;

  const candidates = visibleElements('[role="dialog"], [class*="modal"], [class*="drawer"], [class*="Dialog"]');
  const addTranslationCandidate = candidates.find((element) => normalizeComparable(element.textContent).includes("add translation"));
  if (addTranslationCandidate) return addTranslationCandidate;
  const editAudioCandidate = candidates.find((element) => normalizeComparable(element.textContent).includes("edit audio"));
  if (editAudioCandidate) return editAudioCandidate;

  const addTranslationPanel = findAddTranslationPanelFromHeading();
  if (addTranslationPanel) return addTranslationPanel;

  const headingPanel = findEditAudioPanelFromHeading();
  if (headingPanel) return headingPanel;

  return candidates.sort((a, b) => elementArea(a) - elementArea(b))[0] || document.body;
}

function findButtonByText(text, root = document) {
  const target = normalizeComparable(text);
  return visibleElements("button, [role='button'], a", root).find((button) => normalizeComparable(button.textContent || button.getAttribute("aria-label")).includes(target));
}

function findExactButtonByText(text, root = document) {
  const target = normalizeComparable(text);
  return visibleElements("button, [role='button'], a", root).find((button) => normalizeComparable(button.textContent || button.getAttribute("aria-label")) === target);
}

function isDisabledAction(element) {
  if (!element) return true;
  const disabledAncestor = element.closest("[disabled], [aria-disabled='true']");
  if (disabledAncestor) return true;
  if (element.disabled === true) return true;
  if (element.getAttribute("aria-disabled") === "true") return true;
  const className = String(element.className || "");
  if (/\bdisabled\b/i.test(className)) return true;
  const style = window.getComputedStyle(element);
  return style.pointerEvents === "none" || Number(style.opacity) < 0.2;
}

function findClickableByText(text, root = document) {
  const target = normalizeComparable(text);
  return visibleElements("button, [role='button'], a, [role='option'], [role='menuitem'], li", root)
    .filter((element) => {
      const comparable = normalizeComparable(controlText(element));
      if (!(comparable === target || comparable.includes(target))) return false;
      return controlText(element).length <= Math.max(text.length + 36, 96);
    })
    .sort((a, b) => controlText(a).length - controlText(b).length)[0];
}

function distanceBetweenElements(a, b) {
  if (!(a instanceof Element) || !(b instanceof Element)) return Number.POSITIVE_INFINITY;
  const aRect = a.getBoundingClientRect();
  const bRect = b.getBoundingClientRect();
  const ax = aRect.left + aRect.width / 2;
  const ay = aRect.top + aRect.height / 2;
  const bx = bRect.left + bRect.width / 2;
  const by = bRect.top + bRect.height / 2;
  return Math.hypot(ax - bx, ay - by);
}

function findAddTranslationActionInRoot(root, referenceElement = null) {
  const target = normalizeComparable("Add Translation");
  const candidates = visibleElements("button, [role='button'], a, [role='option'], [role='menuitem'], li", root)
    .filter((element) => {
      const text = normalizeComparable(controlText(element));
      return text === target || text.includes(target);
    })
    .sort((a, b) => {
      if (referenceElement) {
        const distanceDelta = distanceBetweenElements(a, referenceElement) - distanceBetweenElements(b, referenceElement);
        if (distanceDelta !== 0) return distanceDelta;
      }
      return controlText(a).length - controlText(b).length;
    });
  return candidates[0] || null;
}

function findAddTranslationAction(root = document, referenceElement = null) {
  const local = findAddTranslationActionInRoot(root, referenceElement);
  if (local) return local;
  if (root !== document) {
    return findAddTranslationActionInRoot(document, referenceElement);
  }
  return null;
}

function formatElementForDebug(element, referenceElement = null) {
  if (!(element instanceof Element)) return "none";
  const text = controlText(element).slice(0, 80);
  const role = element.getAttribute("role") || "";
  const id = element.id ? `#${element.id}` : "";
  const className = (element.className && typeof element.className === "string" ? `.${element.className.split(/\s+/).slice(0, 2).join(".")}` : "").trim();
  const distance = referenceElement ? Math.round(distanceBetweenElements(element, referenceElement)) : null;
  return `${element.tagName.toLowerCase()}${id}${className ? className : ""}${role ? `[role=${role}]` : ""} text="${text}"${
    Number.isFinite(distance) ? ` dist=${distance}` : ""
  }`;
}

function buildPickerStateDebugInfo() {
  const activeElement = document.activeElement instanceof Element ? document.activeElement : null;
  const active = activeElement ? formatElementForDebug(activeElement) : "none";
  const activeValue = activeElement && "value" in activeElement ? String(activeElement.value || "") : "";
  const ariaExpanded = activeElement ? String(activeElement.getAttribute("aria-expanded") || "") : "";
  const ariaControls = activeElement ? String(activeElement.getAttribute("aria-controls") || "") : "";
  const ariaActiveDescendant = activeElement ? String(activeElement.getAttribute("aria-activedescendant") || "") : "";
  const optionCount = visibleElements("[role='option'], [role='listbox'], [class*='menu'], [class*='Menu']", document).length;
  const modal = findAddTranslationPanelFromHeading() || findModalRoot();
  const addButton = findExactButtonByText("Add", modal) || findButtonByText("Add", modal);
  const addEnabled = addButton ? !addButton.disabled : false;
  const fiberDebug = buildReactSelectFiberDebugInfo(modal);
  return `active=${active}; activeValue="${activeValue.slice(0, 80)}"; ariaExpanded="${ariaExpanded}"; ariaControls="${ariaControls}"; ariaActivedescendant="${ariaActiveDescendant}"; pickerNodes=${optionCount}; addEnabled=${addEnabled}; ${fiberDebug}; mainWorld="${lastMainWorldLanguageSelectionDebug.slice(0, 240)}"`;
}

function buildReactSelectFiberDebugInfo(dialog) {
  const components = collectReactSelectComponents(dialog);
  for (const component of components) {
    const options = flattenReactSelectOptions(component.props.options);
    const sample = options
      .map(optionLabel)
      .filter(Boolean)
      .slice(0, 8)
      .join(" | ");
    return `reactSelect=${component.source}; optionCount=${options.length}; optionSample="${sample}"`;
  }
  return "reactSelect=none";
}

function collectTextCandidatesForDebug(text, root, selector, limit = 4) {
  const target = normalizeComparable(text);
  return visibleElements(selector, root)
    .filter((element) => normalizeComparable(controlText(element)).includes(target))
    .slice(0, limit);
}

function buildLanguageSelectionDebugInfo(referenceElement, modalRoot) {
  const modal = modalRoot || findModalRoot();
  const englishSelector =
    "select, button, input:not([type='hidden']), [role='combobox'], [aria-haspopup='listbox'], [aria-haspopup='menu'], [class*='select'], [class*='Select'], [class*='control'], [class*='Control']";
  const actionSelector = "button, [role='button'], a, [role='option'], [role='menuitem'], li";
  const englishModal = collectTextCandidatesForDebug("english", modal, englishSelector);
  const englishDocument = collectTextCandidatesForDebug("english", document, englishSelector);
  const addModal = collectTextCandidatesForDebug("add translation", modal, actionSelector);
  const addDocument = collectTextCandidatesForDebug("add translation", document, actionSelector);
  const lines = [
    `modal=${modal === document.body ? "document.body" : formatElementForDebug(modal)}`,
    `english(modal)=${englishModal.length ? englishModal.map((el) => formatElementForDebug(el, referenceElement)).join(" | ") : "none"}`,
    `english(document)=${englishDocument.length ? englishDocument.map((el) => formatElementForDebug(el, referenceElement)).join(" | ") : "none"}`,
    `addTranslation(modal)=${addModal.length ? addModal.map((el) => formatElementForDebug(el, referenceElement)).join(" | ") : "none"}`,
    `addTranslation(document)=${addDocument.length ? addDocument.map((el) => formatElementForDebug(el, referenceElement)).join(" | ") : "none"}`,
  ];
  return lines.join("; ");
}

function catalogItemSearchTerms(item) {
  const terms = [];
  const add = (value) => {
    const text = normalizeSpaces(value);
    if (!text) return;
    if (!terms.some((term) => normalizeComparable(term) === normalizeComparable(text))) terms.push(text);
  };

  add(item.fileName);
  add(spaceLetterNumberRuns(item.fileName || ""));
  if (item.stopNumber) {
    const stop = String(item.stopNumber);
    const padded = stop.padStart(3, "0");
    add(padded);
    add(stop);
    const fileTitle = cleanStopTitle(item.fileName || "", item.stopNumber);
    add(`${padded} ${fileTitle}`);
    add(`${stop} ${fileTitle}`);
    add(fileTitle);
    add(spaceLetterNumberRuns(fileTitle));
  }
  add(item.includedIn);
  add(item.title);
  add(item.rawTitle);
  add(item.sourceTitle);
  return terms.filter(Boolean);
}

function findCatalogRowForItem(item) {
  const rows = getCatalogRows();
  const targetStop = item.stopNumber;
  const targetTerms = catalogItemSearchTerms(item);

  const exact = rows.find(({ row, cells, headerMap }) => {
    const title = meaningfulText(cells[headerMap.title]);
    const fileName = meaningfulText(cells[headerMap.fileName]);
    const includedIn = meaningfulText(cells[headerMap.includedIn]);
    const rowStop = parseStopNumber(fileName, title);
    const rowText = `${title} ${fileName} ${includedIn} ${row.textContent || ""}`;
    const stopCompatible = !targetStop || !rowStop || rowStop === targetStop;
    if (!stopCompatible) return false;
    return targetTerms.some((term) => {
      if (/^\d{1,3}$/.test(term)) return rowStop === Number(term);
      return comparableTextMatches(fileName, term) || comparableTextMatches(includedIn, term) || comparableTextMatches(title, term) || comparableTextMatches(rowText, term);
    });
  });
  if (exact) return exact;

  if (targetStop) {
    const byStop = rows.find(
      ({ cells, headerMap }) =>
        parseStopNumber(meaningfulText(cells[headerMap.fileName]), meaningfulText(cells[headerMap.title])) === targetStop
    );
    if (byStop) return byStop;
  }

  return null;
}

function findCatalogSearchInput() {
  const inputs = visibleElements("input, [role='searchbox'], [role='combobox']").filter((input) => {
    const text = normalizeComparable(
      `${input.getAttribute("placeholder") || ""} ${input.getAttribute("aria-label") || ""} ${input.name || ""} ${input.id || ""}`
    );
    return text.includes("search") || text.includes("audio");
  });
  return inputs[0] || null;
}

async function searchCatalogForItem(item) {
  const input = findCatalogSearchInput();
  if (!input) return false;
  const queries = catalogItemSearchTerms(item).slice(0, 8);
  if (!queries.length) return false;

  for (const query of queries) {
    input.focus?.();
    setReactControlledInputValue(input, "");
    input.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward" }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    await delay(120);
    setReactControlledInputValue(input, query);
    input.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: query }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    await delay(700);
    if (findCatalogRowForItem(item)) return true;
  }

  return false;
}

async function clearCatalogSearch() {
  const input = findCatalogSearchInput();
  if (!input || !input.value) return;
  input.focus?.();
  setReactControlledInputValue(input, "");
  input.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "deleteContentBackward" }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
  await delay(600);
}

async function scrollCatalogToItem(item) {
  const container = findCatalogScrollContainer();
  const maxPasses = 90;
  let stagnantPasses = 0;
  let lastScrollTop = -1;

  container.scrollTop = 0;
  await delay(250);
  for (let pass = 0; pass < maxPasses; pass += 1) {
    const match = findCatalogRowForItem(item);
    if (match) return match;

    const maxScrollTop = Math.max(0, container.scrollHeight - container.clientHeight);
    if (container.scrollTop >= maxScrollTop - 4) break;

    const nextScrollTop = Math.min(maxScrollTop, container.scrollTop + Math.max(280, Math.floor(container.clientHeight * 0.85)));
    if (nextScrollTop === lastScrollTop) stagnantPasses += 1;
    else stagnantPasses = 0;
    if (stagnantPasses >= 8) break;

    lastScrollTop = container.scrollTop;
    container.scrollTop = nextScrollTop;
    await delay(220);
  }

  return findCatalogRowForItem(item);
}

async function openCatalogRow(item) {
  await ensureCmsWindowFocused();
  let match = findCatalogRowForItem(item);
  if (!match) {
    await searchCatalogForItem(item);
    match = findCatalogRowForItem(item);
  }
  if (!match) {
    await clearCatalogSearch();
    match = await scrollCatalogToItem(item);
  }
  if (!match) return false;

  const { row, cells, headerMap } = match;
  const editUrl = findEditUrl(row, cells, headerMap);
  if (editUrl) {
    window.location.href = editUrl;
    return true;
  }

  const titleCell = cells[headerMap.title] || row;
  const fileCell = cells[headerMap.fileName] || row;
  const target =
    titleCell.querySelector("a, button, [role='button']") ||
    fileCell.querySelector("a, button, [role='button']") ||
    titleCell ||
    fileCell ||
    row;
  target.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
  await delay(800);
  return true;
}

function findFieldByLabel(labelText, root = document, { preferMultiline = false, preferSingleLine = false } = {}) {
  const singleLineSelector =
    "input:not([type='hidden']):not([type='file']):not([type='checkbox']):not([type='radio']), textarea";
  const multilineSelector = "textarea, [contenteditable='true'], [role='textbox']:not(input)";
  const defaultSelector = `${singleLineSelector}, ${multilineSelector}`;
  const primarySelector = preferMultiline ? multilineSelector : preferSingleLine ? singleLineSelector : defaultSelector;
  const fallbackSelector = preferMultiline ? singleLineSelector : preferSingleLine ? multilineSelector : "";
  const selectFieldFromContainer = (container) => {
    if (!container || !(container instanceof Element)) return null;
    const primary = container.querySelector(primarySelector);
    if (primary) return primary;
    if (fallbackSelector) return container.querySelector(fallbackSelector);
    return null;
  };
  const selectNearestFieldBelow = (anchor, scope) => {
    if (!(anchor instanceof Element)) return null;
    const anchorRect = anchor.getBoundingClientRect();
    const candidates = visibleElements(primarySelector, scope || root)
      .filter((candidate) => {
        if (!(candidate instanceof Element)) return false;
        if (candidate.matches("input[type='file']")) return false;
        const rect = candidate.getBoundingClientRect();
        return rect.top >= anchorRect.bottom - 2;
      })
      .sort((a, b) => {
        const aRect = a.getBoundingClientRect();
        const bRect = b.getBoundingClientRect();
        const aDeltaY = Math.abs(aRect.top - anchorRect.bottom);
        const bDeltaY = Math.abs(bRect.top - anchorRect.bottom);
        if (aDeltaY !== bDeltaY) return aDeltaY - bDeltaY;
        return Math.abs(aRect.left - anchorRect.left) - Math.abs(bRect.left - anchorRect.left);
      });
    return candidates[0] || null;
  };

  const target = normalizeComparable(labelText);
  const labels = visibleElements("label", root);
  for (const label of labels) {
    if (!normalizeComparable(label.textContent).includes(target)) continue;
    const forId = label.getAttribute("for");
    if (forId) {
      const direct = document.getElementById(forId);
      if (direct && !direct.matches("input[type='file']")) return direct;
    }
    const container =
      label.closest("[class*='field'], [class*='Field'], [class*='form'], [class*='Form'], div") || label.parentElement;
    const nearestField = selectNearestFieldBelow(label, container || root);
    if (nearestField) return nearestField;
    const field = selectFieldFromContainer(container);
    if (field) return field;
  }

  // CMS sometimes renders "Transcript"/"Description" headers as plain text rather than <label>.
  const hintNodes = visibleElements("div, span, p, h1, h2, h3, h4, h5, h6", root).filter((node) => {
    const text = normalizeComparable(node.textContent);
    return text === target || text.includes(target);
  });
  for (const hint of hintNodes) {
    // Prefer the nearest matching field below the heading text (important for
    // forms that have Description + Transcript in the same broad container).
    const nearestField = selectNearestFieldBelow(hint, root);
    if (nearestField) return nearestField;

    const candidates = [
      hint.nextElementSibling,
      hint.parentElement,
      hint.closest("[class*='field'], [class*='Field'], [class*='form'], [class*='Form'], section, form, div"),
    ].filter(Boolean);
    for (const container of candidates) {
      const field = selectFieldFromContainer(container);
      if (field) return field;
    }
  }
  return null;
}

function setFieldValue(field, value) {
  if (!field) return false;
  if (field.matches && field.matches("input[type='file']")) return false;
  field.focus();
  if (field.matches("[contenteditable='true'], [role='textbox']")) {
    field.textContent = value;
    field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
    field.dispatchEvent(new Event("change", { bubbles: true }));
    field.dispatchEvent(new Event("blur", { bubbles: true }));
    return true;
  }

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
  return true;
}

function setInputValueWithoutBlur(field, value, { dispatchChange = true } = {}) {
  if (!field || !field.matches("input:not([type='hidden']), textarea, [role='textbox'], [role='combobox']")) {
    return false;
  }
  field.focus?.();
  const prototype = Object.getPrototypeOf(field);
  const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
  if (descriptor && descriptor.set) {
    descriptor.set.call(field, value);
  } else {
    field.value = value;
  }
  field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
  if (dispatchChange) {
    field.dispatchEvent(new Event("change", { bubbles: true }));
  }
  return true;
}

function findTranscriptField(root = findModalRoot()) {
  const target = normalizeComparable("Transcript");
  const ariaSelector =
    "[contenteditable='true'][aria-label='Transcript'], [role='textbox'][aria-label='Transcript'], .tiptap.ProseMirror[aria-label='Transcript']";
  const editorSelector =
    "textarea, [contenteditable='true'], [data-contents='true'], [data-lexical-editor='true'], .public-DraftEditor-content, .ql-editor, .tiptap.ProseMirror, [role='textbox']:not(input)";

  // Primary: Bloomberg's TipTap/ProseMirror editor sets aria-label="Transcript" on
  // the contenteditable node itself. Match that directly first.
  const ariaInRoot = visibleElements(ariaSelector, root);
  if (ariaInRoot.length) return ariaInRoot[0];
  if (root !== document) {
    const ariaInDoc = visibleElements(ariaSelector, document);
    if (ariaInDoc.length) return ariaInDoc[0];
  }

  const headingNodes = visibleElements("label, div, span, p, h1, h2, h3, h4, h5, h6", root)
    .filter((node) => {
      const text = normalizeComparable(node.textContent);
      if (!text) return false;
      // Avoid giant container nodes that include the entire form text.
      if (text.length > 80) return false;
      return text === target || text.startsWith(target);
    })
    .sort((a, b) => {
      const aRect = a.getBoundingClientRect();
      const bRect = b.getBoundingClientRect();
      return aRect.top - bRect.top;
    });

  const normalizeEditableTarget = (candidate) => {
    if (!(candidate instanceof Element)) return null;
    if (candidate.matches("textarea, [contenteditable='true'], [data-lexical-editor='true'], .public-DraftEditor-content, .ql-editor, .tiptap.ProseMirror")) return candidate;
    const editableAncestor = candidate.closest("[contenteditable='true'], [data-lexical-editor='true'], .public-DraftEditor-content, .ql-editor, .tiptap.ProseMirror");
    if (editableAncestor instanceof Element) return editableAncestor;
    const nested = candidate.querySelector(editorSelector);
    return nested instanceof Element ? nested : candidate;
  };

  const pickNearestEditorBelow = (anchor) => {
    if (!(anchor instanceof Element)) return null;
    const anchorRect = anchor.getBoundingClientRect();
    const scope =
      anchor.closest("section, form, [class*='field'], [class*='Field'], [class*='panel'], [class*='Panel'], div") || root;
    const editors = visibleElements(editorSelector, root)
      .filter((editor) => {
        const rect = editor.getBoundingClientRect();
        return rect.top >= anchorRect.bottom - 2;
      })
      .sort((a, b) => {
        const aRect = a.getBoundingClientRect();
        const bRect = b.getBoundingClientRect();
        const aDelta = Math.abs(aRect.top - anchorRect.bottom);
        const bDelta = Math.abs(bRect.top - anchorRect.bottom);
        if (aDelta !== bDelta) return aDelta - bDelta;
        const areaDelta = bRect.width * bRect.height - aRect.width * aRect.height;
        if (areaDelta !== 0) return areaDelta;
        return Math.abs(aRect.left - anchorRect.left) - Math.abs(bRect.left - anchorRect.left);
      });
    return normalizeEditableTarget(editors[0]);
  };

  for (const heading of headingNodes) {
    const nearest = pickNearestEditorBelow(heading);
    if (nearest) return nearest;
  }

  // Fallback: last multiline editor in the language section is typically Transcript.
  const allEditors = visibleElements(editorSelector, root);
  if (allEditors.length) {
    const lastEditor = normalizeEditableTarget(allEditors[allEditors.length - 1]);
    if (lastEditor) return lastEditor;
  }

  // Last resort: look for any contenteditable near a "Transcript" text node anywhere in the document
  const allContentEditables = visibleElements("[contenteditable='true'], [role='textbox']:not(input)", document);
  for (const editable of allContentEditables) {
    // Check if this editable's parent container has "Transcript" text
    const parent = editable.closest("[class*='field'], [class*='Field'], [class*='form'], [class*='Form'], section, form, div");
    if (parent) {
      const parentText = normalizeComparable(parent.textContent);
      if (parentText.includes(target) && parentText.length < 200) {
        return editable;
      }
    }
    // Check previous sibling or parent's previous sibling for "Transcript" text
    const prevSibling = editable.previousElementSibling || (editable.parentElement && editable.parentElement.previousElementSibling);
    if (prevSibling) {
      const prevText = normalizeComparable(prevSibling.textContent);
      if (prevText.includes(target) && prevText.length < 80) {
        return editable;
      }
    }
  }

  return null;
}

function transcriptToPlainText(transcript) {
  return sanitizedTranscriptEntries(transcript)
    .map((entry) => {
      const paragraphs = Array.isArray(entry.paragraphs) && entry.paragraphs.length ? entry.paragraphs : [entry.text || ""].filter(Boolean);
      return [entry.speaker ? `${entry.speaker}:` : "", ...paragraphs].filter(Boolean).join("\n\n");
    })
    .join("\n\n")
    .trim();
}

function transcriptToHtml(transcript) {
  return sanitizedTranscriptEntries(transcript)
    .map((entry) => {
      const paragraphs = Array.isArray(entry.paragraphs) && entry.paragraphs.length ? entry.paragraphs : [entry.text || ""].filter(Boolean);
      const speaker = entry.speaker ? `<strong>${escapeHtml(entry.speaker)}:</strong>` : "";
      const renderedParagraphs = paragraphs.map((paragraph, index) => {
        const text = escapeHtml(paragraph);
        if (index === 0 && speaker) {
          return `<p>${speaker}<br>${text}</p>`;
        }
        return `<p>${text}</p>`;
      });
      return renderedParagraphs.length ? renderedParagraphs.join("") : `<p>${speaker}</p>`;
    })
    .join("");
}

function sanitizedTranscriptEntries(transcript) {
  const entries = transcript && Array.isArray(transcript.entries) ? transcript.entries : [];
  return entries
    .map((entry) => {
      if (isAudioGuidePromptText(entry.speaker || "")) return null;
      const sourceParagraphs = Array.isArray(entry.paragraphs) && entry.paragraphs.length ? entry.paragraphs : [entry.text || ""].filter(Boolean);
      const paragraphs = sourceParagraphs.map(cleanTranscriptParagraphForCms).filter(Boolean);
      if (!paragraphs.length) return null;
      return {
        ...entry,
        paragraphs,
        text: paragraphs.join(" "),
      };
    })
    .filter(Boolean);
}

function cleanTranscriptParagraphForCms(paragraph) {
  const text = normalizeSpaces(paragraph);
  if (!text) return "";
  if (isAudioGuidePromptText(text)) return "";

  const sentences = splitSentences(text);
  if (sentences.length <= 1) return text;
  const kept = sentences.filter((sentence) => !isAudioGuidePromptText(sentence));
  return normalizeSpaces(kept.join(" "));
}

function splitSentences(text) {
  return (
    normalizeSpaces(text).match(/[^.!?]+(?:[.!?]+["”’']*|$)/g) || [text]
  )
    .map((part) => normalizeSpaces(part))
    .filter(Boolean);
}

function isAudioGuidePromptText(value) {
  const text = normalizeComparable(value);
  if (!text) return false;
  if (text.includes("acoustiguide")) return true;

  const hasAction =
    /\b(press|type|enter|dial|select|appuyez|tapez|composez|digite|digitem|tecle|pressione|presione|oprima|marque|pulse|premete|digitate|druecken|drucken|geben)\b/.test(text);
  const hasNumber = /\b\d{1,3}\b/.test(text);
  if (!hasAction || !hasNumber) return false;

  const hasPlayerReference =
    /\b(play|lecture|reproducao|reproduccion|reproduction|botao|bouton|pulsante|knopf|audiofuhrer|audio guide)\b/.test(text);
  const hasPromptLead =
    /^(to hear|if you|when you|para ouvir|para saber|quando|se voces|si desea|si quiere|pour|si vous|quando siete|se volete|wenn|um|gehen sie|unsere fuhrung|la visite|la visita|o tour|el recorrido)/.test(text);

  return hasPlayerReference || hasPromptLead || text.length <= 180;
}

function transcriptVerificationSnippets(text) {
  const normalized = normalizeComparable(text);
  if (!normalized) return [];
  if (normalized.length <= 120) return [normalized];
  // Return beginning, middle, and end snippets for more robust verification.
  const mid = Math.floor(normalized.length / 2);
  return [
    normalized.slice(0, 80),
    normalized.slice(Math.max(0, mid - 40), mid + 40),
    normalized.slice(Math.max(0, normalized.length - 80)),
  ];
}

function transcriptTextAppearsComplete(actualText, expectedText) {
  const actual = normalizeComparable(actualText);
  const expected = normalizeComparable(expectedText);
  if (!expected) return Boolean(actual);
  if (!actual) return false;
  if (actual.includes(expected)) return true;
  // Require at least 85% of the expected length.
  if (actual.length < Math.floor(expected.length * 0.85)) return false;
  // All snippets (beginning, middle, end) must be present.
  return transcriptVerificationSnippets(expected).every((snippet) => actual.includes(snippet));
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function setTranscriptField(root, transcript) {
  const field = findTranscriptField(root) || findFieldByLabel("Transcript", root, { preferMultiline: true });
  if (!field) return false;
  const html = transcriptToHtml(transcript);
  const text = transcriptToPlainText(transcript);

  // If field already has the correct content, skip modification
  const existingText = field.textContent || field.innerText || "";
  if (transcriptTextAppearsComplete(existingText, text)) {
    return true;
  }

  if (field.matches("[contenteditable='true'], [data-lexical-editor='true'], .public-DraftEditor-content, .ql-editor, .tiptap.ProseMirror, [role='textbox']:not(input)")) {
    field.focus();

    const selectAllContents = () => {
      const selection = window.getSelection?.();
      if (!selection || !document.createRange) return;
      const range = document.createRange();
      range.selectNodeContents(field);
      selection.removeAllRanges();
      selection.addRange(range);
    };

    const verify = () => {
      return transcriptTextAppearsComplete(field.textContent || field.innerText || "", text);
    };

    const clearContents = () => {
      selectAllContents();
      try {
        if (document.execCommand) document.execCommand("delete", false);
      } catch (_) {}
    };

    clearContents();

    try {
      const dt = new DataTransfer();
      dt.setData("text/html", html);
      dt.setData("text/plain", text);
      field.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
    } catch (_) {}
    if (verify()) {
      field.dispatchEvent(new Event("change", { bubbles: true }));
      field.dispatchEvent(new Event("blur", { bubbles: true }));
      return true;
    }

    clearContents();
    try {
      const dt2 = new DataTransfer();
      dt2.setData("text/html", html);
      dt2.setData("text/plain", text);
      field.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, inputType: "insertFromPaste", dataTransfer: dt2 }));
      field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertFromPaste", dataTransfer: dt2 }));
    } catch (_) {}
    if (verify()) {
      field.dispatchEvent(new Event("change", { bubbles: true }));
      field.dispatchEvent(new Event("blur", { bubbles: true }));
      return true;
    }

    clearContents();
    try {
      if (document.execCommand) document.execCommand("insertHTML", false, html);
    } catch (_) {}
    if (verify()) {
      field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertFromPaste", data: text }));
      field.dispatchEvent(new Event("change", { bubbles: true }));
      field.dispatchEvent(new Event("blur", { bubbles: true }));
      return true;
    }

    clearContents();
    field.innerHTML = html;
    field.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, inputType: "insertFromPaste", data: text }));
    field.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertFromPaste", data: text }));
    field.dispatchEvent(new Event("change", { bubbles: true }));
    field.dispatchEvent(new Event("blur", { bubbles: true }));
    return verify();
  }

  return setFieldValue(field, text);
}

async function clickLanguageDropdown(root) {
  await ensureCmsWindowFocused();
  const probeControl = (scope) =>
    findLanguageDropdownControl(scope) ||
    findEnglishLanguageControl(scope) ||
    findLanguageDropdownControl(document) ||
    findEnglishLanguageControl(document);

  let control = probeControl(root);
  if (!control) {
    const deadline = Date.now() + 8000;
    while (!control && Date.now() < deadline) {
      await delay(250);
      control = probeControl(root);
    }
  }

  if (!control) {
    const debugInfo = buildLanguageSelectionDebugInfo(null, root || findModalRoot());
    console.warn("Bloomberg Audio Assistant language selector debug:", debugInfo);
    throw statusError("languageSelectionFailed", `Could not find the language selector. Debug: ${debugInfo}`);
  }

  if (control.tagName === "SELECT") {
    return control;
  }

  if (findAddTranslationAction(document)) {
    return control;
  }

  await openLanguageDropdownControl(control);
  return control;
}

function controlText(element) {
  return normalizeSpaces(element.value || element.textContent || element.getAttribute("aria-label") || element.getAttribute("placeholder") || "");
}

function dispatchUserClick(element) {
  element.scrollIntoView?.({ block: "center", inline: "center" });
  element.focus?.();
  const rect = element.getBoundingClientRect();
  const eventInit = {
    bubbles: true,
    cancelable: true,
    view: window,
    clientX: rect.left + Math.min(Math.max(rect.width / 2, 8), Math.max(rect.width - 8, 8)),
    clientY: rect.top + Math.min(Math.max(rect.height / 2, 8), Math.max(rect.height - 8, 8)),
  };
  if (window.PointerEvent) {
    element.dispatchEvent(new PointerEvent("pointerdown", { ...eventInit, pointerId: 1, pointerType: "mouse", isPrimary: true }));
    element.dispatchEvent(new PointerEvent("pointerup", { ...eventInit, pointerId: 1, pointerType: "mouse", isPrimary: true }));
  }
  element.dispatchEvent(new MouseEvent("mousedown", eventInit));
  element.dispatchEvent(new MouseEvent("mouseup", eventInit));
  element.dispatchEvent(new MouseEvent("click", eventInit));
}

function dispatchPointClick(x, y) {
  const element = document.elementFromPoint(x, y);
  if (!element) return false;
  const eventInit = {
    bubbles: true,
    cancelable: true,
    view: window,
    clientX: x,
    clientY: y,
  };
  if (window.PointerEvent) {
    element.dispatchEvent(new PointerEvent("pointerdown", { ...eventInit, pointerId: 1, pointerType: "mouse", isPrimary: true }));
  }
  element.dispatchEvent(new MouseEvent("mousedown", eventInit));
  if (window.PointerEvent) {
    element.dispatchEvent(new PointerEvent("pointerup", { ...eventInit, pointerId: 1, pointerType: "mouse", isPrimary: true }));
  }
  element.dispatchEvent(new MouseEvent("mouseup", eventInit));
  element.dispatchEvent(new MouseEvent("click", eventInit));
  return true;
}

function elementCenterPoint(element) {
  if (!element || !(element instanceof Element)) return false;
  element.scrollIntoView?.({ block: "center", inline: "center" });
  const rect = element.getBoundingClientRect();
  if (!rect.width || !rect.height) return false;
  return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 };
}

function clickElementCenter(element) {
  const point = elementCenterPoint(element);
  if (!point) return false;
  return dispatchPointClick(point.x, point.y);
}

function trustedClickAt(x, y) {
  if (!BCA_ENABLE_DEBUGGER_INPUT) return Promise.resolve(false);
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "bca:trustedClick", x, y }, (response) => {
      if (chrome.runtime.lastError) {
        resolve(false);
        return;
      }
      resolve(Boolean(response && response.ok));
    });
  });
}

function trustedTypeText(text) {
  if (!BCA_ENABLE_DEBUGGER_INPUT) return Promise.resolve(false);
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "bca:trustedType", text }, (response) => {
      if (chrome.runtime.lastError) {
        resolve(false);
        return;
      }
      resolve(Boolean(response && response.ok));
    });
  });
}

function trustedInsertText(text) {
  if (!BCA_ENABLE_DEBUGGER_INPUT) return Promise.resolve(false);
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "bca:trustedInsertText", text }, (response) => {
      if (chrome.runtime.lastError) {
        resolve(false);
        return;
      }
      resolve(Boolean(response && response.ok));
    });
  });
}

function trustedPressKey(key) {
  if (!BCA_ENABLE_DEBUGGER_INPUT) return Promise.resolve(false);
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "bca:trustedPressKey", key }, (response) => {
      if (chrome.runtime.lastError) {
        resolve(false);
        return;
      }
      resolve(Boolean(response && response.ok));
    });
  });
}

async function ensurePageWorldBridgeInstalled() {
  if (pageWorldBridgeInstalledVersion === BCA_PAGE_BRIDGE_EVENT_SUFFIX) return true;
  return new Promise((resolve) => {
    const script = document.createElement("script");
    const bridgeUrl = chrome.runtime.getURL("page-world/react-select-bridge.js");
    script.src = `${bridgeUrl}?v=${encodeURIComponent(BCA_PAGE_BRIDGE_EVENT_SUFFIX)}&t=${Date.now()}`;
    script.async = false;
    script.dataset.bcaReactSelectBridge = "true";
    script.dataset.bcaReactSelectBridgeVersion = BCA_PAGE_BRIDGE_EVENT_SUFFIX;
    script.onload = () => {
      pageWorldBridgeInstalledVersion = BCA_PAGE_BRIDGE_EVENT_SUFFIX;
      script.remove();
      resolve(true);
    };
    script.onerror = () => {
      lastMainWorldLanguageSelectionDebug = "page bridge failed to load";
      script.remove();
      resolve(false);
    };
    (document.head || document.documentElement).appendChild(script);
  });
}

async function directSelectAddTranslationLanguage(languageLabel) {
  const query = languageSearchQuery(languageLabel);
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "bca:mainWorldSelectAddTranslationLanguage", languageLabel, query }, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, method: "main-world-message-error", debug: chrome.runtime.lastError.message });
        return;
      }
      resolve((response && response.response) || { ok: false, method: "main-world-empty", debug: "Empty service worker response." });
    });
  });
}

async function bridgeSelectAddTranslationLanguage(languageLabel) {
  const query = languageSearchQuery(languageLabel);
  const bridgeReady = await ensurePageWorldBridgeInstalled();
  if (!bridgeReady) {
    return { ok: false, method: "page-bridge-load-failed", debug: "Page bridge failed to install." };
  }
  const requestId = `bca-bridge-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  return new Promise((resolve) => {
    let done = false;
    const finish = (result) => {
      if (done) return;
      done = true;
      window.removeEventListener(BCA_PAGE_BRIDGE_RESULT_EVENT, onResult);
      clearTimeout(timeoutId);
      resolve(result);
    };
    const onResult = (event) => {
      const detail = event && event.detail ? event.detail : {};
      if (detail.requestId !== requestId) return;
      const result =
        detail.result && typeof detail.result === "object"
          ? detail.result
          : { ok: false, method: "page-bridge-empty", debug: "Empty page bridge result." };
      finish(result);
    };
    const timeoutId = window.setTimeout(
      () => finish({ ok: false, method: "page-bridge-timeout", debug: "Timed out waiting for page bridge result." }),
      1500
    );
    window.addEventListener(BCA_PAGE_BRIDGE_RESULT_EVENT, onResult);
    window.dispatchEvent(
      new CustomEvent(BCA_PAGE_BRIDGE_REQUEST_EVENT, {
        detail: {
          requestId,
          languageLabel,
          query,
        },
      })
    );
  });
}

async function mainWorldSelectAddTranslationLanguage(languageLabel, dialog = findAddTranslationPanelFromHeading() || findModalRoot()) {
  const directResult = await directSelectAddTranslationLanguage(languageLabel);
  lastMainWorldLanguageSelectionDebug = `${directResult.method || "main-world-unknown"}: ${directResult.debug || ""}`;
  if (directResult && directResult.ok) {
    await delay(700);
    if (isAddTranslationSelectionReady(dialog)) return true;
    const active = document.activeElement instanceof Element ? document.activeElement : null;
    if (active) sendKeyPress(active, "Enter", "Enter");
    await delay(400);
    if (isAddTranslationSelectionReady(dialog)) return true;
  }
  return false;
}

function trustedLanguagePickerFlow(input, languageLabel, dialog = findAddTranslationPanelFromHeading() || findModalRoot()) {
  if (!BCA_ENABLE_DEBUGGER_INPUT) {
    lastMainWorldLanguageSelectionDebug = "trusted-flow disabled";
    return Promise.resolve(false);
  }
  const query = languageSearchQuery(languageLabel);
  const point = input instanceof Element ? elementCenterPoint(input) : null;
  lastMainWorldLanguageSelectionDebug = `trusted-flow attempting ${query}`;
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(
      {
        type: "bca:trustedLanguagePickerFlow",
        text: query,
        x: point ? point.x : null,
        y: point ? point.y : null,
      },
      async (response) => {
        if (chrome.runtime.lastError) {
          lastMainWorldLanguageSelectionDebug = `trusted-flow error: ${chrome.runtime.lastError.message}`;
          resolve(false);
          return;
        }
        if (!response || !response.ok) {
          lastMainWorldLanguageSelectionDebug = `trusted-flow failed: ${(response && response.error) || "unknown"}`;
          resolve(false);
          return;
        }
        await delay(800);
        if (isAddTranslationSelectionReady(dialog)) {
          lastMainWorldLanguageSelectionDebug = "trusted-flow ready";
          resolve(true);
          return;
        }
        if (await selectOptionByText(languageLabel, dialog)) {
          await delay(300);
          if (isAddTranslationSelectionReady(dialog)) {
            lastMainWorldLanguageSelectionDebug = "trusted-flow option click";
            resolve(true);
            return;
          }
        }
        lastMainWorldLanguageSelectionDebug = "trusted-flow dispatched but not selected";
        resolve(false);
      }
    );
  });
}

async function trustedClickElementCenter(element) {
  const point = elementCenterPoint(element);
  if (!point) return false;
  return trustedClickAt(point.x, point.y);
}

async function clickElementCenterWithTrustedFallback(element) {
  if (clickElementCenter(element)) return true;
  return trustedClickElementCenter(element);
}

function languageDropdownClickTargets(control) {
  const targets = [];
  const add = (element) => {
    if (element && element instanceof Element && !targets.includes(element)) targets.push(element);
  };
  add(control);
  add(control.closest("[class*='control'], [class*='Control']"));
  add(control.closest("[class*='singleValue'], [class*='SingleValue'], [class*='valueContainer'], [class*='ValueContainer']"));
  add(control.closest("[class*='select'], [class*='Select'], [class*='dropdown'], [class*='Dropdown'], [role='combobox'], [aria-haspopup='listbox']"));

  const root = findModalRoot();
  const englishTextTarget = findEnglishLanguageControl(root);
  add(englishTextTarget);
  if (englishTextTarget) {
    add(englishTextTarget.closest("[class*='control'], [class*='Control']"));
    add(englishTextTarget.closest("[class*='select'], [class*='Select'], [class*='dropdown'], [class*='Dropdown'], [role='combobox']"));
    let ancestor = englishTextTarget.parentElement;
    const modal = findModalRoot();
    while (ancestor && ancestor !== modal && targets.length < 10) {
      add(ancestor);
      ancestor = ancestor.parentElement;
    }
  }
  return targets.slice(0, 6);
}

async function openLanguageDropdownControl(control) {
  const interactionRoot = findModalRoot();
  const targets = languageDropdownClickTargets(control);
  const hasAddTranslation = () =>
    Boolean(findAddTranslationAction(interactionRoot, control) || findAddTranslationAction(document, control));
  for (const target of targets) {
    target.click?.();
    await delay(450);
    if (hasAddTranslation()) return true;
    dispatchUserClick(target);
    await delay(350);
    if (hasAddTranslation()) return true;
  }

  const focusTarget = targets.find((target) => target.matches("input, [role='combobox'], [role='textbox']")) || control;
  focusTarget.focus?.();
  focusTarget.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "ArrowDown", code: "ArrowDown" }));
  focusTarget.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "ArrowDown", code: "ArrowDown" }));
  await delay(350);
  return hasAddTranslation();
}

async function openEnglishLanguageMenu(root = findModalRoot()) {
  const englishControl = findEnglishLanguageControl(root) || findEnglishLanguageControl(document);
  if (!englishControl) return false;
  const opened = await openLanguageDropdownControl(englishControl);
  if (opened) return true;

  const targets = languageDropdownClickTargets(englishControl);
  for (const target of targets) {
    clickElementCenter(target);
    dispatchUserClick(target);
    await delay(350);
    if (findAddTranslationAction(root, englishControl) || findAddTranslationAction(document, englishControl)) return true;
  }
  return false;
}

async function openCurrentLanguageMenu(root = findEditAudioRoot()) {
  const control =
    findLanguageDropdownControl(root) ||
    findLanguageDropdownControl(document) ||
    findEnglishLanguageControl(root) ||
    findEnglishLanguageControl(document);
  if (!control) return false;
  const opened = await openLanguageDropdownControl(control);
  if (opened) return true;

  const targets = languageDropdownClickTargets(control);
  for (const target of targets) {
    clickElementCenter(target);
    dispatchUserClick(target);
    await delay(350);
    if (findAddTranslationAction(root, control) || findAddTranslationAction(document, control)) return true;
  }
  return Boolean(findAddTranslationAction(root, control) || findAddTranslationAction(document, control));
}

async function activateAddTranslationAction(addTranslation, dropdown) {
  const actionTarget =
    addTranslation &&
    (addTranslation.closest("[role='option'], [role='menuitem'], li") || addTranslation);

  if (actionTarget) {
    clickElementCenter(actionTarget);
    dispatchUserClick(actionTarget);
    await delay(600);
    const opened = await waitForAddTranslationDialog(1200);
    if (opened) return opened;
    await trustedClickElementCenter(actionTarget);
    await delay(600);
    const openedAfterTrustedClick = await waitForAddTranslationDialog(1200);
    if (openedAfterTrustedClick) return openedAfterTrustedClick;
  }

  const menuControl =
    dropdown ||
    findLanguageDropdownControl(findModalRoot()) ||
    findEnglishLanguageControl(findModalRoot()) ||
    findEnglishLanguageControl(document);

  if (menuControl) {
    await openLanguageDropdownControl(menuControl);
    await delay(150);
  }

  const input =
    findReactSelectInput(findModalRoot()) ||
    findReactSelectInput(document) ||
    (menuControl && menuControl.querySelector && menuControl.querySelector("input[id^='react-select-'][id$='-input'], input[role='combobox']"));

  if (input) {
    input.focus?.();
    sendKeyPress(input, "ArrowDown", "ArrowDown");
    await delay(120);
    sendKeyPress(input, "Enter", "Enter");
    await delay(650);
    const opened = await waitForAddTranslationDialog(1500);
    if (opened) return opened;
    await trustedPressKey("ArrowDown");
    await delay(120);
    await trustedPressKey("Enter");
    await delay(650);
    const openedAfterTrustedKeys = await waitForAddTranslationDialog(1500);
    if (openedAfterTrustedKeys) return openedAfterTrustedKeys;
  }

  const refreshedAction = findAddTranslationAction(findModalRoot(), menuControl) || findAddTranslationAction(document, menuControl);
  const refreshedTarget =
    refreshedAction && (refreshedAction.closest("[role='option'], [role='menuitem'], li") || refreshedAction);
  if (refreshedTarget) {
    dispatchUserClick(refreshedTarget);
    clickElementCenter(refreshedTarget);
    await delay(700);
    const opened = await waitForAddTranslationDialog(1500);
    if (opened) return opened;
    await trustedClickElementCenter(refreshedTarget);
    await delay(700);
    const openedAfterTrustedClick = await waitForAddTranslationDialog(1500);
    if (openedAfterTrustedClick) return openedAfterTrustedClick;
  }

  return null;
}

function findLanguageDropdownControl(root) {
  const select = visibleElements("select", root).find((candidate) => isLanguageControlText(controlText(candidate)));
  if (select) return select;

  const directControl = visibleElements(
    "input:not([type='hidden']), [role='combobox'], [role='textbox'], button, [aria-haspopup='listbox'], [aria-haspopup='menu']",
    root
  ).find((element) => {
    const text = normalizeComparable(controlText(element));
    return !text.includes("add translation") && isLanguageControlText(text);
  });
  if (directControl) return directControl;

  return visibleElements(
    "[class*='select'], [class*='Select'], [class*='dropdown'], [class*='Dropdown'], [class*='language'], [class*='Language'], div",
    root
  )
    .filter((element) => {
      const text = normalizeComparable(controlText(element));
      if (!text || text.includes("add translation") || !isKnownLanguageText(text)) return false;
      if (controlText(element).length > 140) return false;
      return Boolean(element.querySelector("input, svg, [class*='indicator'], [class*='Indicator']")) || element.matches("[class*='select'], [class*='Select'], [class*='dropdown'], [class*='Dropdown']");
    })
    .sort((a, b) => controlText(a).length - controlText(b).length)[0];
}

function findEnglishLanguageControl(root = document) {
  return visibleElements(
    "select, button, input:not([type='hidden']), [role='combobox'], [aria-haspopup='listbox'], [aria-haspopup='menu'], [class*='select'], [class*='Select'], [class*='control'], [class*='Control']",
    root
  )
    .filter((element) => {
      const comparable = normalizeComparable(controlText(element));
      if (!comparable.includes("english")) return false;
      if (comparable.includes("add translation")) return false;
      if (controlText(element).length > 120) return false;
      return (
        element.matches("select, button, input:not([type='hidden']), [role='combobox'], [aria-haspopup='listbox'], [aria-haspopup='menu']") ||
        Boolean(element.querySelector("input, svg, [class*='indicator'], [class*='Indicator'], [aria-haspopup='listbox'], [aria-haspopup='menu']"))
      );
    })
    .sort((a, b) => controlText(a).length - controlText(b).length)[0];
}

function isLanguageControlText(value) {
  const comparable = normalizeComparable(value);
  return (
    isKnownLanguageText(comparable) ||
    comparable.includes("language") ||
    comparable.includes("translation")
  );
}

function isKnownLanguageText(value) {
  const comparable = normalizeComparable(value);
  return (
    comparable.includes("english") ||
    comparable.includes("spanish") ||
    comparable.includes("portuguese") ||
    comparable.includes("french") ||
    comparable.includes("italian") ||
    comparable.includes("german")
  );
}

function languageLabelMatches(activeLabel, expectedLabel) {
  const active = normalizeComparable(activeLabel);
  const expected = normalizeComparable(expectedLabel);
  if (!active || !expected) return false;
  return active === expected || active.includes(expected) || expected.includes(active);
}

function languageLabelAliases(label) {
  const comparable = normalizeComparable(label);
  const aliases = [label];
  if (comparable.includes("portuguese")) aliases.push("Portuguese (Brazil)", "Portuguese");
  if (comparable.includes("spanish")) aliases.push("Spanish (Latin America)", "Spanish");
  if (comparable.includes("french")) aliases.push("French");
  if (comparable.includes("italian")) aliases.push("Italian");
  if (comparable.includes("german")) aliases.push("German");
  return Array.from(new Set(aliases.map(normalizeSpaces).filter(Boolean)));
}

function findExistingLanguageMenuOption(languageLabel) {
  const aliases = languageLabelAliases(languageLabel);
  return visibleElements(
    [
      "[role='option']",
      "[role='menuitem']",
      "li",
      "button",
      "div",
      "span",
      "[class*='option']",
      "[class*='Option']",
      "[class*='menu']",
      "[class*='Menu']",
    ].join(", "),
    document
  )
    .filter((element) => {
      if (getPanelRoot()?.contains(element)) return false;
      const text = controlText(element);
      const comparable = normalizeComparable(text);
      if (!comparable || text.length > 120) return false;
      if (comparable.includes("add translation") || comparable === "add" || comparable === "cancel" || comparable.includes("no options")) return false;
      return aliases.some((alias) => languageLabelMatches(text, alias));
    })
    .sort((a, b) => {
      const aText = controlText(a);
      const bText = controlText(b);
      const aExact = languageLabelAliases(languageLabel).some((alias) => normalizeComparable(aText) === normalizeComparable(alias)) ? 0 : 1;
      const bExact = languageLabelAliases(languageLabel).some((alias) => normalizeComparable(bText) === normalizeComparable(alias)) ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      return aText.length - bText.length;
    })[0];
}

function elementOptionText(element) {
  return controlText(element);
}

function findVisibleOptionByText(label, root = document) {
  const target = normalizeComparable(label);
  const targetTokens = target.split(" ").filter((token) => token.length >= 3);
  return visibleElements(
    [
      "[role='option']",
      "[role='menuitem']",
      "li",
      "button",
      "div",
      "span",
      "[class*='option']",
      "[class*='Option']",
      "[class*='menu']",
      "[class*='Menu']",
      "[class*='menuItem']",
      "[class*='MenuItem']",
    ].join(", "),
    root
  )
    .filter((element) => {
      const text = elementOptionText(element);
      const comparable = normalizeComparable(text);
      if (!comparable) return false;
      const tokenMatch = targetTokens.length && targetTokens.every((token) => comparable.includes(token));
      if (!(comparable === target || comparable.includes(target) || tokenMatch)) return false;
      return text.length <= Math.max(label.length + 36, 96);
    })
    .sort((a, b) => elementOptionText(a).length - elementOptionText(b).length)[0];
}

function findOptionByTextAnyVisibility(label, root = document) {
  const target = normalizeComparable(label);
  const targetTokens = target.split(" ").filter((token) => token.length >= 3);
  const nodes = Array.from(
    root.querySelectorAll(
      [
        "[role='option']",
        "[role='menuitem']",
        "li",
        "button",
        "div",
        "span",
        "[class*='option']",
        "[class*='Option']",
        "[class*='menu']",
        "[class*='Menu']",
        "[class*='menuItem']",
        "[class*='MenuItem']",
      ].join(", ")
    )
  );

  return nodes
    .filter((element) => {
      const text = elementOptionText(element);
      const comparable = normalizeComparable(text);
      if (!comparable) return false;
      const tokenMatch = targetTokens.length && targetTokens.every((token) => comparable.includes(token));
      return comparable === target || comparable.includes(target) || tokenMatch;
    })
    .sort((a, b) => elementOptionText(a).length - elementOptionText(b).length)[0];
}

function formatClickTargetForDebug(element) {
  if (!element) return "none";
  const rect = element.getBoundingClientRect();
  return `${element.tagName.toLowerCase()}${element.id ? `#${element.id}` : ""}${element.className ? `.${String(element.className).trim().replace(/\s+/g, ".").slice(0, 80)}` : ""} text="${controlText(element).slice(0, 80)}" rect=${Math.round(rect.left)},${Math.round(rect.top)},${Math.round(rect.width)}x${Math.round(rect.height)}`;
}

function menuOptionCandidates(label) {
  const target = normalizeComparable(label);
  return visibleElements("div, span, li, button, [role='menuitem'], [class*='menu'], [class*='Menu']", document)
    .filter((element) => {
      if (getPanelRoot()?.contains(element)) return false;
      const text = normalizeComparable(controlText(element));
      if (!text) return false;
      if (text === target) return true;
      return text.includes(target) && controlText(element).length <= label.length + 28;
    })
    .sort((a, b) => {
      const aExact = normalizeComparable(controlText(a)) === target ? 0 : 1;
      const bExact = normalizeComparable(controlText(b)) === target ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      return controlText(a).length - controlText(b).length;
    });
}

function clickableMenuRowForOption(option) {
  if (!option) return null;
  const target = normalizeComparable(controlText(option));
  let current = option;
  while (current && current !== document.body) {
    const rect = current.getBoundingClientRect();
    const text = normalizeComparable(controlText(current));
    const role = normalizeComparable(current.getAttribute("role"));
    const style = window.getComputedStyle(current);
    const looksClickable =
      role.includes("menuitem") ||
      current.tagName === "BUTTON" ||
      style.cursor === "pointer" ||
      current.onclick ||
      current.getAttribute("tabindex") !== null;
    const rowSized = rect.width >= 80 && rect.height >= 24 && rect.height <= 80;
    if (rowSized && (text === target || looksClickable)) return current;
    current = current.parentElement;
  }
  return option;
}

function findVisibleMenuPopup() {
  return visibleElements("div, ul, [role='menu'], [role='listbox'], [class*='menu'], [class*='Menu']", document)
    .filter((element) => {
      if (getPanelRoot()?.contains(element)) return false;
      const text = normalizeComparable(controlText(element));
      return text.includes("replace audio") && text.includes("download audio");
    })
    .sort((a, b) => {
      const ar = a.getBoundingClientRect();
      const br = b.getBoundingClientRect();
      return ar.width * ar.height - br.width * br.height;
    })[0];
}

function findElementContainingText(text, root = document) {
  const target = normalizeComparable(text);
  return visibleElements("label, div, span, p, strong, h1, h2, h3, h4", root)
    .filter((element) => normalizeComparable(controlText(element)).includes(target))
    .sort((a, b) => controlText(a).length - controlText(b).length)[0];
}

function findFieldLabel(text, root = document) {
  const target = normalizeComparable(text);
  return visibleElements("label, div, span, p, strong", root)
    .filter((element) => {
      const comparable = normalizeComparable(controlText(element));
      if (!comparable.includes(target)) return false;
      if (comparable.includes("edit audio")) return false;
      if (controlText(element).length > 80) return false;
      return true;
    })
    .sort((a, b) => {
      const ar = a.getBoundingClientRect();
      const br = b.getBoundingClientRect();
      const aExact = normalizeComparable(controlText(a)).startsWith(target) ? 0 : 1;
      const bExact = normalizeComparable(controlText(b)).startsWith(target) ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      return ar.top - br.top;
    })[0];
}

async function selectOptionByText(label, root = document) {
  const option = findVisibleOptionByText(label, root) || findOptionByTextAnyVisibility(label, root);
  if (!option) return false;
  if (!(await trustedClickElementCenter(option))) {
    option.click();
  }
  await delay(500);
  return true;
}

async function waitForAddTranslationSelectionReady(dialog, timeoutMs = 2500) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    if (isAddTranslationSelectionReady(dialog)) return true;
    await delay(120);
  }
  return false;
}

function findAddTranslationLanguageOption(label, dialog) {
  const target = normalizeComparable(label);
  const targetTokens = target.split(" ").filter((token) => token.length >= 3);
  return visibleElements(
    [
      "[role='option']",
      "[role='menuitem']",
      "li",
      "button",
      "div",
      "span",
      "[class*='option']",
      "[class*='Option']",
    ].join(", "),
    document
  )
    .filter((element) => {
      if (getPanelRoot()?.contains(element)) return false;
      const text = controlText(element);
      const comparable = normalizeComparable(text);
      if (!comparable || text.length > Math.max(label.length + 36, 96)) return false;
      const tokenMatch = targetTokens.length && targetTokens.every((token) => comparable.includes(token));
      if (!(comparable === target || comparable.includes(target) || tokenMatch)) return false;
      const role = normalizeComparable(element.getAttribute("role"));
      const menuAncestor = element.closest("[role='listbox'], [role='menu'], [class*='menu'], [class*='Menu']");
      const controlAncestor = element.closest("[class*='select__control'], [class*='Select__control'], [class*='control'], [class*='Control']");
      const inDialog = dialog && dialog.contains(element);
      return role.includes("option") || role.includes("menuitem") || menuAncestor || (inDialog && !controlAncestor);
    })
    .sort((a, b) => {
      const aText = controlText(a);
      const bText = controlText(b);
      const aExact = normalizeComparable(aText) === target ? 0 : 1;
      const bExact = normalizeComparable(bText) === target ? 0 : 1;
      if (aExact !== bExact) return aExact - bExact;
      const aRole = normalizeComparable(a.getAttribute("role"));
      const bRole = normalizeComparable(b.getAttribute("role"));
      const aOption = aRole.includes("option") ? 0 : 1;
      const bOption = bRole.includes("option") ? 0 : 1;
      if (aOption !== bOption) return aOption - bOption;
      return aText.length - bText.length;
    })[0];
}

async function selectAddTranslationOptionByText(label, dialog) {
  const option = findAddTranslationLanguageOption(label, dialog) || findVisibleOptionByText(label, dialog) || findVisibleOptionByText(label, document);
  if (!option) return false;
  const clickTarget = clickableMenuRowForOption(option) || option;
  if (!(await trustedClickElementCenter(clickTarget))) {
    clickElementCenter(clickTarget);
    dispatchUserClick(clickTarget);
  }
  return waitForAddTranslationSelectionReady(dialog);
}

function findAudioOverflowMenuButton(root = findModalRoot()) {
  const buttons = visibleElements("button, [role='button'], [aria-label], [class*='button'], [class*='Button']", root).filter((button) => {
    const text = normalizeComparable(controlText(button));
    return !text || text.includes("more") || text.includes("menu") || text.includes("audio");
  });
  const audioLabel = findFieldLabel("Audio", root);
  const titleLabel = findFieldLabel("Title", root);
  const titleField = findFieldByLabel("Title", root);

  const titleAnchor = titleField || titleLabel;
  if (titleAnchor) {
    const rootRect = root.getBoundingClientRect();
    const titleRect = titleAnchor.getBoundingClientRect();
    const audioRect = audioLabel ? audioLabel.getBoundingClientRect() : null;
    const top = audioRect
      ? Math.max((audioRect.height > 60 ? audioRect.top : audioRect.bottom - 12), rootRect.top + 120)
      : Math.max(titleRect.top - 240, rootRect.top + 160);
    const bottom = titleRect.top - 6;
    const left = titleRect.left - 80;
    const right = titleRect.right + 120;
    const scoped = buttons
      .map((button) => {
        const rect = button.getBoundingClientRect();
        const cy = rect.top + rect.height / 2;
        const cx = rect.left + rect.width / 2;
        return { button, rect, cy, cx, distanceToTitle: Math.abs(titleRect.top - cy) };
      })
      .filter(({ rect, cy, cx }) => {
        if (cy <= top || cy >= bottom) return false;
        if (rect.width < 8 || rect.height < 8 || rect.width > 90 || rect.height > 90) return false;
        if (cx < left || cx > right) return false;
        return true;
      })
      .sort((a, b) => a.distanceToTitle - b.distanceToTitle || b.rect.left - a.rect.left)[0];
    if (scoped) return scoped.button;
  }

  const debug = [
    `titleField=${formatClickTargetForDebug(titleField)}`,
    `audioLabel=${formatClickTargetForDebug(audioLabel)}`,
    `titleLabel=${formatClickTargetForDebug(titleLabel)}`,
    `buttons=${buttons.slice(0, 8).map(formatClickTargetForDebug).join(" | ") || "none"}`,
  ].join("; ");
  lastDownloadAudioClickDebug = `audioMenuNotFound: ${debug}`;
  return null;
}

async function clickDownloadAudioMenuItem() {
  const root = findModalRoot();
  const menuButton = findAudioOverflowMenuButton(root);
  if (!menuButton) throw new Error(`Could not find the audio three-dot menu. Debug: ${lastDownloadAudioClickDebug}`);

  lastDownloadAudioClickDebug = `menuButton=${formatClickTargetForDebug(menuButton)}`;
  let option = menuOptionCandidates("Download Audio")[0];
  if (!option) {
    const opened = await trustedClickElementCenter(menuButton);
    if (!opened) clickElementCenter(menuButton);
    await delay(450);
    option = menuOptionCandidates("Download Audio")[0];
  }
  if (!option) {
    clickElementCenter(menuButton);
    await delay(450);
    option = menuOptionCandidates("Download Audio")[0];
  }

  const candidates = menuOptionCandidates("Download Audio");
  lastDownloadAudioClickDebug = `${lastDownloadAudioClickDebug}; candidates=${candidates.slice(0, 5).map(formatClickTargetForDebug).join(" | ") || "none"}`;
  if (!option) throw new Error(`Could not find Download Audio in the audio menu. Debug: ${lastDownloadAudioClickDebug}`);

  const row = clickableMenuRowForOption(option);
  lastDownloadAudioClickDebug = `${lastDownloadAudioClickDebug}; chosen=${formatClickTargetForDebug(option)}; row=${formatClickTargetForDebug(row)}`;
  const clicked = await trustedClickElementCenter(row);
  if (!clicked) {
    clickElementCenter(row);
  }
  await delay(650);

  if (menuOptionCandidates("Download Audio").length) {
    const popup = findVisibleMenuPopup();
    lastDownloadAudioClickDebug = `${lastDownloadAudioClickDebug}; stillVisible=true; popup=${formatClickTargetForDebug(popup)}`;
    if (popup) {
      const rect = popup.getBoundingClientRect();
      await trustedClickAt(rect.left + rect.width / 2, rect.top + rect.height * 0.73);
      await delay(650);
    }
  }

  if (menuOptionCandidates("Download Audio").length) {
    lastDownloadAudioClickDebug = `${lastDownloadAudioClickDebug}; finalStillVisible=true`;
    throw new Error(`Download Audio menu item stayed open after click. Debug: ${lastDownloadAudioClickDebug}`);
  }
  await delay(500);
}

function sendKeyPress(target, key, code = key) {
  const eventInit = { bubbles: true, cancelable: true, key, code };
  target.dispatchEvent(new KeyboardEvent("keydown", eventInit));
  target.dispatchEvent(new KeyboardEvent("keypress", eventInit));
  target.dispatchEvent(new KeyboardEvent("keyup", eventInit));
}

function isReactSelectInput(element) {
  if (!(element instanceof Element)) return false;
  const id = element.getAttribute("id") || "";
  const className = element.getAttribute("class") || "";
  return /react-select-\d+-input/.test(id) || className.includes("select__input") || className.includes("docent-select__input");
}

function getReactProps(node) {
  if (!(node instanceof Element)) return null;
  const key = Object.keys(node).find((entry) => entry.startsWith("__reactProps$"));
  return key ? node[key] : null;
}

function getReactFiber(node) {
  if (!(node instanceof Element)) return null;
  const key = Object.keys(node).find((entry) => entry.startsWith("__reactFiber$") || entry.startsWith("__reactInternalInstance$"));
  return key ? node[key] : null;
}

function flattenReactSelectOptions(options) {
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
}

function optionLabel(option) {
  if (!option || typeof option !== "object") return "";
  return normalizeSpaces(option.label || option.name || option.title || option.value || "");
}


function findReactSelectComponentFromNode(node) {
  let fiber = getReactFiber(node);
  // Check before traversal: if the original node has aria-haspopup="listbox",
  // use its fiber directly rather than risking returning a mismatched ancestor.
  if (node instanceof Element && node.getAttribute("aria-haspopup") === "listbox" && fiber) {
    const props = fiber.memoizedProps || fiber.pendingProps || {};
    return { fiber, props, source: "aria-haspopup" };
  }
  while (fiber) {
    const props = fiber.memoizedProps || fiber.pendingProps || {};
    const selectProps = props.selectProps || {};
    
    // Check for React Select props (multiple possible patterns)
    const hasOptions = Array.isArray(props.options) || Array.isArray(selectProps.options);
    const hasOnChange = typeof props.onChange === "function" || typeof selectProps.onChange === "function";
    
    // Also check for common React Select prop names
    const isSelectComponent = 
      props.className && typeof props.className === 'string' && 
      (props.className.includes('select__') || props.className.includes('Select__') || 
       props.className.includes('docent-select__'));
    
    if ((hasOptions && hasOnChange) || isSelectComponent) {
      return { 
        fiber, 
        props: hasOptions ? (Array.isArray(props.options) ? props : selectProps) : props,
        source: Array.isArray(props.options) ? "props" : "selectProps"
      };
    }
    
    fiber = fiber.return;
  }
  return null;
}


function collectReactSelectComponents(root = document) {
  const { control, input } = getAddTranslationReactSelectNodes(root);
  const anchors = [
    input,
    control,
    findReactSelectInput(root),
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
    const component = findReactSelectComponentFromNode(anchor);
    if (!component || !component.props || seen.has(component.props)) continue;
    seen.add(component.props);
    components.push(component);
  }
  return components;
}

function languageComparableCandidates(languageLabel) {
  const target = normalizeComparable(languageLabel);
  const query = normalizeComparable(languageSearchQuery(languageLabel));
  const values = new Set([target, query].filter(Boolean));
  if (target.includes("portuguese") || query.includes("portuguese")) values.add("portugues");
  if (target.includes("spanish") || query.includes("spanish")) values.add("espanol");
  if (target.includes("french") || query.includes("french")) values.add("francais");
  if (target.includes("german") || query.includes("german")) values.add("deutsch");
  if (target.includes("italian") || query.includes("italian")) values.add("italiano");
  return Array.from(values);
}

function findReactSelectOption(selectProps, languageLabel) {
  const candidates = languageComparableCandidates(languageLabel);
  const options = flattenReactSelectOptions(selectProps.options);
  return (
    options.find((option) => candidates.includes(normalizeComparable(optionLabel(option)))) ||
    options.find((option) => candidates.some((candidate) => normalizeComparable(optionLabel(option)).includes(candidate))) ||
    options.find((option) => candidates.some((candidate) => candidate.includes(normalizeComparable(optionLabel(option))))) ||
    null
  );
}

async function selectLanguageViaReactFiber(dialog, languageLabel) {
  const components = collectReactSelectComponents(dialog);
  for (const component of components) {
    const option = findReactSelectOption(component.props, languageLabel);
    if (!option) continue;
    try {
      const handler = component.props.onChange;
      const context = component.fiber.stateNode || component.fiber.memoizedProps;
      if (handler && context) {
        handler.call(context, option, {
          action: "select-option",
          name: component.props.name,
          option,
        });
      } else if (handler) {
        handler(option, {
          action: "select-option",
          name: component.props.name,
          option,
        });
      }
      await delay(250);
      return isAddTranslationSelectionReady(dialog);
    } catch (error) {
      console.warn("Bloomberg Audio Assistant React Select fiber selection failed:", error);
    }
  }
  return false;
}

function invokeReactHandler(node, handlerName, eventLike = {}) {
  const props = getReactProps(node);
  const handler = props && typeof props[handlerName] === "function" ? props[handlerName] : null;
  if (!handler) return false;
  const synthetic = {
    target: node,
    currentTarget: node,
    type: handlerName.replace(/^on/, "").toLowerCase(),
    bubbles: true,
    cancelable: true,
    button: 0,
    preventDefault: () => {},
    stopPropagation: () => {},
    nativeEvent: eventLike.nativeEvent || {},
    ...eventLike,
  };
  try {
    handler(synthetic);
    return true;
  } catch (_) {
    return false;
  }
}

function findReactSelectInput(root = document) {
  const scoped = root.querySelector("input[id^='react-select-'][id$='-input'], input.docent-select__input, input[class*='select__input']");
  if (scoped instanceof Element) return scoped;
  const active = document.activeElement instanceof Element ? document.activeElement : null;
  if (active && isReactSelectInput(active)) return active;
  return null;
}

function sendArrowDown(target) {
  const eventInit = { bubbles: true, cancelable: true, key: "ArrowDown", code: "ArrowDown", keyCode: 40, which: 40 };
  target.dispatchEvent(new KeyboardEvent("keydown", eventInit));
  target.dispatchEvent(new KeyboardEvent("keyup", eventInit));
}

function isReactSelectMenuOpen(input) {
  if (!input) return false;
  const expanded = normalizeComparable(input.getAttribute("aria-expanded"));
  if (expanded === "true") return true;
  const controls = input.getAttribute("aria-controls");
  if (controls && document.getElementById(controls)) return true;
  return visibleElements("[role='option'], [role='listbox'], [class*='menu'], [class*='Menu']", document).length > 0;
}

function isAddTranslationDialogOpen() {
  const dialog = findAddTranslationPanelFromHeading() || findModalRoot();
  const text = normalizeComparable(dialog.textContent || "");
  return text.includes("add translation") && text.includes("choose a language");
}

async function waitForAddTranslationDialog(timeoutMs = 2500) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    if (isAddTranslationDialogOpen()) return findAddTranslationPanelFromHeading() || findModalRoot();
    await delay(120);
  }
  return null;
}

function findAddTranslationLanguageControl(dialog) {
  const labeled = visibleElements("label", dialog).find((label) => normalizeComparable(label.textContent).includes("language"));
  if (labeled) {
    const container = labeled.closest("div, section, form");
    const controlInContainer =
      container &&
      container.querySelector(
        "[class*='docent-select__control'], [class*='select__control'], [class*='Select__control'], [role='combobox'], [aria-haspopup='listbox']"
      );
    if (controlInContainer) return controlInContainer;
  }

  const chooseText = visibleElements("div, span", dialog).find((node) => normalizeComparable(controlText(node)).includes("choose a language"));
  if (chooseText) {
    const control =
      chooseText.closest("[class*='docent-select__control'], [class*='select__control'], [class*='Select__control'], [role='combobox'], [aria-haspopup='listbox']") ||
      chooseText.closest("[class*='select'], [class*='Select'], div");
    if (control) return control;
  }
  return null;
}

async function trustedClickElementRightSide(element) {
  if (!(element instanceof Element)) return false;
  element.scrollIntoView?.({ block: "center", inline: "center" });
  const rect = element.getBoundingClientRect();
  if (!rect.width || !rect.height) return false;
  const x = rect.right - Math.min(14, rect.width / 4);
  const y = rect.top + rect.height / 2;
  if (dispatchPointClick(x, y)) return true;
  return trustedClickAt(x, y);
}



async function openAddTranslationPickerDeterministic(dialog) {
  const control = findAddTranslationLanguageControl(dialog);
  const input = findReactSelectInput(dialog) || findReactSelectInput(document);

  // Try to open the menu using React-aware approaches
  const menuOpened = await forceOpenReactSelectMenu(input, control);
  if (menuOpened) return true;

  // Fallback: try clicking with non-trusted methods
  if (control) {
    clickElementCenter(control);
    await delay(100);
    if (await forceOpenReactSelectMenu(input, control)) return true;

    dispatchUserClick(control);
    await delay(100);
    if (await forceOpenReactSelectMenu(input, control)) return true;
  }

  const focusTarget = input || control;
  if (focusTarget) {
    focusTarget.focus?.();
    sendKeyPress(focusTarget, "ArrowDown", "ArrowDown");
    await delay(150);
    if (isReactSelectMenuOpen(focusTarget)) return true;

    sendKeyPress(focusTarget, " ", "Space");
    await delay(150);
    if (isReactSelectMenuOpen(focusTarget)) return true;
  }

  return isReactSelectMenuOpen(input) || isReactSelectMenuOpen(control);
}

async function openReactSelectMenu(input) {
  const control =
    input.closest("[class*='select__control'], [class*='Select__control'], [class*='control'], [class*='Control']") || input.parentElement;
  return forceOpenReactSelectMenu(input, control);
}

function setReactControlledInputValue(input, value) {
  const prototype = Object.getPrototypeOf(input);
  const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
  const previous = input.value;
  if (descriptor && descriptor.set) {
    descriptor.set.call(input, value);
  } else {
    input.value = value;
  }
  const tracker = input._valueTracker;
  if (tracker && typeof tracker.setValue === "function") {
    tracker.setValue(previous);
  }
  input.setAttribute("value", value);
  input.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, data: value, inputType: "insertText" }));
  input.dispatchEvent(new Event("input", { bubbles: true, cancelable: true }));
}

function typeIntoReactSelectInput(input, text) {
  input.focus?.();
  invokeReactHandler(input, "onFocus");
  if (input.matches("input, textarea")) {
    try {
      input.select?.();
      document.execCommand?.("insertText", false, text);
      if (String(input.value || "").length > 0) {
        input.dispatchEvent(new Event("input", { bubbles: true, cancelable: true }));
        invokeReactHandler(input, "onInput", { target: input, currentTarget: input });
        invokeReactHandler(input, "onChange", { target: input, currentTarget: input });
        return;
      }
    } catch (_) {}
  }

  setReactControlledInputValue(input, "");
  let composed = "";
  for (const char of text) {
    composed += char;
    const eventInit = { bubbles: true, cancelable: true, key: char, code: `Key${char.toUpperCase()}` };
    input.dispatchEvent(new KeyboardEvent("keydown", eventInit));
    setReactControlledInputValue(input, composed);
    invokeReactHandler(input, "onInput", { target: input, currentTarget: input });
    invokeReactHandler(input, "onChange", { target: input, currentTarget: input });
    input.dispatchEvent(new KeyboardEvent("keyup", eventInit));
  }
  input.dispatchEvent(new Event("input", { bubbles: true, cancelable: true }));
  invokeReactHandler(input, "onInput", { target: input, currentTarget: input });
  invokeReactHandler(input, "onChange", { target: input, currentTarget: input });
}

async function typeIntoReactSelectInputDeterministic(input, text) {
  input.focus?.();
  await openAddTranslationPickerDeterministic(findAddTranslationPanelFromHeading() || findModalRoot());
  await openReactSelectMenu(input);
  await trustedClickElementCenter(input);
  input.focus?.();
  await trustedPressKey("ArrowDown");
  await delay(120);

  if (await trustedInsertText(text)) {
    await delay(220);
    if (String(input.value || "").length > 0 || isReactSelectMenuOpen(input)) return true;
  }

  if (await trustedTypeText(text)) {
    await delay(220);
    if (String(input.value || "").length > 0 || isReactSelectMenuOpen(input)) return true;
  }

  typeIntoReactSelectInput(input, text);
  await delay(180);
  if (String(input.value || "").length > 0 || isReactSelectMenuOpen(input)) return true;

  // Last keyboard-only fallback when value remains empty.
  sendKeyPress(input, "Home", "Home");
  for (const char of text) {
    sendKeyPress(input, char, `Key${char.toUpperCase()}`);
  }
  await delay(120);
  return String(input.value || "").length > 0 || isReactSelectMenuOpen(input);
}

function languageSearchQuery(languageLabel) {
  const normalized = normalizeSpaces(languageLabel);
  const base = normalized.split("(")[0].trim();
  return base || normalized;
}

async function typeIntoFocusedLanguageInput(languageLabel) {
  const active = document.activeElement instanceof Element ? document.activeElement : null;
  if (!active || !active.matches("input, textarea, [role='combobox'], [role='textbox']")) return false;
  active.focus?.();
  const query = languageSearchQuery(languageLabel);

  if (isReactSelectInput(active)) {
    if (await typeIntoReactSelectInputDeterministic(active, query)) return true;
    return String(active.value || "").length > 0;
  }

  if (setInputValueWithoutBlur(active, query, { dispatchChange: false })) {
    return String(active.value || "").length > 0;
  }
  return false;
}

function findLanguageFieldInAddTranslationDialog(dialog) {
  const explicitLabelField = findFieldByLabel("Language", dialog);
  if (explicitLabelField) return explicitLabelField;

  const placeholderField = visibleElements(
    "input:not([type='hidden']), [role='combobox'], [role='textbox'], button, [aria-haspopup='listbox']",
    dialog
  ).find((element) => {
    const text = normalizeComparable(
      [element.getAttribute("aria-label"), element.getAttribute("placeholder"), element.textContent]
        .filter(Boolean)
        .join(" ")
    );
    return text.includes("choose a language") || text.includes("search");
  });
  return placeholderField || null;
}

function emitTypingToElement(element, text) {
  element.focus?.();
  if (element.matches("input:not([type='hidden']), textarea, [role='textbox'], [role='combobox']")) {
    const prototype = Object.getPrototypeOf(element);
    const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
    if (descriptor && descriptor.set) {
      descriptor.set.call(element, "");
    } else {
      element.value = "";
    }
  }
  let composed = "";
  for (const char of text) {
    composed += char;
    if (element.matches("input:not([type='hidden']), textarea, [role='textbox'], [role='combobox']")) {
      const prototype = Object.getPrototypeOf(element);
      const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");
      if (descriptor && descriptor.set) {
        descriptor.set.call(element, composed);
      } else {
        element.value = composed;
      }
    }
    const eventInit = { bubbles: true, cancelable: true, key: char, code: `Key${char.toUpperCase()}` };
    element.dispatchEvent(new KeyboardEvent("keydown", eventInit));
    element.dispatchEvent(new KeyboardEvent("keypress", eventInit));
    element.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, data: char, inputType: "insertText" }));
    element.dispatchEvent(new KeyboardEvent("keyup", eventInit));
  }
}

function languageSearchCandidates(dialog, languageControl) {
  const candidates = [];
  const add = (element) => {
    if (element && element instanceof Element && !candidates.includes(element)) candidates.push(element);
  };
  add(languageControl);
  add(document.activeElement);

  visibleElements("input:not([type='hidden']), textarea, [contenteditable='true'], [role='textbox'], [role='combobox']", dialog).forEach(add);
  Array.from(dialog.querySelectorAll("input, textarea, [contenteditable='true'], [role='textbox'], [role='combobox']")).forEach(add);

  visibleElements("input:not([type='hidden']), [role='textbox'], [role='combobox']", document)
    .filter((element) => {
      const hint = normalizeComparable(
        [element.getAttribute("placeholder"), element.getAttribute("aria-label"), element.getAttribute("name"), element.getAttribute("id")]
          .filter(Boolean)
          .join(" ")
      );
      return hint.includes("language") || hint.includes("search") || hint.includes("choose");
    })
    .forEach(add);
  Array.from(document.querySelectorAll("input, [role='textbox'], [role='combobox']")).forEach((element) => {
    const hint = normalizeComparable(
      [element.getAttribute("placeholder"), element.getAttribute("aria-label"), element.getAttribute("name"), element.getAttribute("id")]
        .filter(Boolean)
        .join(" ")
    );
    if (hint.includes("language") || hint.includes("search") || hint.includes("choose")) add(element);
  });

  return candidates;
}

async function typeLanguageInAddTranslation(dialog, languageControl, languageLabel) {
  const query = languageSearchQuery(languageLabel);
  const directReactInput = findReactSelectInput(dialog) || findReactSelectInput(document);
  if (directReactInput) {
    directReactInput.click?.();
    directReactInput.focus?.();
    await openReactSelectMenu(directReactInput);
    await typeIntoReactSelectInputDeterministic(directReactInput, query);
    sendArrowDown(directReactInput);
    await delay(200);
    if (String(directReactInput.value || "").length > 0) {
      return true;
    }
  }

  const candidates = languageSearchCandidates(dialog, languageControl);
  if (document.activeElement instanceof Element && !candidates.includes(document.activeElement)) {
    candidates.unshift(document.activeElement);
  }
  for (const candidate of candidates) {
    const isTextEntry =
      candidate.matches("input:not([type='hidden']), textarea, [contenteditable='true'], [role='textbox']") ||
      normalizeComparable(candidate.getAttribute("role")).includes("combobox");
    if (!isTextEntry) continue;
    if (candidate.matches("input:not([type='hidden']), textarea, [contenteditable='true'], [role='textbox']")) {
      if (isReactSelectInput(candidate)) {
        candidate.click?.();
        candidate.focus?.();
        await openReactSelectMenu(candidate);
        await typeIntoReactSelectInputDeterministic(candidate, query);
        sendArrowDown(candidate);
        await delay(200);
        if (String(candidate.value || "").length > 0) {
          return true;
        }
        continue;
      }
      if (setInputValueWithoutBlur(candidate, query) && String(candidate.value || "").length > 0) return true;
      if (setFieldValue(candidate, query) && String(candidate.value || "").length > 0) return true;
    } else {
      candidate.click?.();
      dispatchUserClick(candidate);
      emitTypingToElement(candidate, query);
      candidate.dispatchEvent(new Event("change", { bubbles: true }));
      await delay(120);
      if (String(candidate.value || "").length > 0) return true;
    }
  }

  const active = document.activeElement instanceof Element ? document.activeElement : null;
  if (active) {
    active.focus?.();
    emitTypingToElement(active, query);
    active.dispatchEvent(new Event("change", { bubbles: true }));
    return String(active.value || "").length > 0;
  }
  return false;
}




async function forceOpenReactSelectMenu(input, control) {
  const menuVisible = () =>
    visibleElements("[role='option'], [role='listbox'], [class*='menu'], [class*='Menu']", document).length > 0 ||
    (input && normalizeComparable(input.getAttribute("aria-expanded")) === "true");

  // Approach 1: Use React fiber onMouseDown handler on control
  if (control) {
    invokeReactHandler(control, "onMouseDown");
    control.focus?.();
    dispatchUserClick(control);
    control.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, view: window, button: 0 }));
    control.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, view: window, button: 0 }));
    control.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window, button: 0 }));
    await delay(100);
    if (menuVisible()) return true;
  }

  // Approach 2: Focus input and send ArrowDown
  if (input instanceof Element) {
    input.focus?.();
    sendArrowDown(input);
    await delay(120);
    if (menuVisible()) return true;

    // Approach 3: Send Space to open
    sendKeyPress(input, " ", "Space");
    await delay(120);
    if (menuVisible()) return true;

    // Approach 4: Click the input directly
    dispatchPointClick(input);
    await delay(100);
    if (menuVisible()) return true;
  }

  // Approach 5: Click on dropdown indicator if present
  if (control) {
    const indicator = control.querySelector("[class*='indicator'], [class*='Indicator'], svg, [aria-haspopup='listbox']");
    if (indicator instanceof Element) {
      dispatchPointClick(indicator);
      await delay(120);
      if (menuVisible()) return true;
    }
  }

  return menuVisible();
}
function isAddTranslationSelectionReady(dialog) {
  const addButton = findExactButtonByText("Add", dialog) || findButtonByText("Add", dialog);
  return Boolean(addButton && !isDisabledAction(addButton));
}

function getAddTranslationReactSelectNodes(dialog) {
  const control =
    dialog.querySelector("[class*='docent-select__control'], [class*='select__control'], [class*='Select__control']") || null;
  const input =
    dialog.querySelector("input[id^='react-select-'][id$='-input'], input.docent-select__input, input[class*='select__input'], input[role='combobox']") ||
    null;
  return { control, input };
}

async function selectLanguageViaDialogReactSelect(dialog, languageLabel) {
  const query = languageSearchQuery(languageLabel);
  if (await mainWorldSelectAddTranslationLanguage(languageLabel, dialog)) return true;
  const { control, input } = getAddTranslationReactSelectNodes(dialog);
  if (!control && !input) return false;

  if (await selectLanguageViaReactFiber(dialog, languageLabel)) return true;

  await openMenuSimple(findReactSelectInput(dialog), findAddTranslationLanguageControl(dialog));  // Replaced openAddTranslationPickerDeterministic
  await openMenuSimple(input || control, control || input);  // Replaced openAddTranslationLanguagePicker

  const pickerInput = input || findReactSelectInput(dialog) || findReactSelectInput(document);
  if (pickerInput) {
    await clickElementCenterWithTrustedFallback(pickerInput);
    pickerInput.focus?.();
    setReactControlledInputValue(pickerInput, "");

    // This is the observed manual CMS flow: focus the React Select input,
    // type "Spanish", then press Enter on the first filtered result.
    if (await trustedTypeText(query)) {
      await delay(260);
      if (await selectOptionByText(languageLabel, dialog)) return true;
      if (await selectOptionByText(languageLabel)) return true;
      await trustedPressKey("Enter");
      await delay(320);
      if (isAddTranslationSelectionReady(dialog)) return true;
    }

    sendKeyPress(pickerInput, "ArrowDown", "ArrowDown");
    await delay(80);
    typeIntoReactSelectInput(pickerInput, query);
    await delay(180);
    sendKeyPress(pickerInput, "Enter", "Enter");
    await delay(240);
    if (isAddTranslationSelectionReady(dialog)) return true;
    if (await trustedInsertText(query)) {
      await delay(180);
      await trustedPressKey("Enter");
      await delay(240);
      if (isAddTranslationSelectionReady(dialog)) return true;
    }
    await delay(120);
    if (await selectOptionByText(languageLabel, dialog)) return true;
    if (await selectOptionByText(query, dialog)) return true;
    if (await selectOptionByText(languageLabel)) return true;
    if (await selectOptionByText(query)) return true;
    await typeIntoReactSelectInputDeterministic(pickerInput, query);
    await delay(180);
    await trustedPressKey("Enter");
    await delay(240);
    if (isAddTranslationSelectionReady(dialog)) return true;
  }

  if (await selectOptionByText(languageLabel, dialog)) return true;
  if (await selectOptionByText(query, dialog)) return true;
  if (await selectOptionByText(languageLabel)) return true;
  if (await selectOptionByText(query)) return true;

  if (pickerInput) {
    pickerInput.focus?.();
    await trustedPressKey("ArrowDown");
    await delay(120);
    await trustedPressKey("Enter");
    await delay(220);
  }
  return isAddTranslationSelectionReady(dialog);
}

async function searchAndSelectLanguageInAddTranslation(dialog, languageLabel) {
  const addTranslationDialog = findAddTranslationPanelFromHeading() || dialog;
  const query = languageSearchQuery(languageLabel);
  lastMainWorldLanguageSelectionDebug = "attempting";
  if (await mainWorldSelectAddTranslationLanguage(languageLabel, addTranslationDialog)) return true;

  const languageControl = findLanguageFieldInAddTranslationDialog(addTranslationDialog);
  if (!languageControl) {
    lastMainWorldLanguageSelectionDebug = lastMainWorldLanguageSelectionDebug || "not attempted: no language control";
    return false;
  }

  if (await selectLanguageViaReactFiber(addTranslationDialog, languageLabel)) return true;

  if (languageControl.tagName === "SELECT") {
    const option = Array.from(languageControl.options).find((candidate) =>
      normalizeComparable(candidate.textContent).includes(normalizeComparable(languageLabel))
    );
    if (!option) return false;
    languageControl.value = option.value;
    languageControl.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }

  languageControl.click?.();
  await trustedClickElementCenter(languageControl);
  await trustedClickElementRightSide(languageControl);
  dispatchUserClick(languageControl);
  await delay(250);
  if (await selectLanguageViaReactFiber(addTranslationDialog, languageLabel)) return true;
  if (await selectLanguageViaDialogReactSelect(addTranslationDialog, languageLabel)) return true;
  await openMenuSimple(languageControl || findReactSelectInput(addTranslationDialog), languageControl);  // Replaced openAddTranslationLanguagePicker
  if (await selectLanguageViaDialogReactSelect(addTranslationDialog, languageLabel)) return true;
  await typeIntoFocusedLanguageInput(languageLabel);
  await delay(250);
  if (isAddTranslationSelectionReady(addTranslationDialog)) return true;

  await typeLanguageInAddTranslation(addTranslationDialog, languageControl, languageLabel);
  await delay(320);
  if (await selectLanguageViaDialogReactSelect(addTranslationDialog, languageLabel)) return true;
  if (isAddTranslationSelectionReady(addTranslationDialog)) return true;

  const selectedInDialog = await selectOptionByText(languageLabel, addTranslationDialog);
  if (selectedInDialog) return true;
  if (await selectOptionByText(query, addTranslationDialog)) return true;
  const selectedInDocument = await selectOptionByText(languageLabel);
  if (selectedInDocument) return true;
  if (await selectOptionByText(query)) return true;

  const focusCandidate = findReactSelectInput(addTranslationDialog) || languageControl;
  if (focusCandidate) {
    focusCandidate.focus?.();
    await trustedPressKey("ArrowDown");
    await delay(140);
    await trustedPressKey("Enter");
    await delay(260);
    if (isAddTranslationSelectionReady(addTranslationDialog)) return true;
  }
  return false;
}

function readSelectedLanguageLabel(root = findModalRoot()) {
  const select = visibleElements("select", root).find((candidate) => {
    const selectedText = candidate.options && candidate.selectedIndex >= 0 ? candidate.options[candidate.selectedIndex].textContent : "";
    return normalizeComparable(selectedText).includes("english") || (selectedText && candidate.offsetParent !== null);
  });
  if (select && select.selectedIndex >= 0) return normalizeSpaces(select.options[select.selectedIndex].textContent);

  const languageControl = findLanguageDropdownControl(root);
  if (languageControl) return controlText(languageControl);

  const controls = visibleElements("input:not([type='hidden']), [role='combobox'], [role='textbox'], button, [aria-haspopup='listbox'], [aria-haspopup='menu']", root)
    .map((element) => ({
      element,
      text: controlText(element),
    }))
    .filter(({ text }) => {
      if (!normalizeComparable(text) || text.length > 140) return false;
      return isKnownLanguageText(text);
    })
    .filter(({ element }) => {
      const role = normalizeComparable(element.getAttribute("role"));
      if (role.includes("option") || role.includes("menuitem")) return false;
      const menuAncestor = element.closest("[role='listbox'], [role='menu'], [class*='menu'], [class*='Menu']");
      return !menuAncestor || element.matches("[role='combobox'], [aria-haspopup='listbox'], [aria-haspopup='menu']");
    })
    .sort((a, b) => a.text.length - b.text.length);

  return controls[0] ? controls[0].text : "";
}

async function waitForActiveLanguage(item, timeoutMs = 8000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const selected = readSelectedLanguageLabel(findEditAudioRoot());
    if (languageLabelMatches(selected, item.languageLabel)) return true;
    await delay(250);
  }
  return false;
}

async function waitForAddTranslationDialogClosed(timeoutMs = 8000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    if (!findAddTranslationPanelFromHeading()) return true;
    await delay(250);
  }
  return false;
}

function assertTargetLanguageActive(item) {
  const selected = readSelectedLanguageLabel(findEditAudioRoot());
  if (!languageLabelMatches(selected, item.languageLabel)) {
    throw statusError(
      "languageSelectionFailed",
      `Expected ${item.languageLabel} to be active, but the CMS is showing ${selected || "an unknown language"}. No fields were changed.`
    );
  }
}

async function closeAddTranslationDialogIfOpen() {
  const dialog = findAddTranslationPanelFromHeading();
  if (!dialog) return false;
  const cancel = findExactButtonByText("Cancel", dialog) || findButtonByText("Cancel", dialog);
  if (!cancel) return false;
  cancel.click?.();
  dispatchUserClick(cancel);
  await waitForAddTranslationDialogClosed(3000);
  await delay(300);
  return true;
}

async function selectExistingLanguageFromEditDropdown(item, dropdown) {
  const root = findEditAudioRoot();
  const control = dropdown || findLanguageDropdownControl(root) || findEnglishLanguageControl(root) || findLanguageDropdownControl(document);
  if (!control) return false;

  await openLanguageDropdownControl(control);
  await delay(350);

  for (const alias of languageLabelAliases(item.languageLabel)) {
    const option = findExistingLanguageMenuOption(alias);
    if (!option) continue;
    clickElementCenter(option);
    dispatchUserClick(option);
    await delay(700);
    if (await waitForActiveLanguage(item, 2500)) {
      await closeAddTranslationDialogIfOpen();
      return true;
    }
  }

  return false;
}

async function ensureLanguageSelected(item) {
  let root = findEditAudioRoot();
  if (languageLabelMatches(readSelectedLanguageLabel(root), item.languageLabel)) {
    await closeAddTranslationDialogIfOpen();
    return "selected";
  }
  if (findAddTranslationPanelFromHeading()) {
    await closeAddTranslationDialogIfOpen();
    root = findEditAudioRoot();
    if (languageLabelMatches(readSelectedLanguageLabel(root), item.languageLabel)) return "selected";
  }

  const dropdown = await clickLanguageDropdown(root);

  if (dropdown.tagName === "SELECT") {
    const options = Array.from(dropdown.options);
    const option = options.find((candidate) => languageLabelAliases(item.languageLabel).some((alias) => languageLabelMatches(candidate.textContent, alias)));
    if (option) {
      dropdown.value = option.value;
      dropdown.dispatchEvent(new Event("change", { bubbles: true }));
      await delay(500);
      if (await waitForActiveLanguage(item)) return "selected";
      assertTargetLanguageActive(item);
    }
  } else {
    if (await selectExistingLanguageFromEditDropdown(item, dropdown)) {
      return "selected";
    }
  }

  // Open the language dropdown and wait for the menu to render.
  // The "Add Translation" item may render in a portal outside the modal,
  // so we search the entire document after giving the menu time to appear.
  await openLanguageDropdownControl(dropdown);
  await delay(600);

  let addTranslation = findAddTranslationMenuItem();
  if (!addTranslation) {
    // Retry: close and re-open the dropdown to force a fresh render.
    await closeAddTranslationDialogIfOpen();
    await delay(200);
    await openLanguageDropdownControl(dropdown);
    await delay(600);
    addTranslation = findAddTranslationMenuItem();
  }
  if (!addTranslation) {
    const debugInfo = buildLanguageSelectionDebugInfo(dropdown, findModalRoot());
    console.warn("Bloomberg Audio Assistant language selection debug:", debugInfo);
    throw statusError("languageSelectionFailed", `Could not find Add Translation for ${item.languageLabel}. Debug: ${debugInfo}`);
  }

  // Click the "Add Translation" menu item and wait for the dialog.
  const actionTarget = addTranslation.closest("[role='option'], [role='menuitem'], li") || addTranslation;
  clickElementCenter(actionTarget);
  dispatchUserClick(actionTarget);
  await delay(800);

  const dialog = await waitForAddTranslationDialog(3000);
  if (!dialog) {
    const debugInfo = buildLanguageSelectionDebugInfo(dropdown, findModalRoot());
    console.warn("Bloomberg Audio Assistant Add Translation dialog did not open:", debugInfo);
    throw statusError("languageSelectionFailed", `Could not open Add Translation dialog for ${item.languageLabel}. Debug: ${debugInfo}`);
  }

  // Select the target language in the Add Translation dialog.
  const selected = await selectLanguageInAddTranslationDialog(dialog, item.languageLabel);
  if (!selected) {
    const debugInfo = `${buildLanguageSelectionDebugInfo(dropdown, dialog)}; ${buildPickerStateDebugInfo()}`;
    console.warn("Bloomberg Audio Assistant add-translation selection debug:", debugInfo);
    throw statusError("languageSelectionFailed", `Could not select ${item.languageLabel} in Add Translation. Debug: ${debugInfo}`);
  }

  let addButton = findExactButtonByText("Add", dialog) || findButtonByText("Add", dialog);
  if (!addButton) throw statusError("languageSelectionFailed", "Could not find the Add button in Add Translation.");
  if (isDisabledAction(addButton)) {
    if (!(await selectAddTranslationOptionByText(item.languageLabel, dialog)) && !(await selectAddTranslationOptionByText(languageSearchQuery(item.languageLabel), dialog))) {
      const debugInfo = `${buildLanguageSelectionDebugInfo(dropdown, dialog)}; ${buildPickerStateDebugInfo()}`;
      throw statusError("languageSelectionFailed", `Selected ${item.languageLabel} did not enable Add Translation. Debug: ${debugInfo}`);
    }
    addButton = findExactButtonByText("Add", dialog) || findButtonByText("Add", dialog);
  }
  if (isDisabledAction(addButton)) {
    const debugInfo = `${buildLanguageSelectionDebugInfo(dropdown, dialog)}; ${buildPickerStateDebugInfo()}`;
    throw statusError("languageSelectionFailed", `Add Translation is still disabled for ${item.languageLabel}. Debug: ${debugInfo}`);
  }
  if (!(await trustedClickElementCenter(addButton))) {
    clickElementCenter(addButton);
    dispatchUserClick(addButton);
  }
  const closed = await waitForAddTranslationDialogClosed(8000);
  if (!closed) {
    const debugInfo = `${buildLanguageSelectionDebugInfo(dropdown, dialog)}; ${buildPickerStateDebugInfo()}`;
    throw statusError("languageSelectionFailed", `Add Translation did not close after clicking Add for ${item.languageLabel}. Debug: ${debugInfo}`);
  }
  await delay(1800);
  if (!(await waitForActiveLanguage(item))) {
    assertTargetLanguageActive(item);
  }
  return "added";
}

// Finds the "Add Translation" item in the language dropdown menu.
// Searches the entire document because React Select may render the menu
// in a portal outside the modal DOM tree. Uses flexible text matching.
function findAddTranslationMenuItem() {
  const target = normalizeComparable("Add Translation");
  const candidates = visibleElements("button, [role='button'], a, [role='option'], [role='menuitem'], li, div, span", document)
    .filter((element) => {
      const text = normalizeComparable(controlText(element));
      return text === target || text.includes(target);
    })
    .sort((a, b) => controlText(a).length - controlText(b).length);
  return candidates[0] || null;
}

// Selects a language in the Add Translation dialog using the main-world
// React fiber approach first, then falling back to typing into the input.
async function selectLanguageInAddTranslationDialog(dialog, languageLabel) {
  const query = languageSearchQuery(languageLabel);

  // Primary: use the service worker to run selector code in the MAIN world.
  if (await mainWorldSelectAddTranslationLanguage(languageLabel, dialog)) return waitForAddTranslationSelectionReady(dialog);

  // Fallback: traverse the React fiber tree to find and click the option.
  if (await selectLanguageViaReactFiber(dialog, languageLabel)) return waitForAddTranslationSelectionReady(dialog);

  // Fallback: find the input, type the query, and pick the matching option.
  const input = dialog.querySelector("input[id^='react-select-'][id$='-input'], input.docent-select__input, input[class*='select__input'], input[role='combobox']");
  if (input) {
    input.focus?.();
    input.click?.();
    await delay(200);
    typeIntoReactSelectInput(input, query);
    await delay(400);
    if (await selectAddTranslationOptionByText(languageLabel, dialog)) return true;
    if (await selectAddTranslationOptionByText(query, dialog)) return true;
    sendKeyPress(input, "Enter", "Enter");
    await delay(300);
    if (await waitForAddTranslationSelectionReady(dialog)) return true;
  }

  if (await selectAddTranslationOptionByText(languageLabel, dialog)) return true;
  if (await selectAddTranslationOptionByText(query, dialog)) return true;
  return waitForAddTranslationSelectionReady(dialog);
}

function findLanguageSection(item) {
  const root = findModalRoot();
  const comparableLanguage = normalizeComparable(item.languageLabel);
  const sections = visibleElements("section, form, [class*='tab'], [class*='panel'], [class*='language']", root)
    .filter((element) => normalizeComparable(element.textContent).includes(comparableLanguage))
    .sort((a, b) => a.textContent.length - b.textContent.length);
  return sections[0] || root;
}

async function attachAudioFile(item, record) {
  assertTargetLanguageActive(item);
  const root = findLanguageSection(item);

  const input = findAudioUploadInput(root);
  if (!input) {
    throw statusError(
      "uploadFailed",
      "Could not find a mounted audio file input. The automation will not click Add Audio File because that opens Chrome's file picker."
    );
  }

  const file =
    record.blob instanceof File && record.blob.name === record.fileName
      ? record.blob
      : new File([record.blob], record.fileName, { type: record.contentType || "audio/wav" });
  const transfer = new DataTransfer();
  transfer.items.add(file);
  input.files = transfer.files;
  input.dispatchEvent(new Event("input", { bubbles: true, composed: true }));
  input.dispatchEvent(new Event("change", { bubbles: true, composed: true }));

  const dropTarget =
    input.closest("[class*='drop'], [class*='Drop'], [class*='upload'], [class*='Upload'], section, form, div") || root;
  ["dragenter", "dragover", "drop"].forEach((type) => {
    dropTarget.dispatchEvent(new DragEvent(type, { bubbles: true, composed: true, cancelable: true, dataTransfer: transfer }));
  });

  await waitForUploadComplete(root, record.fileName);
}

function findAudioUploadInput(root) {
  const inputs = Array.from(root.querySelectorAll("input[type='file']"));
  const globalInputs = Array.from(document.querySelectorAll("input[type='file']"));
  const candidates = [...inputs, ...globalInputs].filter((input, index, list) => {
    if (!input || input.disabled) return false;
    return list.indexOf(input) === index;
  });
  return (
    candidates.find((input) => normalizeComparable(input.accept || "").includes("audio")) ||
    candidates.find((input) => !input.accept || /\.(wav|mp3|m4a|aac|ogg|flac)/i.test(input.accept || "")) ||
    candidates[0] ||
    null
  );
}

async function waitForUploadComplete(root, fileName) {
  const normalizedFileName = normalizeComparable(fileName);
  const findUploadErrorNotice = () =>
    visibleElements("[class*='error'], [class*='Error'], [role='alert']", root).find((element) => {
      const text = normalizeComparable(element.textContent);
      if (!text) return false;
      const isUploadScoped =
        text.includes("upload") ||
        text.includes("audio") ||
        text.includes("file") ||
        text.includes(normalizedFileName);
      const isFailure =
        text.includes("failed") ||
        text.includes("error") ||
        text.includes("unsupported") ||
        text.includes("too large");
      return isUploadScoped && isFailure;
    });

  const startedAt = Date.now();
  while (Date.now() - startedAt < BCA_UPLOAD_TIMEOUT_MS) {
    const text = normalizeComparable(root.textContent);
    const hasProgress = /\b\d{1,3}%\b/.test(root.textContent || "") || text.includes("uploading");
    const hasAudio = Boolean(root.querySelector("audio")) || text.includes(normalizeComparable(fileName));
    const uploadErrorNotice = findUploadErrorNotice();
    // Success signal should win over stale, unrelated error labels.
    if (hasAudio && !hasProgress) return;
    if (uploadErrorNotice) {
      throw statusError("uploadFailed", normalizeSpaces(uploadErrorNotice.textContent) || "CMS reported an upload error.");
    }
    await delay(1000);
  }
  throw statusError("uploadFailed", "Timed out waiting for the audio upload to finish.");
}

async function setAccessibilityLanguage(item) {
  assertTargetLanguageActive(item);
  const root = findModalRoot();
  const tab = findButtonByText("Accessibility", root);
  if (!tab) return false;
  tab.click();
  await delay(400);
  const accessibilityRoot = findModalRoot();
  const field = findFieldByLabel("Language", accessibilityRoot);
  if (!field) return false;

  if (field.tagName === "SELECT") {
    const option = Array.from(field.options).find((candidate) =>
      normalizeComparable(candidate.textContent).includes(normalizeComparable(item.accessibilityLabel))
    );
    if (!option) return false;
    field.value = option.value;
    field.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }

  field.click();
  await delay(300);
  return selectOptionByText(item.accessibilityLabel);
}

async function fillTranslationFields(item, transcript) {
  assertTargetLanguageActive(item);
  const root = findLanguageSection(item);
  const titleField = findFieldByLabel("Title", root, { preferSingleLine: true });
  if (!titleField) throw new Error("Could not find the translation Title field.");
  setFieldValue(titleField, item.translatedTitle || translatedTitleFromCatalogTitle(item.sourceTitle || item.rawTitle || "", transcript, item.languageKey));

  // Try the language section first, then fall back to broader roots.
  const editRoot = findEditAudioRoot();
  const modalRoot = findModalRoot();
  const scopes = [root, editRoot, modalRoot, document];
  let transcriptSet = false;

  for (const scope of scopes) {
    const field = findTranscriptField(scope) || findFieldByLabel("Transcript", scope, { preferMultiline: true });
    if (field) {
      transcriptSet = setTranscriptField(scope, transcript);
      if (transcriptSet) break;
    }
  }

  if (!transcriptSet) {
    throw new Error("Could not find the translation Transcript field.");
  }

  await setAccessibilityLanguage(item);
}

function findSaveButton() {
  const root = findModalRoot();
  return findExactButtonByText("Save", root) || findButtonByText("Save", root);
}

function findErrorNotice() {
  return visibleElements("[class*='error'], [class*='Error'], [role='alert']").find((element) => {
    const text = normalizeComparable(element.textContent);
    return text.includes("please correct") || text.includes("error") || text.includes("failed") || text.includes("required");
  });
}

async function saveAndConfirm(item) {
  await ensureCmsWindowFocused();
  const saveButton = findSaveButton();
  if (!saveButton) throw statusError("saveFailed", "Could not find the CMS Save button.");
  saveButton.click();

  const startedAt = Date.now();
  while (Date.now() - startedAt < 45000) {
    const pageText = normalizeComparable(document.body.textContent);
    if (pageText.includes("changes saved") || pageText.includes("saved successfully")) return;
    const errorNotice = findErrorNotice();
    if (errorNotice) {
      throw statusError("saveFailed", normalizeSpaces(errorNotice.textContent) || "CMS validation error.");
    }
    if (isBloombergAudioCatalogPage()) return;
    await delay(BCA_SAVE_CHECK_DELAY_MS);
  }
  throw statusError("saveFailed", "Could not confirm that the CMS saved the item.");
}

async function failAndPause(message, item, status) {
  const state = await readState();
  state.status = "paused";
  state.phase = status || "failed";
  state.current = {
    itemId: item ? item.itemId : "",
    label: itemLabel(item),
    error: message,
  };
  state.counts.failed += 1;
  if (item) {
    const storedItem = state.workItems.find((candidate) => candidate.itemId === item.itemId && candidate.languageKey === item.languageKey);
    if (storedItem) {
      storedItem.status = status || "failed";
      storedItem.error = message;
    }
  }
  pushLog(state, "error", message, item, status || "failed");
  await writeState(state);
  alert(`Bloomberg Audio Assistant (${BCA_BUILD_LABEL}) paused: ${message}`);
}

async function updateCurrentItemWithItemId(state, item, itemId) {
  if (!itemId || item.itemId === itemId) return item;
  const config = await loadConfig();
  const audioCacheKey = `${itemId}:${config.sourceLanguageCode || "en-US"}`;
  const editUrl = buildAudioEditUrl({ itemId, editUrl: window.location.href }, config);
  const record = {
    ...item,
    itemId,
    editUrl,
    audioCacheKey,
    rowText: normalizeSpaces([item.fileName, item.includedIn, item.sourceTitle || item.title].filter(Boolean).join(" | ")),
    lastSeenAt: new Date().toISOString(),
  };
  state.audioItemIndex = mergeAudioItemIndex(state.audioItemIndex || {}, [record], config);

  state.workItems = (state.workItems || []).map((candidate) => {
    const sameRow =
      candidate === item ||
      (candidate.rowKey && item.rowKey && candidate.rowKey === item.rowKey) ||
      (candidate.stopNumber === item.stopNumber && normalizeComparable(candidate.fileName) === normalizeComparable(item.fileName));
    return sameRow
      ? {
          ...candidate,
          itemId,
          audioCacheKey,
          editUrl,
        }
      : candidate;
  });
  state.items = (state.items || []).map((candidate) => {
    const sameRow =
      (candidate.rowKey && item.rowKey && candidate.rowKey === item.rowKey) ||
      (candidate.stopNumber === item.stopNumber && normalizeComparable(candidate.fileName) === normalizeComparable(item.fileName));
    return sameRow
      ? {
          ...candidate,
          itemId,
          editUrl,
        }
      : candidate;
  });
  state.pendingOpenItem = null;
  await writeState(state);
  return state.workItems[state.index] || { ...item, itemId, audioCacheKey };
}

function pendingOpenMatchesItem(state, item, itemId) {
  const pending = state && state.pendingOpenItem;
  if (!pending || !item || !itemId) return false;
  if (pending.index !== state.index) return false;
  if (pending.languageKey !== item.languageKey) return false;
  if (pending.rowKey && item.rowKey && pending.rowKey !== item.rowKey) return false;
  if (pending.stopNumber && item.stopNumber && pending.stopNumber !== item.stopNumber) return false;
  if (pending.fileName && item.fileName && normalizeComparable(pending.fileName) !== normalizeComparable(item.fileName)) return false;
  return true;
}

function currentEditPageMatchesItem(state, item, currentItemId) {
  if (!item || !currentItemId) return false;
  if (item.itemId) return currentItemId === item.itemId;
  return pendingOpenMatchesItem(state, item, currentItemId);
}

async function markTaskComplete(item, status, message) {
  const state = await readState();
  const storedItem = state.workItems[state.index];
  if (storedItem) storedItem.status = status;
  if (status === "processed") state.counts.processed += 1;
  if (status === "alreadyComplete") state.counts.alreadyComplete += 1;
  if (status === "skipped") state.counts.skipped += 1;

  if ((status === "processed" || status === "alreadyComplete") && item && item.rowKey && item.languageKey) {
    const completedTasks = state.completedTasks && typeof state.completedTasks === "object" ? state.completedTasks : {};
    const completedForRow = Array.isArray(completedTasks[item.rowKey]) ? [...completedTasks[item.rowKey]] : [];
    if (!completedForRow.includes(item.languageKey)) {
      completedForRow.push(item.languageKey);
    }
    state.completedTasks = { ...completedTasks, [item.rowKey]: completedForRow };
  }

  state.index += 1;
  state.phase = "ready";
  state.current = null;
  state.pendingOpenItem = null;
  pushLog(state, "success", message, item, status);
  await writeState(state);
  setTimeout(() => readState().then(maybeAutoContinueRun).catch((error) => failAndPause(error.message, item, error.status)), 500);
}

async function completeRun(state) {
  state.status = "complete";
  state.phase = "complete";
  state.current = null;
  pushLog(state, "success", "Run complete.");
  await writeState(state);
  if (!isBloombergAudioCatalogPage()) {
    window.location.href = (await loadConfig()).cmsCatalogUrl;
  }
}

function modalAppearsCompleteForLanguage(item, transcript) {
  assertTargetLanguageActive(item);
  const root = findLanguageSection(item);
  const editRoot = findEditAudioRoot();
  const modalRoot = findModalRoot();
  const scopes = [root, editRoot, modalRoot, document];
  const text = normalizeComparable(root.textContent);
  const title = findFieldByLabel("Title", root, { preferSingleLine: true });

  // Search for transcript field across all scopes
  let transcriptField = null;
  for (const scope of scopes) {
    transcriptField = findTranscriptField(scope) || findFieldByLabel("Transcript", scope, { preferMultiline: true });
    if (transcriptField) break;
  }

  const hasAudio = Boolean(root.querySelector("audio")) || !text.includes("add audio file");
  const titleValue = normalizeSpaces(title ? title.value || title.textContent : "");
  const expectedTitle = normalizeComparable(item.translatedTitle || item.transcriptTitle || item.title);
  const transcriptValue = normalizeSpaces(transcriptField ? transcriptField.value || transcriptField.textContent : "");
  const expectedTranscript = transcriptToPlainText(transcript);
  return (
    hasAudio &&
    titleValue &&
    (!expectedTitle || normalizeComparable(titleValue).includes(expectedTitle)) &&
    transcriptTextAppearsComplete(transcriptValue, expectedTranscript)
  );
}

async function processCurrentEditPage(state) {
  const config = await loadConfig();
  const transcripts = await loadTranscripts();
  let item = state.workItems[state.index];
  if (!item) {
    await completeRun(state);
    return;
  }

  if (!workItemLanguageIsEnabled(item, config)) {
    await markTaskComplete(item, "skipped", "Language disabled in config; skipped.");
    return;
  }

  const currentItemId = extractItemId(window.location.href);
  if (!currentEditPageMatchesItem(state, item, currentItemId)) {
    throw statusError(
      "routeMismatch",
      `Refusing to process ${itemLabel(item)} on CMS audio ${currentItemId || "unknown"}. Returning to the catalog to reopen the correct row.`
    );
  }
  item = await updateCurrentItemWithItemId(state, item, currentItemId);

  state.phase = "processing";
  state.current = { itemId: item.itemId || currentItemId, label: itemLabel(item) };
  pushLog(state, "info", "Processing translation.", item);
  await writeState(state);

  try {
    await ensureLanguageSelected(item);
    await delay(500);
    assertTargetLanguageActive(item);
    const transcript = transcriptFor(transcripts, item.languageKey, item.stopNumber);
    if (!transcript) {
      throw statusError("missingTranscript", "Missing transcript section.");
    }
    if (modalAppearsCompleteForLanguage(item, transcript)) {
      await markTaskComplete(item, "alreadyComplete", "Already complete on edit page; skipped.");
      return;
    }

    const record = await ensureCachedAudio(item, state);
    await attachAudioFile(item, record);
    await fillTranslationFields(item, transcript);
    assertTargetLanguageActive(item);

    state = await readState();
    state.phase = "saving";
    state.current = { itemId: item.itemId, label: itemLabel(item) };
    pushLog(state, "info", "Saving CMS item.", item);
    await writeState(state);

    await saveAndConfirm(item, config);
    await markTaskComplete(item, "processed", "Saved translation.");
  } catch (error) {
    await failAndPause(error.message, item, error.status || "failed");
  }
}

async function resolveRouteForRunItem(state, item, config) {
  if (item.editUrl || item.itemId) return item;

  let resolved = resolveAudioItemFromIndex(state, item, config);
  if (resolved) {
    applyResolvedAudioRouteToState(state, item, resolved, config);
    await writeState(state);
    return resolved;
  }

  if (!isBloombergAudioCatalogPage()) {
    window.location.href = config.cmsCatalogUrl;
    return null;
  }

  const repairKey = routeResolutionKey(item);
  const attempts = state.routeRepairAttempts || {};
  if (!attempts[repairKey]) {
    state.routeRepairAttempts = { ...attempts, [repairKey]: 1 };
    await writeState(state);
    state = await repairAudioItemIndex({ state, silent: true });
    const latestItem = (state.workItems || [])[state.index] || item;
    resolved = resolveAudioItemFromIndex(state, latestItem, config) || (latestItem.editUrl || latestItem.itemId ? latestItem : null);
    if (resolved) {
      applyResolvedAudioRouteToState(state, latestItem, resolved, config);
      await writeState(state);
      return resolved;
    }
  }

  return null;
}

function routeFailureMessage(item) {
  const terms = catalogItemSearchTerms(item).slice(0, 8).join(", ") || "none";
  const rows = isBloombergAudioCatalogPage() ? visibleCatalogRowSamples(8).join(" || ") : "not on catalog page";
  return `Could not resolve CMS item for stop ${item.stopNumber || "unknown"} (${item.fileName || item.title || "unknown item"}). Search terms: ${terms}. Visible rows: ${rows}. Run Repair Item Index from the panel, then resume.`;
}

async function continueRun() {
  assertExtensionBuildCurrent();
  let state = await readState();
  const config = await loadConfig();
  if (state.status !== "running") return;
  if (state.index >= state.workItems.length) {
    await completeRun(state);
    return;
  }

  const item = state.workItems[state.index];
  if (!item) {
    await completeRun(state);
    return;
  }

  if (!workItemLanguageIsEnabled(item, config)) {
    await markTaskComplete(item, "skipped", "Language disabled in config; skipped.");
    return;
  }

  const currentItemId = extractItemId(window.location.href);
  if (isBloombergAudioEditPage()) {
    if (currentEditPageMatchesItem(state, item, currentItemId)) {
      await processCurrentEditPage(state);
      return;
    }
    state.phase = "ready";
    state.pendingOpenItem = null;
    pushLog(state, "warn", `Edit page is for CMS audio ${currentItemId || "unknown"}, not ${itemLabel(item)}; returning to catalog.`, item, "routeMismatch");
    await writeState(state);
    window.location.href = config.cmsCatalogUrl;
    return;
  }

  const routedItem = await resolveRouteForRunItem(state, item, config);
  if (routedItem && routedItem.editUrl) {
    window.location.href = routedItem.editUrl;
    return;
  }

  if (routedItem && routedItem.itemId) {
    window.location.href = buildAudioEditUrl(routedItem, config);
    return;
  }

  if (!isBloombergAudioCatalogPage()) {
    window.location.href = config.cmsCatalogUrl;
    return;
  }

  state.pendingOpenItem = {
    index: state.index,
    rowKey: item.rowKey || "",
    stopNumber: item.stopNumber || null,
    fileName: item.fileName || "",
    languageKey: item.languageKey || "",
    openedAt: new Date().toISOString(),
  };
  await writeState(state);

  const opened = await openCatalogRow(item);
  if (!opened) {
    state.pendingOpenItem = null;
    await writeState(state);
    throw new Error(routeFailureMessage(item));
  }
}

async function maybeAutoContinueRun(state) {
  if (!state || state.status !== "running") return;
  if (["processing", "saving"].includes(state.phase)) return;
  if (bcaContinueInFlight) return;
  const now = Date.now();
  if (now - bcaLastAutoContinueAt < 1200) return;

  bcaLastAutoContinueAt = now;
  bcaContinueInFlight = true;
  try {
    await continueRun();
  } catch (error) {
    const latest = await readState().catch(() => state);
    await failAndPause(error.message, latest.workItems && latest.workItems[latest.index], error.status);
  } finally {
    bcaContinueInFlight = false;
  }
}

async function startRun({ resume = false } = {}) {
  assertExtensionBuildCurrent();
  let state = await readState();
  const hasPriorRun = Boolean(state.workItems && state.workItems.length);

  if (!resume || !hasPriorRun) {
    state = await scanCatalog();
  } else if (isBloombergAudioCatalogPage()) {
    // Reconcile against the live catalog on resume so any languages that
    // finished saving (or were edited externally) before the run was halted
    // are not redone. scanCatalog filters out languages already present in
    // each row's Translations column.
    const previousCounts = state.counts || {};
    const previousLog = Array.isArray(state.log) ? state.log : [];
    state = await scanCatalog();
    const rescanCounts = state.counts || {};
    state.counts = {
      ...rescanCounts,
      processed: (previousCounts.processed || 0) + (rescanCounts.processed || 0),
      alreadyComplete: (previousCounts.alreadyComplete || 0) + (rescanCounts.alreadyComplete || 0),
      skipped: (previousCounts.skipped || 0) + (rescanCounts.skipped || 0),
      failed: (previousCounts.failed || 0) + (rescanCounts.failed || 0),
      missingTranscript: (previousCounts.missingTranscript || 0) + (rescanCounts.missingTranscript || 0),
    };
    state.log = [...previousLog, ...(Array.isArray(state.log) ? state.log : [])].slice(-250);
    pushLog(state, "info", "Resume reconciled against the live catalog; remaining work refreshed.");
    await writeState(state);
  }

  if (!state.workItems.length) {
    state.status = "complete";
    state.phase = "complete";
    pushLog(state, "success", "No ready translation tasks found.");
    await writeState(state);
    return state;
  }

  const remaining = state.workItems.length - (state.index || 0);
  const confirmed = window.confirm(
    `This will add and save translated audio entries in the Bloomberg Connects CMS for ${remaining} remaining tasks. Continue?`
  );
  if (!confirmed) {
    state.status = "scanned";
    state.phase = "dry-run";
    pushLog(state, "info", "Save run cancelled before CMS changes.");
    await writeState(state);
    return state;
  }

  state.status = "running";
  state.phase = "ready";
  state.index = Number.isInteger(state.index) ? state.index : 0;
  pushLog(state, "info", resume ? "Resumed save run." : "Started save run.");
  await writeState(state);

  setTimeout(() => readState().then(maybeAutoContinueRun).catch((error) => failAndPause(error.message, state.workItems[state.index], error.status)), 250);
  return state;
}

async function stopRun() {
  const state = await readState();
  const audioRun = await readAudioDownloadRun();
  if (audioRun) {
    audioRun.status = "stopped";
    audioRun.updatedAt = new Date().toISOString();
    await storageRemove(BCA_AUDIO_DOWNLOAD_KEY);
    state.phase = "audio-download-stopped";
    pushLog(state, "warn", "Audio download run stopped by user.");
    await writeState(state);
    return state;
  }
  state.status = "stopped";
  state.phase = "stopped";
  pushLog(state, "warn", "Run stopped by user.");
  return writeState(state);
}

async function resetCompletionRecord() {
  const confirmed = window.confirm(
    "Reset the extension's completion record? Future scans will rely entirely on the CMS catalog to detect saved translations."
  );
  if (!confirmed) return null;
  const state = await readState();
  state.completedTasks = {};
  pushLog(state, "warn", "Completion record reset by user.");
  return writeState(state);
}

function getPanelRoot() {
  return document.getElementById(BCA_PANEL_ROOT_ID);
}

function createPanelElement(tagName, className, textContent) {
  const element = document.createElement(tagName);
  if (className) element.className = className;
  if (typeof textContent === "string") element.textContent = textContent;
  return element;
}

function ensureInjectedPanel() {
  if (!window.location.href.startsWith("https://cms.bloombergconnects.org/")) return null;
  const existingRoot = getPanelRoot();
  if (existingRoot && bcaPanelElements) return bcaPanelElements;

  const host = document.createElement("div");
  host.id = BCA_PANEL_ROOT_ID;
  const shadowRoot = host.attachShadow({ mode: "open" });
  const style = document.createElement("style");
  style.textContent = BCA_PANEL_STYLE;
  shadowRoot.appendChild(style);

  const panel = createPanelElement("section", "bca-panel");
  const header = createPanelElement("header", "bca-header");
  const titleWrap = createPanelElement("div");
  const title = createPanelElement("div", "bca-title", "Audio Assistant");
  const phase = createPanelElement("div", "bca-phase", "Idle");
  const build = createPanelElement("div", "bca-build", BCA_BUILD_LABEL);
  titleWrap.append(title, phase, build);
  const collapseButton = createPanelElement("button", "bca-collapse", "Hide");
  collapseButton.type = "button";
  header.append(titleWrap, collapseButton);

  const body = createPanelElement("div", "bca-body");
  const statusList = createPanelElement("div", "bca-status-list");
  const cmsStatus = createPanelElement("div", "bca-status", "Checking CMS page");
  const dataStatus = createPanelElement("div", "bca-status", "Checking transcript data");
  statusList.append(cmsStatus, dataStatus);

  const actionRow = createPanelElement("div", "bca-actions");
  const scanButton = createPanelElement("button", "bca-button", "Scan / Dry Run");
  const startButton = createPanelElement("button", "bca-button", "Start");
  const resumeButton = createPanelElement("button", "bca-button", "Resume");
  const stopButton = createPanelElement("button", "bca-button danger", "Stop");
  const resetCompletionButton = createPanelElement("button", "bca-button", "Reset Completion Record");
  [scanButton, startButton, resumeButton, stopButton, resetCompletionButton].forEach((button) => {
    button.type = "button";
    actionRow.appendChild(button);
  });

  const folderRow = createPanelElement("div", "bca-folder-row");
  const folderButton = createPanelElement("button", "bca-button secondary", "Choose Local Audio Folder");
  folderButton.type = "button";
  const downloadAudioButton = createPanelElement("button", "bca-button secondary", "Download Missing Audio");
  downloadAudioButton.type = "button";
  const repairIndexButton = createPanelElement("button", "bca-button secondary", "Repair Item Index");
  repairIndexButton.type = "button";
  const folderStatus = createPanelElement("div", "bca-current bca-folder-status", "No local audio folder selected.");
  folderRow.append(folderButton, downloadAudioButton, repairIndexButton, folderStatus);

  const runStatus = createPanelElement("div", "bca-run-status", "Idle.");
  const currentItem = createPanelElement("div", "bca-current");
  const counts = createPanelElement("div", "bca-counts");
  const countMap = {};
  [
    ["Stops", "stops"],
    ["Tasks", "tasks"],
    ["Ready", "ready"],
    ["Done", "done"],
    ["Skip", "skip"],
    ["Fail", "fail"],
  ].forEach(([label, key]) => {
    const countCard = createPanelElement("div", "bca-count");
    countCard.append(createPanelElement("span", "bca-count-label", label), createPanelElement("span", "bca-count-value", "0"));
    counts.appendChild(countCard);
    countMap[key] = countCard.querySelector(".bca-count-value");
  });

  const logList = createPanelElement("ol", "bca-log");
  logList.appendChild(createPanelElement("li", "", "No activity yet."));
  body.append(statusList, actionRow, folderRow, runStatus, currentItem, counts, logList);
  panel.append(header, body);
  shadowRoot.appendChild(panel);
  document.documentElement.appendChild(host);

  bcaPanelElements = {
    panel,
    phase,
    cmsStatus,
    dataStatus,
    scanButton,
    startButton,
    resumeButton,
    stopButton,
    resetCompletionButton,
    folderButton,
    downloadAudioButton,
    repairIndexButton,
    folderStatus,
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

  scanButton.addEventListener("click", panelCommand(() => scanCatalog()));
  startButton.addEventListener("click", panelCommand(() => startRun()));
  resumeButton.addEventListener("click", panelCommand(() => startRun({ resume: true })));
  stopButton.addEventListener("click", panelCommand(() => stopRun()));
  resetCompletionButton.addEventListener("click", panelCommand(() => resetCompletionRecord()));
  folderButton.addEventListener("click", panelCommand(() => selectLocalAudioFolder()));
  downloadAudioButton.addEventListener("click", panelCommand(() => downloadMissingAudioToFolder()));
  repairIndexButton.addEventListener("click", panelCommand(() => repairAudioItemIndex()));

  if (!bcaStorageListenerRegistered && chrome.storage && chrome.storage.onChanged) {
    chrome.storage.onChanged.addListener((changes, areaName) => {
      if (areaName === "local" && (changes[BCA_STATE_KEY] || changes[BCA_AUDIO_FOLDER_META_KEY] || changes[BCA_AUDIO_DOWNLOAD_KEY])) {
        refreshInjectedPanel();
      }
    });
    bcaStorageListenerRegistered = true;
  }

  if (!bcaPanelRefreshTimer) {
    bcaPanelRefreshTimer = window.setInterval(refreshInjectedPanel, 1000);
  }

  return bcaPanelElements;
}

function panelCommand(callback) {
  return async () => {
    setPanelButtonsBusy(true);
    try {
      await callback();
    } catch (error) {
      window.alert(error.message);
    } finally {
      setPanelButtonsBusy(false);
      refreshInjectedPanel();
    }
  };
}

function setPanelButtonsBusy(disabled) {
  if (!bcaPanelElements) return;
  bcaPanelElements.scanButton.disabled = disabled;
  bcaPanelElements.startButton.disabled = disabled;
  bcaPanelElements.resumeButton.disabled = disabled;
  bcaPanelElements.stopButton.disabled = disabled;
  if (bcaPanelElements.resetCompletionButton) bcaPanelElements.resetCompletionButton.disabled = disabled;
  bcaPanelElements.folderButton.disabled = disabled;
  bcaPanelElements.downloadAudioButton.disabled = disabled;
  if (bcaPanelElements.repairIndexButton) bcaPanelElements.repairIndexButton.disabled = disabled;
}

function setStatusAppearance(element, text, className) {
  if (!element) return;
  element.textContent = text;
  element.className = `bca-status ${className || ""}`.trim();
}

function renderPanelLog(logList, entries) {
  logList.textContent = "";
  if (!entries.length) {
    logList.appendChild(createPanelElement("li", "", "No activity yet."));
    return;
  }
  entries.forEach((entry) => {
    const item = createPanelElement("li");
    if (entry.itemLabel) {
      const strong = createPanelElement("strong", "", entry.itemLabel);
      item.appendChild(strong);
      item.appendChild(document.createTextNode(": "));
    }
    item.appendChild(document.createTextNode(entry.message || ""));
    logList.appendChild(item);
  });
}

async function refreshInjectedPanel() {
  const elements = ensureInjectedPanel();
  if (!elements) return;

  const state = await readState().catch(() => defaultState());
  const audioRun = await readAudioDownloadRun().catch(() => null);
  const folderData = await storageGet(BCA_AUDIO_FOLDER_META_KEY).catch(() => ({}));
  const folderMeta = bcaAudioFolderMeta || (folderData && folderData[BCA_AUDIO_FOLDER_META_KEY]);
  if (isBloombergAudioCatalogPage()) {
    setStatusAppearance(elements.cmsStatus, "Audios catalog is active", "success");
  } else if (isBloombergAudioEditPage()) {
    setStatusAppearance(elements.cmsStatus, "Audio edit page is active", "warning");
  } else {
    setStatusAppearance(elements.cmsStatus, "Open the Audios catalog", "error");
  }

  if (state.transcriptCount) {
    setStatusAppearance(elements.dataStatus, `${state.transcriptCount} transcript sections loaded`, "success");
  } else {
    setStatusAppearance(elements.dataStatus, "Run Scan / Dry Run to load transcripts", "warning");
  }
  elements.folderStatus.textContent =
    folderMeta && folderMeta.name ? `Local audio folder: ${folderMeta.name}` : "No local audio folder selected.";

  const counts = state.counts || {};
  const workTotal = Array.isArray(state.workItems) ? state.workItems.length : 0;
  const index = Number.isInteger(state.index) ? state.index : 0;
  const status = state.status || "idle";
  const phase = state.phase || "idle";

  elements.phase.textContent = `${status} / ${phase}`;
  elements.runStatus.textContent =
    audioRun && audioRun.status === "running"
      ? `downloading audio - ${Math.min(audioRun.index || 0, (audioRun.items || []).length)} of ${(audioRun.items || []).length} stops checked.`
      : status === "idle"
        ? "Idle."
        : `${status} (${phase}) - ${Math.min(index, workTotal)} of ${workTotal} tasks processed.`;
  elements.currentItem.textContent = state.current && state.current.label ? state.current.label : "";
  elements.countMap.stops.textContent = String(counts.eligibleStops || 0);
  elements.countMap.tasks.textContent = String(counts.totalLanguageTasks || 0);
  elements.countMap.ready.textContent = String(counts.ready || 0);
  elements.countMap.done.textContent = String((counts.processed || 0) + (counts.alreadyComplete || 0));
  elements.countMap.skip.textContent = String(counts.skipped || 0);
  elements.countMap.fail.textContent = String((counts.failed || 0) + (counts.missingTranscript || 0));

  elements.scanButton.disabled = !isBloombergAudioCatalogPage();
  elements.startButton.disabled = !isBloombergAudioCatalogPage();
  elements.resumeButton.disabled = !["paused", "stopped"].includes(status);
  elements.stopButton.disabled = status !== "running" && !audioRun;
  elements.folderButton.disabled = false;
  elements.downloadAudioButton.disabled = !isBloombergAudioCatalogPage();
  if (elements.repairIndexButton) elements.repairIndexButton.disabled = !isBloombergAudioCatalogPage();
  renderPanelLog(elements.logList, Array.isArray(state.log) ? state.log.slice(-8).reverse() : []);

  maybeAutoContinueRun(state);
  if (audioRun && audioRun.status === "running" && !bcaAudioDownloadInFlight) {
    setTimeout(() => continueAudioDownloadRun().catch((error) => failAudioDownloadRun(error)), 250);
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") return false;

  if (message.type === "bca:scan") {
    scanCatalog()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:start") {
    startRun()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:resume") {
    startRun({ resume: true })
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:stop") {
    stopRun()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === "bca:repairIndex") {
    repairAudioItemIndex()
      .then((state) => sendResponse({ ok: true, state }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  return false;
});

// On every page load, check if a run was in progress and needs to continue.
// This handles the case where continueRun() navigated the page.
async function checkAndContinueRun() {
  try {
    const state = await readState();
    if (state.status !== "running") return;
    // Don't interfere with processing/saving — those are handled by in-progress code.
    if (["processing", "saving"].includes(state.phase)) return;
    // Small delay to let the page fully settle.
    await delay(300);
    // Re-read state in case it changed during the delay.
    const latest = await readState();
    if (latest.status !== "running") return;
    if (["processing", "saving"].includes(latest.phase)) return;
    await maybeAutoContinueRun(latest);
  } catch (error) {
    console.error("Bloomberg Audio Assistant auto-continue error:", error);
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => {
    ensureInjectedPanel();
    refreshInjectedPanel();
    checkAndContinueRun();
  });
} else {
  ensureInjectedPanel();
  refreshInjectedPanel();
  checkAndContinueRun();
}

// Deterministic menu opener for Add Translation React Select.
async function openMenuSimple(input, control) {
  const menuVisible = () =>
    visibleElements("[role='option'], [role='listbox'], [class*='menu'], [class*='Menu']", document).length > 0 ||
    (input instanceof Element && normalizeComparable(input.getAttribute("aria-expanded")) === "true");

  const tryOpen = async (target) => {
    if (!(target instanceof Element)) return false;
    target.focus?.();
    invokeReactHandler(target, "onMouseDown");
    target.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, view: window, button: 0 }));
    target.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, view: window, button: 0 }));
    target.click?.();
    dispatchUserClick(target);
    await delay(120);
    if (menuVisible()) return true;

    sendKeyPress(target, "ArrowDown", "ArrowDown");
    await delay(120);
    if (menuVisible()) return true;

    sendKeyPress(target, " ", "Space");
    await delay(120);
    return menuVisible();
  };

  if (await tryOpen(control)) return true;
  if (await tryOpen(input)) return true;

  const fallbackInput = findReactSelectInput(document);
  if (fallbackInput && fallbackInput !== input && (await tryOpen(fallbackInput))) return true;

  return menuVisible();
}
