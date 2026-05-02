import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";

function parseStopNumber(...values) {
  for (const value of values) {
    const text = String(value || "");
    const match = text.match(/(?:^|[^\d])0?(\d{1,3})(?:[.\s_-]|$)/);
    if (match) return Number(match[1]);
  }
  return null;
}

function normalizeComparable(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
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

function languageLabelMatches(activeLabel, expectedLabel) {
  const active = normalizeComparable(activeLabel);
  const expected = normalizeComparable(expectedLabel);
  if (!active || !expected) return false;
  return active === expected || active.includes(expected) || expected.includes(active);
}

function normalizeSpaces(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function titleCase(value) {
  return normalizeSpaces(value).replace(/\p{L}[\p{L}'’-]*/gu, (word) => {
    if (word === word.toUpperCase() || word === word.toLowerCase()) {
      return word.charAt(0).toLocaleUpperCase() + word.slice(1).toLocaleLowerCase();
    }
    return word;
  });
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

function chooseCatalogTitle({ fileName, includedIn, transcript, fallbackTitle, stopNumber }) {
  if (fallbackTitle) return normalizeSpaces(fallbackTitle);
  if (includedIn) return titleCase(includedIn);
  if (transcript && transcript.title) return transcript.title;
  return cleanStopTitle(fileName || "", stopNumber);
}

function activeTargetLanguages(config) {
  return (config.targetLanguages || []).filter((language) => language && language.enabled !== false);
}

function firstTranscriptParagraph(transcript) {
  for (const entry of transcript && Array.isArray(transcript.entries) ? transcript.entries : []) {
    for (const paragraph of Array.isArray(entry.paragraphs) ? entry.paragraphs : []) {
      const value = normalizeSpaces(paragraph);
      if (value) return value;
    }
  }
  return "";
}

function stripTitleDecorations(title) {
  return normalizeSpaces(title)
    .replace(/^["“”'‘’]+/, "")
    .replace(/["“”'‘’]+$/, "")
    .replace(/(?:\.{3}|…)+$/, "")
    .trim();
}

function takeWords(text, count) {
  const words = normalizeSpaces(text).split(/\s+/).filter(Boolean);
  if (!words.length) return "";
  if (words.length <= count) return words.join(" ");
  return `${words.slice(0, count).join(" ")}...`;
}

function translatedTitleFromCatalogTitle(sourceTitle, transcript) {
  const source = normalizeSpaces(sourceTitle);
  if (!source) return titleCase(transcript && transcript.title ? transcript.title : "");
  const paragraph = firstTranscriptParagraph(transcript);
  if (!paragraph) return titleCase(transcript && transcript.title ? transcript.title : source);
  const sourceCore = stripTitleDecorations(source);
  const sourceWordCount = sourceCore.split(/\s+/).filter(Boolean).length;
  const isQuoted = /^["“”]/.test(source);
  const isExcerpt = /(?:\.{3}|…)/.test(source) || isQuoted;
  const translated = isExcerpt ? takeWords(paragraph, Math.max(4, sourceWordCount || 8)) : titleCase(transcript.title || paragraph);
  return isQuoted && translated ? `"${translated.replace(/^["“”]+|["“”]+$/g, "")}"` : translated;
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

function cleanTranscriptParagraphForCms(paragraph) {
  const text = normalizeSpaces(paragraph);
  if (!text) return "";
  if (isAudioGuidePromptText(text)) return "";
  const sentences = splitSentences(text);
  if (sentences.length <= 1) return text;
  return normalizeSpaces(sentences.filter((sentence) => !isAudioGuidePromptText(sentence)).join(" "));
}

function sanitizedTranscriptEntries(transcript) {
  const entries = transcript && Array.isArray(transcript.entries) ? transcript.entries : [];
  return entries
    .map((entry) => {
      if (isAudioGuidePromptText(entry.speaker || "")) return null;
      const sourceParagraphs = Array.isArray(entry.paragraphs) && entry.paragraphs.length ? entry.paragraphs : [entry.text || ""].filter(Boolean);
      const paragraphs = sourceParagraphs.map(cleanTranscriptParagraphForCms).filter(Boolean);
      if (!paragraphs.length) return null;
      return { ...entry, paragraphs, text: paragraphs.join(" ") };
    })
    .filter(Boolean);
}

function transcriptVerificationSnippets(text) {
  const normalized = normalizeComparable(text);
  if (!normalized) return [];
  if (normalized.length <= 120) return [normalized];
  const mid = Math.floor(normalized.length / 2);
  return [normalized.slice(0, 80), normalized.slice(Math.max(0, mid - 40), mid + 40), normalized.slice(Math.max(0, normalized.length - 80))];
}

function transcriptTextAppearsComplete(actualText, expectedText) {
  const actual = normalizeComparable(actualText);
  const expected = normalizeComparable(expectedText);
  if (!expected) return Boolean(actual);
  if (!actual) return false;
  if (actual.includes(expected)) return true;
  if (actual.length < Math.floor(expected.length * 0.85)) return false;
  return transcriptVerificationSnippets(expected).every((snippet) => actual.includes(snippet));
}

function isExcludedStop(stopNumber, text, config) {
  if (!stopNumber) return true;
  if ((config.excludeStops || []).includes(stopNumber)) return true;
  const allowlist = Array.isArray(config.stopAllowlist) ? config.stopAllowlist.filter(Boolean) : [];
  if (allowlist.length && !allowlist.includes(stopNumber)) return true;
  const normalized = normalizeComparable(text);
  return (config.excludeTitlePatterns || []).some((pattern) => normalized.includes(normalizeComparable(pattern)));
}

function queueableCatalogItem({ title, fileName, includedIn = "", editUrl = "" }, config) {
  const itemId = (String(editUrl).match(/\/catalog\/audios\/(\d+)/) || [])[1] || "";
  const stopNumber = parseStopNumber(fileName);
  const sourceText = `${fileName} ${includedIn}`;
  if (isExcludedStop(stopNumber, sourceText, config)) return null;
  return {
    itemId,
    stopNumber,
    fileName,
    title: chooseCatalogTitle({ fileName, includedIn, fallbackTitle: title, stopNumber }),
  };
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

function buildRowKey(stopNumber, fileName, title) {
  return [stopNumber || "", normalizeComparable(fileName), normalizeComparable(title)].join("|");
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
  add("included", item.includedIn);
  return Array.from(aliases);
}

function audioIndexRecordFromCatalogRow(rowData) {
  const itemId = rowData.itemId || "";
  const stopNumber = rowData.stopNumber || parseStopNumber(rowData.fileName, rowData.includedIn, rowData.title);
  return {
    itemId,
    editUrl: itemId ? `https://cms.bloombergconnects.org/catalog/audios/${itemId}` : rowData.editUrl || "",
    stopNumber,
    fileName: rowData.fileName || "",
    includedIn: rowData.includedIn || "",
    title: rowData.title || "",
    rowKey: rowData.rowKey || buildRowKey(stopNumber, rowData.fileName, rowData.includedIn || rowData.title),
  };
}

function mergeAudioItemIndex(existingIndex, rowRecords) {
  const nextIndex = { ...(existingIndex || {}) };
  for (const rowData of rowRecords || []) {
    const record = audioIndexRecordFromCatalogRow(rowData);
    for (const alias of audioItemAliases(record)) {
      nextIndex[alias] = { ...(nextIndex[alias] || {}), ...record };
    }
  }
  return nextIndex;
}

function resolveAudioItemFromIndex(state, item) {
  const index = state && state.audioItemIndex ? state.audioItemIndex : {};
  for (const alias of audioItemAliases(item)) {
    const record = index[alias];
    if (record && (record.itemId || record.editUrl)) {
      return { ...item, itemId: record.itemId, editUrl: record.editUrl };
    }
  }
  return null;
}

test("stop numbers parse from filenames and titles", () => {
  assert.equal(parseStopNumber("018 Synagogue.wav"), 18);
  assert.equal(parseStopNumber("Stop 57 - Apokaluptein.wav"), 57);
  assert.equal(parseStopNumber("No number"), null);
});

test("scope exclusions cover requested stops and title variants", () => {
  const config = {
    excludeStops: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 92, 93, 94],
    excludeTitlePatterns: ["pyrrhic defeat", "phyric defeat"],
    stopAllowlist: [],
  };

  assert.equal(isExcludedStop(1, "001 Introduction", config), true);
  assert.equal(isExcludedStop(92, "Cellblock 15", config), true);
  assert.equal(isExcludedStop(18, "018 Synagogue", config), false);
  assert.equal(isExcludedStop(72, "Pyrrhic Defeat", config), true);
  assert.equal(isExcludedStop(72, "Phyric defeat", config), true);
});

test("allowlist narrows dry-run scope", () => {
  const config = {
    excludeStops: [],
    excludeTitlePatterns: [],
    stopAllowlist: [18],
  };
  assert.equal(isExcludedStop(18, "018 Synagogue", config), false);
  assert.equal(isExcludedStop(19, "019 Religion", config), true);
});

test("visible catalog rows remain queueable without hidden edit URLs", () => {
  const config = {
    excludeStops: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 92, 93, 94],
    excludeTitlePatterns: ["pyrrhic defeat", "phyric defeat"],
    stopAllowlist: [],
  };

  assert.deepEqual(queueableCatalogItem({ title: "Synagogue", fileName: "018 Synagogue.wav" }, config), {
    itemId: "",
    stopNumber: 18,
    fileName: "018 Synagogue.wav",
    title: "Synagogue",
  });
});

test("catalog identity uses filename while display title comes from title column", () => {
  const config = {
    excludeStops: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 92, 93, 94],
    excludeTitlePatterns: ["pyrrhic defeat", "phyric defeat"],
    stopAllowlist: [],
  };

  assert.deepEqual(
    queueableCatalogItem(
      {
        title: '"I Find The Beauty Of The Cell To Be Very Contradictory..."',
      fileName: "065 Transient Room.wav",
      includedIn: "Transient Room",
      },
      config
    ),
    {
      itemId: "",
      stopNumber: 65,
      fileName: "065 Transient Room.wav",
      title: '"I Find The Beauty Of The Cell To Be Very Contradictory..."',
    }
  );
});

test("catalog matching tolerates file-name spacing differences", () => {
  assert.equal(comparableTextMatches("012 Cellblock 14 NA.wav", "012 Cellblock14 NA.wav"), true);
  assert.equal(comparableTextMatches("Cellblock 14", "Cellblock14"), true);
  assert.equal(catalogItemSearchTerms({ stopNumber: 12, fileName: "012 Cellblock14 NA.wav" }).includes("012 Cellblock 14 NA.wav"), true);
});

test("audio item index resolves Cellblock spacing variants to the same CMS route", () => {
  const audioItemIndex = mergeAudioItemIndex(
    {},
    [
      {
        itemId: "972360",
        stopNumber: 12,
        fileName: "012 Cellblock 14 NA.wav",
        includedIn: "Cellblock 14",
        title: "Bloco de Celas 14",
      },
    ]
  );

  const resolved = resolveAudioItemFromIndex(
    { audioItemIndex },
    {
      stopNumber: 12,
      fileName: "012 Cellblock14 NA.wav",
      includedIn: "Cellblock14",
      rowKey: buildRowKey(12, "012 Cellblock14 NA.wav", "Cellblock14"),
    }
  );

  assert.equal(resolved.itemId, "972360");
  assert.equal(resolved.editUrl, "https://cms.bloombergconnects.org/catalog/audios/972360");
});

test("disabled target languages are not queued", () => {
  const config = {
    targetLanguages: [
      { key: "spanish", cmsLabel: "Spanish (Latin America)", enabled: false },
      { key: "portuguese", cmsLabel: "Portuguese" },
    ],
  };
  assert.deepEqual(
    activeTargetLanguages(config).map((language) => language.key),
    ["portuguese"]
  );
});

test("translated title follows title-column excerpt style", () => {
  const transcript = {
    title: "Portas",
    entries: [{ paragraphs: ["Que as portas sejam de ferro e as janelas pequenas. O prédio precisava parecer seguro."] }],
  };
  assert.equal(translatedTitleFromCatalogTitle('"Let the doors be of iron..."', transcript), '"Que as portas sejam de ferro..."');
});

test("cms transcript output removes audio guide prompts", () => {
  const transcript = {
    entries: [
      {
        speaker: "SEAN KELLEY, DIRETOR DO PROGRAMA",
        paragraphs: [
          "Eu sou Sean Kelley, Diretor do Programa aqui no conjunto histórico da Penitenciária de Eastern State.",
          "Na época em que este Bloco de Celas foi adicionado, a Penitenciária de Eastern State tinha sua população carcerária mais alta.",
          "Para ouvir sobre a arquitetura deste Bloco de Celas, digite 13.",
        ],
      },
      {
        speaker: "ACOUSTIGUIDE",
        paragraphs: ["Press 13 on your Acoustiguide."],
      },
    ],
  };
  const entries = sanitizedTranscriptEntries(transcript);
  assert.equal(entries.length, 1);
  assert.deepEqual(entries[0].paragraphs, [
    "Eu sou Sean Kelley, Diretor do Programa aqui no conjunto histórico da Penitenciária de Eastern State.",
    "Na época em que este Bloco de Celas foi adicionado, a Penitenciária de Eastern State tinha sua população carcerária mais alta.",
  ]);
});

test("transcript completion check rejects first-paragraph-only inserts", () => {
  const fullText = [
    "SEAN KELLEY, DIRETOR DO PROGRAMA:",
    "Eu sou Sean Kelley, Diretor do Programa aqui no conjunto histórico da Penitenciária de Eastern State.",
    "Na época em que este Bloco de Celas foi adicionado, a Penitenciária de Eastern State tinha sua população carcerária mais alta.",
  ].join("\n\n");
  const partialText = "SEAN KELLEY, DIRETOR DO PROGRAMA: Eu sou Sean Kelley, Diretor do Programa aqui no conjunto histórico da Penitenciária de Eastern State.";

  assert.equal(transcriptTextAppearsComplete(partialText, fullText), false);
  assert.equal(transcriptTextAppearsComplete(fullText, fullText), true);
});

test("edit page processing requires the current CMS item or the exact pending catalog open", () => {
  const stop11 = {
    itemId: "111",
    rowKey: "11|011 sports wav|sports",
    stopNumber: 11,
    fileName: "011 Sports.wav",
    languageKey: "portuguese",
  };
  const stop12 = {
    itemId: "",
    rowKey: "12|012 daily life wav|daily life",
    stopNumber: 12,
    fileName: "012 Daily Life.wav",
    languageKey: "portuguese",
  };

  assert.equal(currentEditPageMatchesItem({ index: 10 }, stop11, "111"), true);
  assert.equal(currentEditPageMatchesItem({ index: 11 }, stop12, "111"), false);
  assert.equal(
    currentEditPageMatchesItem(
      {
        index: 11,
        pendingOpenItem: {
          index: 11,
          rowKey: stop12.rowKey,
          stopNumber: 12,
          fileName: "012 Daily Life.wav",
          languageKey: "portuguese",
        },
      },
      stop12,
      "222"
    ),
    true
  );
});

test("local audio matching rejects same-stop duplicate titles", () => {
  const item = {
    stopNumber: 15,
    fileName: "015 Willie Sutton with Link and DoorNA.wav",
    title: "Willie Sutton With Link And DoorNA",
  };

  assert.equal(localAudioMatchScore("015 Willie Sutton with Link and DoorNA.wav", item), 100);
  assert.equal(localAudioMatchScore("015 1095 Bloque De Celdas 15 Na Spanish Dupe.wav", item), 0);
  assert.equal(localAudioMatchScore("015 1097 (2022) Cellblock 15 Conclusion Spanish Dupe.wav", item), 0);
});

test("language labels must match the active CMS selector before field writes", () => {
  assert.equal(languageLabelMatches("Spanish (Latin America)", "Spanish (Latin America)"), true);
  assert.equal(languageLabelMatches("Spanish (Latin America) - Draft", "Spanish (Latin America)"), true);
  assert.equal(languageLabelMatches("English (United States) - Default", "Spanish (Latin America)"), false);
  assert.equal(languageLabelMatches("", "German"), false);
});

test("transcript verification uses beginning, middle, and end snippets", () => {
  // A long text where the middle differs between expected and actual.
  const expected = "A".repeat(100) + "MIDDLE_UNIQUE" + "Z".repeat(100);
  // Has beginning and end but wrong middle.
  const partialWithWrongMiddle = "A".repeat(100) + "WRONG_MIDDLE" + "Z".repeat(100);

  assert.equal(transcriptTextAppearsComplete(partialWithWrongMiddle, expected), false);
  assert.equal(transcriptTextAppearsComplete(expected, expected), true);
});

test("markTaskComplete schedules auto-continue via maybeAutoContinueRun", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const markFunction = source.match(/async function markTaskComplete[\s\S]*?\n}\n/);
  assert.ok(markFunction, "markTaskComplete should exist");
  assert.match(
    markFunction[0],
    /maybeAutoContinueRun/,
    "markTaskComplete should schedule auto-continue after completing a task"
  );
  assert.match(
    markFunction[0],
    /state\.pendingOpenItem = null/,
    "markTaskComplete should clear pendingOpenItem to prevent cross-stop contamination"
  );
});
