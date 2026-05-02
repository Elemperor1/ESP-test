export function normalizeText(value) {
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\u0000/g, "")
    .replace(/\f/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function normalizeSpaces(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

export function normalizeComparable(value) {
  return normalizeSpaces(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function titleCase(value) {
  return normalizeSpaces(value).replace(/\p{L}[\p{L}'’-]*/gu, (word) => {
    if (word === word.toUpperCase() || word === word.toLowerCase()) {
      return word.charAt(0).toLocaleUpperCase() + word.slice(1).toLocaleLowerCase();
    }
    return word;
  });
}

export function cleanTranscriptTitle(rawTitle) {
  let title = normalizeSpaces(rawTitle)
    .replace(/©.*$/i, "")
    .replace(/\bLOCATION\b.*$/i, "")
    .replace(/\s+/g, " ")
    .trim();

  if (title.includes(":")) {
    const parts = title.split(":");
    title = parts[parts.length - 1];
  }

  return titleCase(title);
}

export function isSpeakerLine(line) {
  const text = normalizeSpaces(line).replace(/©.*$/i, "").trim();
  if (!text.endsWith(":")) return false;
  if (text.length < 3 || text.length > 90) return false;
  if (/\d{3,}/.test(text)) return false;
  const withoutColon = text.slice(0, -1).trim();
  return withoutColon === withoutColon.toLocaleUpperCase();
}

function isFooterOrNoiseLine(line) {
  const text = normalizeSpaces(line);
  if (!text) return true;
  if (/^\d{1,3}$/.test(text)) return true;
  if (/^\((?:location|lugar|local|lieu|ort|posizione)\b/i.test(text)) return true;
  if (/^updated\s+\d/i.test(text)) return true;
  if (/^all rights reserved/i.test(text)) return true;
  if (/^alle rechte vorbehalten/i.test(text)) return true;
  if (/^reservados todos los derechos/i.test(text)) return true;
  if (/^todos os direitos reservados/i.test(text)) return true;
  if (/^tutti i diritti riservati/i.test(text)) return true;
  if (/^tous droits réservés/i.test(text)) return true;
  if (/^©\s*2003/i.test(text)) return true;
  if (/eastern state penitentiary historic site/i.test(text) && /©|rights|reserv/i.test(text)) return true;
  return false;
}

function cleanBodyLine(line) {
  return normalizeSpaces(
    line
      .replace(/©\s*2003.*$/i, "")
      .replace(/Eastern State Penitentiary Historic Site.*$/i, "")
      .replace(/Alle Rechte vorbehalten.*$/i, "")
      .replace(/Todos os direitos reservados.*$/i, "")
      .replace(/Tutti i diritti riservati.*$/i, "")
      .replace(/Tous droits réservés.*$/i, "")
  );
}

export function parseTranscriptEntries(body) {
  const entries = [];
  let current = null;
  let currentParagraph = [];
  let pendingParagraphBreak = false;
  let inNoiseBlock = false;

  function ensureCurrent() {
    if (!current) {
      current = { speaker: "", paragraphs: [] };
    }
  }

  function commitParagraph() {
    const paragraph = normalizeSpaces(currentParagraph.join(" "));
    if (paragraph) {
      ensureCurrent();
      current.paragraphs.push(paragraph);
    }
    currentParagraph = [];
    pendingParagraphBreak = false;
  }

  function commitEntry() {
    if (!current) return;
    commitParagraph();
    current.paragraphs = (current.paragraphs || []).filter(Boolean);
    current.text = normalizeSpaces(current.paragraphs.join(" "));
    if (current.speaker || current.text) {
      entries.push(current);
    }
    current = null;
  }

  for (const rawLine of normalizeText(body).split("\n")) {
    const rawText = normalizeSpaces(rawLine);
    const line = cleanBodyLine(rawLine);
    const isBlank = !rawText || !line;

    if (isBlank) {
      if (!inNoiseBlock && currentParagraph.length) {
        pendingParagraphBreak = true;
      }
      continue;
    }

    if (isFooterOrNoiseLine(line)) {
      inNoiseBlock = true;
      pendingParagraphBreak = false;
      continue;
    }

    inNoiseBlock = false;

    if (isSpeakerLine(line)) {
      commitEntry();
      current = { speaker: line.slice(0, -1).trim(), paragraphs: [] };
      continue;
    }

    ensureCurrent();
    if (pendingParagraphBreak) {
      commitParagraph();
    }
    currentParagraph.push(line);
  }

  commitEntry();

  return entries.filter((entry) => entry.speaker || entry.text);
}

function headingFromLine(line) {
  const cleaned = String(line || "")
    .replace(/©.*$/i, "")
    .trim();
  let match = cleaned.match(/^(\d{1,3})(?:\.|\s{2,})\s*(.{2,})$/);
  if (!match) {
    match = cleaned.match(/^(\d{1,3})\s+(.{2,})$/);
    if (match && match[2] !== match[2].toLocaleUpperCase()) {
      match = null;
    }
  }
  if (!match) return null;
  return {
    stopNumber: Number(match[1]),
    rawTitle: match[2].trim(),
  };
}

function scoreSection(section) {
  const entries = parseTranscriptEntries(section.body);
  const speakerScore = entries.filter((entry) => entry.speaker).length * 1000;
  const bodyScore = normalizeSpaces(section.body).length;
  const titlePenalty = /audio|stop|list|location/i.test(section.rawTitle) ? -250 : 0;
  return speakerScore + bodyScore + titlePenalty;
}

export function parseTranscriptText(rawText, languageKey) {
  const text = normalizeText(rawText);
  const lines = text.split("\n");
  const candidates = [];
  let current = null;

  for (const line of lines) {
    const heading = headingFromLine(line);
    if (heading) {
      if (current) candidates.push(current);
      current = { ...heading, body: "" };
      continue;
    }

    if (current) {
      current.body += `${line}\n`;
    }
  }
  if (current) candidates.push(current);

  const bestByStop = new Map();
  for (const candidate of candidates) {
    const entries = parseTranscriptEntries(candidate.body);
    if (entries.length === 0 || normalizeSpaces(candidate.body).length < 80) continue;
    const previous = bestByStop.get(candidate.stopNumber);
    if (!previous || scoreSection(candidate) > scoreSection(previous)) {
      bestByStop.set(candidate.stopNumber, candidate);
    }
  }

  const stops = {};
  for (const [stopNumber, section] of [...bestByStop.entries()].sort((a, b) => a[0] - b[0])) {
    stops[String(stopNumber)] = {
      stopNumber,
      title: cleanTranscriptTitle(section.rawTitle),
      entries: parseTranscriptEntries(section.body),
    };
  }

  return {
    languageKey,
    stops,
  };
}

export function countTranscriptEntries(transcriptIndex) {
  return Object.values(transcriptIndex.languages || {}).reduce(
    (sum, language) => sum + Object.keys(language.stops || {}).length,
    0
  );
}

export function speakerIdentity(value) {
  const text = normalizeComparable(value)
    .replace(/\b(actor|ator|actora|inmate|interno|presidiario|guard|guarda|guide|guia|tour|director|diretor|direttore|programme|program|programm|coordinator|coordenador|coordinador|supervisor|warden|alcaide|superintendente|artist|artista|guide du tour)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return text.split(" ").filter((part) => part.length > 1).slice(0, 3).join(" ");
}

export function firstSpeakerIdentity(stop) {
  const firstSpeaker = stop && Array.isArray(stop.entries)
    ? stop.entries.find((entry) => entry.speaker)
    : null;
  return firstSpeaker ? speakerIdentity(firstSpeaker.speaker) : "";
}

export function speakersAppearAligned(expectedStop, translatedStop) {
  const expected = firstSpeakerIdentity(expectedStop);
  const translated = firstSpeakerIdentity(translatedStop);
  if (!expected || !translated) return true;
  const expectedTokens = expected.split(" ").filter(Boolean);
  const translatedTokens = translated.split(" ").filter(Boolean);
  if (expectedTokens.length < 2 || translatedTokens.length < 2) return true;
  return expected === translated || expected.includes(translated) || translated.includes(expected) || speakerNamesAreClose(expected, translated);
}

function editDistance(a, b) {
  const left = String(a || "");
  const right = String(b || "");
  const dp = Array.from({ length: left.length + 1 }, () => Array(right.length + 1).fill(0));
  for (let i = 0; i <= left.length; i += 1) dp[i][0] = i;
  for (let j = 0; j <= right.length; j += 1) dp[0][j] = j;
  for (let i = 1; i <= left.length; i += 1) {
    for (let j = 1; j <= right.length; j += 1) {
      const cost = left[i - 1] === right[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[left.length][right.length];
}

export function speakerNamesAreClose(expected, translated) {
  const expectedTokens = String(expected || "").split(" ").filter(Boolean);
  const translatedTokens = String(translated || "").split(" ").filter(Boolean);
  if (expectedTokens.length < 2 || translatedTokens.length < 2) return true;

  const expectedFirst = expectedTokens[0];
  const expectedLast = expectedTokens[expectedTokens.length - 1];
  const translatedFirst = translatedTokens[0];
  const translatedLast = translatedTokens[translatedTokens.length - 1];

  return (
    (editDistance(expectedFirst, translatedFirst) <= 1 && editDistance(expectedLast, translatedLast) <= 2) ||
    (editDistance(expectedFirst, translatedLast) <= 1 && editDistance(expectedLast, translatedFirst) <= 2)
  );
}
