#!/usr/bin/env node
import { execFile } from "node:child_process";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { countTranscriptEntries, firstSpeakerIdentity, parseTranscriptText, speakersAppearAligned } from "./transcript-parser.mjs";

const execFileAsync = promisify(execFile);
const __dirname = dirname(fileURLToPath(import.meta.url));
const extensionRoot = resolve(__dirname, "..");
const configPath = resolve(extensionRoot, "config.default.json");
const outputPath = resolve(extensionRoot, "data/transcripts.json");
const cacheDir = resolve(extensionRoot, ".cache/transcript-pdfs");

async function commandExists(command) {
  try {
    await execFileAsync("which", [command]);
    return true;
  } catch {
    return false;
  }
}

async function downloadFile(url, outputFile) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Download failed for ${url}: ${response.status}`);
  }
  const buffer = Buffer.from(await response.arrayBuffer());
  await writeFile(outputFile, buffer);
}

async function pdfToText(pdfFile, textFile) {
  await execFileAsync("pdftotext", ["-layout", "-enc", "UTF-8", pdfFile, textFile], {
    maxBuffer: 50 * 1024 * 1024,
  });
  return readFile(textFile, "utf8");
}

async function main() {
  if (!(await commandExists("pdftotext"))) {
    throw new Error("pdftotext is required. Install poppler first, then rerun this script.");
  }

  const config = JSON.parse(await readFile(configPath, "utf8"));
  await rm(cacheDir, { force: true, recursive: true });
  await mkdir(cacheDir, { recursive: true });

  const transcriptIndex = {
    generatedAt: new Date().toISOString(),
    source: "Eastern State Accessibility language materials PDFs",
    englishTranscriptUrl: config.englishTranscriptUrl,
    alignmentWarnings: {},
    languages: {},
  };

  const englishPdfPath = resolve(cacheDir, "english.pdf");
  const englishTextPath = resolve(cacheDir, "english.txt");
  console.log("Downloading English canonical transcript...");
  await downloadFile(config.englishTranscriptUrl, englishPdfPath);
  const englishText = await pdfToText(englishPdfPath, englishTextPath);
  const englishParsed = parseTranscriptText(englishText, "english");
  console.log(`  Parsed ${Object.keys(englishParsed.stops).length} canonical stops.`);

  for (const language of config.targetLanguages) {
    const pdfPath = resolve(cacheDir, `${language.key}.pdf`);
    const textPath = resolve(cacheDir, `${language.key}.txt`);
    console.log(`Downloading ${language.cmsLabel} transcript...`);
    await downloadFile(language.transcriptUrl, pdfPath);
    const text = await pdfToText(pdfPath, textPath);
    const parsed = parseTranscriptText(text, language.key);
    const alignmentWarnings = {};

    for (const [stopNumber, translatedStop] of Object.entries(parsed.stops)) {
      const expectedStop = englishParsed.stops[stopNumber];
      if (expectedStop && !speakersAppearAligned(expectedStop, translatedStop)) {
        alignmentWarnings[stopNumber] = {
          title: translatedStop.title,
          firstSpeaker: firstSpeakerIdentity(translatedStop),
          expectedFirstSpeaker: firstSpeakerIdentity(expectedStop),
          reason: "first-speaker-mismatch",
        };
      }
    }

    transcriptIndex.languages[language.key] = {
      cmsLabel: language.cmsLabel,
      accessibilityLabel: language.accessibilityLabel,
      transcriptUrl: language.transcriptUrl,
      stops: parsed.stops,
    };
    if (Object.keys(alignmentWarnings).length) {
      transcriptIndex.alignmentWarnings[language.key] = alignmentWarnings;
    }
    console.log(
      `  Parsed ${Object.keys(parsed.stops).length} stops; ${Object.keys(alignmentWarnings).length} speaker-alignment warnings.`
    );
  }

  await writeFile(outputPath, `${JSON.stringify(transcriptIndex, null, 2)}\n`);
  console.log(`Wrote ${countTranscriptEntries(transcriptIndex)} transcript sections to ${outputPath}.`);
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
