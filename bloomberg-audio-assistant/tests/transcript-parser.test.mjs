import assert from "node:assert/strict";
import test from "node:test";
import {
  cleanTranscriptTitle,
  countTranscriptEntries,
  firstSpeakerIdentity,
  isSpeakerLine,
  parseTranscriptEntries,
  parseTranscriptText,
  speakerNamesAreClose,
  speakersAppearAligned,
} from "../tools/transcript-parser.mjs";

test("speaker labels are detected only when the line is a standalone uppercase label", () => {
  assert.equal(isSpeakerLine("IRV SCHMUCKLER:"), true);
  assert.equal(isSpeakerLine("ALFRED RUNDELL, ALCALDE:"), true);
  assert.equal(isSpeakerLine("This is not a speaker:"), false);
  assert.equal(isSpeakerLine("18. SINAGOGA"), false);
});

test("titles drop speaker prefixes and use readable casing", () => {
  assert.equal(cleanTranscriptTitle("LAURA MASS: SINAGOGA"), "Sinagoga");
  assert.equal(cleanTranscriptTitle("DONALD VAUGHN: RELIGION IM 20. JAHRHUNDERT"), "Religion Im 20. Jahrhundert");
});

test("transcript entries preserve speaker names separately from body text", () => {
  const entries = parseTranscriptEntries(`
IRV SCHMUCKLER:
Me llamo Irving Schmuckler.
Fui el profesor aquí.

ALFRED RUNDELL, ALCALDE:
Construimos una nueva sinagoga.
`);

  assert.deepEqual(entries, [
    {
      speaker: "IRV SCHMUCKLER",
      paragraphs: ["Me llamo Irving Schmuckler. Fui el profesor aquí."],
      text: "Me llamo Irving Schmuckler. Fui el profesor aquí.",
    },
    {
      speaker: "ALFRED RUNDELL, ALCALDE",
      paragraphs: ["Construimos una nueva sinagoga."],
      text: "Construimos una nueva sinagoga.",
    },
  ]);
});

test("transcript entries preserve paragraph breaks without repeating speaker labels", () => {
  const entries = parseTranscriptEntries(`
DONALD VAUGHN, GUARDIA:
Soy Donald Vaughn, encargado de la Institucion Correccional del Estado
en Graterford, y era guardia en la prisión de Eastern State.

Una vez que finalizó el sistema de aislamiento en la prisión de Eastern
State, los deportes en grupo se convirtieron en una parte importante.

Si se dirige hacia la torre de guardia central, verá una valla.
                                     © 2003-2019 Eastern State Penitentiary Historic Site
                                                       Reservados todos los derechos.


para evitar que las pelotas salieran volando de la prisión.
`);

  assert.equal(entries.length, 1);
  assert.equal(entries[0].speaker, "DONALD VAUGHN, GUARDIA");
  assert.deepEqual(entries[0].paragraphs, [
    "Soy Donald Vaughn, encargado de la Institucion Correccional del Estado en Graterford, y era guardia en la prisión de Eastern State.",
    "Una vez que finalizó el sistema de aislamiento en la prisión de Eastern State, los deportes en grupo se convirtieron en una parte importante.",
    "Si se dirige hacia la torre de guardia central, verá una valla. para evitar que las pelotas salieran volando de la prisión.",
  ]);
});

test("parser chooses body sections over table-of-contents headings", () => {
  const parsed = parseTranscriptText(
    `
1. INTRODUCCIÓN
2. LAS PRISIONES ANTES DE EASTERN
18. LAURA MASS: SINAGOGA

18. LAURA MASS: SINAGOGA
IRV SCHMUCKLER:
Me llamo Irving Schmuckler. Fui el profesor aquí en el programa GED.

ALFRED RUNDELL, ALCALDE:
Construimos una nueva sinagoga allí para los internos judíos.

19. DONALD VAUGHN: RELIGIÓN EN EL SIGLO XX
DONALD VAUGHN:
Texto siguiente con suficiente cuerpo para que el parser lo trate como una sección real del documento y no como una línea de índice.
`,
    "spanish"
  );

  assert.equal(parsed.stops["18"].title, "Sinagoga");
  assert.equal(parsed.stops["18"].entries.length, 2);
  assert.equal(parsed.stops["19"].title, "Religión En El Siglo Xx");
});

test("parser recognizes uppercase headings without periods", () => {
  const parsed = parseTranscriptText(
    `
11 ESPORTES

DONALD VAUGHN, GUARDA:
Eu sou Donald Vaughn, superintendente da State Correctional Institution em Graterford.
Os esportes em equipe se tornaram uma parte importante da vida na prisão.

12   BLOCO DE CELAS 14

SEAN KELLEY, DIRETOR DO PROGRAMA:
Eu sou Sean Kelley, Diretor do Programa aqui no conjunto histórico da Penitenciária de Eastern State.
`,
    "portuguese"
  );

  assert.equal(parsed.stops["11"].title, "Esportes");
  assert.equal(parsed.stops["11"].entries[0].speaker, "DONALD VAUGHN, GUARDA");
  assert.deepEqual(parsed.stops["11"].entries[0].paragraphs, [
    "Eu sou Donald Vaughn, superintendente da State Correctional Institution em Graterford. Os esportes em equipe se tornaram uma parte importante da vida na prisão.",
  ]);
  assert.equal(parsed.stops["12"].title, "Bloco De Celas 14");
});


test("transcript entry counts include every language stop", () => {
  assert.equal(
    countTranscriptEntries({
      languages: {
        spanish: { stops: { 18: {}, 19: {} } },
        french: { stops: { 18: {} } },
      },
    }),
    3
  );
});

test("speaker alignment catches mismatched translated sections", () => {
  const canonical = {
    entries: [{ speaker: "LAURA MASS", text: "Welcome." }],
  };
  const aligned = {
    entries: [{ speaker: "LAURA MASS", text: "Bienvenue." }],
  };
  const mismatched = {
    entries: [{ speaker: "IRWIN SCHMUCKLER", text: "Me llamo." }],
  };

  assert.equal(firstSpeakerIdentity(canonical), "laura mass");
  assert.equal(speakersAppearAligned(canonical, aligned), true);
  assert.equal(speakersAppearAligned(canonical, mismatched), false);
  assert.equal(speakerNamesAreClose("donald vaughn", "donald vaughan"), true);
});
