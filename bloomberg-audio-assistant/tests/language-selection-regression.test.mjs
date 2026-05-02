import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";

test("main-world language selection does not fall back to the stale page bridge", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const mainWorldFunction = source.match(/async function mainWorldSelectAddTranslationLanguage[\s\S]*?\n}\n/);
  assert.ok(mainWorldFunction, "mainWorldSelectAddTranslationLanguage should exist");
  assert.doesNotMatch(
    mainWorldFunction[0],
    /bridgeSelectAddTranslationLanguage\(/,
    "mainWorldSelectAddTranslationLanguage should not attempt the stale page-bridge fallback"
  );
});

test("ensureLanguageSelected closes stale add-translation dialog when target language is already active", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const ensureFunction = source.match(/async function ensureLanguageSelected[\s\S]*?\n}\n/);
  assert.ok(ensureFunction, "ensureLanguageSelected should exist");
  assert.match(
    ensureFunction[0],
    /languageLabelMatches\(readSelectedLanguageLabel\(root\), item\.languageLabel\)[\s\S]*closeAddTranslationDialogIfOpen\(\)/,
    "ensureLanguageSelected should close the stale Add Translation dialog when the edit modal already shows the target language"
  );
});

test("scan does not skip solely because local completedTasks says a language was done", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const scanFunction = source.match(/async function scanCatalog[\s\S]*?\n}\n\nfunction openAudioDb/);
  assert.ok(scanFunction, "scanCatalog should exist");
  assert.doesNotMatch(
    scanFunction[0],
    /recordedComplete[\s\S]*continue;/,
    "scanCatalog should queue previously completed tasks for edit-page verification instead of skipping from stale local state"
  );
  assert.match(
    scanFunction[0],
    /Previously marked complete; queued for edit-page verification/,
    "scanCatalog should log that stale local completion records require verification"
  );
});

test("ensureLanguageSelected checks existing language options before Add Translation", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const ensureFunction = source.match(/async function ensureLanguageSelected[\s\S]*?\n}\n/);
  assert.ok(ensureFunction, "ensureLanguageSelected should exist");
  assert.match(
    ensureFunction[0],
    /selectExistingLanguageFromEditDropdown\(item, dropdown\)/,
    "ensureLanguageSelected should select an existing language from the edit dropdown before trying Add Translation"
  );
});

test("ensureLanguageSelected opens the current language menu before Add Translation", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const ensureFunction = source.match(/async function ensureLanguageSelected[\s\S]*?\n}\n/);
  assert.ok(ensureFunction, "ensureLanguageSelected should exist");
  assert.match(
    ensureFunction[0],
    /openLanguageDropdownControl\(dropdown\)/,
    "ensureLanguageSelected should open the language dropdown via openLanguageDropdownControl when looking for Add Translation"
  );
});

test("menu-driven CMS actions request focus before interacting", async () => {
  const contentSource = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  assert.match(contentSource, /async function ensureCmsWindowFocused\(\)/, "content script should expose a CMS focus helper");
  assert.match(
    contentSource,
    /async function clickLanguageDropdown[\s\S]*?ensureCmsWindowFocused\(\)/,
    "language dropdown actions should focus the CMS window first"
  );
  assert.match(
    contentSource,
    /async function openCatalogRow[\s\S]*?ensureCmsWindowFocused\(\)/,
    "catalog row clicks should focus the CMS window first"
  );
  assert.match(
    contentSource,
    /async function saveAndConfirm[\s\S]*?ensureCmsWindowFocused\(\)/,
    "save clicks should focus the CMS window first"
  );

  const backgroundSource = await readFile(new URL("../background/service-worker.js", import.meta.url), "utf8");
  assert.match(backgroundSource, /bca:focusCmsTab/, "background worker should handle CMS tab focus requests");
  assert.match(backgroundSource, /chrome\.windows\.update/, "background worker should focus the CMS window");
  assert.match(backgroundSource, /chrome\.tabs\.update/, "background worker should activate the CMS tab");
});

test("save run resolves CMS routes before falling back to visible catalog rows", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const continueFunction = source.match(/async function continueRun[\s\S]*?\n}\n\nasync function maybeAutoContinueRun/);
  assert.ok(continueFunction, "continueRun should exist");
  assert.match(continueFunction[0], /resolveRouteForRunItem\(state, item, config\)/, "continueRun should resolve route data first");
  assert.ok(
    continueFunction[0].indexOf("resolveRouteForRunItem(state, item, config)") < continueFunction[0].indexOf("openCatalogRow(item)"),
    "visible catalog row lookup must remain the last fallback"
  );
  assert.match(source, /function resolveAudioItemFromIndex/, "durable audio item index resolver should exist");
  assert.match(source, /repairAudioItemIndex\(\{ state, silent: true \}\)/, "missing route data should trigger one repair pass");
  assert.match(source, /function routeFailureMessage/, "route failures should include diagnostic data");
});

test("panel and popup expose Repair Item Index", async () => {
  const contentSource = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const popupHtml = await readFile(new URL("../popup/popup.html", import.meta.url), "utf8");
  const popupJs = await readFile(new URL("../popup/popup.js", import.meta.url), "utf8");
  assert.match(contentSource, /Repair Item Index/, "injected panel should include a Repair Item Index button");
  assert.match(contentSource, /bca:repairIndex/, "content script should accept the repair index command");
  assert.match(popupHtml, /repair-index/, "popup should expose the repair index command");
  assert.match(popupJs, /bca:repairIndex/, "popup should send the repair index command");
});

test("build guard and alerts include the running build", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  assert.match(source, /const BCA_EXTENSION_VERSION = "1\.0\.45"/, "content script should declare the expected manifest version");
  assert.match(source, /assertExtensionBuildCurrent\(\)/, "content script should guard CMS work against stale loaded builds");
  assert.match(source, /Bloomberg Audio Assistant \(\$\{BCA_BUILD_LABEL\}\) paused/, "pause alerts should show the running build label");
});

test("Add Translation selection is not successful until Add is enabled", async () => {
  const source = await readFile(new URL("../content/bloomberg-audio.js", import.meta.url), "utf8");
  const selectFunction = source.match(/async function selectLanguageInAddTranslationDialog[\s\S]*?\n}\n\nfunction findLanguageSection/);
  assert.ok(selectFunction, "selectLanguageInAddTranslationDialog should exist");
  assert.match(selectFunction[0], /waitForAddTranslationSelectionReady\(dialog\)/, "language selection should wait for the Add button to become enabled");
  assert.match(source, /function findAddTranslationLanguageOption/, "Add Translation should use a menu-option-specific language finder");
  assert.doesNotMatch(
    selectFunction[0],
    /selectOptionByText\(languageLabel, dialog\)\) return true/,
    "clicking a visible option should not be treated as success without Add becoming enabled"
  );
  const ensureFunction = source.match(/async function ensureLanguageSelected[\s\S]*?\n}\n\n\/\/ Finds the "Add Translation"/);
  assert.ok(ensureFunction, "ensureLanguageSelected should exist");
  assert.match(ensureFunction[0], /isDisabledAction\(addButton\)/, "ensureLanguageSelected should not click a disabled Add button");
  assert.match(ensureFunction[0], /Selected \$\{item\.languageLabel\} did not enable Add Translation/, "disabled Add should fail before changing fields");
});
