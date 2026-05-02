# Bloomberg Connects Audio Assistant

Chrome extension for adding Eastern State translated audio entries in the Bloomberg Connects CMS.

The extension is separate from the older alt-text extension in the repository. It uses the logged-in CMS browser session and stores resumable run state in `chrome.storage.local`.

## What It Does

- Scans `https://cms.bloombergconnects.org/catalog/audios`.
- Excludes stops `1-10`, `92`, `93`, `94`, and titles matching `Pyrrhic Defeat` / `Phyric defeat`.
- Targets Portuguese, French, Italian, and German by default. Spanish (Latin America) remains configured but is disabled with `"enabled": false` until it is ready to resume.
- Uses the CMS catalog Title column as the source title. For translated entries, excerpt-style titles are generated from the matching translated transcript body, and transcript text comes from the Eastern State Accessibility language-material PDFs.
- Prefers matching audio files from a user-selected local audio folder and can download missing English CMS source audio into that folder. Save mode requires a local audio match and does not download audio into the browser cache.
- Adds missing translations, uploads audio, fills title/transcript, attempts to set Accessibility language, saves, and waits for a saved confirmation.
- Logs processed, skipped, already-complete, missing-transcript, missing-audio, upload-failed, save-failed, and failed states.

## Build Transcript Data

The committed `data/transcripts.json` is generated from the public Eastern State transcript PDFs linked from the Accessibility page.

To regenerate it:

```sh
cd bloomberg-audio-assistant
npm run build:transcripts
```

This requires `pdftotext` from Poppler. On macOS, install it with:

```sh
brew install poppler
```

The parser is intentionally conservative about section boundaries, but it does not reject a translated stop just because it differs from English. Some public translated PDFs legitimately diverge from the English transcript. The generated data keeps those sections and records first-speaker differences in `data/transcripts.json` under `alignmentWarnings` for review.

If a PDF section cannot be parsed at all, the extension marks that stop/language as `missingTranscript` during dry-run instead of guessing.

## Configure

Default config lives in `config.default.json`. Copy `config.example.json` to `config.json` if you need a test allowlist, different transcript URLs, or to re-enable a paused language.

For a one-stop smoke test, set:

```json
"stopAllowlist": [18]
```

Do not add credentials or private local paths. The extension uses the current Chrome login session for CMS access. Local audio folder access is granted through Chrome's folder picker and stored as a browser-managed folder handle, not as a committed path.

Spanish is paused by default:

```json
{
  "key": "spanish",
  "enabled": false
}
```

## Install

1. Open Chrome and go to `chrome://extensions/`.
2. Enable Developer mode.
3. Click Load unpacked.
4. Select this folder: `bloomberg-audio-assistant`.
5. Open the Bloomberg Connects CMS Audios catalog while logged in.

## Run

1. Click the extension icon or use the injected panel on the CMS page.
2. Optional but recommended: click `Choose Local Audio Folder` and select the folder containing already-downloaded CMS audio files. The automation matches audio by stop number in the filename, such as `018 Synagogue.wav`.
3. Click `Download Missing Audio` to fill that folder from the CMS source-audio route. The extension streams the audio into the selected folder without triggering Chrome's normal Downloads-folder workflow. It also normalizes matching local filenames by removing CMS wrapper text such as `easternStatePenitentiary_UNAPPROVED_` and `_sourceUrl`.
4. Click `Scan / Dry Run`.
5. Review counts and logs. Fix missing transcript/config issues before saving.
6. Click `Start` only when ready to make CMS changes.
7. Use `Stop` to pause after the current state checkpoint.
8. Use `Resume` after a stopped or paused run.
9. Use `Export Log` from the popup to download the current run state/log.

The extension confirms before saving CMS changes. It also rechecks the edit modal before adding a language, so reruns should skip translations that are already complete.

## Validation

Run tests:

```sh
cd bloomberg-audio-assistant
npm test
```

Recommended manual validation:

1. Set `stopAllowlist` to `[18]`.
2. Reload the unpacked extension.
3. Open the CMS Audios catalog.
4. Run `Scan / Dry Run` and confirm stop 18 appears.
5. Run `Repair Item Index` and confirm unresolved route warnings are gone before save mode.
6. Run `Start` and verify one enabled language first, preferably Portuguese or French.
7. Run the scan again and confirm the completed language is skipped.

## Notes

- Save mode uploads from the selected local folder only. Run `Download Missing Audio` first if a stop is missing locally.
- Save mode now prefers direct CMS edit URLs from the stored audio item index. If a run pauses on route lookup, run `Repair Item Index` from the panel or popup while on the Audios catalog, then resume.
- Chrome does not allow extensions to reuse a normal local folder path by string. Use `Choose Local Audio Folder` after reloading the extension if Chrome asks for access again. Folder download/rename requires Chrome's read/write folder permission.
- If Chrome opens a file picker, crashes while picking a language, audio files appear in the normal Chrome Downloads folder, the run appears to reuse the previous stop after saving, the assistant cannot find Add Translation even though English is visible, transcripts are truncated/contain audio-guide navigation prompts, an Add Translation dialog remains open after the target language is already active, completed translations are being skipped without edit-page verification, Add Translation cannot be found when the active language is not English, the catalog row cannot be found because CMS spacing differs from the file name, menu actions fail while another browser window is active, or save mode searches/scrolls for a row that was already scanned, reload the unpacked extension and the CMS tab and confirm the panel build label is `selector-v45-add-translation-confirm` or newer. Older builds used crash-prone native picker/debugger fallbacks, a narrower language selector scan, a stale page bridge, could continue on the previous edit page when the next catalog item did not have a CMS id yet, only searched inside the modal even when the language menu rendered elsewhere in the page, accepted partial transcript inserts, retried adding an already-active language, trusted stale local completion records from the catalog scan, tried to open the English selector even after another language was active, required exact file-name spacing, did not focus the CMS window before menu-driven actions, or treated visible catalog rows as the normal navigation path.
- Browser file upload controls vary by CMS release. If the upload control changes, dry-run still works, but save mode may pause with `uploadFailed`.
- The Accessibility language tab is filled on a best-effort basis because CMS field structure may differ by modal state. Missing accessibility controls are logged, but title/transcript/audio save remains the core validation path.
- If dry-run reports `missingTranscript`, do not start a save run for those items until the transcript source has been corrected.
