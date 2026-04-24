# Eastern State Alt Text Assistant

Chrome extension for auditing and updating image descriptions in the Eastern State ExpressionEngine file manager.

The extension targets the `General` file directory and uses an open ChatGPT tab to shorten or generate alt text. It does not use an API key or a backend service.

## What It Does

- Scans the Eastern State `General` upload directory.
- Skips image rows with an existing `Description` under 150 characters.
- Sends existing descriptions of 150+ characters to ChatGPT for shortening.
- Sends missing-description image previews to ChatGPT for generated alt text.
- Validates every returned description locally before saving.
- Pauses instead of saving if ChatGPT, image upload, validation, or CMS save confirmation fails.

## Install

1. Open Chrome and go to `chrome://extensions/`.
2. Enable Developer mode.
3. Click Load unpacked.
4. Select this folder.

## Use

1. Open the Eastern State CMS `General` file directory.
2. Open ChatGPT (`https://chatgpt.com/` or `https://chat.openai.com/`) in another tab and log in.
3. Click the extension icon.
4. Click Scan to dry-run the directory and review counts.
5. Click Start Saving when ready.
6. Confirm the browser prompt before the extension saves public CMS changes.

The extension processes one candidate at a time. If it pauses, check the progress log in the popup, fix the flagged item manually if needed, then scan again before continuing.

## Character Rule

The maximum saved `Description` is 149 characters, including spaces. Existing descriptions with fewer than 150 characters are left unchanged.

## Notes

- Scope is intentionally limited to the Eastern State `General` upload directory.
- Only the ExpressionEngine `Description` field is edited.
- The previous Auto-McGraw scripts remain in the repository history, but they are no longer loaded by the extension manifest.
