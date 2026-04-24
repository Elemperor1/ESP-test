if (!globalThis.__ESP_CHATGPT_BOOTSTRAPPED__) {
  globalThis.__ESP_CHATGPT_BOOTSTRAPPED__ = true;

const RESPONSE_TIMEOUT_MS = 180000;
const RESPONSE_STABLE_MS = 1800;

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || typeof message.type !== "string") {
    return false;
  }

  if (message.type === "esp-chatgpt:ping") {
    sendResponse({ ok: true });
    return false;
  }

  if (message.type !== "esp-chatgpt:requestAltText") {
    return false;
  }

  handleAltTextRequest(message.payload || {})
    .then((result) => sendResponse(result))
    .catch((error) => {
      sendResponse({ ok: false, error: error.message });
    });

  return true;
});

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function waitFor(predicate, timeout = 10000, interval = 150) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const timer = setInterval(() => {
      const value = predicate();
      if (value) {
        clearInterval(timer);
        resolve(value);
        return;
      }

      if (Date.now() - start > timeout) {
        clearInterval(timer);
        reject(new Error("Timed out waiting for ChatGPT UI."));
      }
    }, interval);
  });
}

function findComposer() {
  const selectors = [
    "#prompt-textarea",
    '[data-testid="prompt-textarea"]',
    'div[contenteditable="true"][role="textbox"]',
    '[role="textbox"][contenteditable="true"]',
    ".ProseMirror",
    "textarea",
  ];

  for (const selector of selectors) {
    const element = document.querySelector(selector);
    if (element) return element;
  }

  return null;
}

function findSendButton(composer) {
  const form = composer && typeof composer.closest === "function" ? composer.closest("form") : null;
  const selectors = [
    '[data-testid="send-button"]',
    'button[data-testid="send-button"]',
    'button[type="submit"]',
    '[role="button"][data-testid*="send"]',
    'button[data-testid*="send"]',
    '[data-testid*="composer-send"]',
    'button[aria-label="Send prompt"]',
    'button[aria-label="Send message"]',
    'button[aria-label*="Send"]',
    'button[title*="Send"]',
  ];

  for (const selector of selectors) {
    const button = (form && form.querySelector(selector)) || document.querySelector(selector);
    if (isUsableButton(button)) return button;
  }

  const buttons = Array.from((form || document).querySelectorAll("button, [role='button']"));
  const composerRect = composer && composer.getBoundingClientRect ? composer.getBoundingClientRect() : null;
  return buttons
    .reverse()
    .find((button) => {
      if (!isUsableButton(button)) return false;
      if (composerRect && button.getBoundingClientRect) {
        const rect = button.getBoundingClientRect();
        const nearComposer = Math.abs(rect.top - composerRect.bottom) < 220 && rect.left > composerRect.left - 80;
        if (!nearComposer) return false;
      }

      const label = `${button.getAttribute("aria-label") || ""} ${button.textContent || ""}`.toLowerCase();
      if (label.includes("send")) return true;
      if (label.includes("submit")) return true;

      const title = (button.getAttribute("title") || "").toLowerCase();
      return title.includes("send") || title.includes("submit");
    }) || null;
}

function isUsableButton(button) {
  if (!button) return false;
  if (button.getAttribute("hidden") !== null) return false;
  if (button.disabled) return false;
  if (button.getAttribute("aria-disabled") === "true") return false;
  return true;
}

function setNativeValue(input, value) {
  const prototype = Object.getPrototypeOf(input);
  const descriptor = Object.getOwnPropertyDescriptor(prototype, "value");

  if (descriptor && descriptor.set) {
    descriptor.set.call(input, value);
  } else {
    input.value = value;
  }
}

function setComposerText(composer, text) {
  composer.focus();

  if (composer instanceof HTMLTextAreaElement || composer instanceof HTMLInputElement) {
    setNativeValue(composer, text);
  } else {
    composer.innerHTML = "";
    const paragraph = document.createElement("p");
    paragraph.textContent = text;
    composer.appendChild(paragraph);
  }

  composer.dispatchEvent(new InputEvent("input", {
    bubbles: true,
    cancelable: true,
    inputType: "insertText",
    data: text,
  }));
  composer.dispatchEvent(new Event("change", { bubbles: true }));
}

function submitComposerWithEnter(composer) {
  composer.focus();
  const keyOptions = {
    key: "Enter",
    code: "Enter",
    which: 13,
    keyCode: 13,
    bubbles: true,
    cancelable: true,
  };

  composer.dispatchEvent(new KeyboardEvent("keydown", keyOptions));
  composer.dispatchEvent(new KeyboardEvent("keypress", keyOptions));
  composer.dispatchEvent(new KeyboardEvent("keyup", keyOptions));
}

function submitComposerWithForm(composer) {
  const form = composer.closest("form");
  if (!form) return false;

  if (typeof form.requestSubmit === "function") {
    form.requestSubmit();
    return true;
  }

  form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  return true;
}

async function submitPrompt(composer, assistantSnapshotBefore) {
  const wasSubmitted = () => {
    if (isGenerating()) return true;
    return hasAssistantProgress(assistantSnapshotBefore, getAssistantSnapshot());
  };
  const attempts = [
    () => {
      const sendButton = findSendButton(composer);
      if (!sendButton) return false;
      sendButton.click();
      return true;
    },
    () => submitComposerWithForm(composer),
    () => {
      submitComposerWithEnter(composer);
      return true;
    },
  ];

  for (const attempt of attempts) {
    attempt();
    await delay(550);
    if (wasSubmitted()) {
      return;
    }
  }

  throw new Error("Could not submit the prompt in ChatGPT. The prompt is prepared but not sending.");
}

function dataUrlToFile(dataUrl, filename, mimeType) {
  const base64 = dataUrl.split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return new File([bytes], filename, { type: mimeType || "image/jpeg" });
}

function findFileInput() {
  const inputs = Array.from(document.querySelectorAll('input[type="file"]'));
  return inputs.find((input) => {
    const accept = input.getAttribute("accept") || "";
    return !accept || accept.includes("image") || accept.includes("*");
  }) || inputs[0] || null;
}

function findAttachButton() {
  const selectors = [
    'button[aria-label*="Attach"]',
    'button[aria-label*="Upload"]',
    '[data-testid="composer-plus-btn"]',
    'button[data-testid*="attach"]',
  ];

  for (const selector of selectors) {
    const button = document.querySelector(selector);
    if (isUsableButton(button)) return button;
  }

  return null;
}

async function attachImageToChat(payload) {
  if (!payload.imageDataUrl) {
    throw new Error("No image data was provided for missing alt text.");
  }

  const initialBlobPreviewCount = Array.from(document.querySelectorAll("img"))
    .filter((image) => {
      const source = image.currentSrc || image.src || "";
      return source.startsWith("blob:") || source.startsWith("data:");
    }).length;
  let input = findFileInput();
  if (!input) {
    const attachButton = findAttachButton();
    if (attachButton) {
      attachButton.click();
      await delay(300);
      input = findFileInput();
    }
  }

  if (!input) {
    throw new Error("Could not find ChatGPT's image upload control.");
  }

  const file = dataUrlToFile(
    payload.imageDataUrl,
    payload.fileName || "eastern-state-image.jpg",
    payload.imageMimeType || "image/jpeg"
  );
  const transfer = new DataTransfer();
  transfer.items.add(file);
  input.files = transfer.files;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));

  await waitFor(() => {
    const images = Array.from(document.querySelectorAll("img"));
    const blobPreviewCount = images.filter((image) => {
      const source = image.currentSrc || image.src || "";
      return source.startsWith("blob:") || source.startsWith("data:");
    }).length;
    const hasBlobPreview = blobPreviewCount > initialBlobPreviewCount;
    const hasAttachmentText = document.body.textContent.includes(file.name);
    return hasBlobPreview || hasAttachmentText;
  }, 15000, 300);
}

function buildPrompt(payload) {
  const prefix = "Return only the alt text. Do not add quotation marks, labels, Markdown, or explanation.";
  const lengthRule = `It must be fewer than ${payload.maxLength + 1 || 150} characters including spaces.`;
  const styleRule = "Use concise, objective, plain descriptive language. Avoid 'image of' or 'photo of' unless useful.";
  const context = [
    payload.title ? `Title: ${payload.title}` : "",
    payload.fileName ? `File name: ${payload.fileName}` : "",
  ].filter(Boolean).join("\n");

  if (payload.action === "shorten") {
    return [
      `Rewrite this alt text. ${lengthRule}`,
      styleRule,
      prefix,
      context,
      `Current alt text: ${payload.existingText || ""}`,
      payload.previousOutput ? `Previous output was invalid: ${payload.previousOutput}` : "",
    ].filter(Boolean).join("\n\n");
  }

  return [
    `Generate alt text for the attached image. ${lengthRule}`,
    styleRule,
    prefix,
    payload.imageUrl ? `Image link: ${payload.imageUrl}` : "",
    context,
    payload.previousOutput ? `Previous output was invalid: ${payload.previousOutput}` : "",
  ].filter(Boolean).join("\n\n");
}

function getMessageSignature(message) {
  if (!message) return "";

  return (
    message.getAttribute("data-message-id") ||
    message.getAttribute("data-testid") ||
    message.id ||
    ""
  );
}

function getAssistantMessages() {
  const selectors = [
    '[data-message-author-role="assistant"]',
    '[data-testid^="conversation-turn-"] [data-message-author-role="assistant"]',
    'main article[data-message-author-role="assistant"]',
  ];

  const seen = new Set();
  const messages = [];

  for (const selector of selectors) {
    document.querySelectorAll(selector).forEach((element) => {
      if (!seen.has(element)) {
        seen.add(element);
        messages.push(element);
      }
    });
  }

  return messages;
}

function getAssistantSnapshot() {
  const messages = getAssistantMessages();
  const latestElement = messages[messages.length - 1] || null;
  const latestText = latestElement ? cleanAltText(latestElement.innerText || latestElement.textContent || "") : "";

  return {
    count: messages.length,
    latestElement,
    latestSignature: getMessageSignature(latestElement),
    latestText,
  };
}

function hasAssistantProgress(previousSnapshot, currentSnapshot) {
  if (currentSnapshot.count > previousSnapshot.count) return true;

  if (
    currentSnapshot.latestSignature &&
    previousSnapshot.latestSignature &&
    currentSnapshot.latestSignature !== previousSnapshot.latestSignature
  ) {
    return true;
  }

  if (
    currentSnapshot.latestText &&
    previousSnapshot.latestText &&
    currentSnapshot.latestText !== previousSnapshot.latestText
  ) {
    return true;
  }

  return false;
}

function isGenerating() {
  return Boolean(
    document.querySelector('[data-testid="stop-button"], button[aria-label*="Stop"], button[aria-label*="Cancel"]')
  );
}

function cleanAltText(rawText) {
  let text = String(rawText || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/```[\s\S]*?```/g, (match) => match.replace(/```(?:text|json)?/gi, "").replace(/```/g, ""))
    .trim();

  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      const parsed = JSON.parse(jsonMatch[0]);
      text = parsed.alt || parsed.altText || parsed.answer || parsed.description || text;
    } catch (error) {
      // Keep the plain text if parsing fails.
    }
  }

  text = text
    .replace(/^(?:alt text|description|answer)\s*:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();

  if ((text.startsWith('"') && text.endsWith('"')) || (text.startsWith("'") && text.endsWith("'"))) {
    text = text.slice(1, -1).trim();
  }

  return text;
}

async function waitForAssistantAltText(assistantSnapshotBefore) {
  const startedAt = Date.now();
  let lastText = "";
  let lastChangedAt = Date.now();

  while (Date.now() - startedAt < RESPONSE_TIMEOUT_MS) {
    await delay(700);

    const messages = getAssistantMessages();
    const latest = messages[messages.length - 1];
    if (!latest) {
      continue;
    }

    const latestSignature = getMessageSignature(latest);
    const candidate = cleanAltText(latest.innerText || latest.textContent || "");
    const currentSnapshot = {
      count: messages.length,
      latestSignature,
      latestText: candidate,
    };
    const hasNewAssistantTurn = hasAssistantProgress(assistantSnapshotBefore, currentSnapshot);

    if (!hasNewAssistantTurn) {
      continue;
    }

    if (!candidate) {
      continue;
    }

    if (candidate !== lastText) {
      lastText = candidate;
      lastChangedAt = Date.now();
      continue;
    }

    if (!isGenerating() && Date.now() - lastChangedAt >= RESPONSE_STABLE_MS) {
      return lastText;
    }
  }

  throw new Error("Timed out waiting for ChatGPT to return alt text.");
}

async function handleAltTextRequest(payload) {
  const composer = await waitFor(findComposer, 15000, 250);
  const assistantSnapshotBefore = getAssistantSnapshot();

  if (payload.action === "generate") {
    await attachImageToChat(payload);
  }

  setComposerText(composer, buildPrompt(payload));
  await delay(350);
  await submitPrompt(composer, assistantSnapshotBefore);
  const altText = await waitForAssistantAltText(assistantSnapshotBefore);

  return {
    ok: true,
    altText,
  };
}

// Expose a direct callable API for scripting.executeScript fallback mode.
globalThis.__ESP_CHATGPT_READY__ = true;
globalThis.__ESP_CHATGPT_REQUEST_ALT_TEXT__ = handleAltTextRequest;
}
