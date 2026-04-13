import * as pdfjsLib from "./vendor/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdf.worker.mjs", import.meta.url).toString();

const APP_ASSET_VERSION = "20260413n";
const STORAGE_KEY = "rowcolpage.v4.settings.v3";
const SELECTOR_CHANNEL_NAME = "rowcolpage-v4-selector";
const SELECTOR_DB_NAME = "rowcolpage-v4";
const SELECTOR_DB_STORE = "selectorPayload";
const SELECTOR_DB_KEY = "latest";
const PDF_RENDER_SCALE = 2.2;
const OCR_LANGUAGE = "chi_tra+eng";
const OCR_MODULE_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.esm.min.js";
const DEFAULT_SAMPLE_FILE_NAME = "範例.pdf";
const CONTENT_RGB_THRESHOLD = 244;
const CONTENT_ALPHA_THRESHOLD = 18;
const CONTENT_MARGIN_PX = 28;
const CROPPED_CONTENT_MARGIN_PX = 32;
const MIN_CROP_SIZE = 24;
const TEXT_BOUNDS_HORIZONTAL_PADDING = 24;
const TEXT_BOUNDS_VERTICAL_PADDING = 12;

const DEFAULT_SETTINGS = {
  title: "大南六甲",
  className: "",
  studentName: "",
  date: new Date().toISOString().slice(0, 10),
  startNumber: 1,
  pageCount: 1,
  columnCount: 2,
  rowCount: 4,
  guideMode: "horizontal",
  showSignature: true,
  titlePatterns: "^[（(]\\s*[）)]\\s*\\d+\\s*[.．、]\n^\\d+\\s*[.．、]\n^第\\s*\\d+\\s*題",
  leftPercent: 4,
  rightPercent: 96,
  topPadding: 8,
  bottomPadding: 12,
  nextGap: 10,
};

const uploadedImages = new Map();
const uploadedQuestionTexts = new Map();
const cellBindings = new Map();
let activePasteImageIndex = null;
let sourceFiles = [];
let currentSourceLabel = "";
let extractedQuestions = [];
let latestSourcePages = [];
let selectorWindow = null;
let selectorWindowReady = false;
let ocrWorkerPromise = null;
let demoBootstrapPromise = null;
const selectorChannel =
  typeof window.BroadcastChannel === "function"
    ? new BroadcastChannel(SELECTOR_CHANNEL_NAME)
    : null;
let selectorDbPromise = null;
const sourcePageCanvasCache = new Map();

const titleInput = document.querySelector("#titleInput");
const classInput = document.querySelector("#classInput");
const nameInput = document.querySelector("#nameInput");
const dateInput = document.querySelector("#dateInput");
const startNumberInput = document.querySelector("#startNumberInput");
const pageCountInput = document.querySelector("#pageCountInput");
const rowCountInput = document.querySelector("#rowCountInput");
const guideSelect = document.querySelector("#guideSelect");
const signatureToggle = document.querySelector("#signatureToggle");
const resetButton = document.querySelector("#resetButton");
const printButton = document.querySelector("#printButton");
const pagesRoot = document.querySelector("#pages");
const pageTemplate = document.querySelector("#pageTemplate");
const cellTemplate = document.querySelector("#cellTemplate");

const sourceFileInput = document.querySelector("#sourceFileInput");
const titlePatternsInput = document.querySelector("#titlePatternsInput");
const leftPercentInput = document.querySelector("#leftPercentInput");
const rightPercentInput = document.querySelector("#rightPercentInput");
const topPaddingInput = document.querySelector("#topPaddingInput");
const bottomPaddingInput = document.querySelector("#bottomPaddingInput");
const nextGapInput = document.querySelector("#nextGapInput");
const extractSourceButton = document.querySelector("#extractPdfButton");
const clearExtractedButton = document.querySelector("#clearExtractedButton");
const extractionSummary = document.querySelector("#extractionSummary");
const questionPreviewList = document.querySelector("#questionPreviewList");
const questionReviewActions = document.querySelector("#questionReviewActions");
const questionReviewHint = document.querySelector("#questionReviewHint");
const confirmExtractedButton = document.querySelector("#confirmExtractedButton");
const previewModal = document.querySelector("#previewModal");
const previewModalSource = document.querySelector("#previewModalSource");
const previewModalTitle = document.querySelector("#previewModalTitle");
const previewModalAnchor = document.querySelector("#previewModalAnchor");
const previewModalSourcePreview = document.querySelector("#previewModalSourcePreview");
const previewModalQuestionPreview = document.querySelector("#previewModalQuestionPreview");
const previewModalCloseButton = document.querySelector("#previewModalCloseButton");
const previewZoomOutButton = document.querySelector("#previewZoomOutButton");
const previewZoomResetButton = document.querySelector("#previewZoomResetButton");
const previewZoomInButton = document.querySelector("#previewZoomInButton");
const previewZoomValue = document.querySelector("#previewZoomValue");
const previewResetCropButton = document.querySelector("#previewResetCropButton");
const previewMoveUpButton = document.querySelector("#previewMoveUpButton");
const previewMoveLeftButton = document.querySelector("#previewMoveLeftButton");
const previewMoveRightButton = document.querySelector("#previewMoveRightButton");
const previewMoveDownButton = document.querySelector("#previewMoveDownButton");
const previewWidenButton = document.querySelector("#previewWidenButton");
const previewHeightenButton = document.querySelector("#previewHeightenButton");
const previewNarrowButton = document.querySelector("#previewNarrowButton");
const previewShortenButton = document.querySelector("#previewShortenButton");
const previewRunOcrButton = document.querySelector("#previewRunOcrButton");
const previewReviewStatus = document.querySelector("#previewReviewStatus");
const previewTextEditor = document.querySelector("#previewTextEditor");
const previewSaveButton = document.querySelector("#previewSaveButton");
let previewQuestionZoom = 1;
let previewQuestionStage = null;
let previewSourceStage = null;
let previewCropBox = null;
let previewQuestionId = null;
let previewOriginalBounds = null;
let previewDraftBounds = null;
let previewDraftText = "";
let previewRecognizedText = "";
let previewHasRecognizedText = false;
let previewCropDirty = false;
let previewSyncPromise = null;
let previewPointerState = null;
let previewPointerListenersBound = false;
const PREVIEW_NUDGE_STEP = 6;
const PREVIEW_RESIZE_STEP = 8;

function clampNumber(value, min, max, fallback) {
  const parsed = Number.parseFloat(value);

  if (Number.isNaN(parsed)) {
    return fallback;
  }

  return Math.min(Math.max(parsed, min), max);
}

function formatDisplayDate(dateValue) {
  if (!dateValue) {
    return "____ / ____ / ____";
  }

  const [year, month, day] = dateValue.split("-");

  if (!year || !month || !day) {
    return "____ / ____ / ____";
  }

  return `${year} / ${month} / ${day}`;
}

function withFallback(value, fallback) {
  return value.trim() || fallback;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeText(value) {
  return value.replace(/\s+/g, " ").trim();
}

function formatQuestionTextSpacing(value) {
  return value.replace(/\(\s*\)\s*(\d+\s*[.．、])/gu, "(   )$1")
    .replace(/（\s*）\s*(\d+\s*[.．、])/gu, "（   ）$1");
}

function appendPlainRichText(container, text) {
  const parts = String(text ?? "").split("\n");

  parts.forEach((part, index) => {
    if (index > 0) {
      container.appendChild(document.createElement("br"));
    }

    if (part) {
      container.appendChild(document.createTextNode(part));
    }
  });
}

function createFractionElement(numerator, denominator) {
  const fraction = document.createElement("span");
  const numeratorElement = document.createElement("span");
  const denominatorElement = document.createElement("span");

  fraction.className = "math-fraction";
  fraction.dataset.latex = `\\frac{${numerator}}{${denominator}}`;
  fraction.contentEditable = "false";
  numeratorElement.className = "math-fraction-numerator";
  denominatorElement.className = "math-fraction-denominator";
  numeratorElement.textContent = numerator;
  denominatorElement.textContent = denominator;
  fraction.append(numeratorElement, denominatorElement);

  return fraction;
}

function renderRichQuestionText(container, text) {
  const source = String(text ?? "");
  const fractionPattern = /\\frac\{([^{}]+)\}\{([^{}]+)\}/g;
  let lastIndex = 0;
  let match = fractionPattern.exec(source);

  container.replaceChildren();

  while (match) {
    appendPlainRichText(container, source.slice(lastIndex, match.index));
    container.appendChild(createFractionElement(match[1], match[2]));
    lastIndex = match.index + match[0].length;
    match = fractionPattern.exec(source);
  }

  appendPlainRichText(container, source.slice(lastIndex));
}

function getRichQuestionText(container) {
  if (!container) {
    return "";
  }

  const readNode = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent ?? "";
    }

    if (node.nodeName === "BR") {
      return "\n";
    }

    if (node.nodeType !== Node.ELEMENT_NODE) {
      return "";
    }

    if (node.classList.contains("math-fraction")) {
      return node.dataset.latex ?? "";
    }

    return Array.from(node.childNodes).map(readNode).join("");
  };

  return Array.from(container.childNodes).map(readNode).join("");
}

function getQuestionEditorDisplay() {
  return previewTextEditor?.querySelector(".math-question-display") ?? null;
}

function getQuestionEditorSource() {
  return previewTextEditor?.querySelector(".math-question-source") ?? null;
}

function setQuestionEditorText(text) {
  const display = getQuestionEditorDisplay();
  const source = getQuestionEditorSource();

  if (!display || !source || !previewTextEditor) {
    return;
  }

  previewTextEditor.classList.remove("is-source-editing");
  source.value = normalizeMultilineText(text);
  renderRichQuestionText(display, source.value);
}

function getQuestionEditorText() {
  const source = getQuestionEditorSource();
  const display = getQuestionEditorDisplay();

  if (previewTextEditor?.classList.contains("is-source-editing") && source) {
    return normalizeMultilineText(source.value);
  }

  return normalizeMultilineText(getRichQuestionText(display));
}

function isQuestionEditorDisabled() {
  return !previewTextEditor || previewTextEditor.getAttribute("aria-disabled") === "true";
}

function showQuestionEditorSource() {
  const source = getQuestionEditorSource();

  if (isQuestionEditorDisabled() || previewTextEditor.classList.contains("is-source-editing")) {
    return;
  }

  const sourceText = getQuestionEditorText();
  source.value = sourceText;
  previewTextEditor.classList.add("is-source-editing");
  window.setTimeout(() => {
    source.focus();
    source.setSelectionRange(source.value.length, source.value.length);
  }, 0);
}

function renderQuestionEditorDisplay() {
  const display = getQuestionEditorDisplay();
  const source = getQuestionEditorSource();

  if (!previewTextEditor || !previewTextEditor.classList.contains("is-source-editing")) {
    return;
  }

  previewDraftText = getQuestionEditorText();
  previewTextEditor.classList.remove("is-source-editing");
  source.value = previewDraftText;
  renderRichQuestionText(display, previewDraftText);
  updatePreviewReviewStatus();
}

function setQuestionEditorDisabled(isDisabled) {
  if (!previewTextEditor) {
    return;
  }

  if (isDisabled) {
    renderQuestionEditorDisplay();
  }

  const source = getQuestionEditorSource();

  if (source) {
    source.disabled = isDisabled;
  }

  previewTextEditor.setAttribute("aria-disabled", String(isDisabled));
  previewTextEditor.classList.toggle("is-disabled", isDisabled);
}

function formatStackedMathItemsAsLatex(items) {
  if (!items.length) {
    return "";
  }

  if (items.length === 1) {
    return items[0].text;
  }

  const [numerator, denominator, ...rest] = items.map((item) => item.text);
  const fraction = `\\frac{${numerator}}{${denominator}}`;

  return rest.length ? `${fraction}${rest.join("")}` : fraction;
}

function normalizeMultilineText(value) {
  return String(value ?? "")
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => formatQuestionTextSpacing(line.replace(/\s+/g, " ").trim()))
    .filter(Boolean)
    .join("\n");
}

function findQuestionById(questionId) {
  return extractedQuestions.find((question) => question.id === questionId) ?? null;
}

function getQuestionIndexById(questionId) {
  return extractedQuestions.findIndex((question) => question.id === questionId);
}

function getFirstPendingQuestion() {
  return extractedQuestions.find((question) => !question.isConfirmed) ?? null;
}

function getNextPendingQuestion(afterQuestionId = null) {
  if (!extractedQuestions.length) {
    return null;
  }

  const startIndex = afterQuestionId ? getQuestionIndexById(afterQuestionId) + 1 : 0;

  for (let index = startIndex; index < extractedQuestions.length; index += 1) {
    if (!extractedQuestions[index].isConfirmed) {
      return extractedQuestions[index];
    }
  }

  return getFirstPendingQuestion();
}

function getPendingReviewCount() {
  return extractedQuestions.filter((question) => !question.isConfirmed).length;
}

function getSignatureVisible() {
  return signatureToggle.getAttribute("aria-pressed") === "true";
}

function hasValidTitlePatterns(rawPatterns) {
  const lines = String(rawPatterns)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    return false;
  }

  return lines.every((line) => {
    try {
      new RegExp(line, "iu");
      return true;
    } catch {
      return !/[\\^$.*+?()[\]{}|]/.test(line);
    }
  });
}

function sanitizeSettings(rawSettings = {}) {
  const nextSettings = { ...DEFAULT_SETTINGS, ...rawSettings };

  if (!hasValidTitlePatterns(nextSettings.titlePatterns)) {
    nextSettings.titlePatterns = DEFAULT_SETTINGS.titlePatterns;
  }

  nextSettings.startNumber = clampNumber(nextSettings.startNumber, 1, 1000000, DEFAULT_SETTINGS.startNumber);
  nextSettings.pageCount = clampNumber(nextSettings.pageCount, 1, 50, DEFAULT_SETTINGS.pageCount);
  nextSettings.rowCount = clampNumber(nextSettings.rowCount, 1, 12, DEFAULT_SETTINGS.rowCount);
  nextSettings.leftPercent = clampNumber(nextSettings.leftPercent, 0, 95, DEFAULT_SETTINGS.leftPercent);
  nextSettings.rightPercent = clampNumber(nextSettings.rightPercent, 5, 100, DEFAULT_SETTINGS.rightPercent);
  nextSettings.topPadding = clampNumber(nextSettings.topPadding, 0, 120, DEFAULT_SETTINGS.topPadding);
  nextSettings.bottomPadding = clampNumber(nextSettings.bottomPadding, 0, 180, DEFAULT_SETTINGS.bottomPadding);
  nextSettings.nextGap = clampNumber(nextSettings.nextGap, 0, 120, DEFAULT_SETTINGS.nextGap);
  nextSettings.guideMode = ["horizontal", "dot", "none"].includes(nextSettings.guideMode)
    ? nextSettings.guideMode
    : DEFAULT_SETTINGS.guideMode;
  nextSettings.showSignature = Boolean(nextSettings.showSignature);

  return nextSettings;
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);

    if (!raw) {
      return sanitizeSettings();
    }

    return sanitizeSettings(JSON.parse(raw));
  } catch {
    return sanitizeSettings();
  }
}

function getSelectorDb() {
  if (!("indexedDB" in window)) {
    return Promise.resolve(null);
  }

  if (!selectorDbPromise) {
    selectorDbPromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(SELECTOR_DB_NAME, 1);

      request.onupgradeneeded = () => {
        const db = request.result;

        if (!db.objectStoreNames.contains(SELECTOR_DB_STORE)) {
          db.createObjectStore(SELECTOR_DB_STORE);
        }
      };

      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    }).catch(() => null);
  }

  return selectorDbPromise;
}

async function writeSelectorPayload(payload) {
  const db = await getSelectorDb();

  if (!db) {
    return;
  }

  await new Promise((resolve, reject) => {
    const transaction = db.transaction(SELECTOR_DB_STORE, "readwrite");
    const store = transaction.objectStore(SELECTOR_DB_STORE);

    if (payload) {
      store.put(payload, SELECTOR_DB_KEY);
    } else {
      store.delete(SELECTOR_DB_KEY);
    }

    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error);
    transaction.onabort = () => reject(transaction.error);
  }).catch(() => {});
}

function saveSettings(settings) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

function applySettings(settings) {
  titleInput.value = settings.title;
  if (classInput) {
    classInput.value = settings.className ?? "";
  }
  nameInput.value = settings.studentName;
  dateInput.value = settings.date;
  startNumberInput.value = settings.startNumber;
  pageCountInput.value = settings.pageCount;
  rowCountInput.value = settings.rowCount;
  guideSelect.value = settings.guideMode;
  titlePatternsInput.value = settings.titlePatterns;
  leftPercentInput.value = settings.leftPercent;
  rightPercentInput.value = settings.rightPercent;
  topPaddingInput.value = settings.topPadding;
  bottomPaddingInput.value = settings.bottomPadding;
  nextGapInput.value = settings.nextGap;
  updateSignatureVisibility(settings.showSignature);
}

function collectSettings() {
  return {
    title: titleInput.value.trim() || DEFAULT_SETTINGS.title,
    className: classInput ? classInput.value.trim() : "",
    studentName: nameInput.value.trim(),
    date: dateInput.value,
    startNumber: clampNumber(startNumberInput.value, 1, 1000000, DEFAULT_SETTINGS.startNumber),
    pageCount: clampNumber(pageCountInput.value, 1, 50, DEFAULT_SETTINGS.pageCount),
    columnCount: 2,
    rowCount: clampNumber(rowCountInput.value, 1, 12, DEFAULT_SETTINGS.rowCount),
    guideMode: guideSelect.value,
    showSignature: getSignatureVisible(),
    titlePatterns: titlePatternsInput.value.trim() || DEFAULT_SETTINGS.titlePatterns,
    leftPercent: clampNumber(leftPercentInput.value, 0, 95, DEFAULT_SETTINGS.leftPercent),
    rightPercent: clampNumber(rightPercentInput.value, 5, 100, DEFAULT_SETTINGS.rightPercent),
    topPadding: clampNumber(topPaddingInput.value, 0, 120, DEFAULT_SETTINGS.topPadding),
    bottomPadding: clampNumber(bottomPaddingInput.value, 0, 180, DEFAULT_SETTINGS.bottomPadding),
    nextGap: clampNumber(nextGapInput.value, 0, 120, DEFAULT_SETTINGS.nextGap),
  };
}

function persistSettings() {
  saveSettings(collectSettings());
}

function updateGuideMode(guideMode) {
  pagesRoot.className = "pages";

  if (guideMode === "none") {
    pagesRoot.classList.add("no-guides");
    return;
  }

  pagesRoot.classList.add(`guide-${guideMode}`);
}

function updateSignatureVisibility(showSignature) {
  pagesRoot.classList.toggle("show-signature", showSignature);
  pagesRoot.classList.toggle("hide-signature", !showSignature);
  signatureToggle.textContent = showSignature ? "顯示" : "隱藏";
  signatureToggle.setAttribute("aria-pressed", String(showSignature));
}

function readBlobAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

function getImageFileFromClipboardData(clipboardData) {
  const items = Array.from(clipboardData?.items ?? []);

  for (const item of items) {
    if (item.kind === "file" && item.type.startsWith("image/")) {
      return item.getAsFile();
    }
  }

  return null;
}

async function readClipboardImage() {
  if (!navigator.clipboard?.read) {
    return null;
  }

  const clipboardItems = await navigator.clipboard.read();

  for (const clipboardItem of clipboardItems) {
    const imageType = clipboardItem.types.find((type) => type.startsWith("image/"));

    if (imageType) {
      return clipboardItem.getType(imageType);
    }
  }

  return null;
}

function renderCellContent(container, pane, imageUrl, text = "") {
  container.replaceChildren();
  container.classList.remove("is-empty");
  pane.classList.toggle("has-question-text", Boolean(text));
  pane.classList.toggle("has-image", Boolean(imageUrl));

  if (!imageUrl && !text) {
    container.classList.add("is-empty");
    return;
  }

  const stack = document.createElement("div");
  stack.className = "question-content-stack";

  if (text) {
    const problemText = document.createElement("div");
    problemText.className = "problem-text";
    renderRichQuestionText(problemText, normalizeMultilineText(text));
    stack.appendChild(problemText);
  }

  if (imageUrl) {
    const image = document.createElement("img");
    image.className = "problem-image";
    image.alt = "題目圖片";
    image.src = imageUrl;
    stack.appendChild(image);
  }

  container.appendChild(stack);
}

function updatePasteTargetState() {
  cellBindings.forEach(({ pane, pasteButton }, imageIndex) => {
    const isActive = imageIndex === activePasteImageIndex;
    pane.classList.toggle("is-paste-target", isActive);
    pasteButton.classList.toggle("is-active", isActive);
  });
}

function setActivePasteTarget(imageIndex) {
  activePasteImageIndex = imageIndex;
  updatePasteTargetState();
}

async function applyImageToCell(imageIndex, blobOrFile) {
  const imageUrl = await readBlobAsDataUrl(blobOrFile);
  uploadedImages.set(imageIndex, imageUrl);

  const binding = cellBindings.get(imageIndex);

  if (!binding) {
    return;
  }

  uploadedQuestionTexts.delete(imageIndex);
  renderCellContent(binding.content, binding.pane, imageUrl, "");
  binding.clearButton.hidden = false;
}

function bindCellUpload(cell, imageIndex) {
  const uploadInputs = cell.querySelectorAll(".cell-upload-input");
  const pasteButton = cell.querySelector(".cell-paste-button");
  const clearButton = cell.querySelector(".cell-clear-button");
  const content = cell.querySelector(".cell-content");
  const pane = cell.querySelector(".question-pane");

  pane.tabIndex = 0;
  pane.setAttribute("title", "點一下這格後可直接按 Ctrl+V 貼上截圖");

  cellBindings.set(imageIndex, {
    pane,
    content,
    clearButton,
    pasteButton,
  });

  renderCellContent(content, pane, uploadedImages.get(imageIndex) ?? "", uploadedQuestionTexts.get(imageIndex) ?? "");
  clearButton.hidden = !uploadedImages.has(imageIndex) && !uploadedQuestionTexts.has(imageIndex);
  updatePasteTargetState();

  uploadInputs.forEach((uploadInput) => {
    uploadInput.addEventListener("change", async () => {
      const [file] = uploadInput.files ?? [];

      if (!file) {
        return;
      }

      try {
        setActivePasteTarget(imageIndex);
        await applyImageToCell(imageIndex, file);
      } finally {
        uploadInput.value = "";
      }
    });
  });

  const activatePasteTarget = () => {
    setActivePasteTarget(imageIndex);
  };

  cell.addEventListener("click", activatePasteTarget);
  pane.addEventListener("focus", activatePasteTarget);

  pane.addEventListener("paste", async (event) => {
    const imageFile = getImageFileFromClipboardData(event.clipboardData);

    if (!imageFile) {
      return;
    }

    event.preventDefault();
    setActivePasteTarget(imageIndex);
    await applyImageToCell(imageIndex, imageFile);
  });

  pasteButton.addEventListener("click", async () => {
    setActivePasteTarget(imageIndex);

    try {
      const imageBlob = await readClipboardImage();

      if (!imageBlob) {
        window.alert("剪貼簿中沒有圖片，請先複製截圖後再貼上。");
        return;
      }

      await applyImageToCell(imageIndex, imageBlob);
    } catch (error) {
      console.error(error);
      window.alert("目前瀏覽器沒有允許讀取剪貼簿圖片，請改用 Ctrl+V 貼上，或使用上傳題圖。");
    }
  });

  clearButton.addEventListener("click", () => {
    uploadedImages.delete(imageIndex);
    uploadedQuestionTexts.delete(imageIndex);
    renderCellContent(content, pane, "", "");
    clearButton.hidden = true;
  });
}

function updateReviewActionsVisibility() {
  if (!questionReviewActions) {
    return;
  }

  const pendingCount = getPendingReviewCount();
  const totalCount = extractedQuestions.length;
  questionReviewActions.hidden = !totalCount;

  if (questionReviewHint) {
    const firstPendingQuestion = getFirstPendingQuestion();
    questionReviewHint.textContent = pendingCount
      ? `目前共有 ${totalCount} 題，請從 ${firstPendingQuestion?.numberLabel ?? "第一題"} 開始逐題確認；每儲存一題後會自動前往下一題。`
      : `共 ${totalCount} 題都已確認完成，現在可以前往下一頁勾選並準備列印。`;
  }

  if (confirmExtractedButton) {
    confirmExtractedButton.disabled = pendingCount > 0;
  }
}

function clampPreviewZoom(value) {
  return Math.min(Math.max(value, 0.5), 3);
}

function updatePreviewZoomDisplay() {
  if (!previewQuestionStage) {
    return;
  }

  previewQuestionStage.style.width = `${previewQuestionZoom * 100}%`;

  if (previewZoomValue) {
    previewZoomValue.textContent = `${Math.round(previewQuestionZoom * 100)}%`;
  }
}

function clampQuestionBounds(bounds, pageWidth, pageHeight) {
  const left = clampNumber(bounds.left, 0, Math.max(0, pageWidth - MIN_CROP_SIZE), 0);
  const top = clampNumber(bounds.top, 0, Math.max(0, pageHeight - MIN_CROP_SIZE), 0);
  const right = clampNumber(bounds.right, left + MIN_CROP_SIZE, pageWidth, pageWidth);
  const bottom = clampNumber(bounds.bottom, top + MIN_CROP_SIZE, pageHeight, pageHeight);

  return {
    left,
    top,
    right,
    bottom,
  };
}

function expandBoundsForText(bounds, pageWidth, pageHeight) {
  return clampQuestionBounds(
    {
      left: bounds.left - TEXT_BOUNDS_HORIZONTAL_PADDING,
      right: bounds.right + TEXT_BOUNDS_HORIZONTAL_PADDING,
      top: bounds.top - TEXT_BOUNDS_VERTICAL_PADDING,
      bottom: bounds.bottom + TEXT_BOUNDS_VERTICAL_PADDING,
    },
    pageWidth,
    pageHeight,
  );
}

function resetPreviewDraftState() {
  previewQuestionId = null;
  previewOriginalBounds = null;
  previewDraftBounds = null;
  previewDraftText = "";
  previewRecognizedText = "";
  previewHasRecognizedText = false;
  previewCropDirty = false;
  previewPointerState = null;
  previewSourceStage = null;
  previewCropBox = null;
  previewSyncPromise = null;
}

function bindPreviewPointerListeners() {
  if (previewPointerListenersBound) {
    return;
  }

  window.addEventListener("mousemove", handlePreviewPointerMove);
  window.addEventListener("mouseup", stopPreviewPointer);
  previewPointerListenersBound = true;
}

function unbindPreviewPointerListeners() {
  if (!previewPointerListenersBound) {
    return;
  }

  window.removeEventListener("mousemove", handlePreviewPointerMove);
  window.removeEventListener("mouseup", stopPreviewPointer);
  previewPointerListenersBound = false;
}

async function getSourcePageCanvas(sourceIndex) {
  if (sourcePageCanvasCache.has(sourceIndex)) {
    return sourcePageCanvasCache.get(sourceIndex);
  }

  const sourcePage = latestSourcePages[sourceIndex];

  if (!sourcePage) {
    throw new Error(`Missing source page at index ${sourceIndex}.`);
  }

  const canvas = await sourcePage.renderCanvas();
  sourcePageCanvasCache.set(sourceIndex, canvas);
  return canvas;
}

function getSourcePageScale(sourceIndex) {
  return latestSourcePages[sourceIndex]?.sourceType === "pdf" ? PDF_RENDER_SCALE : 1;
}

async function buildQuestionCropResult(question, bounds) {
  const pageCanvas = await getSourcePageCanvas(question.sourceIndex);
  const scale = getSourcePageScale(question.sourceIndex);
  const croppedCanvas = cropQuestionCanvas(pageCanvas, bounds, scale);

  return {
    canvas: croppedCanvas,
    imageUrl: croppedCanvas.toDataURL("image/png"),
  };
}

function renderPreviewCropBox() {
  if (!previewCropBox || !previewDraftBounds) {
    return;
  }

  const question = findQuestionById(previewQuestionId);

  if (!question) {
    return;
  }

  previewCropBox.style.left = `${(previewDraftBounds.left / question.pageWidth) * 100}%`;
  previewCropBox.style.top = `${(previewDraftBounds.top / question.pageHeight) * 100}%`;
  previewCropBox.style.width = `${((previewDraftBounds.right - previewDraftBounds.left) / question.pageWidth) * 100}%`;
  previewCropBox.style.height = `${((previewDraftBounds.bottom - previewDraftBounds.top) / question.pageHeight) * 100}%`;
}

function updatePreviewReviewStatus(statusText = "") {
  const currentText = previewTextEditor ? getQuestionEditorText() : normalizeMultilineText(previewDraftText);
  const needsCropRecognition = previewCropDirty;
  const canEditText = Boolean(previewQuestionId) && !needsCropRecognition;

  if (previewRunOcrButton) {
    previewRunOcrButton.disabled = previewRunOcrButton.dataset.busy === "true" || !previewQuestionId;
  }

  setQuestionEditorDisabled(!canEditText);

  if (previewSaveButton) {
    previewSaveButton.disabled = !previewQuestionId || needsCropRecognition || !currentText;
  }

  if (previewReviewStatus) {
    if (statusText) {
      previewReviewStatus.textContent = statusText;
    } else if (!previewQuestionId) {
      previewReviewStatus.textContent = "請先選擇題目。";
    } else if (needsCropRecognition) {
      previewReviewStatus.textContent = "裁切範圍已更新，請按「確認裁切並 AI 辨識」。";
    } else if (currentText) {
      previewReviewStatus.textContent = "AI 辨識完成，請檢查文字後儲存此題。";
    } else {
      previewReviewStatus.textContent = "AI 辨識完成，如有需要可手動輸入題目文字。";
    }
  }
}

async function syncPreviewQuestionImage() {
  const question = findQuestionById(previewQuestionId);

  if (!question || !previewDraftBounds || !previewModalQuestionPreview) {
    return;
  }

  const syncPromise = (async () => {
    const cropResult = await buildQuestionCropResult(question, previewDraftBounds);

    if (previewQuestionId !== question.id) {
      return;
    }

    const questionImage = document.createElement("img");
    questionImage.src = cropResult.imageUrl;
    questionImage.alt = `${question.numberLabel} 題目預覽`;

    previewQuestionStage = document.createElement("div");
    previewQuestionStage.className = "preview-modal-question-stage";
    previewQuestionStage.appendChild(questionImage);
    previewModalQuestionPreview.replaceChildren(previewQuestionStage);
    previewQuestionZoom = 1;
    updatePreviewZoomDisplay();
  })();

  previewSyncPromise = syncPromise;
  await syncPromise;
}

function setPreviewDraftBounds(bounds, { markDirty = true } = {}) {
  const question = findQuestionById(previewQuestionId);

  if (!question) {
    return;
  }

  previewDraftBounds = clampQuestionBounds(bounds, question.pageWidth, question.pageHeight);

  if (markDirty) {
    previewCropDirty = true;
    previewHasRecognizedText = false;
  }

  renderPreviewCropBox();
  updatePreviewReviewStatus();
}

function nudgePreviewBounds({ dx = 0, dy = 0, growX = 0, growY = 0 } = {}) {
  if (!previewDraftBounds || !previewQuestionId) {
    return;
  }

  const currentWidth = previewDraftBounds.right - previewDraftBounds.left;
  const currentHeight = previewDraftBounds.bottom - previewDraftBounds.top;
  const nextBounds = {
    left: previewDraftBounds.left + dx - growX,
    right: previewDraftBounds.right + dx + growX,
    top: previewDraftBounds.top + dy - growY,
    bottom: previewDraftBounds.bottom + dy + growY,
  };

  if (growX < 0 || growY < 0) {
    const nextWidth = currentWidth + growX * 2;
    const nextHeight = currentHeight + growY * 2;

    if ((growX < 0 && nextWidth < MIN_CROP_SIZE) || (growY < 0 && nextHeight < MIN_CROP_SIZE)) {
      return;
    }
  }

  setPreviewDraftBounds(nextBounds, { markDirty: true });
  void syncPreviewQuestionImage();
}

function stopPreviewPointer() {
  if (!previewPointerState) {
    return;
  }

  previewPointerState = null;
  unbindPreviewPointerListeners();
  void syncPreviewQuestionImage();
}

function handlePreviewPointerMove(event) {
  if (!previewPointerState) {
    return;
  }

  const question = findQuestionById(previewQuestionId);

  if (!question) {
    return;
  }

  const { startBounds, startX, startY, rectWidth, rectHeight, handle } = previewPointerState;
  const deltaX = ((event.clientX - startX) / rectWidth) * question.pageWidth;
  const deltaY = ((event.clientY - startY) / rectHeight) * question.pageHeight;
  const nextBounds = { ...startBounds };

  if (handle === "move") {
    const width = startBounds.right - startBounds.left;
    const height = startBounds.bottom - startBounds.top;
    nextBounds.left = startBounds.left + deltaX;
    nextBounds.top = startBounds.top + deltaY;
    nextBounds.right = nextBounds.left + width;
    nextBounds.bottom = nextBounds.top + height;

    if (nextBounds.left < 0) {
      nextBounds.right -= nextBounds.left;
      nextBounds.left = 0;
    }

    if (nextBounds.top < 0) {
      nextBounds.bottom -= nextBounds.top;
      nextBounds.top = 0;
    }

    if (nextBounds.right > question.pageWidth) {
      const overflow = nextBounds.right - question.pageWidth;
      nextBounds.left -= overflow;
      nextBounds.right = question.pageWidth;
    }

    if (nextBounds.bottom > question.pageHeight) {
      const overflow = nextBounds.bottom - question.pageHeight;
      nextBounds.top -= overflow;
      nextBounds.bottom = question.pageHeight;
    }
  } else {
    if (handle.includes("n")) {
      nextBounds.top = startBounds.top + deltaY;
    }

    if (handle.includes("s")) {
      nextBounds.bottom = startBounds.bottom + deltaY;
    }

    if (handle.includes("w")) {
      nextBounds.left = startBounds.left + deltaX;
    }

    if (handle.includes("e")) {
      nextBounds.right = startBounds.right + deltaX;
    }
  }

  setPreviewDraftBounds(nextBounds, { markDirty: true });
}

function startPreviewPointer(handle, event) {
  if (!previewSourceStage || !previewDraftBounds) {
    return;
  }

  event.preventDefault();
  event.stopPropagation();

  const rect = previewSourceStage.getBoundingClientRect();
  previewPointerState = {
    handle,
    startX: event.clientX,
    startY: event.clientY,
    rectWidth: Math.max(rect.width, 1),
    rectHeight: Math.max(rect.height, 1),
    startBounds: { ...previewDraftBounds },
  };
  bindPreviewPointerListeners();
}

function getPreviewFallbackText(question, bounds) {
  const sourcePage = latestSourcePages[question?.sourceIndex ?? -1];
  const titleMatchers = compileTitleMatchers(collectSettings().titlePatterns);
  const extractedText = sourcePage
    ? extractQuestionText(
        sourcePage.lines,
        expandBoundsForText(bounds, question.pageWidth, question.pageHeight),
        titleMatchers,
      )
        || (
          question?.textBounds
            ? extractQuestionText(sourcePage.lines, question.textBounds, titleMatchers)
            : ""
        )
    : "";

  return normalizeMultilineText(extractedText || question?.detailText || question?.title || "");
}

function getTextLayerQuestionText(sourcePage, bounds, pageWidth, pageHeight, titleMatchers) {
  if (!sourcePage) {
    return "";
  }

  return normalizeMultilineText(
    extractQuestionText(sourcePage.lines, expandBoundsForText(bounds, pageWidth, pageHeight), titleMatchers),
  );
}

async function runPreviewRecognition() {
  const question = findQuestionById(previewQuestionId);

  if (!question || !previewDraftBounds) {
    return;
  }

  previewRunOcrButton.dataset.busy = "true";
  previewRunOcrButton.disabled = true;
  updatePreviewReviewStatus("AI 辨識中，請稍候...");

  try {
    const cropResult = await buildQuestionCropResult(question, previewDraftBounds);
    const sourcePage = latestSourcePages[question.sourceIndex];
    const fallbackText = getPreviewFallbackText(question, previewDraftBounds);
    let recognizedText = fallbackText;

    if (!recognizedText || sourcePage?.sourceType !== "pdf") {
      const worker = await getOcrWorker();
      const result = await worker.recognize(cropResult.canvas);
      recognizedText = normalizeMultilineText(result.data?.text ?? "") || fallbackText;
    }

    previewRecognizedText = recognizedText;
    previewHasRecognizedText = true;
    previewCropDirty = false;
    previewDraftText = recognizedText || question.detailText || question.title;

    setQuestionEditorText(previewDraftText);

    await syncPreviewQuestionImage();
    updatePreviewReviewStatus("AI 辨識完成，請檢查文字後儲存此題。");
  } catch (error) {
    console.error(error);
    const fallbackText = getPreviewFallbackText(question, previewDraftBounds);

    previewRecognizedText = fallbackText;
    previewHasRecognizedText = true;
    previewCropDirty = false;
    previewDraftText = fallbackText;

    setQuestionEditorText(fallbackText);

    await syncPreviewQuestionImage();
    updatePreviewReviewStatus(
      fallbackText
        ? "AI 辨識失敗，已改用目前抽取文字，請手動修正後儲存此題。"
        : "AI 辨識失敗，請直接輸入題目文字後儲存此題。",
    );
  } finally {
    previewRunOcrButton.dataset.busy = "false";
    previewRunOcrButton.disabled = false;
  }
}

async function savePreviewQuestion() {
  const question = findQuestionById(previewQuestionId);
  const finalizedText = previewTextEditor ? getQuestionEditorText() : normalizeMultilineText(previewDraftText);

  if (!question || !previewDraftBounds || !finalizedText) {
    updatePreviewReviewStatus("請先完成 AI 辨識，並確認文字內容。");
    return;
  }

  previewSaveButton.disabled = true;
  updatePreviewReviewStatus("儲存此題中...");

  try {
    const currentQuestionId = question.id;
    const cropResult = await buildQuestionCropResult(question, previewDraftBounds);
    question.bounds = { ...previewDraftBounds };
    question.overlay = buildOverlayFromBounds(question.bounds, question.pageWidth, question.pageHeight);
    question.imageUrl = cropResult.imageUrl;
    question.detailText = finalizedText;
    question.ocrText = previewRecognizedText || finalizedText;
    question.isConfirmed = true;

    renderQuestionPreviewList();
    sendQuestionsToSelectorWindow();
    updateReviewActionsVisibility();
    closePreviewModal();

    const nextPendingQuestion = getNextPendingQuestion(currentQuestionId);

    if (nextPendingQuestion && !nextPendingQuestion.isConfirmed) {
      extractionSummary.textContent = `已確認 ${question.numberLabel}，請繼續下一題：${nextPendingQuestion.numberLabel}。`;
      void openPreviewModal(nextPendingQuestion);
    } else {
      extractionSummary.textContent = "全部題目都已確認完成，現在可以前往下一頁勾選。";
    }
  } finally {
    previewSaveButton.disabled = false;
  }
}

function closePreviewModal() {
  if (!previewModal) {
    return;
  }

  previewModal.hidden = true;
  document.body.classList.remove("preview-modal-open");
  previewModalSourcePreview?.replaceChildren();
  previewModalQuestionPreview?.replaceChildren();
  previewQuestionZoom = 1;
  previewQuestionStage = null;

  setQuestionEditorText("");
  setQuestionEditorDisabled(true);

  if (previewRunOcrButton) {
    previewRunOcrButton.dataset.busy = "false";
    previewRunOcrButton.disabled = false;
  }

  resetPreviewDraftState();
  unbindPreviewPointerListeners();
  updatePreviewReviewStatus();
}

async function openPreviewModal(question) {
  if (!previewModal || !question) {
    return;
  }

  const firstPendingQuestion = getFirstPendingQuestion();

  if (
    firstPendingQuestion
    && question.id !== firstPendingQuestion.id
    && !question.isConfirmed
    && getPendingReviewCount() > 0
  ) {
    window.alert(`請依序確認題目，先完成 ${firstPendingQuestion.numberLabel}。`);
    question = firstPendingQuestion;
  }

  previewQuestionId = question.id;
  previewOriginalBounds = { ...question.bounds };
  previewDraftBounds = { ...question.bounds };
  previewDraftText = question.detailText || "";
  previewRecognizedText = question.ocrText || "";
  previewHasRecognizedText = Boolean(question.isConfirmed && (question.ocrText || question.detailText));
  previewCropDirty = false;

  const sourceImage = document.createElement("img");
  sourceImage.src = question.pagePreviewUrl;
  sourceImage.alt = `${question.numberLabel} 原始頁面`;

  previewCropBox = document.createElement("div");
  previewCropBox.className = "question-source-rect question-source-rect-editable";
  previewCropBox.addEventListener("mousedown", (event) => startPreviewPointer("move", event));

  ["nw", "ne", "sw", "se"].forEach((handle) => {
    const cropHandle = document.createElement("button");
    cropHandle.type = "button";
    cropHandle.className = `crop-handle crop-handle-${handle}`;
    cropHandle.setAttribute("aria-label", `調整 ${handle} 裁切角`);
    cropHandle.addEventListener("mousedown", (event) => startPreviewPointer(handle, event));
    previewCropBox.appendChild(cropHandle);
  });

  previewSourceStage = document.createElement("div");
  previewSourceStage.className = "preview-modal-source-stage";
  previewSourceStage.append(sourceImage, previewCropBox);

  previewModalSource.textContent = question.sourceLabel;
  previewModalTitle.textContent = question.numberLabel;
  renderRichQuestionText(previewModalAnchor, question.detailText || question.title);
  previewModalSourcePreview.replaceChildren(previewSourceStage);

  setQuestionEditorText(previewDraftText);

  renderPreviewCropBox();
  const questionIndex = getQuestionIndexById(question.id) + 1;
  updatePreviewReviewStatus(
    question.isConfirmed
      ? `第 ${questionIndex} / ${extractedQuestions.length} 題已確認，可繼續微調後重新儲存。`
      : `現在處理第 ${questionIndex} / ${extractedQuestions.length} 題，請先確認裁切，再進行 AI 辨識。`,
  );
  previewModal.hidden = false;
  document.body.classList.add("preview-modal-open");
  await syncPreviewQuestionImage();
}

function renderQuestionPreviewList() {
  questionPreviewList.replaceChildren();
  updateReviewActionsVisibility();

  if (!extractedQuestions.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "question-preview-empty";
    emptyState.textContent = "尚未產生抽題結果。";
    questionPreviewList.appendChild(emptyState);
    return;
  }

  const fragment = document.createDocumentFragment();

  extractedQuestions.forEach((question) => {
    const card = document.createElement("article");
    card.className = "question-preview-card";
    card.dataset.confirmed = String(Boolean(question.isConfirmed));
    card.tabIndex = 0;
    card.setAttribute("role", "button");
    card.setAttribute("aria-label", `${question.numberLabel}，點一下放大檢視`);

    const thumb = document.createElement("img");
    thumb.className = "question-preview-image";
    thumb.src = question.imageUrl;
    thumb.alt = `${question.numberLabel} 預覽`;

    const sourcePreview = document.createElement("div");
    sourcePreview.className = "question-source-preview";

    const sourceImage = document.createElement("img");
    sourceImage.className = "question-source-image";
    sourceImage.src = question.pagePreviewUrl;
    sourceImage.alt = `${question.numberLabel} 原始頁面`;

    const sourceRect = document.createElement("div");
    sourceRect.className = "question-source-rect";
    sourceRect.style.left = `${question.overlay.leftPercent}%`;
    sourceRect.style.top = `${question.overlay.topPercent}%`;
    sourceRect.style.width = `${question.overlay.widthPercent}%`;
    sourceRect.style.height = `${question.overlay.heightPercent}%`;

    sourcePreview.append(sourceImage, sourceRect);

    const meta = document.createElement("div");
    meta.className = "question-preview-meta";

    const heading = document.createElement("h3");
    heading.textContent = question.numberLabel;

    const source = document.createElement("p");
    source.textContent = question.sourceLabel;

    const title = document.createElement("p");
    title.className = "question-preview-title";
    renderRichQuestionText(title, question.detailText || question.title);

    const status = document.createElement("p");
    status.className = "question-preview-status";
    status.textContent = question.isConfirmed ? "已確認裁切與文字" : "待確認裁切與文字";

    meta.append(heading, source, title, status);
    card.append(sourcePreview, thumb, meta);
    card.addEventListener("click", () => {
      void openPreviewModal(question);
    });
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        void openPreviewModal(question);
      }
    });
    fragment.appendChild(card);
  });

  questionPreviewList.appendChild(fragment);
}

function openSelectorWindow() {
  selectorWindowReady = false;
  selectorWindow = window.open(`./selector.html?v=${APP_ASSET_VERSION}`, "v4-question-selector");

  if (!selectorWindow) {
    window.alert("瀏覽器擋下了新的勾選頁，請允許彈出視窗後再試。");
    return null;
  }

  return selectorWindow;
}

function buildSelectorPayload(questions) {
  if (!questions.length) {
    return null;
  }

  return {
    type: "v4-selector-data",
    questions,
    updatedAt: Date.now(),
  };
}

function sendQuestionsToSelectorWindow() {
  const readyQuestions = extractedQuestions.length && getPendingReviewCount() === 0 ? extractedQuestions : [];
  const payload = buildSelectorPayload(readyQuestions);

  if (!payload) {
    void writeSelectorPayload(null);
    selectorChannel?.postMessage({ type: "v4-selector-clear", updatedAt: Date.now() });
    return;
  }

  void writeSelectorPayload(payload);
  selectorChannel?.postMessage(payload);

  if (!selectorWindow || selectorWindow.closed || !selectorWindowReady) {
    return;
  }

  selectorWindow.postMessage(payload, window.location.origin);
}

function handleSelectorRequest() {
  if (!extractedQuestions.length || getPendingReviewCount() > 0) {
    selectorChannel?.postMessage({ type: "v4-selector-clear", updatedAt: Date.now() });
    return;
  }

  sendQuestionsToSelectorWindow();
}

function renderPages() {
  const settings = collectSettings();
  const title = settings.title;
  const studentName = withFallback(settings.studentName, "________________");
  const displayDate = formatDisplayDate(settings.date);
  const cellsPerPage = settings.rowCount;
  const totalCells = settings.pageCount * cellsPerPage;
  const layoutLabel = `單欄 / 每頁 ${cellsPerPage} 題`;

  if (activePasteImageIndex !== null && activePasteImageIndex >= totalCells) {
    activePasteImageIndex = null;
  }

  pagesRoot.replaceChildren();
  cellBindings.clear();
  updateGuideMode(settings.guideMode);
  updateSignatureVisibility(settings.showSignature);
  saveSettings(settings);

  for (let pageIndex = 0; pageIndex < settings.pageCount; pageIndex += 1) {
    const pageFragment = pageTemplate.content.cloneNode(true);
    const pageTitle = pageFragment.querySelector(".page-title");
    const pageLayout = pageFragment.querySelector(".page-layout");
    const pageName = pageFragment.querySelector(".page-name");
    const pageDate = pageFragment.querySelector(".page-date");
    const pageMeta = pageFragment.querySelector(".page-meta");
    const pageHeader = pageFragment.querySelector(".page-header");
    const grid = pageFragment.querySelector(".grid");
    const page = pageFragment.querySelector(".page");

    pageTitle.textContent = title;
    pageLayout.textContent = layoutLabel;
    pageName.textContent = studentName;
    pageDate.textContent = displayDate;
    pageMeta.textContent = `第 ${pageIndex + 1} / ${settings.pageCount} 頁`;
    grid.style.gridTemplateColumns = "repeat(1, minmax(0, 1fr))";
    grid.style.gridTemplateRows = `repeat(${settings.rowCount}, minmax(0, 1fr))`;

    if (pageIndex > 0) {
      page.classList.add("page-following");
      pageHeader.remove();
    }

    for (let cellIndex = 0; cellIndex < cellsPerPage; cellIndex += 1) {
      const imageIndex = pageIndex * cellsPerPage + cellIndex;
      const cellFragment = cellTemplate.content.cloneNode(true);
      const cell = cellFragment.querySelector(".cell");
      const cellNumber = cellFragment.querySelector(".cell-number");
      const number = settings.startNumber + imageIndex;

      cellNumber.textContent = number;
      bindCellUpload(cell, imageIndex);
      grid.appendChild(cellFragment);
    }

    pagesRoot.appendChild(page);
  }
}

function compileTitleMatchers(rawPatterns) {
  const lines = rawPatterns
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  return lines.map((line) => {
    try {
      return new RegExp(line, "iu");
    } catch {
      return new RegExp(escapeRegExp(line), "iu");
    }
  });
}

function deriveQuestionNumberLabel(title, fallbackNumber) {
  const match = title.match(/第\s*\d+\s*題|\d+[.、)]|[（(]\d+[）)]|\d+/u);
  return match ? normalizeText(match[0]) : `題號 ${fallbackNumber}`;
}

async function loadPdfDocument(bytes) {
  const loadingTask = pdfjsLib.getDocument({ data: bytes });
  return loadingTask.promise;
}

async function extractPdfLines(page) {
  const viewport = page.getViewport({ scale: 1 });
  const textContent = await page.getTextContent();
  const items = textContent.items
    .map((item) => {
      if (!("str" in item)) {
        return null;
      }

      const text = normalizeText(item.str);

      if (!text) {
        return null;
      }

      const transformed = pdfjsLib.Util.transform(viewport.transform, item.transform);
      const x = transformed[4];
      const y = transformed[5];
      const width = Math.max(item.width || 0, 1);
      const height = Math.max(Math.abs(item.height || transformed[0] || 0), 8);

      return {
        text,
        x,
        y,
        width,
        height,
      };
    })
    .filter(Boolean)
    .sort((a, b) => a.y - b.y || a.x - b.x);

  const lines = [];

  items.forEach((item) => {
    const line = lines.at(-1);
    const tolerance = Math.max(4, item.height * 0.65);

    if (!line || Math.abs(line.y - item.y) > tolerance) {
      lines.push({
        y: item.y,
        items: [item],
      });
      return;
    }

    line.items.push(item);
    line.y = (line.y + item.y) / 2;
  });

  return {
    pageWidth: viewport.width,
    pageHeight: viewport.height,
    lines: lines.map((line) => {
      const sortedItems = line.items.sort((a, b) => a.x - b.x);
      let text = "";

      sortedItems.forEach((item, index) => {
        const previous = sortedItems[index - 1];

        if (!previous) {
          text = item.text;
          return;
        }

        const previousRight = previous.x + previous.width;
        const gap = item.x - previousRight;
        text += gap > Math.max(6, item.height * 0.5) ? ` ${item.text}` : item.text;
      });

      const left = Math.min(...sortedItems.map((item) => item.x));
      const right = Math.max(...sortedItems.map((item) => item.x + item.width));
      const top = Math.min(...sortedItems.map((item) => item.y - item.height));
      const bottom = Math.max(...sortedItems.map((item) => item.y + item.height * 0.35));

      return {
        text: normalizeText(text),
        left,
        right,
        top,
        bottom,
        items: sortedItems.map((item) => ({
          text: item.text,
          x: item.x,
          y: item.y,
          width: item.width,
          height: item.height,
          top: item.y - item.height,
          bottom: item.y + item.height * 0.35,
        })),
      };
    }),
  };
}

async function getOcrWorker() {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = (async () => {
      const { createWorker } = await import(OCR_MODULE_URL);
      return createWorker(OCR_LANGUAGE);
    })();
  }

  return ocrWorkerPromise;
}

async function loadImageElement(file) {
  const objectUrl = URL.createObjectURL(file);

  try {
    return await new Promise((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = reject;
      image.src = objectUrl;
    });
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function buildCanvasFromImage(image) {
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  canvas.width = image.naturalWidth || image.width;
  canvas.height = image.naturalHeight || image.height;
  context.drawImage(image, 0, 0);
  return canvas;
}

async function extractImageLines(file) {
  const image = await loadImageElement(file);
  const canvas = buildCanvasFromImage(image);
  const worker = await getOcrWorker();
  const result = await worker.recognize(canvas);
  const lines = (result.data?.lines ?? [])
    .map((line) => ({
      text: normalizeText(line.text || ""),
      left: line.bbox?.x0 ?? 0,
      right: line.bbox?.x1 ?? 0,
      top: line.bbox?.y0 ?? 0,
      bottom: line.bbox?.y1 ?? 0,
    }))
    .filter((line) => line.text);

  return {
    pageWidth: canvas.width,
    pageHeight: canvas.height,
    lines,
    canvas,
  };
}

function matchTitleLines(lines, patterns, pageInfo) {
  return lines
    .filter((line) => patterns.some((pattern) => pattern.test(line.text)))
    .map((line) => ({
      ...line,
      title: line.text,
      ...pageInfo,
    }));
}

function isStackedMathLine(text) {
  return /^[\d\s.,]+$/.test(normalizeText(text));
}

function shouldRebuildStackedMathLine(text, titleMatchers) {
  return titleMatchers.some((pattern) => pattern.test(text))
    || /[？?]|\([A-Da-dＡ-Ｄ]\)|（[A-Da-dＡ-Ｄ]）/.test(text);
}

function buildBoundedLine(line, bounds) {
  const boundedItems = Array.isArray(line.items) && line.items.length
    ? line.items
        .filter((item) => {
          const itemRight = item.x + Math.max(item.width || 0, 1);
          return itemRight >= bounds.left && item.x <= bounds.right;
        })
        .map((item) => ({
          ...item,
          top: item.top ?? line.top,
          bottom: item.bottom ?? line.bottom,
        }))
        .sort((a, b) => a.x - b.x)
    : [];

  return {
    ...line,
    boundedItems,
  };
}

function joinPositionedItems(items) {
  let text = "";

  items.forEach((item, index) => {
    const previous = items[index - 1];

    if (!previous) {
      text = item.text;
      return;
    }

    const previousRight = previous.x + Math.max(previous.width || 0, 1);
    const gap = item.x - previousRight;
    text += gap > Math.max(6, item.height * 0.5) ? ` ${item.text}` : item.text;
  });

  return normalizeText(text);
}

function buildStackedMathText(anchorLine, nearbyMathLines) {
  const tokens = anchorLine.boundedItems.map((item) => ({
    text: item.text,
    x: item.x,
    width: item.width,
    height: item.height,
    top: item.top,
  }));

  const mathItems = nearbyMathLines
    .flatMap((line) => line.boundedItems)
    .filter((item) => /^\d+$/.test(item.text))
    .sort((a, b) => a.x - b.x || a.top - b.top);
  const groupedMathItems = [];

  for (const item of mathItems) {
    const center = item.x + Math.max(item.width || 0, 1) / 2;
    const group = groupedMathItems.find((candidate) =>
      Math.abs(candidate.center - center) <= Math.max(7, item.width * 0.75),
    );

    if (group) {
      group.items.push(item);
      group.center = (group.center * (group.items.length - 1) + center) / group.items.length;
      continue;
    }

    groupedMathItems.push({
      center,
      items: [item],
    });
  }

  groupedMathItems.forEach((group) => {
    const sortedItems = group.items.sort((a, b) => a.top - b.top);
    const text = formatStackedMathItemsAsLatex(sortedItems);
    const left = Math.min(...sortedItems.map((item) => item.x));
    const right = Math.max(...sortedItems.map((item) => item.x + Math.max(item.width || 0, 1)));
    const top = Math.min(...sortedItems.map((item) => item.top));

    tokens.push({
      text,
      x: left,
      width: Math.max(1, right - left),
      height: sortedItems[0]?.height ?? anchorLine.height ?? 8,
      top,
    });
  });

  return joinPositionedItems(tokens.sort((a, b) => a.x - b.x || a.top - b.top));
}

function buildTextLinesWithStackedMath(lines, bounds, titleMatchers) {
  const boundedLines = lines
    .filter((line) => line.bottom >= bounds.top && line.top <= bounds.bottom)
    .filter((line) => line.right >= bounds.left && line.left <= bounds.right)
    .map((line) => buildBoundedLine(line, bounds))
    .filter((line) => !Array.isArray(line.items) || !line.items.length || line.boundedItems.length)
    .sort((a, b) => a.top - b.top || a.left - b.left);
  const consumedLineIndexes = new Set();

  return boundedLines
    .map((line, index) => {
      if (consumedLineIndexes.has(index)) {
        return "";
      }

      if (isStackedMathLine(line.text)) {
        const belongsToAnchor = boundedLines.some((candidate, candidateIndex) => {
          if (candidateIndex === index || !candidate.boundedItems.length) {
            return false;
          }

          const candidateText = Array.isArray(candidate.items) && candidate.items.length
            ? joinPositionedItems(candidate.boundedItems)
            : normalizeText(candidate.text);
          const verticalGap = Math.max(
            0,
            Math.max(line.top, candidate.top) - Math.min(line.bottom, candidate.bottom),
          );
          const overlapsHorizontally = line.right >= candidate.left - 8 && line.left <= candidate.right + 8;

          return shouldRebuildStackedMathLine(candidateText, titleMatchers)
            && overlapsHorizontally
            && verticalGap <= Math.max(18, (candidate.bottom - candidate.top) * 1.7);
        });

        if (belongsToAnchor) {
          return "";
        }
      }

      if (!Array.isArray(line.items) || !line.items.length) {
        return normalizeText(line.text);
      }

      const lineText = joinPositionedItems(line.boundedItems);

      if (!shouldRebuildStackedMathLine(lineText, titleMatchers)) {
        return lineText;
      }

      const nearbyMathLines = boundedLines.filter((candidate, candidateIndex) => {
        if (candidateIndex === index || !candidate.boundedItems.length || !isStackedMathLine(candidate.text)) {
          return false;
        }

        const verticalGap = Math.max(
          0,
          Math.max(line.top, candidate.top) - Math.min(line.bottom, candidate.bottom),
        );
        const overlapsHorizontally = candidate.right >= line.left - 8 && candidate.left <= line.right + 8;
        const isNearby = overlapsHorizontally && verticalGap <= Math.max(18, (line.bottom - line.top) * 1.7);

        if (isNearby) {
          consumedLineIndexes.add(candidateIndex);
        }

        return isNearby;
      });

      return nearbyMathLines.length ? buildStackedMathText(line, nearbyMathLines) : lineText;
    })
    .filter(Boolean);
}

function extractQuestionText(lines, bounds, titleMatchers = []) {
  const textLines = lines
    .filter((line) => line.bottom >= bounds.top && line.top <= bounds.bottom)
    .filter((line) => line.right >= bounds.left && line.left <= bounds.right)
    .sort((a, b) => a.top - b.top || a.left - b.left)
    .map((line) => {
      if (!Array.isArray(line.items) || !line.items.length) {
        return normalizeText(line.text);
      }

      const pageWidth = Number.POSITIVE_INFINITY;
      const filteredItems = line.items
        .filter((item) => item.x < pageWidth * 0.66 || item.text.length > 2 || /[()（）]/.test(item.text))
        .sort((a, b) => a.x - b.x);

      const boundedItems = line.items
        .filter((item) => {
          const itemRight = item.x + Math.max(item.width || 0, 1);
          return itemRight >= bounds.left && item.x <= bounds.right;
        })
        .sort((a, b) => a.x - b.x);

      if (!boundedItems.length) {
        return "";
      }

      let text = "";

      boundedItems.forEach((item, index) => {
        const previous = boundedItems[index - 1];

        if (!previous) {
          text = item.text;
          return;
        }

        const previousRight = previous.x + previous.width;
        const gap = item.x - previousRight;
        text += gap > Math.max(6, item.height * 0.5) ? ` ${item.text}` : item.text;
      });

      return normalizeText(text);
    })
    .filter(Boolean);

  const stackedTextLines = buildTextLinesWithStackedMath(lines, bounds, titleMatchers);
  const readableTextLines = stackedTextLines.length ? stackedTextLines : textLines;
  const firstTitleLineIndex = readableTextLines.findIndex((textLine) =>
    titleMatchers.some((pattern) => pattern.test(textLine)),
  );
  const candidateLines = firstTitleLineIndex > 0 ? readableTextLines.slice(firstTitleLineIndex) : readableTextLines;
  const mergedLines = [];

  for (const textLine of candidateLines) {
    if (mergedLines.length > 0 && titleMatchers.some((pattern) => pattern.test(textLine))) {
      break;
    }

    mergedLines.push(textLine);
  }

  const merged = normalizeText(mergedLines.join(" "));
  const startIndex = merged.search(/[（(]\s*[）)]\s*\d+\s*[.．、]|第\s*\d+\s*題|\d+\s*[.．、]/u);
  const normalized = startIndex >= 0 ? merged.slice(startIndex).trim() : merged;
  const optionMatch = normalized.match(/^(.*?\(\s*[AaＡ]\s*\).*?\(\s*[DdＤ]\s*\)[^。]*(?:。|$))/u);
  const optionNormalized = optionMatch ? optionMatch[1].trim() : normalized;

  return optionNormalized
    .replace(/([？?])\s+\d+\s+(\(\s*[AaＡ]\s*\))/u, "$1 $2")
    .replace(/\(\s*\)\s*(\d+\s*[.．、])/gu, "(   )$1")
    .replace(/（\s*）\s*(\d+\s*[.．、])/gu, "（   ）$1")
    .trim();
}

async function buildSourcePages(files) {
  const sourcePages = [];

  for (const file of files) {
    if (file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf")) {
      const pdfBytes = await file.arrayBuffer();
      const pdf = await loadPdfDocument(pdfBytes);

      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
        const page = await pdf.getPage(pageNumber);
        const extracted = await extractPdfLines(page);
        sourcePages.push({
          sourceType: "pdf",
          sourceName: file.name,
          sourceLabel: `${file.name} / 第 ${pageNumber} 頁`,
          pageNumber,
          pageWidth: extracted.pageWidth,
          pageHeight: extracted.pageHeight,
          lines: extracted.lines,
          renderCanvas: () => renderPdfPageToCanvas(page, PDF_RENDER_SCALE),
        });
      }

      continue;
    }

    if (file.type.startsWith("image/")) {
      const extracted = await extractImageLines(file);
      sourcePages.push({
        sourceType: "image",
        sourceName: file.name,
        sourceLabel: file.name,
        pageNumber: 1,
        pageWidth: extracted.pageWidth,
        pageHeight: extracted.pageHeight,
        lines: extracted.lines,
        renderCanvas: async () => extracted.canvas,
      });
    }
  }

  return sourcePages;
}

async function renderPdfPageToCanvas(page, scale) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");

  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);

  await page.render({
    canvasContext: context,
    viewport,
  }).promise;

  return canvas;
}

function cropQuestionCanvas(pageCanvas, bounds, scale) {
  const sx = Math.max(0, Math.floor(bounds.left * scale));
  const sy = Math.max(0, Math.floor(bounds.top * scale));
  const sw = Math.max(10, Math.floor((bounds.right - bounds.left) * scale));
  const sh = Math.max(10, Math.floor((bounds.bottom - bounds.top) * scale));

  const croppedCanvas = document.createElement("canvas");
  const context = croppedCanvas.getContext("2d");

  croppedCanvas.width = sw;
  croppedCanvas.height = sh;
  context.fillStyle = "white";
  context.fillRect(0, 0, sw, sh);
  context.drawImage(pageCanvas, sx, sy, sw, sh, 0, 0, sw, sh);

  return trimCanvasToVisibleContent(croppedCanvas, CROPPED_CONTENT_MARGIN_PX);
}

function cropQuestionImage(pageCanvas, bounds, scale) {
  return cropQuestionCanvas(pageCanvas, bounds, scale).toDataURL("image/png");
}

function trimCanvasToVisibleContent(canvas, marginPx) {
  const context = canvas.getContext("2d", { willReadFrequently: true });

  if (!context) {
    return canvas;
  }

  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
  const { data, width, height } = imageData;
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      const alpha = data[offset + 3];

      if (alpha < CONTENT_ALPHA_THRESHOLD) {
        continue;
      }

      const red = data[offset];
      const green = data[offset + 1];
      const blue = data[offset + 2];

      if (red > CONTENT_RGB_THRESHOLD && green > CONTENT_RGB_THRESHOLD && blue > CONTENT_RGB_THRESHOLD) {
        continue;
      }

      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  }

  if (maxX < 0 || maxY < 0) {
    return canvas;
  }

  const left = Math.max(0, minX - marginPx);
  const top = Math.max(0, minY - marginPx);
  const right = Math.min(width, maxX + marginPx + 1);
  const bottom = Math.min(height, maxY + marginPx + 1);
  const trimmedCanvas = document.createElement("canvas");
  const trimmedContext = trimmedCanvas.getContext("2d");

  trimmedCanvas.width = Math.max(10, right - left);
  trimmedCanvas.height = Math.max(10, bottom - top);
  trimmedContext.fillStyle = "white";
  trimmedContext.fillRect(0, 0, trimmedCanvas.width, trimmedCanvas.height);
  trimmedContext.drawImage(
    canvas,
    left,
    top,
    trimmedCanvas.width,
    trimmedCanvas.height,
    0,
    0,
    trimmedCanvas.width,
    trimmedCanvas.height,
  );

  return trimmedCanvas;
}

function refineBoundsToVisibleContent(pageCanvas, bounds, scale, pageWidth, pageHeight) {
  const sx = Math.max(0, Math.floor(bounds.left * scale));
  const sy = Math.max(0, Math.floor(bounds.top * scale));
  const sw = Math.max(10, Math.floor((bounds.right - bounds.left) * scale));
  const sh = Math.max(10, Math.floor((bounds.bottom - bounds.top) * scale));
  const context = pageCanvas.getContext("2d", { willReadFrequently: true });

  if (!context) {
    return bounds;
  }

  const imageData = context.getImageData(sx, sy, sw, sh);
  const { data, width, height } = imageData;
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;

  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      const alpha = data[offset + 3];

      if (alpha < CONTENT_ALPHA_THRESHOLD) {
        continue;
      }

      const red = data[offset];
      const green = data[offset + 1];
      const blue = data[offset + 2];

      if (red > CONTENT_RGB_THRESHOLD && green > CONTENT_RGB_THRESHOLD && blue > CONTENT_RGB_THRESHOLD) {
        continue;
      }

      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  }

  if (maxX < 0 || maxY < 0) {
    return bounds;
  }

  const margin = CONTENT_MARGIN_PX / scale;

  return {
    left: Math.max(0, Math.min(pageWidth, bounds.left + minX / scale - margin)),
    top: Math.max(0, Math.min(pageHeight, bounds.top + minY / scale - margin)),
    right: Math.max(
      bounds.left + 20 / scale,
      Math.min(pageWidth, bounds.left + (maxX + 1) / scale + margin),
    ),
    bottom: Math.max(
      bounds.top + 24 / scale,
      Math.min(pageHeight, bounds.top + (maxY + 1) / scale + margin),
    ),
  };
}

function buildOverlayFromBounds(bounds, pageWidth, pageHeight) {
  return {
    leftPercent: (bounds.left / pageWidth) * 100,
    topPercent: (bounds.top / pageHeight) * 100,
    widthPercent: ((bounds.right - bounds.left) / pageWidth) * 100,
    heightPercent: ((bounds.bottom - bounds.top) / pageHeight) * 100,
  };
}

function computeQuestionBounds(anchor, nextAnchor, settings) {
  const left = anchor.pageWidth * (settings.leftPercent / 100);
  const right = anchor.pageWidth * (settings.rightPercent / 100);
  const top = Math.max(0, anchor.top - settings.topPadding);
  const pageBottom = Math.max(top + 24, anchor.pageHeight - settings.bottomPadding);

  let bottom = pageBottom;

  if (nextAnchor && nextAnchor.sourceIndex === anchor.sourceIndex) {
    bottom = Math.min(bottom, nextAnchor.top - settings.nextGap);
  }

  return {
    left: Math.max(0, Math.min(left, anchor.pageWidth - 20)),
    right: Math.max(left + 20, Math.min(right, anchor.pageWidth)),
    top,
    bottom: Math.max(top + 24, Math.min(bottom, anchor.pageHeight)),
  };
}

function isFooterLikeLine(text, top, pageHeight) {
  if (top < pageHeight * 0.82) {
    return false;
  }

  const normalized = normalizeText(text);
  return /^\d+$/.test(normalized) || /題目結束|背面還有題目/u.test(normalized);
}

function trimBoundsBeforeFooter(bounds, lines, pageHeight, nextGap) {
  const footerTop = lines
    .filter((line) => line.top > bounds.top)
    .find((line) => isFooterLikeLine(line.text, line.top, pageHeight))?.top;

  if (!footerTop) {
    return bounds;
  }

  return {
    ...bounds,
    bottom: Math.max(bounds.top + 24, Math.min(bounds.bottom, footerTop - nextGap)),
  };
}

function applyExtractedQuestions(questions) {
  const settings = collectSettings();
  const requiredPages = Math.max(1, Math.ceil(questions.length / settings.rowCount));

  if (requiredPages > settings.pageCount) {
    pageCountInput.value = requiredPages;
  }

  uploadedImages.clear();
  uploadedQuestionTexts.clear();

  questions.forEach((question, index) => {
    uploadedImages.set(index, question.imageUrl);
    uploadedQuestionTexts.set(index, question.detailText || "");
  });

  activePasteImageIndex = questions.length ? 0 : null;
  renderPages();
}

function applySelectedQuestionIds(selectedIds) {
  const selectedIdSet = new Set(selectedIds);
  const selectedQuestions = extractedQuestions.filter((question) => selectedIdSet.has(question.id));

  if (!selectedQuestions.length) {
    uploadedImages.clear();
    uploadedQuestionTexts.clear();
    activePasteImageIndex = null;
    renderPages();
    extractionSummary.textContent = "勾選頁沒有選任何題目，主頁未帶入題圖。";
    return;
  }

  applyExtractedQuestions(selectedQuestions);
  extractionSummary.textContent = `已依序帶入 ${selectedQuestions.length} 題到左側題目區（題目文字 + 題圖）。`;
}

function clearExtractedQuestions() {
  extractedQuestions = [];
  latestSourcePages = [];
  sourcePageCanvasCache.clear();
  uploadedImages.clear();
  uploadedQuestionTexts.clear();
  activePasteImageIndex = null;
  sendQuestionsToSelectorWindow();
  renderQuestionPreviewList();
  extractionSummary.textContent = currentSourceLabel
    ? `已載入 ${currentSourceLabel}，尚未抽題。`
    : "尚未載入來源檔案。";
  renderPages();
}

async function extractQuestionsFromSources() {
  if (!sourceFiles.length) {
    window.alert("請先上傳含有題號的圖片或 PDF。");
    return;
  }

  const settings = collectSettings();
  const titleMatchers = compileTitleMatchers(settings.titlePatterns);

  if (!titleMatchers.length) {
    window.alert("請至少輸入一條題號規則。");
    return;
  }

  extractSourceButton.disabled = true;
  extractionSummary.textContent = "來源分析中，請稍候...";

  try {
    const sourcePages = await buildSourcePages(sourceFiles);
    latestSourcePages = sourcePages;
    sourcePageCanvasCache.clear();
    const anchors = [];

    sourcePages.forEach((pageInfo, sourceIndex) => {
      anchors.push(
        ...matchTitleLines(pageInfo.lines, titleMatchers, {
          sourceIndex,
          sourceLabel: pageInfo.sourceLabel,
          sourceType: pageInfo.sourceType,
          pageNumber: pageInfo.pageNumber,
          pageWidth: pageInfo.pageWidth,
          pageHeight: pageInfo.pageHeight,
        }),
      );
    });

    if (!anchors.length) {
      extractedQuestions = [];
      renderQuestionPreviewList();
      extractionSummary.textContent = "沒有找到符合題號規則的題目，請調整規則後再試。";
      return;
    }

    anchors.sort((a, b) => a.sourceIndex - b.sourceIndex || a.top - b.top);

    const pageCanvasCache = new Map();
    const pagePreviewCache = new Map();
    const questions = [];

    for (let index = 0; index < anchors.length; index += 1) {
      const anchor = anchors[index];
      const nextAnchor = anchors[index + 1] ?? null;
      let pageCanvas = pageCanvasCache.get(anchor.sourceIndex);

      if (!pageCanvas) {
        pageCanvas = await sourcePages[anchor.sourceIndex].renderCanvas();
        pageCanvasCache.set(anchor.sourceIndex, pageCanvas);
      }

      const scale =
        sourcePages[anchor.sourceIndex].sourceType === "pdf"
          ? PDF_RENDER_SCALE
          : 1;
      const roughBounds = trimBoundsBeforeFooter(
        computeQuestionBounds(anchor, nextAnchor, settings),
        sourcePages[anchor.sourceIndex].lines,
        anchor.pageHeight,
        settings.nextGap,
      );
      const bounds = refineBoundsToVisibleContent(
        pageCanvas,
        roughBounds,
        scale,
        anchor.pageWidth,
        anchor.pageHeight,
      );
      const imageUrl = cropQuestionImage(pageCanvas, bounds, scale);
      let pagePreviewUrl = pagePreviewCache.get(anchor.sourceIndex);

      if (!pagePreviewUrl) {
        pagePreviewUrl = pageCanvas.toDataURL("image/jpeg", 0.82);
        pagePreviewCache.set(anchor.sourceIndex, pagePreviewUrl);
      }

      const numberLabel = deriveQuestionNumberLabel(anchor.title, index + 1);
      const sourcePage = sourcePages[anchor.sourceIndex];
      const textBounds = expandBoundsForText(bounds, anchor.pageWidth, anchor.pageHeight);
      const roughTextBounds = expandBoundsForText(roughBounds, anchor.pageWidth, anchor.pageHeight);
      const detailText =
        getTextLayerQuestionText(sourcePage, bounds, anchor.pageWidth, anchor.pageHeight, titleMatchers)
        || extractQuestionText(sourcePage.lines, roughTextBounds, titleMatchers)
        || anchor.title;

      questions.push({
        id: `q-${index + 1}`,
        numberLabel,
        title: anchor.title,
        detailText,
        sourceLabel: anchor.sourceLabel,
        sourceIndex: anchor.sourceIndex,
        pageNumber: anchor.pageNumber,
        pageWidth: anchor.pageWidth,
        pageHeight: anchor.pageHeight,
        bounds,
        textBounds,
        imageUrl,
        pagePreviewUrl,
        overlay: buildOverlayFromBounds(bounds, anchor.pageWidth, anchor.pageHeight),
        ocrText: "",
        isConfirmed: false,
      });
    }

    extractedQuestions = questions;
    renderQuestionPreviewList();
    extractionSummary.textContent = `已從 ${currentSourceLabel} 抽出 ${questions.length} 題，現在開始逐題確認。`;
    sendQuestionsToSelectorWindow();
    document.querySelector(".question-bank")?.scrollIntoView({ behavior: "smooth", block: "start" });
    const firstPendingQuestion = getFirstPendingQuestion();

    if (firstPendingQuestion) {
      window.setTimeout(() => {
        void openPreviewModal(firstPendingQuestion);
      }, 80);
    }
  } catch (error) {
    console.error(error);
    extractionSummary.textContent = "來源分析失敗，請確認檔案內容清楚且含有題號。";
    window.alert("來源分析失敗，請確認檔案內容清楚且含有題號。");
  } finally {
    extractSourceButton.disabled = false;
  }
}

async function loadExampleSourceFile() {
  const response = await fetch(`./${encodeURIComponent(DEFAULT_SAMPLE_FILE_NAME)}?v=${APP_ASSET_VERSION}`, { cache: "no-store" });

  if (!response.ok) {
    throw new Error(`Unable to load ${DEFAULT_SAMPLE_FILE_NAME}: ${response.status}`);
  }

  const blob = await response.blob();
  const file = new File([blob], DEFAULT_SAMPLE_FILE_NAME, {
    type: blob.type || "application/pdf",
    lastModified: Date.now(),
  });

  sourceFiles = [file];
  currentSourceLabel = file.name;
  extractedQuestions = [];
  renderQuestionPreviewList();
  extractionSummary.textContent = `已載入 ${currentSourceLabel}，可開始分析。`;

  if (typeof DataTransfer === "function") {
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(file);
    sourceFileInput.files = dataTransfer.files;
  }
}

function bootstrapExampleDemo() {
  if (demoBootstrapPromise) {
    return demoBootstrapPromise;
  }

  demoBootstrapPromise = (async () => {
    await loadExampleSourceFile();
    await extractQuestionsFromSources();
  })();

  return demoBootstrapPromise;
}

function loadDefaultExampleSource() {
  return loadExampleSourceFile().catch((error) => {
    console.warn(error);
    extractionSummary.textContent = "尚未載入來源檔案。";
  });
}

[titleInput, classInput, nameInput, dateInput, startNumberInput, pageCountInput, rowCountInput, guideSelect]
  .filter(Boolean)
  .forEach((element) => {
    element.addEventListener("input", renderPages);
    element.addEventListener("change", renderPages);
  });

[titlePatternsInput, leftPercentInput, rightPercentInput, topPaddingInput, bottomPaddingInput, nextGapInput]
  .filter(Boolean)
  .forEach((element) => {
    element.addEventListener("input", persistSettings);
    element.addEventListener("change", persistSettings);
  });

signatureToggle.addEventListener("click", () => {
  updateSignatureVisibility(!getSignatureVisible());
  renderPages();
});

resetButton.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  closePreviewModal();
  uploadedImages.clear();
  uploadedQuestionTexts.clear();
  activePasteImageIndex = null;
  extractedQuestions = [];
  latestSourcePages = [];
  sourcePageCanvasCache.clear();
  sourceFiles = [];
  currentSourceLabel = "";
  sourceFileInput.value = "";
  applySettings({ ...DEFAULT_SETTINGS });
  renderQuestionPreviewList();
  void loadDefaultExampleSource();
  extractionSummary.textContent = "尚未載入來源檔案。";
  renderPages();
});

printButton.addEventListener("click", () => {
  window.print();
});

sourceFileInput.addEventListener("change", () => {
  closePreviewModal();
  sourceFiles = Array.from(sourceFileInput.files ?? []);
  latestSourcePages = [];
  sourcePageCanvasCache.clear();
  currentSourceLabel = sourceFiles.length === 1
    ? sourceFiles[0].name
    : sourceFiles.length
      ? `${sourceFiles.length} 個來源檔案`
      : "";
  extractedQuestions = [];
  renderQuestionPreviewList();
  extractionSummary.textContent = currentSourceLabel
    ? `已載入 ${currentSourceLabel}，可開始分析。`
    : "尚未載入來源檔案。";
});

extractSourceButton.addEventListener("click", () => {
  extractQuestionsFromSources();
});

clearExtractedButton.addEventListener("click", () => {
  closePreviewModal();
  clearExtractedQuestions();
});

confirmExtractedButton?.addEventListener("click", () => {
  const openedWindow = openSelectorWindow();

  if (!openedWindow) {
    return;
  }

  sendQuestionsToSelectorWindow();
  openedWindow.focus();
});

previewModal?.addEventListener("click", (event) => {
  const target = event.target;

  if (target instanceof HTMLElement && target.dataset.previewClose === "true") {
    closePreviewModal();
  }
});

previewModalCloseButton?.addEventListener("click", closePreviewModal);
previewResetCropButton?.addEventListener("click", () => {
  if (!previewOriginalBounds) {
    return;
  }

  setPreviewDraftBounds({ ...previewOriginalBounds }, { markDirty: true });
  void syncPreviewQuestionImage();
});
previewMoveUpButton?.addEventListener("click", () => {
  nudgePreviewBounds({ dy: -PREVIEW_NUDGE_STEP });
});
previewMoveLeftButton?.addEventListener("click", () => {
  nudgePreviewBounds({ dx: -PREVIEW_NUDGE_STEP });
});
previewMoveRightButton?.addEventListener("click", () => {
  nudgePreviewBounds({ dx: PREVIEW_NUDGE_STEP });
});
previewMoveDownButton?.addEventListener("click", () => {
  nudgePreviewBounds({ dy: PREVIEW_NUDGE_STEP });
});
previewWidenButton?.addEventListener("click", () => {
  nudgePreviewBounds({ growX: PREVIEW_RESIZE_STEP });
});
previewHeightenButton?.addEventListener("click", () => {
  nudgePreviewBounds({ growY: PREVIEW_RESIZE_STEP });
});
previewNarrowButton?.addEventListener("click", () => {
  nudgePreviewBounds({ growX: -PREVIEW_RESIZE_STEP });
});
previewShortenButton?.addEventListener("click", () => {
  nudgePreviewBounds({ growY: -PREVIEW_RESIZE_STEP });
});
previewRunOcrButton?.addEventListener("click", () => {
  void runPreviewRecognition();
});
previewTextEditor?.addEventListener("click", () => {
  showQuestionEditorSource();
});
previewTextEditor?.addEventListener("input", () => {
  previewDraftText = getQuestionEditorText();
  updatePreviewReviewStatus();
});
previewTextEditor?.addEventListener("focusin", (event) => {
  if (event.target?.classList?.contains("math-question-display")) {
    showQuestionEditorSource();
  }
});
previewTextEditor?.addEventListener("focusout", (event) => {
  if (!previewTextEditor.contains(event.relatedTarget)) {
    renderQuestionEditorDisplay();
  }
});
previewTextEditor?.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && previewTextEditor.classList.contains("is-source-editing")) {
    event.preventDefault();
    renderQuestionEditorDisplay();
    getQuestionEditorDisplay()?.focus();
  }
});
previewSaveButton?.addEventListener("click", () => {
  renderQuestionEditorDisplay();
  void savePreviewQuestion();
});
previewZoomOutButton?.addEventListener("click", () => {
  previewQuestionZoom = clampPreviewZoom(previewQuestionZoom - 0.25);
  updatePreviewZoomDisplay();
});
previewZoomResetButton?.addEventListener("click", () => {
  previewQuestionZoom = 1;
  updatePreviewZoomDisplay();
});
previewZoomInButton?.addEventListener("click", () => {
  previewQuestionZoom = clampPreviewZoom(previewQuestionZoom + 0.25);
  updatePreviewZoomDisplay();
});

window.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && previewModal && !previewModal.hidden) {
    closePreviewModal();
    return;
  }

  if (!previewModal || previewModal.hidden || previewTextEditor?.contains(document.activeElement)) {
    return;
  }

  if (event.key === "ArrowUp") {
    event.preventDefault();
    nudgePreviewBounds({ dy: -PREVIEW_NUDGE_STEP });
    return;
  }

  if (event.key === "ArrowDown") {
    event.preventDefault();
    nudgePreviewBounds({ dy: PREVIEW_NUDGE_STEP });
    return;
  }

  if (event.key === "ArrowLeft") {
    event.preventDefault();
    nudgePreviewBounds({ dx: -PREVIEW_NUDGE_STEP });
    return;
  }

  if (event.key === "ArrowRight") {
    event.preventDefault();
    nudgePreviewBounds({ dx: PREVIEW_NUDGE_STEP });
    return;
  }

  if (event.key.toLowerCase() === "w") {
    event.preventDefault();
    nudgePreviewBounds({ growX: PREVIEW_RESIZE_STEP });
    return;
  }

  if (event.key.toLowerCase() === "h") {
    event.preventDefault();
    nudgePreviewBounds({ growY: PREVIEW_RESIZE_STEP });
  }
});

document.addEventListener("paste", async (event) => {
  if (activePasteImageIndex === null) {
    return;
  }

  const imageFile = getImageFileFromClipboardData(event.clipboardData);

  if (!imageFile) {
    return;
  }

  event.preventDefault();
  await applyImageToCell(activePasteImageIndex, imageFile);
});

window.addEventListener("message", (event) => {
  if (event.origin !== window.location.origin) {
    return;
  }

  if (event.data?.type === "v4-selector-ready") {
    selectorWindowReady = true;
    sendQuestionsToSelectorWindow();
    return;
  }

  if (event.data?.type === "v4-selector-request") {
    handleSelectorRequest();
    return;
  }

  if (event.data?.type === "v4-selector-apply") {
    applySelectedQuestionIds(Array.isArray(event.data.selectedIds) ? event.data.selectedIds : []);
  }

  if (event.data?.type === "v4-selector-apply-and-print") {
    applySelectedQuestionIds(Array.isArray(event.data.selectedIds) ? event.data.selectedIds : []);
    window.setTimeout(() => window.print(), 120);
  }
});

selectorChannel?.addEventListener("message", (event) => {
  if (event.data?.type === "v4-selector-request") {
    handleSelectorRequest();
    return;
  }

  if (event.data?.type === "v4-selector-apply") {
    applySelectedQuestionIds(Array.isArray(event.data.selectedIds) ? event.data.selectedIds : []);
  }

  if (event.data?.type === "v4-selector-apply-and-print") {
    applySelectedQuestionIds(Array.isArray(event.data.selectedIds) ? event.data.selectedIds : []);
    window.setTimeout(() => window.print(), 120);
  }
});

window.addEventListener("beforeunload", () => {
  if (ocrWorkerPromise) {
    ocrWorkerPromise
      .then((worker) => worker.terminate())
      .catch(() => {});
  }

  selectorChannel?.close();
});

applySettings(loadSettings());
renderQuestionPreviewList();
renderPages();

const searchParams = new URLSearchParams(window.location.search);
if (searchParams.get("demo") === "example") {
  void bootstrapExampleDemo();
} else {
  void loadDefaultExampleSource();
}
