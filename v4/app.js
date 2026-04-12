import * as pdfjsLib from "./vendor/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdf.worker.mjs", import.meta.url).toString();

const APP_ASSET_VERSION = "20260412e";
const STORAGE_KEY = "rowcolpage.v4.settings.v3";
const SELECTOR_CHANNEL_NAME = "rowcolpage-v4-selector";
const SELECTOR_DB_NAME = "rowcolpage-v4";
const SELECTOR_DB_STORE = "selectorPayload";
const SELECTOR_DB_KEY = "latest";
const PDF_RENDER_SCALE = 2.2;
const OCR_MODULE_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.esm.min.js";
const CONTENT_RGB_THRESHOLD = 244;
const CONTENT_ALPHA_THRESHOLD = 18;
const CONTENT_MARGIN_PX = 28;
const CROPPED_CONTENT_MARGIN_PX = 32;

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
const cellBindings = new Map();
let activePasteImageIndex = null;
let sourceFiles = [];
let currentSourceLabel = "";
let extractedQuestions = [];
let selectorWindow = null;
let selectorWindowReady = false;
let ocrWorkerPromise = null;
let demoBootstrapPromise = null;
const selectorChannel =
  typeof window.BroadcastChannel === "function"
    ? new BroadcastChannel(SELECTOR_CHANNEL_NAME)
    : null;
let selectorDbPromise = null;

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
let previewQuestionZoom = 1;
let previewQuestionStage = null;

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

function renderCellImage(container, pane, imageUrl) {
  container.replaceChildren();
  container.classList.remove("is-empty");
  pane.classList.toggle("has-image", Boolean(imageUrl));

  if (!imageUrl) {
    container.classList.add("is-empty");
    return;
  }

  const image = document.createElement("img");
  image.className = "problem-image";
  image.alt = "題目圖片";
  image.src = imageUrl;
  container.appendChild(image);
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

  renderCellImage(binding.content, binding.pane, imageUrl);
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

  renderCellImage(content, pane, uploadedImages.get(imageIndex) ?? "");
  clearButton.hidden = !uploadedImages.has(imageIndex);
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
    renderCellImage(content, pane, "");
    clearButton.hidden = true;
  });
}

function updateReviewActionsVisibility() {
  if (!questionReviewActions) {
    return;
  }

  questionReviewActions.hidden = !extractedQuestions.length;
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
}

function openPreviewModal(question) {
  if (!previewModal || !question) {
    return;
  }

  const sourceImage = document.createElement("img");
  sourceImage.src = question.pagePreviewUrl;
  sourceImage.alt = `${question.numberLabel} 原始頁面`;

  const sourceRect = document.createElement("div");
  sourceRect.className = "question-source-rect";
  sourceRect.style.left = `${question.overlay.leftPercent}%`;
  sourceRect.style.top = `${question.overlay.topPercent}%`;
  sourceRect.style.width = `${question.overlay.widthPercent}%`;
  sourceRect.style.height = `${question.overlay.heightPercent}%`;

  const sourceStage = document.createElement("div");
  sourceStage.className = "preview-modal-source-stage";

  sourceStage.append(sourceImage, sourceRect);

  const questionImage = document.createElement("img");
  questionImage.src = question.imageUrl;
  questionImage.alt = `${question.numberLabel} 題目預覽`;

  previewQuestionStage = document.createElement("div");
  previewQuestionStage.className = "preview-modal-question-stage";
  previewQuestionStage.appendChild(questionImage);

  previewModalSource.textContent = question.sourceLabel;
  previewModalTitle.textContent = question.numberLabel;
  previewModalAnchor.textContent = question.detailText || question.title;
  previewModalSourcePreview.replaceChildren(sourceStage);
  previewModalQuestionPreview.replaceChildren(previewQuestionStage);
  previewQuestionZoom = 1;
  updatePreviewZoomDisplay();
  previewModal.hidden = false;
  document.body.classList.add("preview-modal-open");
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
    title.textContent = question.detailText || question.title;

    meta.append(heading, source, title);
    card.append(sourcePreview, thumb, meta);
    card.addEventListener("click", () => openPreviewModal(question));
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openPreviewModal(question);
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
  const payload = buildSelectorPayload(extractedQuestions);

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
  if (!extractedQuestions.length) {
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
          width: item.width,
          height: item.height,
        })),
      };
    }),
  };
}

async function getOcrWorker() {
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = (async () => {
      const { createWorker } = await import(OCR_MODULE_URL);
      return createWorker("eng");
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

function extractQuestionText(lines, bounds, pageWidth) {
  const textLines = lines
    .filter((line) => line.bottom >= bounds.top && line.top <= bounds.bottom)
    .filter((line) => line.right >= bounds.left && line.left <= bounds.right)
    .sort((a, b) => a.top - b.top || a.left - b.left)
    .map((line) => {
      if (!Array.isArray(line.items) || !line.items.length) {
        return normalizeText(line.text);
      }

      const filteredItems = line.items
        .filter((item) => item.x < pageWidth * 0.66 || item.text.length > 2 || /[()（）]/.test(item.text))
        .sort((a, b) => a.x - b.x);

      if (!filteredItems.length) {
        return "";
      }

      let text = "";

      filteredItems.forEach((item, index) => {
        const previous = filteredItems[index - 1];

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

  const merged = normalizeText(textLines.join(" "));
  const startIndex = merged.search(/[（(]\s*[）)]\s*\d+\s*[.．、]|第\s*\d+\s*題|\d+\s*[.．、]/u);
  const normalized = startIndex >= 0 ? merged.slice(startIndex).trim() : merged;
  const optionMatch = normalized.match(/^(.*?\(\s*[AaＡ]\s*\).*?\(\s*[DdＤ]\s*\)[^。]*(?:。|$))/u);
  const optionNormalized = optionMatch ? optionMatch[1].trim() : normalized;

  return optionNormalized.replace(/([？?])\s+\d+\s+(\(\s*[AaＡ]\s*\))/u, "$1 $2").trim();
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

function cropQuestionImage(pageCanvas, bounds, scale) {
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

  return trimCanvasToVisibleContent(croppedCanvas, CROPPED_CONTENT_MARGIN_PX).toDataURL("image/png");
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

  questions.forEach((question, index) => {
    uploadedImages.set(index, question.imageUrl);
  });

  activePasteImageIndex = questions.length ? 0 : null;
  renderPages();
}

function applySelectedQuestionIds(selectedIds) {
  const selectedIdSet = new Set(selectedIds);
  const selectedQuestions = extractedQuestions.filter((question) => selectedIdSet.has(question.id));

  if (!selectedQuestions.length) {
    uploadedImages.clear();
    activePasteImageIndex = null;
    renderPages();
    extractionSummary.textContent = "勾選頁沒有選任何題目，主頁未帶入題圖。";
    return;
  }

  applyExtractedQuestions(selectedQuestions);
  extractionSummary.textContent = `已依序帶入 ${selectedQuestions.length} 題到左側題目區。`;
}

function clearExtractedQuestions() {
  extractedQuestions = [];
  uploadedImages.clear();
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
      const detailText = extractQuestionText(sourcePages[anchor.sourceIndex].lines, bounds, anchor.pageWidth) || anchor.title;

      questions.push({
        id: `q-${index + 1}`,
        numberLabel,
        title: anchor.title,
        detailText,
        sourceLabel: anchor.sourceLabel,
        pageNumber: anchor.pageNumber,
        imageUrl,
        pagePreviewUrl,
        overlay: buildOverlayFromBounds(bounds, anchor.pageWidth, anchor.pageHeight),
      });
    }

    extractedQuestions = questions;
    renderQuestionPreviewList();
    extractionSummary.textContent = `已從 ${currentSourceLabel} 抽出 ${questions.length} 題，請先點縮圖放大檢查，確認後再前往下一頁。`;
    sendQuestionsToSelectorWindow();
    document.querySelector(".question-bank")?.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    console.error(error);
    extractionSummary.textContent = "來源分析失敗，請確認檔案內容清楚且含有題號。";
    window.alert("來源分析失敗，請確認檔案內容清楚且含有題號。");
  } finally {
    extractSourceButton.disabled = false;
  }
}

async function loadExampleSourceFile() {
  const response = await fetch(`./example.pdf?v=${APP_ASSET_VERSION}`, { cache: "no-store" });

  if (!response.ok) {
    throw new Error(`Unable to load example.pdf: ${response.status}`);
  }

  const blob = await response.blob();
  const file = new File([blob], "example.pdf", {
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
  activePasteImageIndex = null;
  extractedQuestions = [];
  sourceFiles = [];
  currentSourceLabel = "";
  sourceFileInput.value = "";
  applySettings({ ...DEFAULT_SETTINGS });
  renderQuestionPreviewList();
  extractionSummary.textContent = "尚未載入來源檔案。";
  renderPages();
});

printButton.addEventListener("click", () => {
  window.print();
});

sourceFileInput.addEventListener("change", () => {
  closePreviewModal();
  sourceFiles = Array.from(sourceFileInput.files ?? []);
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
}
