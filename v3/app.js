const STORAGE_KEY = "rowcolpage.v3.settings";

const DEFAULT_SETTINGS = {
  title: "大南六甲",
  className: "",
  studentName: "",
  date: new Date().toISOString().slice(0, 10),
  startNumber: 1,
  pageCount: 1,
  columnCount: 2,
  rowCount: 4,
  guideMode: "none",
  showSignature: true,
};

const uploadedImages = new Map();
const cellBindings = new Map();
let activePasteImageIndex = null;

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

function clampNumber(value, min, max, fallback) {
  const parsed = Number.parseInt(value, 10);

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

function getSignatureVisible() {
  return signatureToggle.getAttribute("aria-pressed") === "true";
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);

    if (!raw) {
      return { ...DEFAULT_SETTINGS };
    }

    return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
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
  };
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
  image.alt = "幾何題目圖片";
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

[titleInput, classInput, nameInput, dateInput, startNumberInput, pageCountInput, rowCountInput, guideSelect]
  .filter(Boolean)
  .forEach((element) => {
    element.addEventListener("input", renderPages);
    element.addEventListener("change", renderPages);
  });

signatureToggle.addEventListener("click", () => {
  updateSignatureVisibility(!getSignatureVisible());
  renderPages();
});

resetButton.addEventListener("click", () => {
  localStorage.removeItem(STORAGE_KEY);
  uploadedImages.clear();
  activePasteImageIndex = null;
  applySettings({ ...DEFAULT_SETTINGS });
  renderPages();
});

printButton.addEventListener("click", () => {
  window.print();
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

applySettings(loadSettings());
renderPages();
