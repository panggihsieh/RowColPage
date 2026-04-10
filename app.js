const CELLS_PER_PAGE = 8;

const titleInput = document.querySelector("#titleInput");
const startNumberInput = document.querySelector("#startNumberInput");
const pageCountInput = document.querySelector("#pageCountInput");
const guideSelect = document.querySelector("#guideSelect");
const signatureToggle = document.querySelector("#signatureToggle");
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

function getSignatureVisible() {
  return signatureToggle.getAttribute("aria-pressed") === "true";
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

function renderPages() {
  const title = titleInput.value.trim() || "大南國小";
  const startNumber = clampNumber(startNumberInput.value, 1, 1000000, 1);
  const pageCount = clampNumber(pageCountInput.value, 1, 50, 1);
  const guideMode = guideSelect.value;
  const showSignature = getSignatureVisible();

  pagesRoot.replaceChildren();
  updateGuideMode(guideMode);
  updateSignatureVisibility(showSignature);

  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const pageFragment = pageTemplate.content.cloneNode(true);
    const pageTitle = pageFragment.querySelector(".page-title");
    const pageMeta = pageFragment.querySelector(".page-meta");
    const grid = pageFragment.querySelector(".grid");
    const page = pageFragment.querySelector(".page");

    pageTitle.textContent = title;
    pageMeta.textContent = `第 ${pageIndex + 1} 頁`;

    for (let cellIndex = 0; cellIndex < CELLS_PER_PAGE; cellIndex += 1) {
      const cellFragment = cellTemplate.content.cloneNode(true);
      const cellNumber = cellFragment.querySelector(".cell-number");
      const number = startNumber + pageIndex * CELLS_PER_PAGE + cellIndex;

      cellNumber.textContent = number;
      grid.appendChild(cellFragment);
    }

    pagesRoot.appendChild(page);
  }
}

[titleInput, startNumberInput, pageCountInput, guideSelect].forEach((element) => {
  element.addEventListener("input", renderPages);
  element.addEventListener("change", renderPages);
});

signatureToggle.addEventListener("click", () => {
  updateSignatureVisibility(!getSignatureVisible());
  renderPages();
});

printButton.addEventListener("click", () => {
  window.print();
});

renderPages();
