const selectorStatus = document.querySelector("#selectorStatus");
const selectorSummary = document.querySelector("#selectorSummary");
const selectorQuestionList = document.querySelector("#selectorQuestionList");
const selectAllButton = document.querySelector("#selectAllButton");
const clearAllButton = document.querySelector("#clearAllButton");
const applySelectionButton = document.querySelector("#applySelectionButton");
const applyAndPrintButton = document.querySelector("#applyAndPrintButton");
const SELECTOR_CHANNEL_NAME = "rowcolpage-v4-selector";
const SELECTOR_DB_NAME = "rowcolpage-v4";
const SELECTOR_DB_STORE = "selectorPayload";
const SELECTOR_DB_KEY = "latest";
const selectorChannel =
  typeof window.BroadcastChannel === "function"
    ? new BroadcastChannel(SELECTOR_CHANNEL_NAME)
    : null;
let selectorDbPromise = null;

let questions = [];
let checkedIds = new Set();

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

async function loadPersistedQuestions() {
  const db = await getSelectorDb();

  if (!db) {
    return [];
  }

  return new Promise((resolve) => {
    const transaction = db.transaction(SELECTOR_DB_STORE, "readonly");
    const store = transaction.objectStore(SELECTOR_DB_STORE);
    const request = store.get(SELECTOR_DB_KEY);

    request.onsuccess = () => {
      const payload = request.result;
      resolve(Array.isArray(payload?.questions) ? payload.questions : []);
    };

    request.onerror = () => resolve([]);
  });
}

function applyQuestions(nextQuestions, statusText) {
  questions = Array.isArray(nextQuestions) ? nextQuestions : [];
  checkedIds = new Set(questions.map((question) => question.id));

  if (questions.length) {
    selectorStatus.textContent = statusText || "已收到主頁送來的題目，請勾選要插入的題目。";
  } else if (!window.opener) {
    selectorStatus.textContent = "尚未收到題目；回主頁分析後，這一頁會自動更新。";
  }

  renderQuestions();
}

function renderQuestions() {
  selectorQuestionList.replaceChildren();

  if (!questions.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "question-preview-empty";
    emptyState.textContent = "等待抽題資料。";
    selectorQuestionList.appendChild(emptyState);
    selectorSummary.textContent = "尚未收到題目。";
    return;
  }

  selectorSummary.textContent = `共 ${questions.length} 題，目前勾選 ${checkedIds.size} 題。`;

  const fragment = document.createDocumentFragment();

  questions.forEach((question, index) => {
    const label = document.createElement("label");
    label.className = "selector-question-card";
    label.dataset.checked = String(checkedIds.has(question.id));

    const checkbox = document.createElement("input");
    checkbox.className = "selector-checkbox";
    checkbox.type = "checkbox";
    checkbox.checked = checkedIds.has(question.id);

    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        checkedIds.add(question.id);
      } else {
        checkedIds.delete(question.id);
      }

      label.dataset.checked = String(checkbox.checked);
      selectorSummary.textContent = `共 ${questions.length} 題，目前勾選 ${checkedIds.size} 題。`;
    });

    const sourcePreview = document.createElement("div");
    sourcePreview.className = "selector-source-preview";

    const sourceImage = document.createElement("img");
    sourceImage.className = "selector-source-image";
    sourceImage.src = question.pagePreviewUrl;
    sourceImage.alt = `${question.numberLabel || `第 ${index + 1} 題`} 原始頁面`;

    const sourceRect = document.createElement("div");
    sourceRect.className = "selector-source-rect";
    sourceRect.style.left = `${question.overlay.leftPercent}%`;
    sourceRect.style.top = `${question.overlay.topPercent}%`;
    sourceRect.style.width = `${question.overlay.widthPercent}%`;
    sourceRect.style.height = `${question.overlay.heightPercent}%`;

    sourcePreview.append(sourceImage, sourceRect);

    const thumb = document.createElement("img");
    thumb.className = "selector-question-image";
    thumb.src = question.imageUrl;
    thumb.alt = `第 ${index + 1} 題預覽`;

    const meta = document.createElement("div");
    meta.className = "selector-question-meta";

    const title = document.createElement("h2");
    title.textContent = question.numberLabel || `第 ${index + 1} 題`;

    const source = document.createElement("p");
    source.textContent = question.sourceLabel;

    const anchorText = document.createElement("p");
    anchorText.className = "selector-question-title";
    anchorText.textContent = question.title;

    meta.append(title, source, anchorText);
    label.append(checkbox, sourcePreview, thumb, meta);
    fragment.appendChild(label);
  });

  selectorQuestionList.appendChild(fragment);
}

function postReadyMessage() {
  if (!window.opener) {
    if (!questions.length) {
      selectorStatus.textContent = "請回主頁分析來源檔案；完成後這一頁會自動出現題目。";
    }
    return;
  }

  window.opener.postMessage({ type: "v4-selector-ready" }, window.location.origin);
}

function requestQuestionsFromMainPage() {
  selectorChannel?.postMessage({ type: "v4-selector-request", requestedAt: Date.now() });

  if (window.opener) {
    window.opener.postMessage({ type: "v4-selector-request" }, window.location.origin);
  }
}

window.addEventListener("message", (event) => {
  if (event.origin !== window.location.origin) {
    return;
  }

  if (event.data?.type !== "v4-selector-data") {
    return;
  }

  applyQuestions(event.data.questions, "已收到主頁送來的題目，請勾選要插入的題目。");
});

selectorChannel?.addEventListener("message", (event) => {
  if (event.data?.type === "v4-selector-data") {
    applyQuestions(event.data.questions, "已同步主頁最新抽題結果，可直接勾選。");
    return;
  }

  if (event.data?.type === "v4-selector-clear") {
    applyQuestions([], "主頁已清空抽題結果。");
  }
});

selectAllButton.addEventListener("click", () => {
  checkedIds = new Set(questions.map((question) => question.id));
  renderQuestions();
});

clearAllButton.addEventListener("click", () => {
  checkedIds = new Set();
  renderQuestions();
});

applySelectionButton.addEventListener("click", () => {
  const selectedIds = questions.filter((question) => checkedIds.has(question.id)).map((question) => question.id);

  if (window.opener) {
    window.opener.postMessage(
      {
        type: "v4-selector-apply",
        selectedIds,
      },
      window.location.origin,
    );
  }

  selectorChannel?.postMessage({
    type: "v4-selector-apply",
    selectedIds,
  });

  selectorStatus.textContent = "已將勾選結果送回主頁。";
});

applyAndPrintButton.addEventListener("click", () => {
  const selectedIds = questions.filter((question) => checkedIds.has(question.id)).map((question) => question.id);
  const payload = {
    type: "v4-selector-apply-and-print",
    selectedIds,
  };

  if (window.opener) {
    window.opener.postMessage(payload, window.location.origin);
  }

  selectorChannel?.postMessage(payload);
  selectorStatus.textContent = "已將勾選結果送回主頁，準備列印。";
});

renderQuestions();
postReadyMessage();
requestQuestionsFromMainPage();
void loadPersistedQuestions().then((persistedQuestions) => {
  if (persistedQuestions.length && !questions.length) {
    applyQuestions(persistedQuestions, "已載入最近一次抽題結果，可直接勾選。");
  }
});

window.addEventListener("focus", () => {
  if (!questions.length) {
    requestQuestionsFromMainPage();
  }
});

window.addEventListener("beforeunload", () => {
  selectorChannel?.close();
});
