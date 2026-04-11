let isProcessing = false;
let currentController = null;
let activeRequestId = 0;
let activeTimeouts = new Set();
let latestImageResult = null;
let latestTextResult = null;
let currentHost = "office";
const OFFICEXADD_CONFIG = window.__OFFICEXADD_CONFIG__ || {};
const OFFICEXADD_API_BASE_URL = (OFFICEXADD_CONFIG.apiBaseUrl || window.location.origin || "").replace(/\/$/, "");
const OFFICEXADD_API_TOKEN = OFFICEXADD_CONFIG.apiToken || "";
const WEB_SEARCH_TIMEOUT_MS = 120000;
const DEFAULT_TIMEOUT_MS = 45000;
const IMAGE_TIMEOUT_MS = 120000;
const MAX_CONTEXT_CHARS = 12000;
const CONTEXT_MARKER_START = "[[EDIT_START]]";
const CONTEXT_MARKER_END = "[[EDIT_END]]";
const CONTEXT_MARKER_CURSOR = "[[CURSOR]]";
const CONTEXT_MODE_NONE = "none";
const CONTEXT_MODE_FULL = "full";
const CONTEXT_MODE_CHARS = "chars";
const CONTEXT_MODE_PAGES = "pages";
const APPROX_PAGE_CHARS = 1500;
const MODE_TEXT = "text";
const MODE_IMAGE = "image";
const UI_PREFERENCES_KEY_PREFIX = "officexadd_ui_preferences_v1";

const HOST_CONFIG = {
    office: {
        appTitle: "AI Assistant",
        subtitle: "Rewrite selected content or generate images for Office.",
        inputLabel: "引入內容",
        inputPlaceholder: "Enter text to rewrite...",
        rewriteButton: "Rewrite & Replace",
        resultTitle: "Rewritten Content",
        insertButton: "填入",
        skipPasteHelp: "勾選後只顯示結果，不會自動覆蓋目前選取內容。",
        emptyState: "Please enter instructions or select text in Office.",
        selectionLoading: "Reading selection...",
        replacing: "Replacing selection in Office...",
        replaced: "Content replaced in Office!",
        replaceError: "Error replacing content",
        imageInputPlaceholder: "例如：一張高質感產品海報，主角是咖啡杯，暖色系，16:9",
        imageHostHint: "Office",
    },
    word: {
        appTitle: "AI Assistant for Word",
        subtitle: "Rewrite selected text, or generate images and insert into Word.",
        inputLabel: "引入文章",
        inputPlaceholder: "Enter text from Word...",
        rewriteButton: "Rewrite & Replace",
        resultTitle: "Rewritten Text",
        insertButton: "填入",
        skipPasteHelp: "勾選後只顯示結果，不會自動取代 Word 內選取文字。",
        emptyState: "Please enter instructions or select text in Word.",
        selectionLoading: "Reading selection...",
        replacing: "Replacing selection in Word...",
        replaced: "Text replaced in Word!",
        replaceError: "Error replacing text",
        imageInputPlaceholder: "例如：一隻戴太空頭盔的柴犬，在月球上喝珍珠奶茶，電影感，16:9",
        imageHostHint: "Word",
    },
    powerpoint: {
        appTitle: "AI Assistant for PowerPoint",
        subtitle: "Rewrite selected slide text, or generate images and insert into PowerPoint.",
        inputLabel: "引入投影片文字",
        inputPlaceholder: "Enter text from PowerPoint...",
        rewriteButton: "Rewrite & Insert",
        resultTitle: "Rewritten Slide Text",
        insertButton: "插入投影片",
        skipPasteHelp: "勾選後只顯示結果，不會自動覆蓋目前投影片選取的文字。",
        emptyState: "Please enter instructions or select text in PowerPoint.",
        selectionLoading: "Reading selected slide text...",
        replacing: "Replacing selected slide text...",
        replaced: "Text replaced in PowerPoint!",
        replaceError: "Error replacing slide text",
        imageInputPlaceholder: "例如：科技感簡報封面背景，藍綠漸層，幾何線條，16:9",
        imageHostHint: "PowerPoint",
    },
};

function getHostConfig() {
    return HOST_CONFIG[currentHost] || HOST_CONFIG.office;
}

function isWordHost() {
    return currentHost === "word";
}

function isPowerPointHost() {
    return currentHost === "powerpoint";
}

function supportsDocumentContext() {
    return isWordHost();
}

function buildApiHeaders() {
    const headers = { "Content-Type": "application/json" };
    if (OFFICEXADD_API_TOKEN) {
        headers.Authorization = `Bearer ${OFFICEXADD_API_TOKEN}`;
        headers["X-API-Key"] = OFFICEXADD_API_TOKEN;
    }
    return headers;
}

function setStatus(message) {
    const statusBar = document.getElementById("statusBar");
    if (statusBar) {
        statusBar.textContent = message;
    }
}

function setElementText(id, value) {
    const node = document.getElementById(id);
    if (node) {
        node.textContent = value;
    }
}

function escapeHtml(value) {
    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function getGenerationMode() {
    const modeSelect = document.getElementById("generationMode");
    if (!modeSelect) {
        return MODE_TEXT;
    }
    return modeSelect.value === MODE_IMAGE ? MODE_IMAGE : MODE_TEXT;
}

function canUseLocalStorage() {
    try {
        if (typeof window === "undefined" || !window.localStorage) {
            return false;
        }
        const testKey = "__officexadd_storage_test__";
        window.localStorage.setItem(testKey, "1");
        window.localStorage.removeItem(testKey);
        return true;
    } catch (error) {
        return false;
    }
}

function getUiPreferencesStorageKey() {
    return `${UI_PREFERENCES_KEY_PREFIX}_${currentHost || "office"}`;
}

function loadUiPreferences() {
    const defaults = {
        generationMode: MODE_TEXT,
        provider: "openai",
        model: "",
        webSearchEnabled: false,
        skipPasteEnabled: false,
        contextMode: CONTEXT_MODE_CHARS,
        contextSize: "1",
        imageModel: "gemini-3.1-flash-image-preview",
        imageAspectRatio: "1:1",
        imageSize: "",
    };

    if (!canUseLocalStorage()) {
        return defaults;
    }

    try {
        const raw = window.localStorage.getItem(getUiPreferencesStorageKey());
        if (!raw) {
            return defaults;
        }
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== "object") {
            return defaults;
        }
        return {
            ...defaults,
            ...parsed,
        };
    } catch (error) {
        console.warn("Failed to read UI preferences", error);
        return defaults;
    }
}

function setSelectValueIfExists(selectNode, value) {
    if (!selectNode || typeof value !== "string") {
        return;
    }
    const hasValue = Array.from(selectNode.options || []).some((option) => option.value === value);
    if (hasValue) {
        selectNode.value = value;
    }
}

function applyUiPreferences(preferences) {
    if (!preferences || typeof preferences !== "object") {
        return;
    }

    const generationMode = document.getElementById("generationMode");
    setSelectValueIfExists(generationMode, preferences.generationMode);

    const providerSelect = document.getElementById("providerSelect");
    setSelectValueIfExists(providerSelect, preferences.provider);

    const modelInput = document.getElementById("modelInput");
    if (modelInput && typeof preferences.model === "string") {
        modelInput.value = preferences.model;
    }

    const webSearchToggle = document.getElementById("webSearchToggle");
    if (webSearchToggle && typeof preferences.webSearchEnabled === "boolean") {
        webSearchToggle.checked = preferences.webSearchEnabled;
    }

    const skipPasteToggle = document.getElementById("skipPasteToggle");
    if (skipPasteToggle && typeof preferences.skipPasteEnabled === "boolean") {
        skipPasteToggle.checked = preferences.skipPasteEnabled;
    }

    const contextMode = document.getElementById("contextMode");
    setSelectValueIfExists(contextMode, preferences.contextMode);

    const contextSize = document.getElementById("contextSize");
    if (contextSize && preferences.contextSize !== undefined && preferences.contextSize !== null) {
        const value = String(preferences.contextSize).trim();
        if (value) {
            contextSize.value = value;
        }
    }

    const imageModelInput = document.getElementById("imageModelInput");
    if (imageModelInput && typeof preferences.imageModel === "string") {
        imageModelInput.value = preferences.imageModel;
    }

    const imageAspectRatio = document.getElementById("imageAspectRatio");
    setSelectValueIfExists(imageAspectRatio, preferences.imageAspectRatio);

    const imageSize = document.getElementById("imageSize");
    setSelectValueIfExists(imageSize, preferences.imageSize);
}

function collectUiPreferences() {
    const generationMode = document.getElementById("generationMode");
    const providerSelect = document.getElementById("providerSelect");
    const modelInput = document.getElementById("modelInput");
    const webSearchToggle = document.getElementById("webSearchToggle");
    const skipPasteToggle = document.getElementById("skipPasteToggle");
    const contextMode = document.getElementById("contextMode");
    const contextSize = document.getElementById("contextSize");
    const imageModelInput = document.getElementById("imageModelInput");
    const imageAspectRatio = document.getElementById("imageAspectRatio");
    const imageSize = document.getElementById("imageSize");

    return {
        generationMode: generationMode ? generationMode.value : MODE_TEXT,
        provider: providerSelect ? providerSelect.value : "openai",
        model: modelInput ? modelInput.value.trim() : "",
        webSearchEnabled: webSearchToggle ? Boolean(webSearchToggle.checked) : false,
        skipPasteEnabled: skipPasteToggle ? Boolean(skipPasteToggle.checked) : false,
        contextMode: contextMode ? contextMode.value : CONTEXT_MODE_NONE,
        contextSize: contextSize ? String(contextSize.value || "").trim() : "1",
        imageModel: imageModelInput ? imageModelInput.value.trim() : "",
        imageAspectRatio: imageAspectRatio ? imageAspectRatio.value : "1:1",
        imageSize: imageSize ? imageSize.value : "",
    };
}

function persistUiPreferences() {
    if (!canUseLocalStorage()) {
        return;
    }
    try {
        const preferences = collectUiPreferences();
        window.localStorage.setItem(getUiPreferencesStorageKey(), JSON.stringify(preferences));
    } catch (error) {
        console.warn("Failed to save UI preferences", error);
    }
}

function setPrimaryButtonText() {
    const button = document.getElementById("rewriteBtn");
    if (!button || isProcessing) {
        return;
    }
    const hostConfig = getHostConfig();
    button.textContent = getGenerationMode() === MODE_IMAGE ? "Generate Image & Insert" : hostConfig.rewriteButton;
}

function updateModeUI() {
    const hostConfig = getHostConfig();
    const mode = getGenerationMode();
    const isImageMode = mode === MODE_IMAGE;
    const canUseContext = !isImageMode && supportsDocumentContext();

    setElementText("appTitle", hostConfig.appTitle);
    setElementText("appSubtitle", hostConfig.subtitle);

    const imageCard = document.getElementById("imageCard");
    if (imageCard) {
        imageCard.classList.toggle("context-hidden", !isImageMode);
    }
    const providerCard = document.getElementById("providerCard");
    if (providerCard) {
        providerCard.classList.toggle("context-hidden", isImageMode);
    }
    const modelCard = document.getElementById("modelCard");
    if (modelCard) {
        modelCard.classList.toggle("context-hidden", isImageMode);
    }
    const webSearchCard = document.getElementById("webSearchCard");
    if (webSearchCard) {
        webSearchCard.classList.remove("context-hidden");
    }
    const webSearchRow = document.getElementById("webSearchRow");
    if (webSearchRow) {
        webSearchRow.classList.toggle("context-hidden", isImageMode);
    }
    const webSearchHelp = document.getElementById("webSearchHelp");
    if (webSearchHelp) {
        webSearchHelp.classList.toggle("context-hidden", isImageMode);
    }
    const contextCard = document.getElementById("contextCard");
    const contextMode = document.getElementById("contextMode");
    const contextSizeRow = document.getElementById("contextSizeRow");
    const contextHelp = document.getElementById("contextHelp");
    if (contextCard) {
        contextCard.classList.toggle("context-hidden", !canUseContext);
    }
    if (!supportsDocumentContext()) {
        if (contextMode) {
            contextMode.value = CONTEXT_MODE_NONE;
        }
        if (contextSizeRow) {
            contextSizeRow.classList.add("context-hidden");
        }
        if (contextHelp) {
            contextHelp.textContent = "PowerPoint 目前只會讀取選取文字，不會抓整份文件/投影片上下文。";
        }
    }

    const title = document.getElementById("resultTitle");
    if (title) {
        title.textContent = isImageMode ? "Generated Image" : hostConfig.resultTitle;
    }

    const inputLabel = document.getElementById("inputLabel");
    if (inputLabel) {
        inputLabel.textContent = isImageMode ? "生圖需求" : hostConfig.inputLabel;
    }
    const instructionLabel = document.getElementById("instructionLabel");
    if (instructionLabel) {
        instructionLabel.textContent = isImageMode ? "附加指令 (可選)" : "Instructions:";
    }
    const inputText = document.getElementById("inputText");
    if (inputText) {
        inputText.placeholder = isImageMode
            ? hostConfig.imageInputPlaceholder
            : hostConfig.inputPlaceholder;
    }
    const instructionText = document.getElementById("instructionText");
    if (instructionText) {
        instructionText.placeholder = isImageMode
            ? "例如：偏寫實、保留暖色光影、細節清晰"
            : "e.g., Make it more formal";
    }
    const insertBtn = document.getElementById("insertBtn");
    if (insertBtn) {
        insertBtn.textContent = isImageMode ? "插入圖片" : hostConfig.insertButton;
    }
    const skipPasteHelp = document.getElementById("skipPasteHelp");
    if (skipPasteHelp) {
        skipPasteHelp.textContent = isImageMode
            ? `勾選後只顯示圖片，不會自動插入 ${hostConfig.imageHostHint}。`
            : hostConfig.skipPasteHelp;
    }
    if (canUseContext) {
        updateContextControls();
    }

    setPrimaryButtonText();
}

function setResultContent(content, options = {}) {
    const { isHtml = true, allowActions = true } = options;
    const resultContent = document.getElementById("resultContent");
    const copyBtn = document.getElementById("copyBtn");
    const insertBtn = document.getElementById("insertBtn");
    latestImageResult = null;
    latestTextResult = null;
    if (!resultContent) {
        return;
    }
    if (isHtml) {
        resultContent.innerHTML = content;
    } else {
        resultContent.textContent = content;
    }
    const hasText = resultContent.textContent.trim().length > 0;
    if (copyBtn) {
        copyBtn.disabled = !hasText || !allowActions;
    }
    if (insertBtn) {
        insertBtn.disabled = !hasText || !allowActions;
    }

    if (window.renderMathInElement) {
        window.renderMathInElement(resultContent, {
            delimiters: [
                { left: "$$", right: "$$", display: true },
                { left: "$", right: "$", display: false },
                { left: "\\(", right: "\\)", display: false },
                { left: "\\[", right: "\\]", display: true },
            ],
        });
    }
}

function setImageResult(imageBase64, mimeType, prompt, modelName) {
    const resultContent = document.getElementById("resultContent");
    const copyBtn = document.getElementById("copyBtn");
    const insertBtn = document.getElementById("insertBtn");
    if (!resultContent) {
        return;
    }
    const safePrompt = escapeHtml(prompt || "");
    const safeModel = escapeHtml(modelName || "");
    const mediaType = mimeType || "image/png";
    const dataUrl = `data:${mediaType};base64,${imageBase64}`;
    resultContent.innerHTML = `
        <div style="display:flex; flex-direction:column; gap:8px;">
            <img src="${dataUrl}" alt="Generated image" style="width:100%; border-radius:10px; border:1px solid #e1d6c7;" />
            ${safeModel ? `<small>Model: ${safeModel}</small>` : ""}
            ${safePrompt ? `<small>Prompt: ${safePrompt}</small>` : ""}
        </div>
    `;
    latestImageResult = {
        imageBase64,
        mimeType: mediaType,
    };
    latestTextResult = null;
    if (copyBtn) {
        copyBtn.disabled = true;
    }
    if (insertBtn) {
        insertBtn.disabled = false;
    }
}

function setProcessingState(active) {
    isProcessing = active;
    const button = document.getElementById("rewriteBtn");
    if (!button) {
        return;
    }
    if (active) {
        button.textContent = "Stop";
        button.classList.add("is-stop");
    } else {
        setPrimaryButtonText();
        button.classList.remove("is-stop");
        currentController = null;
    }
}

function cancelCurrentRequest() {
    if (!isProcessing) {
        return;
    }
    activeRequestId += 1;
    if (currentController) {
        currentController.abort();
        currentController = null;
    }
    setProcessingState(false);
    setStatus("Canceled by user.");
}

function cleanupResources() {
    // 取消當前請求
    if (currentController) {
        currentController.abort();
        currentController = null;
    }
    // 清除所有 timeout
    activeTimeouts.forEach(timeoutId => {
        clearTimeout(timeoutId);
    });
    activeTimeouts.clear();
    // 重置狀態
    isProcessing = false;
}

function updateContextControls() {
    const modeSelect = document.getElementById("contextMode");
    const contextCard = document.getElementById("contextCard");
    const sizeRow = document.getElementById("contextSizeRow");
    const unitSpan = document.getElementById("contextUnit");
    const help = document.getElementById("contextHelp");
    if (!modeSelect || !contextCard || !sizeRow || !unitSpan || !help) {
        return;
    }
    if (!supportsDocumentContext()) {
        contextCard.classList.add("context-hidden");
        modeSelect.value = CONTEXT_MODE_NONE;
        sizeRow.classList.add("context-hidden");
        help.textContent = "PowerPoint 目前只會讀取選取文字，不會抓整份文件/投影片上下文。";
        return;
    }
    contextCard.classList.remove("context-hidden");
    const mode = modeSelect.value;
    if (mode === CONTEXT_MODE_CHARS) {
        sizeRow.classList.remove("context-hidden");
        unitSpan.textContent = "字";
        help.textContent = "會取選取位置前後 N 字的內容當上下文。";
    } else if (mode === CONTEXT_MODE_PAGES) {
        sizeRow.classList.remove("context-hidden");
        unitSpan.textContent = "頁";
        help.textContent = "會取當前頁前後 N 頁的內容當上下文。";
    } else if (mode === CONTEXT_MODE_FULL) {
        sizeRow.classList.add("context-hidden");
        help.textContent = "會把全文送出，並標記目前游標或選取區間。";
    } else {
        sizeRow.classList.add("context-hidden");
        help.textContent = "不使用文件上下文。";
    }
}

function parseContextSize() {
    const sizeInput = document.getElementById("contextSize");
    if (!sizeInput) {
        return 0;
    }
    const parsed = parseInt(sizeInput.value, 10);
    if (Number.isNaN(parsed) || parsed < 1) {
        return 0;
    }
    return parsed;
}

function getSelectedTextFromOffice() {
    return new Promise((resolve) => {
        if (!Office || !Office.context || !Office.context.document) {
            resolve("");
            return;
        }
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || "");
            } else {
                resolve("");
            }
        });
    });
}

async function getPowerPointSnapshot() {
    const fallbackText = await getSelectedTextFromOffice();
    return { selectionText: fallbackText, documentText: "", paragraphHints: [] };
}

async function getDocumentSnapshot(options = {}) {
    const includeDocumentText = Boolean(options.includeDocumentText);
    if (isPowerPointHost()) {
        return getPowerPointSnapshot();
    }
    if (!isWordHost() || typeof Word === "undefined") {
        const fallbackText = await getSelectedTextFromOffice();
        return { selectionText: fallbackText, documentText: "", paragraphHints: [] };
    }
    try {
        return await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load("text");

            let body;
            let paragraphs;
            if (includeDocumentText) {
                body = context.document.body;
                paragraphs = selection.paragraphs;
                body.load("text");
                paragraphs.load("text");
            }

            await context.sync().catch(err => {
                console.warn("Context sync error:", err);
                throw err;
            });
            const paragraphHints = includeDocumentText && paragraphs
                ? (paragraphs.items || []).map((item) => item.text || "").filter(Boolean)
                : [];
            return {
                selectionText: selection.text || "",
                documentText: includeDocumentText && body ? body.text || "" : "",
                paragraphHints
            };
        });
    } catch (error) {
        console.warn("Word snapshot failed", error);
        const fallbackText = await getSelectedTextFromOffice();
        return { selectionText: fallbackText, documentText: "", paragraphHints: [] };
    }
}

function findAllOccurrences(text, search) {
    if (!search) {
        return [];
    }
    const positions = [];
    let index = text.indexOf(search);
    while (index !== -1) {
        positions.push(index);
        index = text.indexOf(search, index + 1);
    }
    return positions;
}

function findCursorIndex(documentText, paragraphHints) {
    const hint = paragraphHints.find((text) => text.trim());
    if (!hint) {
        return 0;
    }
    const index = documentText.indexOf(hint);
    if (index === -1) {
        return 0;
    }
    return index;
}

function findSelectionIndex(documentText, selectionText, paragraphHints) {
    if (!selectionText) {
        return findCursorIndex(documentText, paragraphHints);
    }
    const occurrences = findAllOccurrences(documentText, selectionText);
    if (!occurrences.length) {
        return -1;
    }
    if (occurrences.length === 1) {
        return occurrences[0];
    }
    const primaryParagraph = paragraphHints.find((text) => text.trim());
    if (!primaryParagraph) {
        return occurrences[0];
    }
    const paragraphOccurrences = findAllOccurrences(documentText, primaryParagraph);
    if (!paragraphOccurrences.length) {
        return occurrences[0];
    }
    for (const occ of occurrences) {
        for (const paraStart of paragraphOccurrences) {
            const paraEnd = paraStart + primaryParagraph.length;
            if (occ >= paraStart && occ <= paraEnd) {
                return occ;
            }
        }
    }
    return occurrences[0];
}

function limitDocumentText(documentText, selectionIndex) {
    if (!documentText) {
        return { text: documentText, selectionIndex, note: "" };
    }
    if (documentText.length <= MAX_CONTEXT_CHARS) {
        return { text: documentText, selectionIndex, note: "" };
    }

    if (selectionIndex < 0) {
        return {
            text: documentText.slice(0, MAX_CONTEXT_CHARS),
            selectionIndex,
            note: `Context truncated to ${MAX_CONTEXT_CHARS} characters from start.`,
        };
    }

    const half = Math.floor(MAX_CONTEXT_CHARS / 2);
    const start = Math.max(0, selectionIndex - half);
    const end = Math.min(documentText.length, start + MAX_CONTEXT_CHARS);
    const sliceStart = Math.max(0, end - MAX_CONTEXT_CHARS);
    const sliceEnd = Math.min(documentText.length, sliceStart + MAX_CONTEXT_CHARS);
    const trimmedText = documentText.slice(sliceStart, sliceEnd);
    const adjustedIndex = Math.max(0, selectionIndex - sliceStart);

    return {
        text: trimmedText,
        selectionIndex: adjustedIndex,
        note: `Context truncated to ${MAX_CONTEXT_CHARS} characters around selection.`,
    };
}

function buildMarkedContext(beforeText, selectionText, afterText) {
    if (selectionText) {
        return `${beforeText}${CONTEXT_MARKER_START}${selectionText}${CONTEXT_MARKER_END}${afterText}`;
    }
    return `${beforeText}${CONTEXT_MARKER_CURSOR}${afterText}`;
}

function buildFullContext(documentText, selectionText, selectionIndex) {
    if (!documentText) {
        return buildMarkedContext("", selectionText, "");
    }
    if (selectionIndex < 0) {
        return buildMarkedContext("", selectionText, documentText);
    }
    const start = selectionIndex;
    const end = selectionIndex + (selectionText ? selectionText.length : 0);
    const before = documentText.slice(0, start);
    const after = documentText.slice(end);
    return buildMarkedContext(before, selectionText, after);
}

function buildCharContext(documentText, selectionText, selectionIndex, size) {
    if (!documentText) {
        return buildMarkedContext("", selectionText, "");
    }
    const start = selectionIndex < 0 ? 0 : selectionIndex;
    const end = selectionIndex < 0 ? 0 : selectionIndex + (selectionText ? selectionText.length : 0);
    const before = documentText.slice(Math.max(0, start - size), start);
    const after = documentText.slice(end, end + size);
    return buildMarkedContext(before, selectionText, after);
}

function splitPages(documentText) {
    const pageBreak = "\f";
    if (documentText.includes(pageBreak)) {
        return { pages: documentText.split(pageBreak), delimiter: pageBreak, approx: false };
    }
    const pages = [];
    for (let i = 0; i < documentText.length; i += APPROX_PAGE_CHARS) {
        pages.push(documentText.slice(i, i + APPROX_PAGE_CHARS));
    }
    return { pages, delimiter: "", approx: true };
}

function buildPageContext(documentText, selectionText, selectionIndex, size) {
    if (!documentText) {
        return { context: buildMarkedContext("", selectionText, ""), approx: false };
    }
    const { pages, delimiter, approx } = splitPages(documentText);
    const joiner = delimiter || "\n\n";
    if (!pages.length) {
        return { context: buildMarkedContext("", selectionText, ""), approx };
    }
    let offset = 0;
    const pageStarts = pages.map((page) => {
        const start = offset;
        offset += page.length + delimiter.length;
        return start;
    });
    const safeIndex = selectionIndex < 0 ? 0 : selectionIndex;
    let pageIndex = 0;
    for (let i = 0; i < pageStarts.length; i += 1) {
        if (safeIndex >= pageStarts[i]) {
            pageIndex = i;
        } else {
            break;
        }
    }
    const startPage = Math.max(0, pageIndex - size);
    const endPage = Math.min(pages.length - 1, pageIndex + size);
    const beforePages = pages.slice(startPage, pageIndex).join(joiner);
    const afterPages = pages.slice(pageIndex + 1, endPage + 1).join(joiner);
    const currentPage = pages[pageIndex] || "";
    const pageStartOffset = pageStarts[pageIndex] || 0;
    const relativeIndex = Math.max(0, safeIndex - pageStartOffset);
    const selectionLength = selectionText ? selectionText.length : 0;
    const before = currentPage.slice(0, relativeIndex);
    const after = currentPage.slice(relativeIndex + selectionLength);
    const currentWithMarker = buildMarkedContext(before, selectionText, after);
    const context = [beforePages, currentWithMarker, afterPages].filter(Boolean).join(joiner);
    return { context, approx };
}

function buildContextFromSnapshot(snapshot, mode, size) {
    const originalDocumentText = snapshot.documentText || "";
    const selectionText = snapshot.selectionText || "";
    const rawSelectionIndex = findSelectionIndex(originalDocumentText, selectionText, snapshot.paragraphHints || []);
    let contextNote = rawSelectionIndex < 0 ? "Selection location not found in document text." : "";
    const limited = limitDocumentText(originalDocumentText, rawSelectionIndex);
    const documentText = limited.text || "";
    const selectionIndex = limited.selectionIndex;
    if (limited.note) {
        contextNote = contextNote ? `${contextNote} ${limited.note}` : limited.note;
    }
    if (mode === CONTEXT_MODE_FULL) {
        return { contextText: buildFullContext(documentText, selectionText, selectionIndex), note: contextNote };
    }
    if (mode === CONTEXT_MODE_CHARS) {
        const safeSize = size > 0 ? size : 200;
        return { contextText: buildCharContext(documentText, selectionText, selectionIndex, safeSize), note: contextNote };
    }
    if (mode === CONTEXT_MODE_PAGES) {
        const safeSize = size > 0 ? size : 1;
        const result = buildPageContext(documentText, selectionText, selectionIndex, safeSize);
        let note = result.approx ? "Page boundaries approximated by characters." : "";
        if (contextNote) {
            note = note ? `${note} ${contextNote}` : contextNote;
        }
        return { contextText: result.context, note };
    }
    return { contextText: "", note: contextNote };
}

function formatRequestError(error, didTimeout, timeoutMessage, provider) {
    if (didTimeout) {
        return timeoutMessage || "Request timed out. Try again or disable web search.";
    }
    if (error && error.name === "AbortError") {
        return "Request canceled.";
    }
    const message = error && error.message ? error.message : "Unknown error";
    if (message.includes("Load failed") || message.includes("Failed to fetch")) {
        return "Network error. Please check the backend service and try again.";
    }
    const lowered = message.toLowerCase();
    if (lowered.includes("unauthorized") || lowered.includes("api error: 401")) {
        if (provider === "ollama") {
            return "Authorization failed for Ollama. Check AI_BASE_URL / AI_API_KEY, or switch provider to OpenAI.";
        }
        if (provider === "google") {
            return "Authorization failed for Google image API. Check GOOGLE_API_KEY in backend .env.";
        }
        return "Authorization failed for the selected provider. Check server-side API credentials.";
    }
    return `Error: ${message}`;
}

function htmlToPlainText(html) {
    const container = document.createElement("div");
    container.innerHTML = html;
    return (container.textContent || "").trim();
}

async function copyResult() {
    const resultContent = document.getElementById("resultContent");
    const copyBtn = document.getElementById("copyBtn");
    if (!resultContent) {
        return;
    }
    const html = resultContent.innerHTML.trim();
    const text = resultContent.textContent.trim();
    if (!text) {
        return;
    }
    try {
        if (navigator.clipboard && window.ClipboardItem) {
            const item = new ClipboardItem({
                "text/html": new Blob([html], { type: "text/html" }),
                "text/plain": new Blob([text], { type: "text/plain" }),
            });
            await navigator.clipboard.write([item]);
        } else if (navigator.clipboard) {
            await navigator.clipboard.writeText(text);
        } else {
            throw new Error("Clipboard not available");
        }
        setStatus("Copied to clipboard.");
        if (copyBtn) {
            copyBtn.textContent = "Copied";
            const timeoutId = setTimeout(() => {
                copyBtn.textContent = "Copy";
                activeTimeouts.delete(timeoutId);
            }, 1200);
            activeTimeouts.add(timeoutId);
        }
    } catch (error) {
        setStatus("Copy failed.");
    }
}

function setSelectedDataAsyncPromise(data, options) {
    return new Promise((resolve, reject) => {
        if (!Office || !Office.context || !Office.context.document) {
            reject(new Error("Office context is not available."));
            return;
        }
        Office.context.document.setSelectedDataAsync(data, options, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(new Error(asyncResult.error.message));
            } else {
                resolve();
            }
        });
    });
}

async function insertTextIntoOffice(html, plainText) {
    if (isWordHost()) {
        await setSelectedDataAsyncPromise(html, { coercionType: Office.CoercionType.Html });
        return;
    }
    const textPayload = plainText && plainText.trim() ? plainText : htmlToPlainText(html);
    await setSelectedDataAsyncPromise(textPayload, { coercionType: Office.CoercionType.Text });
}

async function insertImageIntoOffice(imageBase64) {
    if (isWordHost() && typeof Word !== "undefined") {
        try {
            await Word.run(async (context) => {
                const range = context.document.getSelection();
                range.insertInlinePictureFromBase64(imageBase64, Word.InsertLocation.replace);
                await context.sync();
            });
            return;
        } catch (error) {
            console.warn("Word image insertion fallback triggered", error);
        }
    }
    await setSelectedDataAsyncPromise(imageBase64, { coercionType: Office.CoercionType.Image });
}

function getResultPayload() {
    if (latestTextResult && latestTextResult.html) {
        return latestTextResult;
    }
    const resultContent = document.getElementById("resultContent");
    if (!resultContent) {
        return null;
    }
    const html = resultContent.innerHTML.trim();
    const text = resultContent.textContent.trim();
    if (!text) {
        return null;
    }
    return { html, text };
}

async function insertResultIntoDocument() {
    const hostConfig = getHostConfig();
    try {
        if (latestImageResult && latestImageResult.imageBase64) {
            setStatus(`Inserting image into ${hostConfig.imageHostHint}...`);
            await insertImageIntoOffice(latestImageResult.imageBase64);
            setStatus(`Image inserted into ${hostConfig.imageHostHint}.`);
            return true;
        }

        const payload = getResultPayload();
        if (!payload) {
            return false;
        }

        setStatus(hostConfig.replacing);
        await insertTextIntoOffice(payload.html, payload.text);
        setStatus(hostConfig.replaced);
        return true;
    } catch (error) {
        setStatus(`Insert failed: ${error.message}`);
        return false;
    }
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        currentHost = "word";
    } else if (info.host === Office.HostType.PowerPoint) {
        currentHost = "powerpoint";
    } else {
        currentHost = "office";
    }

    if (info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint) {
        window.addEventListener("unload", cleanupResources);
        window.addEventListener("beforeunload", cleanupResources);
        document.getElementById("rewriteBtn").onclick = rewriteText;
        const copyBtn = document.getElementById("copyBtn");
        if (copyBtn) {
            copyBtn.onclick = copyResult;
        }
        const insertBtn = document.getElementById("insertBtn");
        if (insertBtn) {
            insertBtn.onclick = insertResultIntoDocument;
        }

        const contextMode = document.getElementById("contextMode");
        if (contextMode) {
            contextMode.addEventListener("change", () => {
                updateContextControls();
                persistUiPreferences();
            });
        }
        const providerSelect = document.getElementById("providerSelect");
        if (providerSelect) {
            providerSelect.addEventListener("change", () => {
                refreshModelOptions(providerSelect.value);
                persistUiPreferences();
            });
        }

        const generationMode = document.getElementById("generationMode");
        if (generationMode) {
            generationMode.addEventListener("change", () => {
                updateModeUI();
                persistUiPreferences();
            });
        }

        applyUiPreferences(loadUiPreferences());
        refreshModelOptions(providerSelect ? providerSelect.value : "openai");
        updateModeUI();

        const fieldsToPersist = [
            "webSearchToggle",
            "skipPasteToggle",
            "modelInput",
            "contextSize",
            "imageModelInput",
            "imageAspectRatio",
            "imageSize",
        ];
        fieldsToPersist.forEach((id) => {
            const node = document.getElementById(id);
            if (!node) {
                return;
            }
            const eventName = node.tagName === "INPUT" && node.type !== "checkbox" ? "input" : "change";
            node.addEventListener(eventName, persistUiPreferences);
            if (eventName !== "change") {
                node.addEventListener("change", persistUiPreferences);
            }
        });
    }
});

function buildImagePrompt(inputText, instructionText, selectionText) {
    const chunks = [];
    if (inputText && inputText.trim()) {
        chunks.push(inputText.trim());
    } else if (selectionText && selectionText.trim()) {
        chunks.push(selectionText.trim());
    }
    if (instructionText && instructionText.trim()) {
        chunks.push(instructionText.trim());
    }
    return chunks.join("\n\n");
}

async function rewriteText() {
    if (isProcessing) {
        cancelCurrentRequest();
        return;
    }

    const hostConfig = getHostConfig();
    const generationMode = getGenerationMode();
    const isImageMode = generationMode === MODE_IMAGE;
    const inputTextElement = document.getElementById("inputText");
    const inputText = inputTextElement ? inputTextElement.value : "";
    const instructionNode = document.getElementById("instructionText");
    const instructionText = instructionNode ? instructionNode.value : "";
    const providerNode = document.getElementById("providerSelect");
    const providerChoice = providerNode ? providerNode.value : "openai";
    const modelNode = document.getElementById("modelInput");
    const modelChoice = modelNode ? modelNode.value.trim() : "";
    const webSearchNode = document.getElementById("webSearchToggle");
    const useWebSearch = webSearchNode ? webSearchNode.checked : false;
    const skipPasteNode = document.getElementById("skipPasteToggle");
    const skipPaste = skipPasteNode ? skipPasteNode.checked : false;
    const contextModeNode = document.getElementById("contextMode");
    const requestedContextMode = contextModeNode ? contextModeNode.value : CONTEXT_MODE_NONE;
    const contextMode = supportsDocumentContext() ? requestedContextMode : CONTEXT_MODE_NONE;
    const contextSize = parseContextSize();
    const requestId = activeRequestId + 1;
    activeRequestId = requestId;
    currentController = new AbortController();

    setResultContent("Processing...", { isHtml: false, allowActions: false });
    setProcessingState(true);
    setStatus("Preparing request...");

    try {
        setStatus(hostConfig.selectionLoading);
        const needsDocument = !isImageMode && contextMode !== CONTEXT_MODE_NONE;
        const snapshot = await getDocumentSnapshot({ includeDocumentText: needsDocument });
        const selectedText = snapshot.selectionText || "";

        if (isImageMode) {
            const imagePrompt = buildImagePrompt(inputText, instructionText, selectedText);
            const imageModelNode = document.getElementById("imageModelInput");
            const imageAspectNode = document.getElementById("imageAspectRatio");
            const imageSizeNode = document.getElementById("imageSize");
            const imageModel = imageModelNode ? imageModelNode.value.trim() : "";
            const imageAspect = imageAspectNode ? imageAspectNode.value : "";
            const imageSize = imageSizeNode ? imageSizeNode.value : "";

            if (!imagePrompt.trim()) {
                setResultContent(`請輸入生圖需求，或先在 ${hostConfig.imageHostHint} 選取一段文字。`, { isHtml: false, allowActions: false });
                setProcessingState(false);
                setStatus("Idle");
                return;
            }

            let didTimeout = false;
            try {
                setStatus("Calling Google image model...");
                const payload = {
                    prompt: imagePrompt,
                };
                if (imageModel) {
                    payload.model = imageModel;
                }
                if (imageAspect) {
                    payload.aspect_ratio = imageAspect;
                }
                if (imageSize) {
                    payload.image_size = imageSize;
                }

                const timeoutId = setTimeout(() => {
                    didTimeout = true;
                    if (currentController) {
                        currentController.abort();
                    }
                    activeTimeouts.delete(timeoutId);
                }, IMAGE_TIMEOUT_MS);
                activeTimeouts.add(timeoutId);

                let response;
                try {
                    response = await fetch(`${OFFICEXADD_API_BASE_URL}/api/generate-image`, {
                        method: "POST",
                        headers: buildApiHeaders(),
                        body: JSON.stringify(payload),
                        signal: currentController.signal,
                    });
                } finally {
                    clearTimeout(timeoutId);
                    activeTimeouts.delete(timeoutId);
                }

                let data;
                if (!response.ok) {
                    try {
                        data = await response.json();
                    } catch (parseError) {
                        data = {};
                    }
                    const errorDetail = data && data.error ? data.error : "";
                    const errorMessage = errorDetail
                        ? `API error: ${response.status} - ${errorDetail}`
                        : `API error: ${response.status}`;
                    throw new Error(errorMessage);
                } else {
                    data = await response.json();
                }

                if (requestId !== activeRequestId || !isProcessing) {
                    return;
                }

                const imageBase64 = data.image_base64 || "";
                if (!imageBase64) {
                    throw new Error("API returned empty image data.");
                }
                const mimeType = data.mime_type || "image/png";
                setImageResult(imageBase64, mimeType, imagePrompt, data.model || imageModel);

                if (skipPaste) {
                    setStatus("Done.");
                    setProcessingState(false);
                    return;
                }

                const inserted = await insertResultIntoDocument();
                if (inserted && inputTextElement) {
                    inputTextElement.value = "";
                }
                setStatus(inserted ? "Done." : "Insert failed.");
                setProcessingState(false);
                return;
            } catch (apiError) {
                if (requestId !== activeRequestId || !isProcessing) {
                    return;
                }
                const message = formatRequestError(
                    apiError,
                    didTimeout,
                    "Image generation timed out. Try a shorter prompt.",
                    "google"
                );
                setResultContent(message, { isHtml: false, allowActions: false });
                if (message.startsWith("Image generation timed out")) {
                    setStatus("Image generation timed out.");
                } else if (message.startsWith("Request canceled")) {
                    setStatus("Canceled by user.");
                } else if (message.startsWith("Network error")) {
                    setStatus("Network error.");
                } else {
                    setStatus("Error during image generation.");
                }
                setProcessingState(false);
                return;
            }
        }

        let textToRewrite = inputText;
        if (selectedText.trim()) {
            textToRewrite = selectedText;
            if (inputTextElement) {
                inputTextElement.value = textToRewrite;
            }
        }

        if (!textToRewrite.trim() && !instructionText.trim()) {
            setResultContent(hostConfig.emptyState, { isHtml: false, allowActions: false });
            setProcessingState(false);
            setStatus("Idle");
            return;
        }

        let contextText = "";
        let contextNote = "";
        if (contextMode !== CONTEXT_MODE_NONE) {
            if (snapshot.documentText && snapshot.documentText.trim()) {
                setStatus("Building context...");
                const contextResult = buildContextFromSnapshot(snapshot, contextMode, contextSize);
                contextText = contextResult.contextText;
                contextNote = contextResult.note;
            } else {
                contextNote = "Context unavailable from document.";
            }
        }

        if (!textToRewrite.trim()) {
            setStatus("Generating from instruction...");
        }

        let didTimeout = false;
        try {
            setStatus(useWebSearch ? "Using web search tool..." : "Calling AI model...");
            const payload = {
                text: textToRewrite,
                instruction: instructionText,
                provider: providerChoice,
                use_web_search: useWebSearch
            };
            if (modelChoice) {
                payload.model = modelChoice;
            }
            if (contextMode !== CONTEXT_MODE_NONE && contextText) {
                payload.context_mode = contextMode;
                payload.context_text = contextText;
                if (contextNote) {
                    payload.context_note = contextNote;
                }
            } else if (contextMode !== CONTEXT_MODE_NONE && contextNote) {
                payload.context_mode = contextMode;
                payload.context_note = contextNote;
            }

            const timeoutMs = useWebSearch ? WEB_SEARCH_TIMEOUT_MS : DEFAULT_TIMEOUT_MS;
            const timeoutId = setTimeout(() => {
                didTimeout = true;
                if (currentController) {
                    currentController.abort();
                }
                activeTimeouts.delete(timeoutId);
            }, timeoutMs);
            activeTimeouts.add(timeoutId);

            let response;
            try {
                response = await fetch(`${OFFICEXADD_API_BASE_URL}/api/rewrite`, {
                    method: "POST",
                    headers: buildApiHeaders(),
                    body: JSON.stringify(payload),
                    signal: currentController.signal,
                });
            } finally {
                clearTimeout(timeoutId);
                activeTimeouts.delete(timeoutId);
            }

            let data;
            if (!response.ok) {
                try {
                    data = await response.json();
                } catch (parseError) {
                    data = {};
                }
                const errorDetail = data && data.error ? data.error : "";
                const errorMessage = errorDetail
                    ? `API error: ${response.status} - ${errorDetail}`
                    : `API error: ${response.status}`;
                throw new Error(errorMessage);
            } else {
                data = await response.json();
            }

            const newText = data.rewritten_text;
            if (requestId !== activeRequestId || !isProcessing) {
                return;
            }
            if (!newText || !newText.trim()) {
                throw new Error("API returned empty rewritten text.");
            }

            const plainText = htmlToPlainText(newText);
            setResultContent(newText, { isHtml: true, allowActions: true });
            latestTextResult = { html: newText, text: plainText };

            if (skipPaste) {
                setStatus("Done.");
                setProcessingState(false);
                return;
            }

            const inserted = await insertResultIntoDocument();
            if (inserted && inputTextElement) {
                inputTextElement.value = "";
            }
            setStatus(inserted ? "Done." : "Insert failed.");
            setProcessingState(false);
        } catch (apiError) {
            if (requestId !== activeRequestId || !isProcessing) {
                return;
            }
            const message = formatRequestError(apiError, didTimeout, undefined, providerChoice);
            setResultContent(message, { isHtml: false, allowActions: false });
            if (message.startsWith("Request timed out")) {
                setStatus("Request timed out.");
            } else if (message.startsWith("Request canceled")) {
                setStatus("Canceled by user.");
            } else if (message.startsWith("Network error")) {
                setStatus("Network error.");
            } else {
                setStatus("Error during AI request.");
            }
            setProcessingState(false);
        }
    } catch (error) {
        setResultContent(`Error: ${error.message}`, { isHtml: false, allowActions: false });
        setStatus("Unexpected error.");
        setProcessingState(false);
    }
}

async function refreshModelOptions(provider) {
    const modelList = document.getElementById("modelList");
    const modelStatus = document.getElementById("modelStatus");
    if (!modelList || !modelStatus) {
        return;
    }

    modelList.innerHTML = "";
    modelStatus.textContent = "Loading available models...";

    try {
        const response = await fetch(`${OFFICEXADD_API_BASE_URL}/api/models?provider=${encodeURIComponent(provider)}`, {
            headers: buildApiHeaders()
        });
        if (!response.ok) {
            throw new Error(`Server returned ${response.status}`);
        }

        const data = await response.json();
        const models = Array.isArray(data.models) ? data.models : [];
        if (!models.length) {
            modelStatus.textContent = `No models returned for ${provider}.`;
            return;
        }

        models.forEach((modelId) => {
            const option = document.createElement("option");
            option.value = modelId;
            modelList.appendChild(option);
        });
        modelStatus.textContent = `Loaded ${models.length} models for ${provider}.`;
    } catch (error) {
        console.error("Failed to load models", error);
        modelStatus.textContent = `Unable to load models (${error.message}).`;
    }
}
