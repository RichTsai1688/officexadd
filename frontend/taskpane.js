let isProcessing = false;
let currentController = null;
let activeRequestId = 0;
let activeTimeouts = new Set();
let lastRewriteResult = null;
const OFFICEXADD_CONFIG = window.__OFFICEXADD_CONFIG__ || {};
const OFFICEXADD_API_BASE_URL = (OFFICEXADD_CONFIG.apiBaseUrl || window.location.origin || "https://fcu.labelnine.app:2053").replace(/\/$/, "");
const OFFICEXADD_API_TOKEN = OFFICEXADD_CONFIG.apiToken || "";
const WEB_SEARCH_TIMEOUT_MS = 120000;
const DEFAULT_TIMEOUT_MS = 45000;
const MAX_CONTEXT_CHARS = 12000;
const CONTEXT_MARKER_START = "[[EDIT_START]]";
const CONTEXT_MARKER_END = "[[EDIT_END]]";
const CONTEXT_MARKER_CURSOR = "[[CURSOR]]";
const CONTEXT_MODE_NONE = "none";
const CONTEXT_MODE_FULL = "full";
const CONTEXT_MODE_CHARS = "chars";
const CONTEXT_MODE_PAGES = "pages";
const APPROX_PAGE_CHARS = 1500;
let currentHost = "office";

const HOST_CONFIG = {
    office: {
        appTitle: "AI Text Rewriter",
        subtitle: "Rewrite selected text with the tone you choose.",
        inputLabel: "引入內容",
        inputPlaceholder: "Enter text to rewrite...",
        rewriteButton: "Rewrite & Replace",
        resultTitle: "Rewritten Text",
        insertButton: "填入",
        skipPasteHelp: "勾選後只顯示結果，不會自動覆蓋目前選取內容。",
        emptyState: "Please enter instructions or select text in Office.",
        selectionLoading: "Reading selection...",
        selectionMissing: "Please enter instructions or select text in Office.",
        generating: "Generating from instruction...",
        inserting: "Inserting into Office...",
        inserted: "Inserted into Office.",
        replacing: "Replacing selection in Office...",
        replaced: "Content replaced in Office!",
        replaceError: "Error replacing content",
    },
    word: {
        appTitle: "AI Text Rewriter for Word",
        subtitle: "Rewrite selected Word text with the tone you choose.",
        inputLabel: "引入文章",
        inputPlaceholder: "Enter text from Word...",
        rewriteButton: "Rewrite & Replace",
        resultTitle: "Rewritten Text",
        insertButton: "填入",
        skipPasteHelp: "勾選後只顯示結果，不會自動取代 Word 內選取文字。",
        emptyState: "Please enter instructions or select text in Word.",
        selectionLoading: "Reading selection...",
        selectionMissing: "Please enter instructions or select text in Word.",
        generating: "Generating from instruction...",
        inserting: "Inserting into Word...",
        inserted: "Inserted into Word.",
        replacing: "Replacing selection in Word...",
        replaced: "Text replaced in Word!",
        replaceError: "Error replacing text",
    },
    powerpoint: {
        appTitle: "AI Text Rewriter for PowerPoint",
        subtitle: "Rewrite selected slide text and send the polished copy back into PowerPoint.",
        inputLabel: "引入投影片文字",
        inputPlaceholder: "Enter text from PowerPoint...",
        rewriteButton: "Rewrite & Insert",
        resultTitle: "Rewritten Slide Text",
        insertButton: "插入投影片",
        skipPasteHelp: "勾選後只顯示結果，不會自動覆蓋目前投影片選取的文字方塊內容。",
        emptyState: "Please enter instructions or select text in PowerPoint.",
        selectionLoading: "Reading selected slide text...",
        selectionMissing: "Please enter instructions or select text in PowerPoint.",
        generating: "Generating slide text...",
        inserting: "Inserting into PowerPoint...",
        inserted: "Inserted into PowerPoint.",
        replacing: "Replacing selected slide text...",
        replaced: "Text replaced in PowerPoint!",
        replaceError: "Error replacing slide text",
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
    const element = document.getElementById(id);
    if (element) {
        element.textContent = value;
    }
}

function setHostSpecificText() {
    const hostConfig = getHostConfig();
    setElementText("appTitle", hostConfig.appTitle);
    setElementText("appSubtitle", hostConfig.subtitle);
    setElementText("inputLabel", hostConfig.inputLabel);
    setElementText("resultTitle", hostConfig.resultTitle);
    setElementText("skipPasteHelp", hostConfig.skipPasteHelp);

    const inputText = document.getElementById("inputText");
    if (inputText) {
        inputText.placeholder = hostConfig.inputPlaceholder;
    }

    const rewriteBtn = document.getElementById("rewriteBtn");
    if (rewriteBtn && !isProcessing) {
        rewriteBtn.textContent = hostConfig.rewriteButton;
    }

    const insertBtn = document.getElementById("insertBtn");
    if (insertBtn) {
        insertBtn.textContent = hostConfig.insertButton;
    }
}

function setResultContent(content, options = {}) {
    const { isHtml = true, allowActions = true } = options;
    const resultContent = document.getElementById("resultContent");
    const copyBtn = document.getElementById("copyBtn");
    const insertBtn = document.getElementById("insertBtn");
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

function htmlToPlainText(html) {
    const container = document.createElement("div");
    container.innerHTML = html;
    return container.textContent.trim();
}

function setProcessingState(active) {
    isProcessing = active;
    const button = document.getElementById("rewriteBtn");
    const hostConfig = getHostConfig();
    if (!button) {
        return;
    }
    if (active) {
        button.textContent = "Stop";
        button.classList.add("is-stop");
    } else {
        button.textContent = hostConfig.rewriteButton;
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
        help.textContent = "PowerPoint 目前只會讀取選取的文字，不會抓整份投影片上下文。";
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

function formatRequestError(error, didTimeout, provider) {
    if (didTimeout) {
        return "Request timed out. Try again or disable web search.";
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
            return "Authorization failed for Ollama. Check AI_BASE_URL / AI_API_KEY on the server, or switch the provider to OpenAI.";
        }
        return "Authorization failed for the selected provider. Check the server-side API credentials and try again.";
    }
    return `Error: ${message}`;
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

function replaceSelectedContent(content, callbacks = {}) {
    const { onSuccess, onError } = callbacks;
    const payload = isWordHost() ? content.html : content.text;
    const coercionType = isWordHost() ? Office.CoercionType.Html : Office.CoercionType.Text;
    Office.context.document.setSelectedDataAsync(payload, { coercionType }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            if (onError) {
                onError(asyncResult.error);
            }
            return;
        }
        if (onSuccess) {
            onSuccess();
        }
    });
}

function getResultPayload() {
    if (lastRewriteResult && lastRewriteResult.text) {
        return lastRewriteResult;
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

function insertResultIntoDocument() {
    const hostConfig = getHostConfig();
    const payload = getResultPayload();
    if (!payload) {
        return;
    }
    setStatus(hostConfig.inserting);
    replaceSelectedContent(payload, {
        onSuccess: () => setStatus(hostConfig.inserted),
        onError: (error) => setStatus(`Insert failed: ${error.message}`),
    });
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        currentHost = "word";
    } else if (info.host === Office.HostType.PowerPoint) {
        currentHost = "powerpoint";
    }
    setHostSpecificText();

    if (info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint) {
        // 使用 cleanupResources 而不是 cancelCurrentRequest 確保完全清理
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
            contextMode.addEventListener("change", updateContextControls);
        }
        updateContextControls();

        const providerSelect = document.getElementById("providerSelect");
        if (providerSelect) {
            providerSelect.addEventListener("change", () => refreshModelOptions(providerSelect.value));
            refreshModelOptions(providerSelect.value);
        }
    }
});

async function rewriteText() {
    if (isProcessing) {
        cancelCurrentRequest();
        return;
    }

    const hostConfig = getHostConfig();
    const inputTextElement = document.getElementById("inputText");
    const inputText = inputTextElement ? inputTextElement.value : "";
    const instructionText = document.getElementById("instructionText").value;
    const providerChoice = document.getElementById("providerSelect").value;
    const modelChoice = document.getElementById("modelInput").value.trim();
    const useWebSearch = document.getElementById("webSearchToggle").checked;
    const skipPaste = document.getElementById("skipPasteToggle").checked;
    const requestedContextMode = document.getElementById("contextMode").value;
    const contextMode = supportsDocumentContext() ? requestedContextMode : CONTEXT_MODE_NONE;
    const contextSize = parseContextSize();
    const requestId = activeRequestId + 1;
    activeRequestId = requestId;
    currentController = new AbortController();

    // If no manual input, we'll try to get selection, but we need to handle the case where both are empty later

    lastRewriteResult = null;
    setResultContent("Processing...", { isHtml: false, allowActions: false });
    setProcessingState(true);
    setStatus("Preparing request...");

    try {
        setStatus(hostConfig.selectionLoading);
        const needsDocument = contextMode !== CONTEXT_MODE_NONE;
        const snapshot = await getDocumentSnapshot({ includeDocumentText: needsDocument });
        let textToRewrite = inputText;
        if (snapshot.selectionText && snapshot.selectionText.trim()) {
            textToRewrite = snapshot.selectionText;
            if (inputTextElement) {
                inputTextElement.value = textToRewrite;
            }
        }

        if (!textToRewrite.trim() && !instructionText.trim()) {
            setResultContent(hostConfig.selectionMissing, { isHtml: false, allowActions: false });
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
            setStatus(hostConfig.generating);
        }

        // Call backend API
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

            // Display result
            lastRewriteResult = { html: newText, text: htmlToPlainText(newText) };
            setResultContent(newText, { isHtml: true, allowActions: true });

            // Replace selection in the current Office host
            if (skipPaste) {
                setStatus("Done.");
                setProcessingState(false);
                return;
            }
            setStatus(hostConfig.replacing);
            replaceSelectedContent({ html: newText, text: htmlToPlainText(newText) }, {
                onSuccess: () => {
                    if (requestId !== activeRequestId) {
                        setProcessingState(false);
                        return;
                    }
                    setResultContent(`${newText}<br><span style="color:green">${hostConfig.replaced}</span>`, { isHtml: true, allowActions: true });
                    if (inputTextElement) {
                        inputTextElement.value = "";
                    }
                    setStatus("Done.");
                    setProcessingState(false);
                },
                onError: (error) => {
                    if (requestId !== activeRequestId) {
                        setProcessingState(false);
                        return;
                    }
                    setResultContent(`${newText}<br><span style="color:red">${hostConfig.replaceError}: ${error.message}</span>`, { isHtml: true, allowActions: true });
                    setStatus(hostConfig.replaceError);
                    setProcessingState(false);
                },
            });

        } catch (apiError) {
            if (requestId !== activeRequestId || !isProcessing) {
                return;
            }
            const message = formatRequestError(apiError, didTimeout, providerChoice);
            lastRewriteResult = null;
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
        lastRewriteResult = null;
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
