let isProcessing = false;
let currentController = null;
let activeRequestId = 0;
let activeTimeouts = new Set();
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

function setStatus(message) {
    const statusBar = document.getElementById("statusBar");
    if (statusBar) {
        statusBar.textContent = message;
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
        button.textContent = "Rewrite & Replace";
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
    const sizeRow = document.getElementById("contextSizeRow");
    const unitSpan = document.getElementById("contextUnit");
    const help = document.getElementById("contextHelp");
    if (!modeSelect || !sizeRow || !unitSpan || !help) {
        return;
    }
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
        help.textContent = "不使用 Word 上下文。";
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

function getSelectedTextFallback() {
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

async function getWordSnapshot(options = {}) {
    const includeDocumentText = Boolean(options.includeDocumentText);
    if (typeof Word === "undefined") {
        const fallbackText = await getSelectedTextFallback();
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
        const fallbackText = await getSelectedTextFallback();
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

function formatRequestError(error, didTimeout) {
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

function insertResultIntoDocument() {
    const resultContent = document.getElementById("resultContent");
    if (!resultContent) {
        return;
    }
    const html = resultContent.innerHTML.trim();
    const text = resultContent.textContent.trim();
    if (!text) {
        return;
    }
    setStatus("Inserting into Word...");
    Office.context.document.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setStatus(`Insert failed: ${asyncResult.error.message}`);
        } else {
            setStatus("Inserted into Word.");
        }
    });
}

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
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
            updateContextControls();
        }
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

    const inputTextElement = document.getElementById("inputText");
    const inputText = inputTextElement ? inputTextElement.value : "";
    const instructionText = document.getElementById("instructionText").value;
    const providerChoice = document.getElementById("providerSelect").value;
    const modelChoice = document.getElementById("modelInput").value.trim();
    const useWebSearch = document.getElementById("webSearchToggle").checked;
    const skipPaste = document.getElementById("skipPasteToggle").checked;
    const contextMode = document.getElementById("contextMode").value;
    const contextSize = parseContextSize();
    const requestId = activeRequestId + 1;
    activeRequestId = requestId;
    currentController = new AbortController();

    // If no manual input, we'll try to get selection, but we need to handle the case where both are empty later

    setResultContent("Processing...", { isHtml: false, allowActions: false });
    setProcessingState(true);
    setStatus("Preparing request...");

    try {
        setStatus("Reading selection...");
        const needsDocument = contextMode !== CONTEXT_MODE_NONE;
        const snapshot = await getWordSnapshot({ includeDocumentText: needsDocument });
        let textToRewrite = inputText;
        if (snapshot.selectionText && snapshot.selectionText.trim()) {
            textToRewrite = snapshot.selectionText;
            if (inputTextElement) {
                inputTextElement.value = textToRewrite;
            }
        }

        if (!textToRewrite.trim() && !instructionText.trim()) {
            setResultContent("Please enter instructions or select text in Word.", { isHtml: false, allowActions: false });
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
                response = await fetch("http://localhost:5001/rewrite", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                    },
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
            setResultContent(newText, { isHtml: true, allowActions: true });

            // Replace selection in Word
            if (skipPaste) {
                setStatus("Done.");
                setProcessingState(false);
                return;
            }
            setStatus("Replacing selection in Word...");
            Office.context.document.setSelectedDataAsync(newText, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
                if (requestId !== activeRequestId) {
                    setProcessingState(false);
                    return;
                }
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    setResultContent(`${newText}<br><span style="color:red">Error replacing text: ${asyncResult.error.message}</span>`, { isHtml: true, allowActions: true });
                    setStatus("Error replacing text.");
                } else {
                    setResultContent(`${newText}<br><span style="color:green">Text replaced in Word!</span>`, { isHtml: true, allowActions: true });
                    if (inputTextElement) {
                        inputTextElement.value = "";
                    }
                    setStatus("Done.");
                }
                setProcessingState(false);
            });

        } catch (apiError) {
            if (requestId !== activeRequestId || !isProcessing) {
                return;
            }
            const message = formatRequestError(apiError, didTimeout);
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
        const response = await fetch(`http://localhost:5001/models?provider=${encodeURIComponent(provider)}`);
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
