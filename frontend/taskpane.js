let isProcessing = false;
let currentController = null;
let activeRequestId = 0;
const WEB_SEARCH_TIMEOUT_MS = 120000;
const DEFAULT_TIMEOUT_MS = 45000;

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
            setTimeout(() => {
                copyBtn.textContent = "Copy";
            }, 1200);
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
        document.getElementById("rewriteBtn").onclick = rewriteText;
        const copyBtn = document.getElementById("copyBtn");
        if (copyBtn) {
            copyBtn.onclick = copyResult;
        }
        const insertBtn = document.getElementById("insertBtn");
        if (insertBtn) {
            insertBtn.onclick = insertResultIntoDocument;
        }

        // Register selection change event handler
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange);
        const providerSelect = document.getElementById("providerSelect");
        if (providerSelect) {
            providerSelect.addEventListener("change", () => refreshModelOptions(providerSelect.value));
            refreshModelOptions(providerSelect.value);
        }
    }
});

async function onSelectionChange(eventArgs) {
    await Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("inputText").value = result.value;
        }
    });
}

async function rewriteText() {
    if (isProcessing) {
        cancelCurrentRequest();
        return;
    }

    const inputText = document.getElementById("inputText").value;
    const instructionText = document.getElementById("instructionText").value;
    const providerChoice = document.getElementById("providerSelect").value;
    const modelChoice = document.getElementById("modelInput").value.trim();
    const useWebSearch = document.getElementById("webSearchToggle").checked;
    const skipPaste = document.getElementById("skipPasteToggle").checked;
    const requestId = activeRequestId + 1;
    activeRequestId = requestId;
    currentController = new AbortController();

    // If no manual input, we'll try to get selection, but we need to handle the case where both are empty later

    setResultContent("Processing...", { isHtml: false, allowActions: false });
    setProcessingState(true);
    setStatus("Preparing request...");

    try {
        // Get selected text if available
        let textToRewrite = inputText;

        // We need to wrap the Office call in a promise to await it properly if we want to use the selection as input
        // However, for the "Replace" functionality, we primarily want to operate on the selection.

        await Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (result) => {
            if (requestId !== activeRequestId || !isProcessing) {
                return;
            }
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value.trim()) {
                textToRewrite = result.value;
                document.getElementById("inputText").value = textToRewrite; // Update UI
            }

            if (!textToRewrite.trim()) {
                setResultContent("Please select text in Word or enter text above.", { isHtml: false, allowActions: false });
                setProcessingState(false);
                setStatus("Idle");
                return;
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

                const timeoutMs = useWebSearch ? WEB_SEARCH_TIMEOUT_MS : DEFAULT_TIMEOUT_MS;
                const timeoutId = setTimeout(() => {
                    didTimeout = true;
                    if (currentController) {
                        currentController.abort();
                    }
                }, timeoutMs);

                let response;
                try {
                    response = await fetch("http://localhost:5000/rewrite", {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                        },
                        body: JSON.stringify(payload),
                        signal: currentController.signal,
                    });
                } finally {
                    clearTimeout(timeoutId);
                }

                if (!response.ok) {
                    throw new Error(`API error: ${response.status}`);
                }

                const data = await response.json();
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
                    if (requestId !== activeRequestId || !isProcessing) {
                        return;
                    }
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        setResultContent(`${newText}<br><span style="color:red">Error replacing text: ${asyncResult.error.message}</span>`, { isHtml: true, allowActions: true });
                        setStatus("Error replacing text.");
                    } else {
                        setResultContent(`${newText}<br><span style="color:green">Text replaced in Word!</span>`, { isHtml: true, allowActions: true });
                        document.getElementById("inputText").value = ""; // Clear input after success
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
        });
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
        const response = await fetch(`http://localhost:5000/models?provider=${encodeURIComponent(provider)}`);
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
