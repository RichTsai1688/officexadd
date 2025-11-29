Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("rewriteBtn").onclick = rewriteText;

        // Register selection change event handler
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChange);
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
    const inputText = document.getElementById("inputText").value;
    const instructionText = document.getElementById("instructionText").value;
    const resultDiv = document.getElementById("result");
    const button = document.getElementById("rewriteBtn");

    // If no manual input, we'll try to get selection, but we need to handle the case where both are empty later

    button.disabled = true;
    button.textContent = "Rewriting...";
    resultDiv.innerHTML = "Processing...";

    try {
        // Get selected text if available
        let textToRewrite = inputText;

        // We need to wrap the Office call in a promise to await it properly if we want to use the selection as input
        // However, for the "Replace" functionality, we primarily want to operate on the selection.

        await Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value.trim()) {
                textToRewrite = result.value;
                document.getElementById("inputText").value = textToRewrite; // Update UI
            }

            if (!textToRewrite.trim()) {
                resultDiv.innerHTML = "Please select text in Word or enter text above.";
                button.disabled = false;
                button.textContent = "Rewrite & Replace";
                return;
            }

            // Call backend API
            try {
                const response = await fetch("http://localhost:5000/rewrite", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                    },
                    body: JSON.stringify({
                        text: textToRewrite,
                        instruction: instructionText
                    }),
                });

                if (!response.ok) {
                    throw new Error(`API error: ${response.status}`);
                }

                const data = await response.json();
                const newText = data.rewritten_text;

                // Display result
                resultDiv.innerHTML = `<strong>Rewritten Text:</strong><br>${newText}`;

                // Replace selection in Word
                Office.context.document.setSelectedDataAsync(newText, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        resultDiv.innerHTML += `<br><span style="color:red">Error replacing text: ${asyncResult.error.message}</span>`;
                    } else {
                        resultDiv.innerHTML += `<br><span style="color:green">Text replaced in Word!</span>`;
                        document.getElementById("inputText").value = ""; // Clear input after success
                    }
                });

            } catch (apiError) {
                resultDiv.innerHTML = `Error: ${apiError.message}`;
            } finally {
                button.disabled = false;
                button.textContent = "Rewrite & Replace";
            }
        });
    } catch (error) {
        resultDiv.innerHTML = `Error: ${error.message}`;
        button.disabled = false;
        button.textContent = "Rewrite & Replace";
    }
}
