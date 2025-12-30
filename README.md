# OfficeXAdd - Word AI Rewrite Assistant

This is an Office.js add-in for Microsoft Word that provides AI-powered text rewriting capabilities using OpenAI's GPT-4o-mini model.

**Key Feature**: Eliminates the "Copy-Paste" workflow. Select text in Word, give an instruction, and the AI directly replaces the selection with the polished version.

## Project Structure

```
officexadd/
├── frontend/
│   ├── manifest.xml
│   ├── taskpane.html
│   ├── taskpane.js
│   └── package.json
├── backend/
│   ├── app.py
│   ├── requirements.txt
│   ├── .env (create this file)
│   └── Dockerfile
├── local_setup_guide_zh.md (Detailed Chinese Setup Guide)
└── README.md
```

## Prerequisites

- **Python 3.x** (for Backend)
- **Node.js** (Recommended for Frontend & Debugging)
- **OpenAI API key**
- **Microsoft Word** (Mac or Windows)

## Quick Start (Local Development)

For a detailed guide in Chinese, please see [local_setup_guide_zh.md](local_setup_guide_zh.md).

## One-click Install & Start (Mac/Linux)

1.  Ensure `backend/.env` exists. If not, copy `backend/.env.example` and fill in your keys.
2.  Run:
    ```bash
    ./one_click.sh
    ```
3.  To stop background servers later:
    ```bash
    ./stop.sh
    ```

### 1. Backend Setup

1.  Navigate to `backend/`.
2.  Create a `.env` file and add your API keys/configuration:
    ```
    OPENAI_API_KEY=sk-proj-your-key-here
    AI_BASE_URL=https://ollama.labelnine.app:5016/v1
    AI_API_KEY=ollama-your-key-here
    MODEL_NAME=mistral-large-3:675b-cloud
    OLLAMA_WEB_SEARCH_API_KEY=ollama-web-search-key-here
    ```
    * `AI_BASE_URL` and `AI_API_KEY` let the backend talk to OpenAI-compatible hosts such as Ollama; omit them if you only target `api.openai.com`.
    * `OLLAMA_WEB_SEARCH_API_KEY` is required if you enable web search while using the Ollama provider (if omitted, the backend falls back to `AI_API_KEY`).
    * If your environment cannot reach `api.ollama.com`, set `OLLAMA_WEB_SEARCH_URL` to a reachable web search endpoint.
3.  Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
4.  Start the server:
    ```bash
    python app.py
    ```

### 2. Frontend Setup

1.  Navigate to `frontend/`.
2.  Start a simple HTTP server (requires Node.js `http-server` or Python `http.server`):
    ```bash
    # Node.js (Recommended)
    npx http-server -p 3000 --cors
    
    # Python alternative
    python3 -m http.server 3000
    ```

### 3. Sideload to Word

**Option A: Automatic (Node.js required)**
From the project root:
```bash
npx office-addin-debugging start frontend/manifest.xml
```

**Option B: Manual (Mac)**
1.  Copy `frontend/manifest.xml` to `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`.
2.  Restart Word.
3.  Go to **Insert** > **My Add-ins** > **Developer Add-ins** > **OfficeXAdd**.

## Usage Guide

1.  **Open the Taskpane**: Click the **"Show AI Assistant"** button on the **Home** tab (or use `Ctrl+Alt+I`).
2.  **Select Text**: Highlight any text in your Word document that you want to rewrite.
3.  **Give Instructions (Optional)**: In the taskpane, type what you want the AI to do (e.g., "Make it more professional", "Fix grammar", "Translate to English").
4.  **Rewrite & Replace**: Click the button.
    - The AI will process your text.
    - The selected text in Word will be **automatically replaced** with the new version.
    - Formatting (bold, lists, etc.) is preserved/applied where possible.
    - Use the **AI Provider** dropdown to switch between the official OpenAI endpoint and an OpenAI-compatible host such as Ollama (which requires `AI_BASE_URL`/`AI_API_KEY` in `.env`).
    - The **Model (optional)** field now pulls a provider-specific list of models from the backend; change the provider to refresh the suggestion list or type any other name you need.
    - Enable **Use web search** to let the backend call the provider's web search tool (if supported by the chosen model/provider).

## Troubleshooting

- **500 Error**: Check the backend terminal for error messages. Ensure your OpenAI API key is valid and you are using a compatible `openai` library version (the code supports v1.x).
- **Add-in not showing**: Try the Manual Sideload method if the automatic command fails.
