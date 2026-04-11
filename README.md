# OfficeXAdd - Word / PowerPoint AI Assistant

This is an Office.js add-in for Microsoft Word and Microsoft PowerPoint that provides AI-powered text rewriting and image generation capabilities.

**Key Feature**: Eliminates the "Copy-Paste" workflow. Select text in Word/PowerPoint, give an instruction, and the AI directly replaces the selection with the polished version.

## Project Structure

```text
officexadd/
├── frontend/
│   ├── manifest.template.xml
│   ├── manifest-powerpoint.template.xml
│   ├── manifest.xml              # Word manifest, generated from OFFICEXADD_PUBLIC_ORIGIN
│   ├── manifest-powerpoint.xml   # PowerPoint manifest, generated from OFFICEXADD_PUBLIC_ORIGIN
│   ├── taskpane.html
│   ├── taskpane.js
│   └── config.js            # generated from OFFICEXADD_PUBLIC_ORIGIN / OFFICEXADD_API_TOKEN
├── backend/
│   ├── app.py
│   ├── requirements.txt
│   ├── .env (create this file for local backend runs)
│   └── Dockerfile
├── .env.example             # docker compose / Cloudflare Tunnel settings
├── render_frontend_assets.sh
└── README.md
```

## Prerequisites

- **Python 3.x** (for Backend)
- **Node.js** (Recommended for Frontend & Debugging)
- **OpenAI API key** (for text rewrite mode)
- **Google Gemini API key** (for Nano Banana image generation mode)
- **Microsoft Word or Microsoft PowerPoint** (Mac or Windows)
- **Cloudflare Tunnel** (for public access to the add-in in deployment)

## Quick Start

### Docker + Cloudflare Tunnel

1. Copy `.env.example` to `.env`.
2. Fill in these required values in `.env`:
   - `OFFICEXADD_PUBLIC_ORIGIN=https://your-addin-hostname.example.com`
   - `OFFICEXADD_API_TOKEN=your-random-secret`
   - `CLOUDFLARE_TUNNEL_TOKEN=your-cloudflare-tunnel-token`
3. Ensure your Cloudflare Tunnel public hostname points to `http://nginx:80` for this service.
4. Start the stack:
   ```bash
   docker compose up -d
   ```
5. Regenerate local frontend files any time the public hostname changes:
   ```bash
   ./render_frontend_assets.sh
   ```

### Local Development

For a detailed guide in Chinese, please see [local_setup_guide_zh.md](local_setup_guide_zh.md).

1. Ensure `backend/.env` exists. If not, copy `backend/.env.example` and fill in your backend AI settings.
2. Ensure root `.env` exists with `OFFICEXADD_PUBLIC_ORIGIN` and `OFFICEXADD_API_TOKEN` so the manifest/config can be generated.
3. Run:
   ```bash
   ./one_click.sh
   ```
4. To stop background servers later:
   ```bash
   ./stop.sh
   ```

## Backend Setup

1. Navigate to `backend/`.
2. Create a `.env` file and add your API keys/configuration:
   ```text
   OPENAI_API_KEY=sk-proj-your-key-here
   AI_BASE_URL=https://ollama.com/v1
   AI_API_KEY=ollama-your-key-here
   MODEL_NAME=gpt-oss
   OLLAMA_WEB_SEARCH_API_KEY=ollama-web-search-key-here
   GOOGLE_API_KEY=your-google-api-key
   GOOGLE_IMAGE_MODEL=gemini-3.1-flash-image-preview
   GOOGLE_IMAGE_ASPECT_RATIO=1:1
   # GOOGLE_IMAGE_SIZE=1K
   # GOOGLE_API_BASE_URL=https://generativelanguage.googleapis.com/v1beta
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Start the server:
   ```bash
   python app.py
   ```

## Frontend / Manifest Generation

The taskpane config and Office manifest are generated from the root `.env` file.

```bash
./render_frontend_assets.sh
```

This updates:
- `frontend/config.js`
- `frontend/manifest.xml`
- `frontend/manifest-powerpoint.xml`

## Sideload to Word / PowerPoint

**Option A: Automatic (Node.js required)**
From the project root:
```bash
npx office-addin-debugging start frontend/manifest.xml
```
For PowerPoint:
```bash
npx office-addin-debugging start frontend/manifest-powerpoint.xml
```

**Option B: Manual (Mac)**
1. Copy the matching manifest:
   - Word: `frontend/manifest.xml` -> `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   - PowerPoint: `frontend/manifest-powerpoint.xml` -> `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`
2. Restart the Office app.
3. Go to **Insert** > **My Add-ins** > **Developer Add-ins** > **OfficeXAdd**.

**Option C: Remote Mac install script**
Run this on the target Mac:
```bash
bash install_word_addin_mac.sh
```

## Usage Guide

1. **Open the Taskpane**: Click the **"Show AI Assistant"** button on the **Home** tab (or use `Ctrl+Alt+I`).
2. **Choose Mode**:
   - `文字改寫`: rewrites selected text with OpenAI/Ollama.
   - `Google 生圖 (Nano Banana)`: generates image with Gemini and inserts into Word/PowerPoint.
3. **Provide Prompt/Instruction**:
   - Text mode: select Word/PowerPoint text and give rewrite instructions.
   - Image mode: enter image requirements (prompt) and optional style instructions.
4. **Run**:
   - Text mode button: `Rewrite & Replace`
   - Image mode button: `Generate Image & Insert`
5. **PowerPoint behavior**:
   - Uses selected slide text as input/output target.
   - Document context modes are disabled in PowerPoint (selection-only flow).

## Troubleshooting

- **Manifest points to the wrong hostname**: Update `OFFICEXADD_PUBLIC_ORIGIN` in root `.env`, then run `./render_frontend_assets.sh` again.
- **Cloudflare Tunnel is connected but the app is unreachable**: Verify the tunnel's public hostname points to `http://nginx:80`.
- **500 Error**: Check the backend logs and API credentials.
