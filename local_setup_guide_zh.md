# 本機測試與安裝指南 (Local Setup Guide)

這份指南將協助您在本機執行後端伺服器，產生可 sideload 的 Office Add-in manifest，並透過 Cloudflare Tunnel 公開提供前端入口。

## 一鍵安裝與啟動 (Mac/Linux)

1. 確認 `backend/.env` 存在，若沒有可先複製 `backend/.env.example` 後填入 AI 相關金鑰。
2. 確認專案根目錄 `.env` 已填好：
   ```text
   OFFICEXADD_PUBLIC_ORIGIN=https://your-addin.example.com
   OFFICEXADD_API_TOKEN=replace-with-a-random-secret
   CLOUDFLARE_TUNNEL_TOKEN=replace-with-your-cloudflare-tunnel-token
   ```
3. 在專案根目錄執行：
   ```bash
   ./one_click.sh
   ```
4. 要停止背景服務時：
   ```bash
   ./stop.sh
   ```

## 1. 後端設定 (Backend Setup)

1. 進入 `backend` 資料夾：
   ```bash
   cd /path/to/officexadd/backend
   ```
2. 建立 `.env` 並填入後端 AI 設定：
   ```text
   OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxxxxxxxxxxxxx
   AI_BASE_URL=https://ollama.com/v1
   AI_API_KEY=ollama-xxxxxxxxxxxxxxxxxxxxxxxx
   MODEL_NAME=mistral-large-3:675b-cloud
   OLLAMA_WEB_SEARCH_API_KEY=ollama-web-search-key-here
   GOOGLE_API_KEY=your-google-api-key
   GOOGLE_IMAGE_MODEL=gemini-3.1-flash-image-preview
   GOOGLE_IMAGE_ASPECT_RATIO=1:1
   # GOOGLE_IMAGE_SIZE=1K
   ```
3. 安裝依賴套件：
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```
4. 啟動後端伺服器：
   ```bash
   python app.py
   ```

## 2. 產生前端設定與 Manifest

Office Add-in 的公開網址不再寫死在 repo 內，而是從根目錄 `.env` 產生。

```bash
./render_frontend_assets.sh
```

這會更新：
- `frontend/config.js`
- `frontend/manifest.xml`
- `frontend/manifest-powerpoint.xml`

若您更換 Cloudflare Tunnel 綁定的公開網址，請重新執行一次。

## 3. Docker + Cloudflare Tunnel 部署

1. 複製 `.env.example` 成 `.env`，並填入實際值。
2. 在 Cloudflare Zero Trust / Tunnel 後台，將這個服務的 public hostname 指向：
   ```text
   http://nginx:80
   ```
3. 啟動：
   ```bash
   docker compose up -d
   ```
4. 此版本不再對外開放舊的主機 port，也不再依賴固定網域寫死在 repo 內。

## 4. 如何匯入 Word / PowerPoint (Sideloading on Mac)

### 方法一：手動放入 manifest

1. 打開 Finder，按下 `Cmd + Shift + G`，依測試目標貼上：
   ```text
   /Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef
   ```
   或：
   ```text
   /Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   ```
2. 複製對應 manifest：
   - Word：`frontend/manifest.xml`
   - PowerPoint：`frontend/manifest-powerpoint.xml`
3. 完全關閉對應 Office App 再重新打開。
4. 到 **Insert** > **My Add-ins** > **Developer Add-ins** 找到 **OfficeXAdd**。

### 方法二：使用 Node.js 自動安裝

```bash
npx office-addin-debugging start frontend/manifest.xml
```
PowerPoint：
```bash
npx office-addin-debugging start frontend/manifest-powerpoint.xml
```

## 5. 測試流程

1. 確認後端 (`python app.py` 或 Docker backend) 正在執行。
2. 確認已執行 `./render_frontend_assets.sh`。
3. 在 Word 或 PowerPoint 中載入對應 manifest。
4. 在文件或投影片文字方塊中選取一段文字並執行改寫。
5. 若需要網路搜尋，勾選 **Use web search**。
6. 若要生圖，切換到 **Google 生圖 (Nano Banana)** 模式，輸入需求後按 **Generate Image & Insert**。
7. PowerPoint 目前為「選取文字」流程，不使用全文上下文模式。
