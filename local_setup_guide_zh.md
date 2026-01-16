# 本機測試與安裝指南 (Local Setup Guide)

這份指南將協助您在不使用 Docker 的情況下，於本機執行後端伺服器，並將 Office Add-in 安裝 (Sideload) 到 Word 中。

## 一鍵安裝與啟動 (Mac/Linux)

1.  確認 `backend/.env` 存在，若沒有可先複製 `backend/.env.example` 後填入金鑰。
2.  在專案根目錄執行：
    ```bash
    ./one_click.sh
    ```
3.  要停止背景服務時：
    ```bash
    ./stop.sh
    ```

## 1. 後端設定 (Backend Setup)

我們已經將設定改為使用 `.env` 檔案，方便您管理 API Key。

1.  **進入後端資料夾**：
    打開終端機 (Terminal)，進入 `backend` 資料夾：
    ```bash
    cd /Users/rich_imac/Downloads/officexadd/backend
    ```

2.  **設定 API Key 與自訂終端**：
    我們已經為您建立了一個 `.env` 檔案。請使用文字編輯器打開它，並填入您的金鑰與模型設定。
    ```text
    OPENAI_API_KEY=sk-proj-xxxxxxxxxxxxxxxxxxxxxxxx  <-- 替換成您的 Key
    AI_BASE_URL=https://ollama.labelnine.app:5016/v1
    AI_API_KEY=ollama-xxxxxxxxxxxxxxxxxxxxxxxx
    MODEL_NAME=mistral-large-3:675b-cloud
    OLLAMA_WEB_SEARCH_API_KEY=ollama-web-search-key-here
    ```
    * `AI_BASE_URL` 與 `AI_API_KEY` 可讓後端發送請求到 OpenAI 相容服務（例如 Ollama）。如果您只使用官方 OpenAI，這兩行可以留空或刪除。
    * 使用 Ollama 的 web search 時需要提供 `OLLAMA_WEB_SEARCH_API_KEY`（若未提供，會改用 `AI_API_KEY`）。
    * 若無法連到 `api.ollama.com`，請設定 `OLLAMA_WEB_SEARCH_URL` 指向可用的 web search 服務。

3.  **安裝依賴套件**：
    建議建立一個虛擬環境 (Virtual Environment) 來安裝套件：
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt
    ```

4.  **啟動後端伺服器**：
    ```bash
    python app.py
    ```
    看到 `Running on http://0.0.0.0:5010` 表示啟動成功。

---

## 2. 前端設定 (Frontend Setup)

由於 Office Add-in 需要透過 HTTPS 存取，我們需要一個簡單的方式來提供前端檔案。

### 選項 A：使用 Python (最簡單，不用安裝新東西)

您剛才指令打錯了，正確指令是 `http.server` (點號) 而不是 `http-server` (橫線)。

1.  **進入前端資料夾**：
    ```bash
    cd /Users/rich_imac/Downloads/officexadd/frontend
    ```

2.  **啟動伺服器**：
    ```bash
    python3 -m http.server 3010
    ```
    *(注意：是 `http.server`，中間是點號)*

### 選項 B：使用 Node.js (推薦，功能更強大)

如果您想要使用 `http-server` 或微軟官方的除錯工具，您需要安裝 Node.js。

#### 1. 安裝 Node.js
有兩種方式安裝：

*   **方法一：使用 Homebrew (推薦 Mac 使用者)**
    如果您有安裝 Homebrew，在終端機輸入：
    ```bash
    brew install node
    ```

*   **方法二：下載安裝檔**
    前往 [Node.js 官網](https://nodejs.org/) 下載 "LTS" 版本並安裝。

安裝完成後，檢查是否成功：
```bash
node -v
npm -v
```

#### 2. 安裝並使用 http-server
安裝 Node.js 後，您可以安裝全域的 `http-server` 工具：

```bash
npm install --global http-server
```

然後就可以使用您原本想用的指令了：
```bash
http-server -p 3010 --cors
```
*(加上 `--cors` 可以避免一些跨域問題)*

---

**為了最順暢的體驗，建議使用 Node.js 的 `office-addin-debugging`**:
如果您安裝好了 Node.js，直接執行這個指令最方便，它會自動處理憑證並開啟 Word：
```bash
npx office-addin-debugging start manifest.xml
```

---

## 3. 如何匯入 Word (Sideloading on Mac)

這是您最關心的部分：如何把這個 Add-in 放進 Word 裡測試。

### 方法一：直接將 manifest.xml 放入特定資料夾 (最推薦 Mac 使用)

1.  **找到 Weff 檔案夾**：
    打開 Finder，按下 `Cmd + Shift + G`，貼上以下路徑：
    ```
    /Users/rich_imac/Library/Containers/com.microsoft.Word/Data/Documents/wef
    ```
    *(如果 `wef` 資料夾不存在，請手動建立它)*

2.  **複製 manifest.xml**：
    將 `/Users/rich_imac/Downloads/officexadd/frontend/manifest.xml` 這個檔案複製到上面的 `wef` 資料夾中。

3.  **重啟 Word**：
    完全關閉 Word 再重新打開。

4.  **找到 Add-in**：
    - 開啟一個空白文件。
    - 點選上方選單的 **「插入 (Insert)」** > **「我的增益集 (My Add-ins)」**。
    - 點選上方的小箭頭或是 **「開發人員 (Developer)」** 分頁 (如果有出現)。
    - 您應該會看到 **"OfficeXAdd"** 出現在列表中。點選它即可開啟側邊欄。

### 方法二：使用 Node.js 自動安裝 (如果您有 Node.js)

如果您電腦有安裝 Node.js，這是最快的方法：

1.  在 `frontend` 資料夾中執行：
    ```bash
    npx office-addin-debugging start manifest.xml
    ```
2.  這會自動啟動 Word 並載入您的 Add-in。

---

## 4. 測試流程

1.  確認後端 (`python app.py`) 正在執行。
2.  確認前端伺服器 (Port 3010) 正在執行。
3.  在 Word 中開啟 "OfficeXAdd" 側邊欄。
4.  在文件中輸入一段文字，例如：「這是一個測試文句，請幫我改寫。」
5.  選取這段文字。
6.  在側邊欄的 Instructions 輸入：「變得更專業一點」。
7.  按下 **"Rewrite & Replace"**。
8.  文字應該會自動變更！
9.  在側邊欄的 **AI Provider** 下拉選單中選擇官方 OpenAI 或 Ollama (選 Ollama 時請先在 `.env` 填好 `AI_BASE_URL` 與 `AI_API_KEY`)，需要不同模型時可在「Model (optional)」欄位輸入名稱。
10. 「Model (optional)」欄位會從後端讀取並顯示該提供者目前可用的模型，切換提供者會重新載入選單內容，但仍可自行輸入任意模型名稱。
11. 若需要網路搜尋，勾選 **Use web search**，後端會呼叫對應提供者的網頁搜尋工具（前提是該模型/服務支援此能力）。
