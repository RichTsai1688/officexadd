# Word 無法關閉問題修復說明

## 問題描述
使用 Office Add-in 後，Word 無法正常關閉。

## 根本原因
1. **未清理的事件監聽器**：`unload` 和 `beforeunload` 事件監聽器中的操作阻止了 Word 關閉
2. **未清除的計時器**：`setTimeout` 沒有被追蹤和清理，導致 JavaScript 引擎持續運行
3. **未終止的 API 請求**：AbortController 沒有在所有情況下正確清理
4. **異步操作未完成**：Office.js 的異步操作可能在關閉時仍在執行

## 已修復的問題

### 1. 添加資源追蹤機制
```javascript
let activeTimeouts = new Set();
```
- 追蹤所有活動的 timeout
- 確保在卸載時可以清理所有計時器

### 2. 實現完整的資源清理函數
```javascript
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
```

### 3. 修改事件監聽器
```javascript
// 修改前
window.addEventListener("unload", cancelCurrentRequest);
window.addEventListener("beforeunload", cancelCurrentRequest);

// 修改後
window.addEventListener("unload", cleanupResources);
window.addEventListener("beforeunload", cleanupResources);
```

### 4. 追蹤所有 setTimeout
- Copy 按鈕的 timeout
- API 請求的 timeout
- 所有 timeout 都會被追蹤並在清理時移除

### 5. 改進異步操作處理
- 在 Office.js 異步回調中檢查 `requestId` 而不是 `isProcessing`
- 確保即使在操作完成後也能正確設置狀態

### 6. 添加錯誤處理
- 在 `context.sync()` 中添加 catch 處理
- 防止未捕獲的 Promise 錯誤

## 使用建議

### 1. 正常關閉 Word
- 在關閉 Word 前，確保沒有正在進行的 AI 請求
- 如果有正在進行的請求，點擊 "Stop" 按鈕先停止

### 2. 強制關閉（如果需要）
如果 Word 仍然無法關閉：
1. 打開任務管理器（macOS: Activity Monitor）
2. 找到 Microsoft Word 進程
3. 強制結束進程

### 3. 防止問題再次發生
- 避免在有大量未完成操作時關閉 Word
- 定期檢查是否有 "Processing..." 狀態
- 使用 "Skip Auto-Paste" 選項可以減少異步操作

## 技術細節

### 資源生命週期
```
使用者操作 → 創建資源（timeout, controller） → 追蹤資源 → 使用資源 → 清理資源
```

### 清理觸發時機
1. 使用者點擊 Stop 按鈕
2. 頁面卸載 (unload)
3. 瀏覽器關閉前 (beforeunload)
4. API 請求完成或失敗

### Word.run 的正確使用
```javascript
try {
    return await Word.run(async (context) => {
        // 操作
        await context.sync().catch(err => {
            console.warn("Context sync error:", err);
            throw err;
        });
        // 返回結果
    });
} catch (error) {
    // 錯誤處理
}
```

## 測試步驟

1. **正常流程測試**
   - 開啟 Word 和 Add-in
   - 執行一次完整的重寫操作
   - 關閉 Word（應該可以正常關閉）

2. **中斷流程測試**
   - 開啟 Word 和 Add-in
   - 開始一個重寫操作
   - 點擊 Stop 按鈕
   - 關閉 Word（應該可以正常關閉）

3. **多次操作測試**
   - 執行多次重寫操作
   - 有些成功，有些取消
   - 關閉 Word（應該可以正常關閉）

4. **Web Search 測試**
   - 啟用 Web Search
   - 執行長時間操作
   - 等待完成或取消
   - 關閉 Word（應該可以正常關閉）

## 後續建議

1. **添加日誌記錄**
   - 記錄資源創建和清理
   - 方便追蹤潛在問題

2. **添加超時保護**
   - 為所有異步操作設置最大時間限制
   - 自動清理超時的操作

3. **監控內存使用**
   - 定期檢查是否有內存洩漏
   - 使用瀏覽器開發工具監控

4. **使用者提示**
   - 當有未完成操作時，提示使用者
   - 防止意外關閉

## 相關資源

- [Office.js API 文件](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [AbortController MDN](https://developer.mozilla.org/en-US/docs/Web/API/AbortController)
- [事件監聽器最佳實踐](https://developer.mozilla.org/en-US/docs/Web/API/EventTarget/addEventListener)
