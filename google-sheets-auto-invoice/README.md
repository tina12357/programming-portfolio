# 📑 自動化報支憑證生成系統 (GAS)

## 💡系統概述
本系統透過 Google Apps Script (GAS) 實作，旨在將 Google Sheets 中的報支資料，自動填入指定的 Google Docs 模板中，並合併生成一份可供列印的 PDF/Doc 檔案。

## ✨ 核心特色
- **選取即產出**：支援多選儲存格，批次生成憑證。
- **智慧模板**：根據金額（>5,000元）自動判定並切換「大額/一般」模板。
- **自動轉換**：內建數字轉中文大寫金額、日期格式化功能。
- **排版優化**：每兩筆自動插入裁切線，減少紙張浪費。
- **安全機制**：產生連結 3 分鐘後自動清空，保護個人隱私。

## 🚀 快速開始
1. 在 Google Sheets 建立資料表（欄位順序：日期(C)、主旨(D)、專案(E)、類別(F)...）。
2. 開啟 **延伸功能 > Apps Script**。
3. 貼上 `Code.gs` 中的代碼。
4. 在代碼中填入你的 `TEMPLATE_ID` 與 `PRINT_DOC_ID`。
5. 重新整理試算表，點擊上方選單 `🛠️ 權限修復` 進行初次授權。

## 🛠️ 技術細節
- **語言**: Google Apps Script (JavaScript)
- **主要 API**: `SpreadsheetApp`, `DocumentApp`, `ScriptApp` (Triggers)
- **邏輯設計**: 使用 `getActiveRangeList()` 抓取選取區塊，透過 `replaceText` 進行動態標籤替換。

## 🏗️ 系統架構與資料流使用者操作
  1.選取試算表列 -> 點擊自定義選單「🛠️ 權限修復」。
  2.資料讀取：程式抓取所選列的第 3 欄至第 17 欄資料。
  3.邏輯判斷：根據 Amount 與 Type 決定使用哪一個 Template ID。
  4.內容填充：利用 replaceText 將模板中的 {{Tag}} 替換為實際數值。
  5.輸出結果：將內容寫入 PRINT_DOC_ID 並在試算表 C1 格生成超連結。
## 🛠️ 維護者準備工作 (Setup)若要部署到新的試算表，請修改程式碼中的 // 設定區：
變數名稱|說明|TEMPLATE_ID_SMALL|小額或一般報支的 Google Doc 模板 ID。|TEMPLATE_ID_LARGE|大額報支專用的 Google Doc 模板 ID。|
PRINT_DOC_I|D最終生成的列印用暫存檔 ID（程式會每次清空它）。|EXPIRE_MINUTES|連結消失的時間（預設 3 分鐘）。|

## 📝 寫作要注意的細節 (Tips)
在寫這份文件時，我特別注意了以下幾點，這也是你以後寫技術文件可以參考的：

1.列出「副作用」 (Side Effects)：
- 例如：程式會清空 PRINT_DOC_ID 的所有內容。這很重要，要提醒使用者不要把重要資料存在那個檔案裡。

2.解釋「特殊邏輯」：
  - 為什麼要判斷 data.amount > 5000 && data.type !== "代墊款"？這是業務邏輯的精髓，一定要寫出來。

3.錯誤處理說明：
- 程式碼中有 try...catch，文件應說明當 C1 出現紅字「❌ 錯誤」時，通常是 ID 填錯或權限未開。
