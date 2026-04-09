# 學生任務自動化派發與一鍵回報系統
An automated system built with Google Apps Script (GAS) that dispatches personalized tasks to students via Email. Students can report progress through a single click in their email or via a customized web dashboard—no Google Forms required.

這是一套基於 Google Apps Script (GAS) 開發的自動化系統，能透過 Email 將個人化任務派發給學生。學生只需在信件中點擊或透過專屬網頁儀表板即可回報進度，完全不需要填寫 Google 表單。
## ✨ Key Features | 系統亮點
- Zero-Entry Reporting: Students click a button in the email, and the Sheet status updates instantly.

  - 零摩擦回報：學生在 Email 點擊按鈕，試算表狀態立即更新。

Smart Dispatch: Automatically skips students who have already received the task to avoid spam.

智慧派發：自動跳過已派發過的任務，避免重複寄信。

Personal Dashboard: A mobile-friendly web UI for students to view all pending tasks and request a resend of their task list.

個人儀表板：提供適合手機閱讀的網頁介面，供學生查看所有待辦事項並可自主補發清單信。

Anti-Spam Logic: Only sends emails when new tasks are added (limited to the top 5 tasks in the email for better readability).

防干擾邏輯：僅在有新任務時發信，且 Email 僅列出前 5 項以維持閱讀體驗。

## 🛠️ Setup Instructions | 安裝步驟
1. Spreadsheet Preparation | 準備試算表
Create a Google Sheet with the following 3 tabs (case sensitive):
在 Google 試算表中建立以下三個分頁（名稱需完全一致）：

- 學生名單 (Student List): ``` A: Name, B: Email ```

- 固定任務庫 (Task Library):```  A: Task Name, B: Description ```

- 全體進度總表 (Master Progress): ``` A: Name, B: Task, C: Status, D: Date, E: Desc, F: URL ```

2. Apps Script Deployment | 部署腳本
In Google Sheets, go to Extensions > Apps Script.
在試算表中點擊「擴充功能 > Apps Script」。

Copy the code from WebApp.gs and Code.gs into the editor.
將本專案的代碼分別貼入腳本編輯器。

Click Deploy > New Deployment.
點擊「部署 > 新增部署」。

Select Web App:
選擇「網頁應用程式」：

Execute as: Me (Your Email) / 執行身分：我

Who has access: Anyone / 誰可以存取：所有人

- Copy the Web App URL and paste it into the WEB_APP_URL variable in Code.gs.
複製產生的「網頁應用程式網址」，回貼至 Code.gs 中的 WEB_APP_URL 變數。

## 📖 Usage | 使用方式
- Add Data: Fill in students and tasks in their respective tabs.
增加資料：在名單與任務庫填入資料。

- Dispatch: Use the custom menu 🛠️ Admin Tools > 🚀 Dispatch Tasks to send emails.
派發任務：點擊選單「🛠️ 管理工具 > 🚀 執行：派發任務與寄送確認信」。

- Tracking: Monitor the Master Progress tab to see real-time updates as students click the buttons.
追蹤進度：當學生點擊回報時，可在「全體進度總表」看到狀態即時更新為「✅ 已完成」。

## 📂 Project Structure | 檔案結構
Code.gs: Main logic for scanning data, appending rows, and sending emails.

Code.gs: 處理資料掃描、寫入總表及發送郵件的主邏輯。

WebApp.gs: Handles the Web UI, CSS styling, and the progress update API (doGet).

WebApp.gs: 處理網頁介面、CSS 樣式美化及進度更新介面。

## ⚠️ Requirements | 注意事項
The system uses MailApp, which is subject to Google's daily quotas (usually 100 emails/day for free accounts).
本系統使用 Google 每日發信配額（一般帳號每日約 100 封）。

Always create a New Version when re-deploying the Web App after modifying doGet.
修改網頁介面後，必須選擇「新版本」進行重新部署才會生效。

## 📝 License
MIT License. Feel free to fork and modify!
