# 📄 Excel Salary Slip PDF Generator / 薪資單 PDF 批次匯出工具

## 📖 Overview / 專案概述
This project is an Excel VBA macro that converts each worksheet into a separate PDF file.  
這是一個 Excel VBA 巨集，能將每個工作表轉換為獨立的 PDF 檔案。

Each PDF is automatically named based on the workbook name and worksheet name.  
每個 PDF 檔案會依據「檔案名稱 + 分頁名稱」自動命名。

---

## 🎯 Use Case / 使用情境
- Batch export salary slips / 批次匯出薪資單  
- Generate reports per employee / 每位員工一份報表  
- Automate repetitive Excel export tasks / 自動化重複性 Excel 操作  

---

## ✨ Features / 功能特色

- 📄 Export each worksheet as a PDF / 每個分頁輸出成 PDF  
- 🏷 Auto naming based on sheet name / 自動依分頁命名  
- 📁 Auto create output folder / 自動建立輸出資料夾  
- ⚡ One-click execution / 一鍵執行  

---

## 🧠 How It Works / 核心邏輯

```text
Excel Workbook
      ↓
Loop through Worksheets
      ↓
Generate PDF for each sheet
      ↓
Save to Folder (WorkbookName/)

--
📂 Output Example / 輸出範例

假設檔名為：Salary.xlsx

輸出結果：

/Salary/
├── Salary_A.pdf
├── Salary_B.pdf
├── Salary_C.pdf


##🚀 How to Use / 使用方式
- Open Excel file / 開啟 Excel 檔案
- Press Alt + F11 to open VBA editor / 開啟 VBA 編輯器
- Insert → Module / 插入模組
- Paste the code / 貼上程式碼
- Run pdf轉檔() / 執行巨集
##⚠️ Notes / 注意事項
- Worksheet names will be used as PDF names
分頁名稱會直接作為 PDF 檔名
- Avoid illegal file characters (e.g. \ / : * ? " < > |)
請避免分頁名稱含非法字元
- All worksheets will be exported
所有分頁都會被輸出
##📈 My Contribution / 我的貢獻
- Designed automation logic for batch PDF export
- Implemented dynamic file naming and folder creation
- Reduced manual export workload significantly
##🔮 Future Improvements / 未來優化
- Export selected sheets only / 選擇性匯出分頁
- Add UI button in Excel / 加入按鈕操作
- Customize file naming rules / 自訂命名規則
- Add logging / 匯出紀錄
