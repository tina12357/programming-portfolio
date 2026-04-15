# 📊 Project Expense Analytics System | 專案開支分析系統
This sheet automatically aggregates data from the "Expenditure Log," enabling real-time tracking of project lifetime totals and monthly trends using Google Sheets QUERY functions.
本分頁旨在自動彙整「支出流水帳」中的數據，透過 Google Sheets 的 QUERY 函數，實現專案總額加總與月度支出趨勢的即時追蹤。

## 1. Lifetime Project Totals | 專案總累計加總
Function: Calculates the total expenditure for each project/category since inception.
功能： 計算各專案類別自記錄以來的累計支出總額。


- Formula / 公式:
```
=QUERY('支出流水帳共筆（當月在最上面）'!A:H, "select E, sum(H) where E is not null and E != '類別標題1' and E != '專案' group by E label E '類別', sum(H) '總金額'", 1) 
```
 - Logic / 運作邏輯:

   - Data Source / 資料來源: Columns A:H in the Expenditure Log.


   - Aggregation / 加總條件: Groups by Category (Col E) and sums Amount (Col H).


   - Filters / 過濾條件: Excludes empty rows and header text like "Project" or "Category Header".

## 2. Monthly Pivot Analysis | 月份交叉分析統計

Function: Generates a matrix report showing monthly expenditure trends for each project.
功能： 產生矩陣式報表，呈現各專案在不同月份的支出變化。

- Formula / 公式:
```
=QUERY(INDIRECT("'支出流水帳共筆（當月在最上面）'!A2:R"), "select Col5, sum(Col8) where Col5 is not null and Col18 is not null group by Col5 pivot Col18 label Col5 '類別/會計月'", 1)
```
- Logic / 運作邏輯:

   -  Dynamic Range / 動態範圍: Uses INDIRECT to prevent range shifting when new rows are inserted at the top of the log.

   - Pivot Table / 樞紐轉換: Rows show categories (Col5), and Columns show months (Col18).

   - Data Integrity / 資料完整性: Requires both category and accounting month to be present.

### 💡 Maintenance Notes | 維護注意事項
- Automatic Month-Tagging / 月份自動化: Ensure Column R in the Log uses ARRAYFORMULA to convert dates (Col C) into "yyyy-mm" format.
請確保流水帳 R 欄已套用公式，將 C 欄日期自動轉換為「yyyy-mm」格式。

- Array Expansion / 陣列展開: Keep the cells to the right and below the formulas empty to avoid #REF! errors.
請保持公式右方與下方的空間空白，避免出現 #REF! 錯誤（未展開陣列）。

### ⚠️ Troubleshooting | 常見問題排除
 |Issue |Cause |Solution |
 | -----| ------ | ----------- |
 |#REF!  Error |Array cannot expand |.Clear all cells below and to the right of the formula.  |
|Missing Data |Null values in Col18 (R). |Ensure the date is entered in Col C to trigger the auto-month formula.  |
|Duplicate Categories |Leading/trailing spaces in names. |Use TRIM() or data validation to unify category names. |
