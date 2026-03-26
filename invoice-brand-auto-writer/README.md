# Invoice Automation / 發票自動化專案

## Overview / 專案概述
This project automates the process of parsing invoices and mapping vendor-brand data in Google Sheets using Apps Script.  
這個專案利用 Google Sheets 與 Apps Script，自動化發票解析及廠商品牌對應流程。

**Key Features / 核心功能**
- Parse invoice data dynamically / 解析發票資料（不依賴固定格式）  
- Normalize codes and item IDs / 正規化料號與品項代碼  
- Map vendors to brands with fallback logic / 廠商品牌對應與備援策略  
- Write-back automation / 自動回寫到最新 INVOICE 表格  

---
## Architecture / 系統架構
```text
Google Sheets (Invoice & Order & Brand)
       ↓
Apps Script Automation
 ├─ parseInvoice()      → Extract and normalize invoice data
 ├─ parseOrder()        → Extract order/vendor info
 ├─ parseBrand()        → Vendor → Brand mapping
 ├─ normalizeKey()      → Key standardization
 ├─ normalizeBrand()    → Brand string formatting
 └─ writeBack()         → Write results with deduplication (〃)
       ↓
Updated Invoice Sheet


