function runProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const invoiceSheet = findInvoiceSheet(ss);
  const orderSheet = ss.getSheetByName("點貨單");
  const brandSheet = ss.getSheetByName("商標");

  const invoiceData = parseInvoice(invoiceSheet);
  const orderMap = parseOrder(orderSheet);
  const brandMap = parseBrand(brandSheet);

  Logger.log("=== INVOICE DATA ===");
  Logger.log(JSON.stringify(invoiceData, null, 2));

  // ✅ 這行「一定要有」
  writeBack(invoiceSheet, invoiceData, orderMap, brandMap);
}
function writeBack(sheet, invoiceData, orderMap, brandMap) {
  invoiceData.sort((a, b) => a.row - b.row);
let prevOutput = "";
  invoiceData.forEach(item => {

  let code = normalizeKey(item.code);
  let rawItemCode = item.itemCode || "";
  let itemCode = normalizeKey(rawItemCode);

  let key = code + "_" + itemCode;

  // ✅ 先用精準 key
  let recordList = orderMap[key];

  // ✅ fallback 1：用 code（最重要）
  if (!recordList || recordList.length === 0) {
    recordList = orderMap[code] || [];
  }

  let record = (recordList.length > 0) ? recordList.shift() : {};

  let vendor = normalizeKey(record.vendor || "");

  let brand = "";

  if (vendor === "東陽") {
    brand = "肉眼";
  } else if (!vendor) {
    brand = "no brand";
  } else {
    brand = brandMap[vendor] ? normalizeBrand(brandMap[vendor]) : "OTN";
  }
Logger.log({
  row: item.row,
  code: code,
  itemCode: itemCode,
  key: key,
  recordExists: !!orderMap[key],
  fallbackExists: !!orderMap[code]
});
  // ✅ 用 output 比較（關鍵）
  let output = (brand === prevOutput) ? "〃" : brand;

  sheet.getRange(item.row, 1).setValue(output);

  prevOutput = output;
});
}
/**
 * 找最新 INVOICE 表
 */
function findInvoiceSheet(ss) {
  let sheets = ss.getSheets();

  let invoiceSheets = sheets.filter(s => s.getName().startsWith("INVOICE"));

  invoiceSheets.sort((a, b) => a.getName().localeCompare(b.getName()));

  return invoiceSheets[invoiceSheets.length - 1];
}

/**
 * 正規化字串（超重要）
 */
function normalizeKey(str) {
  return (str || "")
    .toString()
    .replace(/（.*?）/g, "")
    .replace(/\(.*?\)/g, "")
    .replace(/\u00A0/g, "")
    .replace(/\s+/g, "")   // ✅ 移除空白（key需要）
    .trim();
}
function normalizeBrand(str) {
  return (str || "")
    .toString()
    .replace(/\u00A0/g, " ") // 非斷行空白轉正常空白
    .trim();                // ✅ 保留空白
}


/**
 * 解析 Invoice（不依賴固定格式）
 */
function parseInvoice(sheet) {
  let map = {};
  let currentCode = "";

  const startRow = 10;
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 4).getValues();

  for (let i = startRow - 1; i < lastRow; i++) {
    let text = data[i][2];
    if (!text) continue;

    let str = text.toString();

    let codeMatch = str.match(/\d{1,4}-\d{1,4}/);
    if (codeMatch) {
      currentCode = normalizeKey(codeMatch[0]);
    }

    let itemMatch = str.match(/^[A-Z0-9\-]+/i);
    let itemCode = itemMatch ? normalizeKey(itemMatch[0]) : "";

    if (!currentCode) continue;

    let key = currentCode + "_" + itemCode;

    if (!map[key]) {
      map[key] = {
        row: i + 1,
        code: currentCode,
        itemCode: itemCode,
        count: 0
      };
    }

    // ✅ 累加（代表同一料號拆單）
    map[key].count += 1;
  }

  return Object.values(map);
}
/**
 * 解析點貨單（去掉 i % 4 假設）
 */
function parseOrder(sheet) {
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 2, lastRow, 3).getValues(); // B:D

  let map = {};
  let currentVendor = "";

  for (let i = 0; i < data.length; i++) {
   let vendorCell = data[i][0]; // B欄
    let descCell = data[i][2];    // D欄（當列）
    let descCellNext = data[i + 1] ? data[i + 1][2] : ""; // ⭐ 下一列 D欄
    // ✅ 更新 vendor（往下延續）
    if (vendorCell && vendorCell.toString().trim() !== "") {
      currentVendor = normalizeKey(vendorCell);
    }

    if (!descCell) continue;

    let text = normalizeKey(descCell);

    // ✅ 抓 code（放寬，避免漏掉）
    let match = text.match(/\d{1,4}-\d{1,4}/);

 if (match) {
  let code = normalizeKey(match[0]);
  let itemCode = normalizeKey(descCellNext || "");

  let key = code + "_" + itemCode;

  // ✅ key map
  if (!map[key]) map[key] = [];
  map[key].push({
    vendor: currentVendor,
    itemCode: itemCode
  });

  // ✅ code map（fallback 用）
  if (!map[code]) map[code] = [];
  map[code].push({
    vendor: currentVendor,
    itemCode: itemCode
  });
}
  }

  return map;
}
function getBrandByCode(code) {
  if (!code) return "OTN";
let c = code.trim();

  // ✅ 只要結尾是 S（允許前面有其他字）
  return c.endsWith("S") ? "V-WIN" : "OTN";
}
