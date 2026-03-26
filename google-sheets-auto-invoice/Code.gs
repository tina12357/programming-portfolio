function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ 權限修復')
    .addItem('👉 點我進行初次授權', 'generateAndOpenDoc') // 這裡的 'main' 請改為你原本執行程式的函數名稱
    .addToUi();
}
function generateAndOpenDoc() {
  // =======【設定區：請務必填入 ID】=======
  const TEMPLATE_ID_SMALL = 'IDDD'; 
  const TEMPLATE_ID_LARGE = 'IDDDD'; 
  const PRINT_DOC_ID = 'IDDDD'; 
  const EXPIRE_MINUTES = 3; // <--- 設定連結幾分鐘後自動消失
  // ======================================

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const linkCell = sheet.getRange("C1"); 
  linkCell.clearContent(); // 每次開始跑程式時，先清掉舊的連結

  try {
    const rangeList = sheet.getActiveRangeList().getRanges();
    let rowsToProcess = [];
    
    rangeList.forEach(range => {
      let startRow = range.getRow();
      let numRows = range.getNumRows();
      for (let r = 0; r < numRows; r++) {
        let rowIdx = startRow + r;
        if (rowIdx >= 2) rowsToProcess.push(rowIdx);
      }
    });

    if (rowsToProcess.length === 0) {
      linkCell.setValue("⚠️ 請先選取資料列再按按鈕");
      return;
    }

    const printDoc = DocumentApp.openById(PRINT_DOC_ID);
    const printBody = printDoc.getBody();
    printBody.clear(); 
    printBody.setMarginTop(20).setMarginBottom(20).setMarginLeft(40).setMarginRight(40);

    for (let i = 0; i < rowsToProcess.length; i++) {
      let currentRow = rowsToProcess[i];
      const data = {
        date: formatMyDate(sheet.getRange(currentRow, 3).getDisplayValue()),
        subject: sheet.getRange(currentRow, 4).getValue(),
        project: sheet.getRange(currentRow, 5).getValue(),
        type: sheet.getRange(currentRow, 6).getValue().toString().trim(),
        desc: sheet.getRange(currentRow, 7).getValue(),
        amount: parseFloat(sheet.getRange(currentRow, 8).getValue()) || 0,
        payee: sheet.getRange(currentRow, 9).getValue(),
        note: sheet.getRange(currentRow, 14).getValue(),
        bank: sheet.getRange(currentRow, 15).getValue(),
        account: sheet.getRange(currentRow, 16).getDisplayValue(),
        recipient: sheet.getRange(currentRow, 17).getValue()
      };

      let targetTemplateId = (data.amount > 5000 && data.type !== "代墊款") ? TEMPLATE_ID_LARGE : TEMPLATE_ID_SMALL;
      const templateDoc = DocumentApp.openById(targetTemplateId);
      const tempBody = templateDoc.getBody();

      for (let j = 0; j < tempBody.getNumChildren(); j++) {
        const child = tempBody.getChild(j).copy();
        const type = child.getType();
        let newElement;
        if (type == DocumentApp.ElementType.PARAGRAPH) newElement = printBody.appendParagraph(child);
        else if (type == DocumentApp.ElementType.TABLE) newElement = printBody.appendTable(child);
        else if (type == DocumentApp.ElementType.LIST_ITEM) newElement = printBody.appendListItem(child);

        if (newElement) {
          newElement.replaceText('{{Date}}', data.date);
          newElement.replaceText('{{Subject}}', data.subject);
          newElement.replaceText('{{Project}}', data.project);
          newElement.replaceText('{{Description}}', data.desc);
          newElement.replaceText('{{Amount}}', data.amount);
          newElement.replaceText('{{ChineseAmount}}', numberToChinese(data.amount));
          newElement.replaceText('{{Payee}}', data.payee);
          newElement.replaceText('{{Note}}', data.note || "");
          newElement.replaceText('{{Bank}}', data.bank || "");
          newElement.replaceText('{{Account}}', data.account || "");
          newElement.replaceText('{{Recipient}}', data.recipient || "");
        }
      }

      if (i < rowsToProcess.length - 1) {
        if (i % 2 === 0) {
          printBody.appendParagraph("").setSpacingBefore(0).setSpacingAfter(0);
          printBody.appendParagraph("- - - - - - - - - - - - - - - - - 裁 切 線 - - - - - - - - - - - - - - - - -");
        } else {
          printBody.appendPageBreak();
        }
      }
    }
    
    printDoc.saveAndClose();

    // --- 成功生成後的處理 ---
    const url = printDoc.getUrl();
    const richValue = SpreadsheetApp.newRichTextValue()
      .setText("👉 點我開啟並列印 (" + rowsToProcess.length + "筆)")
      .setLinkUrl(url)
      .build();
    
    linkCell.setRichTextValue(richValue).setHorizontalAlignment("right");
    ss.toast('已成功處理 ' + rowsToProcess.length + ' 筆資料。連結將在 ' + EXPIRE_MINUTES + ' 分鐘後失效。', '✅ 完成');



  } catch (e) {
    linkCell.setValue("❌ 錯誤：" + e.toString());
  }
}

// 這個函式會被觸發器呼叫，用來清空 C1 並刪除自己（避免累積太多觸發器）
function clearLinkCell() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getRange("C1");
  cell.setValue("⚠️ 請先選取資料列再按按鈕");
  
  // 清理已經執行完的觸發器，維持專案整潔
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'clearLinkCell') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// --- 輔助函式：日期與金額轉換 (確保這些都在，才不會報錯) ---
function formatMyDate(dStr) {
  dStr = dStr.toString().trim();
  if (dStr.length === 6 && !isNaN(dStr)) {
    return "20" + dStr.substring(0, 2) + "年" + dStr.substring(2, 4) + "月" + dStr.substring(4, 6) + "日";
  }
  return dStr;
}

function numberToChinese(n) {
  if (isNaN(n) || n === "" || n === null) return "零元整";
  const digit = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'];
  const unit = [['元', '萬', '億'], ['', '拾', '佰', '仟']];
  let s = '';
  let num = Math.floor(Math.abs(n));
  if (num === 0) return "零元整";
  for (let i = 0; i < unit[0].length && num > 0; i++) {
    let p = '';
    for (let j = 0; j < unit[1].length && num > 0; j++) {
      p = digit[num % 10] + unit[1][j] + p;
      num = Math.floor(num / 10);
    }
    s = p.replace(/(零.)*零$/, '').replace(/^$/, '零') + unit[0][i] + s;
  }
  return s.replace(/(零.)*零元/, '元').replace(/(零.)+/g, '零').replace(/^$/, '零元') + '整';
}
