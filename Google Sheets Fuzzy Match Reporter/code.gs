function generateFuzzyMatchReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  const lastRow = sourceSheet.getLastRow();
  
  if (lastRow < 2) return;

  // 取得原始資料（假設在 A 欄，從 A2 開始）
  const data = sourceSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // 1. 定義排除稱謂與清理函數
  const titles = ["先生", "小姐", "女士", "律師", "教授", "長", "員", "醫師"]; 
  const titleRegex = new RegExp(titles.join("|"), "g");

  const cleanedData = data.map((row, index) => {
    let originalText = String(row[0]);
    let cleanText = originalText.replace(titleRegex, "").replace(/\s+/g, "");
    return {
      rowNum: index + 2, // 實際在工作表中的列號
      original: originalText,
      clean: cleanText
    };
  });

  // 2. 進行比對並收集結果
  let reportData = [["群組", "原始列號", "原始內容", "清理後比對字串", "相似對象列號"]];
  let matchedIndices = new Set();
  let groupCounter = 1;

  for (let i = 0; i < cleanedData.length; i++) {
    if (cleanedData[i].clean.length < 2 || matchedIndices.has(i)) continue;

    let currentMatches = [];
    for (let j = i + 1; j < cleanedData.length; j++) {
      if (isFuzzyMatch(cleanedData[i].clean, cleanedData[j].clean)) {
        currentMatches.push(cleanedData[j]);
        matchedIndices.add(j);
      }
    }

    // 如果有找到重複項，把原始項也加入
    if (currentMatches.length > 0) {
      // 加入基準項
      reportData.push([
        "Group " + groupCounter, 
        cleanedData[i].rowNum, 
        cleanedData[i].original, 
        cleanedData[i].clean, 
        "主項目"
      ]);
      
      // 加入所有匹配項
      currentMatches.forEach(m => {
        reportData.push([
          "Group " + groupCounter, 
          m.rowNum, 
          m.original, 
          m.clean, 
          "與第 " + cleanedData[i].rowNum + " 列相似"
        ]);
      });
      
      groupCounter++;
      reportData.push(["", "", "", "", ""]); // 組與組之間留一個空行
    }
  }

  // 3. 輸出報告分頁
  let reportSheet = ss.getSheetByName("重複檢查報告");
  if (reportSheet) {
    reportSheet.clear(); // 如果已存在則清空
  } else {
    reportSheet = ss.insertSheet("重複檢查報告");
  }

  if (reportData.length > 1) {
    reportSheet.getRange(1, 1, reportData.length, 5).setValues(reportData);
    reportSheet.setFrozenRows(1); // 凍結標題列
    reportSheet.getRange("A1:E1").setBackground("#f3f3f3").setFontWeight("bold");
    reportSheet.autoResizeColumns(1, 5);
    SpreadsheetApp.getUi().alert("報告已產生！請查看「重複檢查報告」分頁。");
  } else {
    SpreadsheetApp.getUi().alert("恭喜！未發現符合條件的重複項目。");
  }
}

// 核心模糊比對邏輯 (保持不變)
function isFuzzyMatch(s1, s2) {
  const minLen = Math.min(s1.length, s2.length);
  let matchCount = 0;
  const set2 = s2.split("");
  s1.split("").forEach(char => {
    const idx = set2.indexOf(char);
    if (idx !== -1) {
      matchCount++;
      set2.splice(idx, 1);
    }
  });

  const ratio = matchCount / minLen;
  if (minLen >= 5 && matchCount >= 4) return true;
  if (minLen === 3 && matchCount >= 2) return true;
  if (minLen === 2 && matchCount >= 2) return true;
  if (minLen > 3 && minLen < 5 && ratio >= 0.75) return true;
  return false;
}
