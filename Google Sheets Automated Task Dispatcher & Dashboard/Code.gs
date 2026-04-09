const WEB_APP_URL = "你的網頁應用程式網址"; 

/**
 * 自動建立選單
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ 管理工具')
      .addItem('🚀 執行：派發任務與寄送確認信', 'main_Process')
      .addToUi();
}

function main_Process() {
  generateAndEmailTasks(SpreadsheetApp.getActiveSpreadsheet());
}

// 供 WebApp 呼叫的補寄功能
function triggerResendEmail(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("學生名單");
  const data = studentSheet.getRange(2, 1, studentSheet.getLastRow()-1, 2).getValues();
  const student = data.find(r => r[0] === name);
  if (student) {
    generateAndEmailTasks(ss, name, true); 
    return "已將清單寄送至您的信箱！";
  }
  return "找不到您的 Email 資訊。";
}

/**
 * 核心功能
 */
function generateAndEmailTasks(ss, targetName = null, isForced = false) {
  const studentSheet = ss.getSheetByName("學生名單");
  const templateSheet = ss.getSheetByName("固定任務庫");
  const masterSheet = ss.getSheetByName("全體進度總表");
  
  const lastRow = studentSheet.getLastRow();
  const lastTemplateRow = templateSheet.getLastRow();
  
  console.log("目前的學生名單最後一行為：" + lastRow);
  if (lastRow < 2) {
    console.warn("❌ 學生名單內沒有資料");
    return;
  }

  // === 修正點：務必定義 templates ===
  const students = studentSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const templates = templateSheet.getRange(2, 1, lastTemplateRow - 1, 2).getValues(); 
  console.log("成功抓取學生人數：" + students.length);
  console.log("成功抓取任務總數：" + templates.length);

  students.forEach(sRow => {
    const sName = String(sRow[0]).trim();
    const sEmail = String(sRow[1]).trim();
    if (!sName || !sEmail) return;
    if (targetName && sName !== targetName) return; 

    // 每次比對前抓取最新的總表內容
    let existingData = masterSheet.getDataRange().getValues();
    let hasNewTask = false;

    // 1. 檢查並寫入新任務
    templates.forEach(tRow => {
      const tName = String(tRow[0]).trim();
      const tDesc = String(tRow[1]).trim();
      const isExist = existingData.some(row => String(row[0]).trim() === sName && String(row[1]).trim() === tName);
      
      if (!isExist) {
        const url = `${WEB_APP_URL}?name=${encodeURIComponent(sName)}&task=${encodeURIComponent(tName)}`;
        masterSheet.appendRow([sName, tName, "❌ 未開始", "", tDesc, url]);
        hasNewTask = true; 
      }
    });

    // 強制更新緩存
    SpreadsheetApp.flush();

    // 2. 寄信邏輯
    if (hasNewTask || isForced) {
      const latestData = masterSheet.getDataRange().getValues();
      const currentUnfinished = latestData.filter(row => String(row[0]).trim() === sName && row[2] === "❌ 未開始");

      if (currentUnfinished.length > 0) {
        const top5 = currentUnfinished.slice(0, 5);
        let html = `<h3>你好 ${sName}，您有 ${currentUnfinished.length} 項未完成任務：</h3>`;
        
        top5.forEach(task => {
          html += `<div style="border:1px solid #eee; padding:10px; margin-bottom:5px; border-radius:5px;">
                    <b style="font-size:16px;">${task[1]}</b><br>
                    <span style="color:#666;">${task[4] || ""}</span><br><br>
                    <a href="${task[5]}" style="background-color:#4CAF50; color:white; padding:5px 10px; text-decoration:none; border-radius:3px;">單項回報</a>
                  </div>`;
        });

        const dashUrl = `${WEB_APP_URL}?name=${encodeURIComponent(sName)}&action=list`;
        
        if (currentUnfinished.length > 5) {
          html += `<p style="color:red;"><b>還有 ${currentUnfinished.length - 5} 個任務未列出...</b></p>`;
        }
        
        html += `<br><a href="${dashUrl}" style="display:inline-block; padding:15px; background:#4CAF50; color:white; text-decoration:none; border-radius:10px; font-weight:bold;">進入「個人任務儀表板」查看全部</a>`;

        try {
          MailApp.sendEmail({ 
            to: sEmail, 
            subject: isForced ? `【補寄】您的任務清單` : `【重要通知】您有 ${currentUnfinished.length} 項工作待處理`, 
            htmlBody: html 
          });
          console.log(`✅ 已成功寄信給：${sName} (${sEmail})`);
        } catch (e) {
          console.error(`❌ 寄信給 ${sName} 時出錯：${e.message}`);
        }
      }
    } else {
      console.log(`ℹ️ 學生 ${sName} 無新任務，跳過寄信。`);
    }
  });
  
  if (!targetName) {
    SpreadsheetApp.getUi().alert("✅ 處理完成！請查看「執行紀錄」確認寄信明細。");
  }
}
