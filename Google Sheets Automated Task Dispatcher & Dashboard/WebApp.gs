function doGet(e) {
  const studentName = e.parameter.name;
  const taskName = e.parameter.task;
  const action = e.parameter.action;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("全體進度總表");
  const data = masterSheet.getDataRange().getValues();
  const WEB_APP_URL = ScriptApp.getService().getUrl();

  // CSS 樣式表
  const style = `
    <style>
      body { font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; background-color: #f5f7fa; color: #333; margin: 0; padding: 20px; line-height: 1.6; }
      .container { max-width: 500px; margin: 0 auto; }
      .header { text-align: center; margin-bottom: 30px; }
      .task-card { background: white; border-radius: 12px; padding: 20px; margin-bottom: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-left: 5px solid #4CAF50; }
      .task-title { font-size: 18px; font-weight: bold; color: #2c3e50; margin-bottom: 8px; }
      .task-desc { font-size: 14px; color: #666; margin-bottom: 15px; }
      .btn { display: block; text-align: center; text-decoration: none; padding: 12px; border-radius: 8px; font-weight: bold; transition: 0.3s; cursor: pointer; border: none; width: 100%; box-sizing: border-box; }
      .btn-report { background-color: #4CAF50; color: white; }
      .btn-report:hover { background-color: #45a049; }
      .btn-resend { background-color: #2196F3; color: white; margin-top: 30px; }
      .btn-back { background-color: #9e9e9e; color: white; margin-top: 10px; }
      .success-msg { text-align: center; color: #2e7d32; padding: 40px 20px; }
    </style>
  `;

  // 動作 A：處理單項回報
  if (taskName && !action) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == studentName && data[i][1] == taskName) {
        masterSheet.getRange(i + 1, 3).setValue("✅ 已完成");
        found = true; break;
      }
    }
    const msg = found ? `<h2>✅ 回報成功！</h2><p>任務「${taskName}」已更新。</p>` : `<h2>❌ 找不到任務</h2>`;
    return HtmlService.createHtmlOutput(`${style}<div class="container success-msg">${msg}<a href="${WEB_APP_URL}?name=${encodeURIComponent(studentName)}&action=list" class="btn btn-back">回到清單</a></div>`);
  }

  // 動作 B：顯示個人清單
  const unfinishedTasks = data.filter(row => row[0] === studentName && row[2] === "❌ 未開始");
  
  let listHtml = `${style}<div class="container">
    <div class="header">
      <h2>📋 任務儀表板</h2>
      <p>你好 <b>${studentName}</b>，請完成以下任務</p>
    </div>`;

  if (unfinishedTasks.length === 0) {
    listHtml += `<div class="task-card" style="text-align:center;"><p>🎉 太棒了！目前沒有待辦任務。</p></div>`;
  } else {
    unfinishedTasks.forEach(row => {
      const reportUrl = `${WEB_APP_URL}?name=${encodeURIComponent(studentName)}&task=${encodeURIComponent(row[1])}`;
      listHtml += `
        <div class="task-card">
          <div class="task-title">${row[1]}</div>
          <div class="task-desc">${row[4] || '無任務說明'}</div>
          <a href="${reportUrl}" class="btn btn-report">回報完畢</a>
        </div>`;
    });
  }

  // 補寄按鈕
  const resendUrl = `${WEB_APP_URL}?name=${encodeURIComponent(studentName)}&action=resend`;
  listHtml += `<a href="${resendUrl}" class="btn btn-resend">📧 補寄清單到 Email</a></div>`;

  // 處理補寄請求
  if (action === "resend") {
    // 呼叫主程式發信邏輯
    const result = triggerResendEmail(studentName);
    return HtmlService.createHtmlOutput(`${style}<div class="container success-msg"><h2>📩 ${result}</h2><a href="${WEB_APP_URL}?name=${encodeURIComponent(studentName)}&action=list" class="btn btn-back">回到清單</a></div>`);
  }

  return HtmlService.createHtmlOutput(listHtml).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
