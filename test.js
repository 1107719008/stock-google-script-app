// === 個股資料處理 ===
function updateStockSheet(code, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `${code} ${name}`;
  let sheet = ss.getSheetByName(sheetName);

  // 🎯 呼叫新的歷史法人數據 API
  const foreignDataMap = fetchFinMindForeignHistory(code); // 假設此函數已正確定義

  if (!sheet) {
    sheet = ss.insertSheet();
    sheet.setName(sheetName);
    // 🚩 新表頭 (17欄)
    sheet.appendRow(["日期", "開盤", "最高", "最低", "收盤", "成交量", "MA5", "MA10", "MA20", 
                 "外資淨額", "投信淨額", "自營商淨額", // <-- 新增
                 "技術訊號", "建議進場價", "止損價", "目標價", "操作建議"]);
  }

  const data = fetchYahooHistory(code);
  if (!data || data.length === 0) return;

  const lastRow = sheet.getLastRow();
  // 🚨 修正 A: 欄位清除範圍應為 17 欄
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 17).clearContent();

  const closes = data.map(d => d.close);

  data.forEach((d, i) => {
    d.MA5 = avg(closes, i, 5);
    d.MA10 = avg(closes, i, 10);
    d.MA20 = avg(closes, i, 20);

    const last5 = data.slice(Math.max(0, i - 4), i + 1);
    const last10 = data.slice(Math.max(0, i - 9), i + 1);
    const low5 = Math.min(...last5.map(x => x.low));
    const high10 = Math.max(...last10.map(x => x.high));
    const avgVol5 = last5.reduce((a, b) => a + b.volume, 0) / last5.length;

    const isBullish = d.MA5 && d.MA10 && d.MA20 && (d.MA5 > d.MA10 && d.MA10 > d.MA20);
    const isBearish = d.MA5 && d.MA10 && d.MA20 && (d.MA5 < d.MA10 && d.MA10 < d.MA20);
    const isBigRed = (d.close > d.open) && (d.volume > avgVol5 * 1.5);
    const isBigBlack = (d.close < d.open) && (d.volume > avgVol5 * 1.5); // 主力出場

    // 🎯 整合法人數據
    // 嘗試將 K 線日期格式化為 YYYY-MM-DD 以便對齊 (這段邏輯沒有問題)
    const dateKey = (d.date.includes('/')) ? new Date(d.date).toISOString().split('T')[0] : d.date;
    const fData = foreignDataMap[dateKey] || { foreignNet: '', trustNet: '', dealerNet: '' };

    d.foreignNet = fData.foreignNet;
    d.trustNet = fData.trustNet;
    d.dealerNet = fData.dealerNet;

    // 技術面訊號 (邏輯不變)
    if (isBullish) {
      d.signal = "多頭排列";
      d.entry = parseFloat(d.MA5).toFixed(2);
      d.stop = (Math.min(low5, d.MA10) * 0.99).toFixed(2);
      d.target = (Math.max(high10, d.MA20) * 1.03).toFixed(2);
      d.advice = "多頭格局，可沿 MA5 觀察進場";
    } else if (isBearish) {
      d.signal = "空頭排列";
      d.entry = parseFloat(d.MA10).toFixed(2);
      d.stop = (Math.max(high10, d.MA5) * 1.01).toFixed(2);
      d.target = (Math.min(low5, d.MA20) * 0.97).toFixed(2);
      d.advice = "空頭格局，建議觀望或反彈減碼";
    } else {
      d.signal = "";
      d.entry = "";
      d.stop = "";
      d.target = "";
      d.advice = "";
    }

    // 主力進出場判斷 (邏輯不變)
    if (isBigRed) {
      d.signal += (d.signal ? "＋" : "") + "主力進場";
      d.advice += (d.advice ? "；" : "") + "主力紅K留意隔日續攻";
    }

    if (isBigBlack) {
      d.signal += (d.signal ? "＋" : "") + "主力出場";
      d.advice += (d.advice ? "；" : "") + "主力黑K留意隔日回檔或出貨";
    }
  });

  // 🎯 數據輸出 (17個元素)
  const rows = data.map(d => [
      d.date, d.open, d.high, d.low, d.close, d.volume, d.MA5, d.MA10, d.MA20, 
      d.foreignNet, d.trustNet, d.dealerNet, // <-- 新增的法人數據
      d.signal, d.entry, d.stop, d.target, d.advice
  ]);

  // 🚩 更新寫入範圍的欄位數 (17 欄)
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // 🚩 標示範圍：技術訊號欄位現在是第 13 欄
  // 🚨 修正 B & C: 移除重複聲明，並修正欄位索引
  const sigRange = sheet.getRange(2, 13, rows.length, 1); // <-- 技術訊號欄位是第 13 欄
  const sigValues = sigRange.getValues();
  for (let i = 0; i < sigValues.length; i++) {
    // 🚨 修正：這裡的 cell 應該是標示技術訊號的儲存格
    const cell = sheet.getRange(i + 2, 13, 1, 1); 
    const text = sigValues[i][0].toString();
    
    // 標示主力紅K / 黑K
    if (text.includes("主力進場")) {
      cell.setBackground("#ffcccc"); // 紅底
    } else if (text.includes("主力出場")) {
      cell.setBackground("#cce5ff"); // 淡藍底
    } else {
      cell.setBackground(null);
    }
  }
}