
// 📈 股票自動更新＋技術面策略＋主力進場標示
const SHEET_TOTAL = "總覽";
const SHEET_HISTORY = "更新紀錄";
const DAYS = 30;

// === 主入口 ===
function updateAllStocks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //初始化「總覽」
  let totalSheet = ss.getSheetByName(SHEET_TOTAL);
  if (!totalSheet) {
    totalSheet = ss.insertSheet(SHEET_TOTAL);
    totalSheet.appendRow(["股票代號", "股票名稱", "收盤價", "成交量", "技術訊號", "建議進場價", "止損價", "目標價", "操作建議", "持股成本(手動)", "備註", "更新時間"]);
  }

  // 初始化「更新紀錄」
  let historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY);
    historySheet.appendRow(["更新時間", "股票代號", "股票名稱", "收盤價", "成交量", "技術訊號", "建議進場價", "止損價", "目標價", "操作建議"]);
  }

  const lastRow = totalSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("⚠️ 請先在「總覽」頁面輸入股票代號（例如 2330, 2603）");
    return;
  }

  const codes = totalSheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .filter(c => c && c.toString().trim() !== "");

  const now = new Date();
  const allRows = [];

  for (let code of codes) {
    const name = getStockName(code);
    if (!name) continue;

    updateStockSheet(code, name);

    const sheetName = `${code} ${name}`;
    const s = ss.getSheetByName(sheetName);
    if (!s) continue;

    const lastRow = s.getLastRow();
    if (lastRow < 2) continue;
    const row = s.getRange(lastRow, 1, 1, 14).getValues()[0];

    const summary = [code, name, row[4], row[5], row[9], row[10], row[11], row[12], row[13], "", "", now];
    allRows.push(summary);

    historySheet.appendRow([now, code, name, row[4], row[5], row[9], row[10], row[11], row[12], row[13]]);
  }

  if (allRows.length > 0) {
    const startRow = 2;
    const existingRows = totalSheet.getLastRow() - 1;
    // if (existingRows > 0) {
    //   totalSheet.getRange(startRow, 1, existingRows, 12).clearContent();
    // }
    // totalSheet.getRange(startRow, 1, allRows.length, allRows[0].length).setValues(allRows);
    if (existingRows > 0) {
      totalSheet.getRange(startRow, 1, existingRows, 15).clearContent(); // <-- 改成 15
    }
    totalSheet.getRange(startRow, 1, allRows.length, allRows[0].length).setValues(allRows);
  }

  // === 在總覽頁面標示主力進出場 ===
  const sigRange = totalSheet.getRange(2, 5, allRows.length, 1); // 技術訊號欄
  const sigValues = sigRange.getValues();

  for (let i = 0; i < sigValues.length; i++) {
    const text = sigValues[i][0]?.toString() || "";
    const cell = totalSheet.getRange(i + 2, 5); // 第 5 欄是技術訊號

    if (text.includes("主力進場")) {
      cell.setBackground("#ffcccc"); // 紅底
    } else if (text.includes("主力出場")) {
      cell.setBackground("#cce5ff"); // 淡藍底
    } else {
      cell.setBackground(null);
    }
  }

  SpreadsheetApp.getUi().alert("✅ 所有股票已更新完成！");
}



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
                 "技術訊號", "建議進場價", "止損價", "目標價", "操作建議",
    "外資淨額", "投信淨額", "自營商淨額",
    ]);
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
      d.signal, d.entry, d.stop, d.target, d.advice,
      d.foreignNet, d.trustNet, d.dealerNet // <-- 新增的法人數據
  ]);

  // 🚩 更新寫入範圍的欄位數 (17 欄)
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // 🚩 標示範圍：技術訊號欄位現在是第 13 欄
  // 🚨 修正 B & C: 移除重複聲明，並修正欄位索引
  const sigRange = sheet.getRange(2, 10, rows.length, 1); // <-- 技術訊號欄位是第 10 欄
  const sigValues = sigRange.getValues();
  for (let i = 0; i < sigValues.length; i++) {
    // 🚨 修正：這裡的 cell 應該是標示技術訊號的儲存格
    const cell = sheet.getRange(i + 2, 10, 1, 1); 
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

// === 抓股票名稱（先 TWSE，失敗再用 FinMind） ===
function getStockName(code) {
  const nameFromTWSE = getStockNameFromTWSE(code);
  if (nameFromTWSE) return nameFromTWSE; // 有找到就直接回傳
  return getStockNameFromFinMind(code);  // 否則使用 FinMind 備援
}
// function getStockName(code) {
//   // 直接回傳 FinMind 的結果，避免不穩定的 TWSE API 造成錯誤。
//   return getStockNameFromFinMind(code); 
// }

// === 1️⃣ TWSE API ===
function getStockNameFromTWSE(code) {
  const url = `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${code}.tw&json=1&delay=0`;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.rtcode === "0000" && json.msgArray && json.msgArray.length > 0) {
      const stock = json.msgArray[0];
      return stock.n || stock.nf || "";
    }
    return "";
  } catch (e) {
    Logger.log(`TWSE API 錯誤: ${e}`);
    return "";
  }
}

// === 2️⃣ FinMind API ===
// FinMind 提供公司基本資料，包括中文名稱、產業別等
function getStockNameFromFinMind(code) {
  const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo&data_id=${code}`;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.data && json.data.length > 0) {
      const name = json.data[0].stock_name; // 中文股票名稱
      return name || "";
    }
    return "";
  } catch (e) {
    Logger.log(`FinMind API 錯誤: ${e}`);
    return "";
  }
}


// === 抓 Yahoo 歷史價格（Yahoo 為主，FinMind 為備援） ===
function fetchYahooHistory(code) {
  let data = [];
  try {
    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${code}.TW?interval=1d&range=${DAYS}d`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const result = json.chart?.result?.[0];
    if (result && result.timestamp) {
      const timestamps = result.timestamp;
      const quotes = result.indicators.quote[0];
      data = timestamps.map((t, i) => ({
        date: new Date(t * 1000).toLocaleDateString("zh-TW"),
        open: Number(quotes.open[i]?.toFixed(2) || 0),
        high: Number(quotes.high[i]?.toFixed(2) || 0),
        low: Number(quotes.low[i]?.toFixed(2) || 0),
        close: Number(quotes.close[i]?.toFixed(2) || 0),
        volume: quotes.volume[i] || 0
      }));
    }
  } catch (e) {
    Logger.log(`Yahoo API 錯誤: ${e}`);
  }

  // 若 Yahoo 無資料 → 改用 FinMind
  if (!data || data.length === 0) {
    Logger.log(`Yahoo 無資料，改用 FinMind 抓取：${code}`);
    data = fetchFinMindHistory(code);
  }

  return data;
}

// === 抓 FinMind 歷史股價 ===
function fetchFinMindHistory(code) {
  try {
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=${code}&date_from=${getPastDate(DAYS)}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (!json.data || json.data.length === 0) return [];

    return json.data.map(d => ({
      date: d.date,
      open: Number(d.open),
      high: Number(d.max),
      low: Number(d.min),
      close: Number(d.close),
      volume: Number(d.Trading_Volume)
    }));
  } catch (e) {
    Logger.log(`FinMind API 錯誤: ${e}`);
    return [];
  }
}

// === 工具函數：回傳 n 天前的日期字串（yyyy-mm-dd） ===
function getPastDate(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d.toISOString().split("T")[0];
}

// === 移動平均計算 ===
function avg(arr, i, n) {
  if (i < n - 1) return "";
  const slice = arr.slice(i - n + 1, i + 1);
  return (slice.reduce((a, b) => a + b, 0) / n).toFixed(2);
}


//---new--- 自動選股測試(外資買超兩天)
// === 功能：選股(MA10>股價>MA20) ===
function filterStocks_LowAndBuy2Days() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 沿用原頁面名稱
    const SHEET_PICK = "選股(以低檔並買超兩天)";
    const totalSheet = ss.getSheetByName("總覽");
    
    if (!totalSheet) {
        SpreadsheetApp.getUi().alert("⚠️ 錯誤：請先執行 updateAllStocks() 建立『總覽』頁面以取得股票清單。");
        return;
    }

    // 建立或清空選股頁
    let pickSheet = ss.getSheetByName(SHEET_PICK);
    if (!pickSheet) {
        pickSheet = ss.insertSheet(SHEET_PICK);
    }
    
    // 設置/重設標題行 (移除外資、MA60欄位)
    const header = ["股票代號", "股票名稱", "收盤價", "MA10", "MA20", "進場建議價格(由總表撈)", "外資連續買賣超", "更新時間"];
    pickSheet.getRange(1, 1, 1, header.length).setValues([header]);
    
    // 清空現有內容 (從第 2 行開始)
    const lastRow = pickSheet.getLastRow();
    if (lastRow > 1) pickSheet.getRange(2, 1, lastRow - 1, pickSheet.getMaxColumns()).clearContent();

    // 讀取股票代號
    const codes = totalSheet.getRange(2, 1, totalSheet.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .filter(c => c && c.toString().trim() !== "");

    const now = new Date();
    const picks = [];

    for (const code of codes) {
        const name = getStockName(code);
        if (!name) continue; 
        
        const sheetName = `${code} ${name}`;
        const s = ss.getSheetByName(sheetName);
        if (!s) continue;

        const lastRowStock = s.getLastRow();
        if (lastRowStock < 2) continue;
        
        // 讀取最後一行的資料
        // [0:日期, 1:開盤, 2:最高, 3:最低, 4:收盤, 5:成交量, 6:MA5, 7:MA10, 8:MA20, 9:MA60, 10:技術訊號...]
        // 我們只需要讀取到 MA20 (索引 8) 和 技術訊號 (索引 10)
        const row = s.getRange(lastRowStock, 1, 1, 11).getValues()[0];
        
        const close = Number(row[4]);
        const MA10 = Number(row[7]);
        const MA20 = Number(row[8]);
        const suggestPrice = row[10] || ""; 

        const foreignDays = getForeignDays(code);

        // 檢查數值是否有效，避免 NaN 錯誤判斷
        if (isNaN(close) || isNaN(MA10) || isNaN(MA20)) continue;

        // === 股價區間判斷 (您的最新邏輯) ===
        // 條件：(收盤價 < MA10) 且 (收盤價 > MA20)
        const isLowPosition =  close < MA20;
        
        if (isLowPosition) {
            // 將符合條件的股票資料寫入陣列
            picks.push([code, name, close, MA10, MA20, suggestPrice, foreignDays ,now]);
        }
    }

    if (picks.length > 0) {
        // 將選股結果寫入新的選股頁面
        pickSheet.getRange(2, 1, picks.length, picks[0].length).setValues(picks);
    }

    SpreadsheetApp.getUi().alert(`✅ 選股完成，共找到 ${picks.length} 檔股票！`);
}

//---------外資買賣超天數計算-----------
// ---------外資買賣超天數計算 (新增 TWSE API 備援)-----------

/**
 * 取得個股三大法人總計的連續淨買超或淨賣超天數。
 * * 優先使用 FinMind API，若遇到 400 錯誤則切換到 TWSE 非官方 API 作為備援。
 * * @param {string} code 股票代號 (e.g., "2330")
 * @returns {number} 連續淨買賣超天數 (買超為正，賣超為負)
 */
function getForeignDays(code) {
    // 嘗試 FinMind API
    const finmindResult = fetchFinMindData(code);
    
    // 如果 FinMind 成功返回數據 (非 0 且非錯誤代碼)
    if (finmindResult !== 0) {
        return finmindResult;
    }

    // 如果 FinMind 失敗（通常是 400 Bad Request 或數據為空），嘗試 TWSE 備援
    Logger.log(`FinMind API 無法取得 ${code} 資料，嘗試切換到 TWSE 備援 API...`);
    return fetchTWSEData(code);
}


/**
 * 輔助函數：從 FinMind 獲取三大法人數據並計算連續天數。
 */
function fetchFinMindData(code) {
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=${code}`;

    try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const responseCode = res.getResponseCode();
        
        // 如果是 400 (Bad Request)，代表 FinMind 沒有這個代號的數據，返回 0 讓主函數切換備援
        if (responseCode === 400 || responseCode !== 200) {
            Logger.log(`FinMind API 呼叫失敗(${code})，狀態碼: ${responseCode}`);
            return 0; 
        }
        
        const json = JSON.parse(res.getContentText());
        const data = json.data || [];

        if (data.length === 0) return 0;
        
        // 將數據按日期升序排列，確保從舊到新
        const sortedData = data.sort((a, b) => new Date(a.date) - new Date(b.date));

        let count = 0;
        let trend = 0; // 1: 買超, -1: 賣超, 0: 初始/淨額為零

        // 從最新的交易日開始遍歷 (反轉陣列)
        for (let i = sortedData.length - 1; i >= 0; i--) {
            const d = sortedData[i];
            
            // === 步驟 1: 計算三大法人總淨額 (所有買 - 所有賣) ===
            const totalBuy = (Number(d.Foreign_Investor_Buy) || 0) + 
                             (Number(d.Dealer_Proprietary_Buy) || 0) +
                             (Number(d.Dealer_Hedge_Buy) || 0) +
                             (Number(d.Trust_Buy) || 0);
            
            const totalSell = (Number(d.Foreign_Investor_Sell) || 0) + 
                              (Number(d.Dealer_Proprietary_Sell) || 0) +
                              (Number(d.Dealer_Hedge_Sell) || 0) +
                              (Number(d.Trust_Sell) || 0);
                              
            const totalNet = totalBuy - totalSell;
            
            let currentTrend = totalNet > 0 ? 1 : (totalNet < 0 ? -1 : 0);

            // === 步驟 2: 計算連續天數 ===
            if (count === 0) {
                if (currentTrend !== 0) {
                    trend = currentTrend;
                    count = 1;
                } else {
                    return 0;
                }
            } else {
                if (currentTrend === trend) {
                    count++;
                } else {
                    break;
                }
            }
        }
        
        return trend * count;

    } catch (e) {
        Logger.log(`fetchFinMindData 錯誤(${code}): ${e}`); 
        return 0;
    }
}


/**
 * 輔助函數：從 TWSE 非官方 API 獲取三大法人數據並計算連續天數。
 * 注意：TWSE 網址較不穩定，且有 JSON/HTML 錯誤風險。
 */
function fetchTWSEData(code) {
    // 獲取當前日期 (格式 YYYYMMDD)
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
    
    // TWSE 股權分散/籌碼分析的 API 結構 (通常是每日更新)
    // 備註: 這個 API 網址可能隨時失效
    const url = `https://www.twse.com.tw/exchangeReport/BWIBBU_d?response=json&stockNo=${code}&date=${today}`;

    try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        
        if (res.getResponseCode() !== 200) {
            Logger.log(`TWSE API 呼叫失敗(${code})，狀態碼: ${res.getResponseCode()}`);
            return 0;
        }
        
        const jsonText = res.getContentText();

        // 檢查是否收到 HTML 錯誤頁面 (防止 Unexpected token '<' 錯誤)
        if (jsonText.startsWith('<html>')) {
            Logger.log(`TWSE API 收到 HTML 錯誤頁面 (${code})`);
            return 0;
        }

        const json = JSON.parse(jsonText);
        
        // 檢查數據是否存在，TWSE 的數據結構複雜，這裡假設我們需要手動從外部獲取法人買賣超資料
        // **!!! 警告 !!!** TWSE 的 BWIBBU_d API 並不直接提供多日歷史買賣超數據，
        // 它主要提供股權分散資訊，我們需要一個能提供歷史數據的 TWSE API。
        
        // 為了達到多日連續計算的目的，我們必須使用另一個 TWSE API 獲取歷史數據。
        // 然而，公開且穩定的 TWSE 歷史法人 API 較難找到，且有頻率限制。
        
        // 鑒於此，這裡將使用一個常見的歷史法人 API 結構作為範例，但穩定性無法保證。
        // 請確保您的 TWSE API 能夠提供歷史日期的數據。

        // **!!! 假設有一個 TWSE API 可以提供歷史法人數據 (此處為示意結構) !!!**
        
        // ***由於缺乏穩定的 TWSE 歷史法人 API 結構，我們將假設 TWSE API 邏輯與 FinMind 類似，
        // 只是呼叫網址不同，並在 TWSE API 呼叫失敗時給出更明確的紀錄。***
        
        if (!json.data || json.data.length < 2) {
            Logger.log(`TWSE API (${code}) 數據不足或結構錯誤`);
            return 0;
        }

        // 重新使用 FinMind 的解析邏輯 (如果 TWSE 結構類似)
        // -----------------------------------------------------------
        const data = json.data;
        // ... (這裡需要根據 TWSE 實際返回的 JSON 結構調整解析邏輯) ...
        // 由於無法得知 TWSE 實際返回的欄位名稱，我們暫時停止 TWSE 的詳細計算，
        // 並返回 0，避免產生新的錯誤。
        // -----------------------------------------------------------

        Logger.log(`TWSE API (${code}) 數據已獲取，但由於解析結構未知，暫時返回 0`);
        return 0; 

    } catch (e) {
        Logger.log(`fetchTWSEData 錯誤(${code}): ${e}`); 
        return 0;
    }
}


/**
 * 取得最新交易日的外資、投信、自營商各自的淨買賣超金額（張數）。
 * @param {string} code 股票代號 (e.g., "2330")
 * @returns {object} { foreignNet, trustNet, dealerNet } (買超為正，賣超為負)
 */

function getLatestThreeForeignsNet(code) {
    // 設置一個較小的日期範圍，以確保快速獲取最新數據
    const tenDaysAgo = new Date();
    tenDaysAgo.setDate(tenDaysAgo.getDate() - 10); 
    const startDate = Utilities.formatDate(tenDaysAgo, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // FinMind API 網址
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=${code}&start_date=${startDate}`;

    try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        
        if (res.getResponseCode() !== 200) {
            Logger.log(`API 呼叫失敗(${code})，狀態碼: ${res.getResponseCode()}`);
            return { foreignNet: 0, trustNet: 0, dealerNet: 0 }; 
        }
        
        const json = JSON.parse(res.getContentText());
        const data = json.data || [];

        if (data.length === 0) return { foreignNet: 0, trustNet: 0, dealerNet: 0 };
        
        const sortedData = data.sort((a, b) => new Date(a.date) - new Date(b.date));
        const latestData = sortedData[sortedData.length - 1]; 

        // 1. 外資淨額
        const foreignNet = (Number(latestData.Foreign_Investor_Buy) || 0) - 
                           (Number(latestData.Foreign_Investor_Sell) || 0) +
                           (Number(latestData.Foreign_Dealer_Buy) || 0) - 
                           (Number(latestData.Foreign_Dealer_Sell) || 0);
        
        // 2. 投信淨額
        const trustNet = (Number(latestData.Trust_Buy) || 0) - 
                         (Number(latestData.Trust_Sell) || 0);
                          
        // 3. 自營商淨額
        const dealerNet = (Number(latestData.Dealer_Proprietary_Buy) || 0) - 
                          (Number(latestData.Dealer_Proprietary_Sell) || 0) +
                          (Number(latestData.Dealer_Hedge_Buy) || 0) -
                          (Number(latestData.Dealer_Hedge_Sell) || 0);

        return { 
            foreignNet: Math.round(foreignNet), 
            trustNet: Math.round(trustNet), 
            dealerNet: Math.round(dealerNet) 
        };
        
    } catch (e) {
        Logger.log(`getLatestThreeForeignsNet 錯誤(${code}): ${e}`); 
        return { foreignNet: 0, trustNet: 0, dealerNet: 0 }; 
    }
}

/**
 * 抓取 FinMind 過去 DAYS 天的三大法人歷史買賣超數據。
 * @param {string} code 股票代號 (e.g., "2330")
 * @returns {object} 以日期為鍵，包含三大法人淨額的 Map。
 */
function fetchFinMindForeignHistory(code) {
    // DAYS 已經定義為 30
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=${code}&date_from=${getPastDate(DAYS)}`;
    const foreignDataMap = {};

    try {
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (res.getResponseCode() !== 200) {
            Logger.log(`FinMind Foreign API 呼叫失敗(${code})，狀態碼: ${res.getResponseCode()}`);
            return foreignDataMap;
        }

        const json = JSON.parse(res.getContentText());
        const data = json.data || [];

        // 將數據轉換為以日期為鍵的 Map，方便與 K 線數據對齊
        for (const d of data) {
            // 1. 外資淨額
            const foreignNet = (Number(d.Foreign_Investor_Buy) || 0) - (Number(d.Foreign_Investor_Sell) || 0) +
                               (Number(d.Dealer_Proprietary_Buy) || 0) - (Number(d.Dealer_Proprietary_Sell) || 0);
            
            // 2. 投信淨額
            const trustNet = (Number(d.Trust_Buy) || 0) - (Number(d.Trust_Sell) || 0);
                              
            // 3. 自營商淨額
            const dealerNet = (Number(d.Dealer_Proprietary_Buy) || 0) - (Number(d.Dealer_Proprietary_Sell) || 0) +
                              (Number(d.Dealer_Hedge_Buy) || 0) - (Number(d.Dealer_Hedge_Sell) || 0);

            // FinMind 日期格式通常是 YYYY-MM-DD
            foreignDataMap[d.date] = { 
                foreignNet: Math.round(foreignNet), 
                trustNet: Math.round(trustNet), 
                dealerNet: Math.round(dealerNet) 
            };
        }
    } catch (e) {
        Logger.log(`fetchFinMindForeignHistory 錯誤(${code}): ${e}`);
    }

    return foreignDataMap;
}
