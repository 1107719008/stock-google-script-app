// =================================================================
// 台灣股市三大法人買賣超數據抓取 (FinMind API v4)
// =================================================================

// ⚠️ 設定區：請務必填寫您的 FinMind API 金鑰
// 如果您沒有金鑰，請留空 ("")，但數據穩定性會降低。
const FINMIND_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJkYXRlIjoiMjAyNS0xMC0yNyAxNzo1OToyNSIsInVzZXJfaWQiOiJEZWxhc2siLCJpcCI6IjYxLjIyOC43Ni4yMzUifQ.vgVvInO6tK2YdLSizK4ZK7w2VaXz8sZKgUCGlRDQv9k"; 

// 🎯 設定區：要查詢的股票代號 (上市 TWSE)
const STOCK_CODES = ["2330"];

// 🎯 設定區：查詢過去多少天的數據
// 根據您的要求，我們將起始日期設為從現在倒數 5 天前。
// 注意：由於 FinMind API 只回傳實際交易日數據，所以結果會自動跳過週末和國定假日。
const DAYS_TO_FETCH = 5; 

// =================================================================
// 輔助函數
// =================================================================

/**
 * 獲取 N 天前的日期，格式為 YYYY-MM-DD。
 * @param {number} days 過去的天數。
 * @returns {string} 日期字串。
 */
function getPastDate(days) {
    const date = new Date();
    date.setDate(date.getDate() - days); 
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd"); 
}

/**
 * 獲取今天的日期，格式為 YYYY-MM-DD。
 * @returns {string} 日期字串。
 */
function getTodayDate() {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); 
}


// =================================================================
// 核心 API 函數 (無變動，僅獲取原始數據)
// =================================================================

/**
 * 抓取 FinMind 過去 DAYS_TO_FETCH 天的三大法人歷史買賣超數據。
 * @param {string} code 股票代號 (e.g., "2330")
 * @returns {object[]} FinMind 回傳的數據陣列，或 null (如果失敗)。
 */
function fetchFinMindInstitutionalHistory(code) {
    const startDate = getPastDate(DAYS_TO_FETCH); 
    const endDate = getTodayDate(); 
    
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=${code}&start_date=${startDate}&end_date=${endDate}`;
    
    Logger.log(`[查詢] 股票代碼: ${code}, 查詢區間: ${startDate} 至 ${endDate}`);

    const options = {
        'method': 'get',
        'muteHttpExceptions': true
    };

    if (FINMIND_TOKEN) {
        options.headers = {
            "Authorization": "Bearer " + FINMIND_TOKEN
        };
        Logger.log(`[Token] 正在使用 API Token 進行授權查詢。`);
    } else {
        Logger.log(`[Token] 未設置 API Token，可能因限流導致數據為 0 或查詢失敗。`);
    }
    
    try {
        const res = UrlFetchApp.fetch(url, options);
        
        if (res.getResponseCode() !== 200) {
            Logger.log(`[失敗] FinMind API 呼叫失敗(${code})，狀態碼: ${res.getResponseCode()}，回應: ${res.getContentText()}`);
            return null;
        }

        const json = JSON.parse(res.getContentText());
        
        if (json.status !== 200) {
             Logger.log(`[失敗] FinMind API 回傳非 200 狀態，代碼: ${json.status}，訊息: ${json.msg}`);
             return null;
        }

        return json.data || [];
        
    } catch (e) {
        Logger.log(`[錯誤] 抓取 ${code} 時發生例外: ${e.message}`);
        return null;
    }
}

// =================================================================
// 主控函數：將結果寫入 Google Sheet (新增數據彙總邏輯)
// =================================================================

/**
 * 主要執行函數：抓取多檔股票的法人買賣超數據，並寫入 Google Sheet。
 */
function updateInstitutionalDataSheet() {
    Logger.log("--- 程式開始執行，新增數據彙總邏輯 ---");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- 1. 處理歷史紀錄分頁 (舊邏輯) ---
    const historicalSheetName = "TWSE 法人買賣超歷史紀錄";
    let historicalSheet = ss.getSheetByName(historicalSheetName);
    
    if (!historicalSheet) {
        historicalSheet = ss.insertSheet(historicalSheetName);
    } else {
        historicalSheet.clearContents();
    }
    
    const historicalHeader = [
        "日期", "證券代碼", "外資買進張數", "外資賣出張數", "外資淨買賣超",
        "投信買進張數", "投信賣出張數", "投信淨買賣超",
        "自營商買進張數", "自營商賣出張數", "自營商淨買賣超",
        "三大法人淨買賣超合計"
    ];
    historicalSheet.appendRow(historicalHeader);
    
    const allRows = [];
    // 🎯 總結數據結構：用於累積每檔股票的總買賣超
    const summaryMap = {}; 

    for (const code of STOCK_CODES) {
        Logger.log(`---> 正在抓取 ${code} 的法人原始數據...`);
        const rawData = fetchFinMindInstitutionalHistory(code);

        // 初始化該股票的總計數據
        summaryMap[code] = {
            foreignNet: 0,
            trustNet: 0,
            dealerNet: 0,
            totalNet: 0
        };

        if (!rawData || rawData.length === 0) {
            Logger.log(`[注意] ${code} 未抓到數據或數據為空。`);
            continue;
        }

        // --- 核心彙總邏輯：將 Long Format 轉為 Wide Format (每日一筆) ---
        
        // dailyData 的 Key 為日期字串 (e.g., '2025-10-27')
        const dailyData = {}; 

        // Step 1: Group by date and institution type
        for (const record of rawData) {
            const key = record.date;
            
            if (!dailyData[key]) {
                dailyData[key] = {
                    date: record.date,
                    stock_id: record.stock_id,
                    foreign: { buy: 0, sell: 0 },
                    trust: { buy: 0, sell: 0 },
                    dealerSelf: { buy: 0, sell: 0 },
                    dealerHedging: { buy: 0, sell: 0 },
                };
            }
            
            // 根據 FinMind 的 'name' 欄位，將買賣超數據加總
            switch (record.name) {
                case 'Foreign_Investor':
                    dailyData[key].foreign.buy += record.buy;
                    dailyData[key].foreign.sell += record.sell;
                    break;
                case 'Investment_Trust':
                    dailyData[key].trust.buy += record.buy;
                    dailyData[key].trust.sell += record.sell;
                    break;
                case 'Dealer_self':
                    dailyData[key].dealerSelf.buy += record.buy;
                    dailyData[key].dealerSelf.sell += record.sell;
                    break;
                case 'Dealer_Hedging':
                    dailyData[key].dealerHedging.buy += record.buy;
                    dailyData[key].dealerHedging.sell += record.sell;
                    break;
                default:
                    break;
            }
        }
        
        // Step 2: Convert grouped data to rows and calculate daily net totals
        for (const dateKey in dailyData) {
            const d = dailyData[dateKey];
            
            // a. 計算自營商總計 (自營商自行買賣 + 避險)
            const dealerBuy = d.dealerSelf.buy + d.dealerHedging.buy;
            const dealerSell = d.dealerSelf.sell + d.dealerHedging.sell;
            const dealerNet = dealerBuy - dealerSell;
            
            // b. 計算其他法人淨額
            const foreignNet = d.foreign.buy - d.foreign.sell;
            const trustNet = d.trust.buy - d.trust.sell;
            
            // c. 計算三大法人總淨額
            const totalNet = foreignNet + trustNet + dealerNet;

            // 🎯 累積到總結數據 (實現加總需求)
            summaryMap[code].foreignNet += foreignNet;
            summaryMap[code].trustNet += trustNet;
            summaryMap[code].dealerNet += dealerNet;
            summaryMap[code].totalNet += totalNet;


            // d. 加入歷史紀錄行
            allRows.push([
                d.date, 
                d.stock_id, 
                d.foreign.buy, 
                d.foreign.sell, 
                foreignNet,
                d.trust.buy, 
                d.trust.sell, 
                trustNet,
                dealerBuy,
                dealerSell,
                dealerNet,
                totalNet
            ]);
        }
        
        Utilities.sleep(500); 
    }
    
    // 3. 一次性寫入所有歷史數據
    if (allRows.length > 0) {
        allRows.sort((a, b) => new Date(a[0]) - new Date(b[0]));
        
        historicalSheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
        Logger.log(`歷史紀錄更新完成！共寫入 ${allRows.length} 筆紀錄。`);
    } else {
        Logger.log("未抓取到任何有效的法人買賣超歷史數據。");
    }

    // --- 4. 處理總結分頁 (新邏輯：個股資料總計) ---
    
    const summarySheetName = "個股法人買賣超總計";
    let summarySheet = ss.getSheetByName(summarySheetName);

    if (!summarySheet) {
        summarySheet = ss.insertSheet(summarySheetName);
    } else {
        summarySheet.clearContents();
    }
    
    // 🎯 總結分頁的標頭
    const summaryHeader = [
        "代碼", "名稱 (請手動填寫)", "外資總買賣超", "投信總買賣超", "自營商總買賣超", "三大法人總買賣超合計"
    ];
    summarySheet.appendRow(summaryHeader);
    
    const summaryRows = [];
    
    for (const code of STOCK_CODES) {
        const data = summaryMap[code];
        
        summaryRows.push([
            code, // 代碼
            "", // 名稱 (API 不提供，留空讓使用者手動填寫)
            data.foreignNet, // 外資總買賣超
            data.trustNet, // 投信總買賣超
            data.dealerNet, // 自營商總買賣超
            data.totalNet  // 總買賣超合計
        ]);
    }
    
    // 寫入總結數據
    if (summaryRows.length > 0) {
        summarySheet.getRange(2, 1, summaryRows.length, summaryRows[0].length).setValues(summaryRows);
        Logger.log(`總計分頁更新完成！共寫入 ${summaryRows.length} 筆總結。`);
    }

    Browser.msgBox(`數據更新完成！已更新歷史紀錄 (${historicalSheetName}) 和總結 (${summarySheetName}) 兩個分頁。`);
    Logger.log("--- 程式執行結束 ---");
}
