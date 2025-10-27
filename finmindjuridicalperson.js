// =================================================================
// 台灣股市三大法人買賣超數據抓取 (FinMind API v4)
// =================================================================

// ⚠️ 設定區：請務必填寫您的 FinMind API 金鑰
// 如果您沒有金鑰，請留空 ("")，但數據穩定性會降低。
const FINMIND_TOKEN = getMyUserToken();

// 🎯 設定區：要查詢的股票代號 (上市 TWSE) - 此為首次運行或清單為空時的預設值
const STOCK_CODES_DEFAULT = ["2330"];

// 🎯 設定區：查詢過去多少天的數據
const DAYS_TO_FETCH = 5; 

//get token from 外部
function getMyUserToken() {
  // 取得使用者屬性服務實例
  const properties = PropertiesService.getUserProperties();
  const userToken = properties.getProperty('FINMIND_API_TOKEN');
  return userToken;
}

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

/**
 * 抓取所有台灣股票的代碼和名稱對應表 (使用 TaiwanStockInfo)。
 * @returns {object} { 'code': 'name', ... }
 */
function fetchStockNamesMap() {
    Logger.log("[查詢] 正在從 FinMind 抓取所有股票名稱資訊...");
    // 抓取所有股票資訊，用於建立代碼/名稱對應表
    const url = "https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo";

    const options = {
        'method': 'get',
        'muteHttpExceptions': true
    };

    if (FINMIND_TOKEN) {
        options.headers = {
            "Authorization": "Bearer " + FINMIND_TOKEN
        };
    }
    
    try {
        const res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() !== 200) {
            Logger.log(`[失敗] FinMind StockInfo API 呼叫失敗，狀態碼: ${res.getResponseCode()}`);
            return {};
        }

        const json = JSON.parse(res.getContentText());
        if (json.status !== 200) {
             Logger.log(`[失敗] FinMind StockInfo API 回傳非 200 狀態，訊息: ${json.msg}`);
             return {};
        }

        const nameMap = {};
        if (json.data) {
            json.data.forEach(item => {
                // item 結構預期有 stock_id 和 stock_name
                if (item.stock_id && item.stock_name) {
                    nameMap[item.stock_id] = item.stock_name;
                }
            });
        }
        Logger.log(`[成功] 總共抓取到 ${Object.keys(nameMap).length} 筆股票名稱資料。`);
        return nameMap;
    } catch (e) {
        Logger.log(`[錯誤] 抓取股票名稱時發生例外: ${e.message}`);
        return {};
    }
}

/**
 * 依照 A 欄的股票代碼，自動填寫 B 欄的股票名稱。
 */
function updateStockNames() {
    Logger.log("--- 程式開始執行：更新股票名稱 ---");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheetName = "個股法人買賣超總計";
    let summarySheet = ss.getSheetByName(summarySheetName);

    if (!summarySheet) {
        Browser.msgBox(`錯誤：找不到分頁名稱 "${summarySheetName}"。請確認分頁已存在。`);
        return;
    }

    // 1. 抓取所有股票名稱對應表
    const stockNameMap = fetchStockNamesMap();
    if (Object.keys(stockNameMap).length === 0) {
        Browser.msgBox("警告：未能從 FinMind API 獲取股票名稱資料，請檢查 API 金鑰或網路連線。");
        return;
    }

    const lastDataRow = summarySheet.getLastRow();
    if (lastDataRow < 2) {
        Logger.log("總結分頁沒有資料行 (從第 2 行開始)，無須更新名稱。");
        Browser.msgBox("總結分頁沒有股票代碼，請先在 A 欄填寫代碼。");
        return;
    }

    // 2. 讀取 A 欄 (代碼) 和 B 欄 (名稱) 的現有數據
    // 範圍從 A2 開始，高度為 lastDataRow - 1，寬度為 2 (A欄和B欄)
    const dataRange = summarySheet.getRange(2, 1, lastDataRow - 1, 2);
    const values = dataRange.getValues();
    let updatedCount = 0;

    // 3. 遍歷並更新名稱
    values.forEach((row, index) => {
        const code = String(row[0]).trim();
        const currentName = String(row[1]).trim();
        const fetchedName = stockNameMap[code];

        if (code && fetchedName) {
            // 只有當現有名稱是空的，或現有名稱不等於抓取到的名稱時才更新
            if (!currentName || currentName !== fetchedName) {
                values[index][1] = fetchedName; // 更新名稱
                updatedCount++;
            }
        }
    });

    // 4. 一次性寫回更新後的名稱
    if (updatedCount > 0) {
        dataRange.setValues(values);
        Browser.msgBox(`股票名稱更新完成！共更新 ${updatedCount} 筆名稱。`);
    } else {
        Browser.msgBox("股票名稱無需更新，所有代碼的名稱均已正確填寫或未找到匹配名稱。");
    }

    Logger.log("--- 程式執行結束：更新股票名稱 ---");
}

/**
 * 計算最近連續買超天數。
 * @param {object[]} dailyNets 每日淨額數據陣列，已按日期排序 (最新在後)。
 * @param {string} type 法人類型鍵 ('foreign', 'trust', or 'dealer').
 * @returns {number} 連續買超天數。
 */
function calculateConsecutiveBuyDays(dailyNets, type) {
    let consecutiveDays = 0;

    // 將陣列按日期倒序排列 (最新的日期在前)，確保從最近的交易日開始計算
    dailyNets.sort((a, b) => new Date(b.date) - new Date(a.date));

    for (const day of dailyNets) {
        const netValue = day[type];
        
        // 買超 (Net > 0)，則連續天數增加
        if (netValue > 0) {
            consecutiveDays++;
        } 
        // 賣超或持平 (Net <= 0)，則中斷連續買超，停止計算
        else {
            break; 
        }
    }
    return consecutiveDays;
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

/**
 * 讀取「個股法人買賣超總計」分頁 A 欄中的股票代碼清單。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet 總結分頁物件。
 * @param {string[]} defaultCodes 預設的股票代碼清單。
 * @returns {string[]} 股票代碼陣列。
 */
function getStockCodesFromSheet(summarySheet, defaultCodes) {
    const lastRow = summarySheet.getLastRow();
    
    // 如果只有標頭或分頁是空的，則先寫入預設代碼，並使用預設代碼
    if (lastRow < 2 || summarySheet.getRange(2, 1).isBlank()) {
        Logger.log(`[代碼清單] 總結分頁為空或無代碼，寫入並使用預設代碼清單。`);
        const initialCodes = defaultCodes.map(code => [code]);
        if (initialCodes.length > 0) {
            // 從 A2 開始寫入預設代碼
            summarySheet.getRange(2, 1, initialCodes.length, 1).setValues(initialCodes); 
        }
        return defaultCodes;
    }

    // 讀取 A2 到 A[LastRow] 的範圍
    const codeRange = summarySheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
    
    const codes = [];
    
    codeRange.forEach(row => {
        const code = String(row[0]).trim(); // 確保是字串並去除空白
        // 檢查是否為非空且是數字組成的字串
        if (code && code.length >= 4 && !isNaN(code)) { 
            codes.push(code);
        }
    });

    if (codes.length === 0) {
         Logger.log(`[代碼清單] 從總結頁面讀取到 0 個有效的股票代碼，使用預設代碼清單。`);
         return defaultCodes;
    }
    
    Logger.log(`[代碼清單] 從總結頁面讀取到 ${codes.length} 個股票代碼。`);
    return codes;
}

// =================================================================
// 主控函數：將結果寫入 Google Sheet (新增連續買超邏輯)
// =================================================================

function runAllCode(){
    Logger.log("------runnning code start")
    updateInstitutionalDataSheet();

    Logger.log("--- 執行 runCode 函數結束 ---");
}

/**
 * 主要執行函數：抓取多檔股票的法人買賣超數據，並寫入 Google Sheet。
 */
function updateInstitutionalDataSheet() {
    Logger.log("--- 程式開始執行，新增數據彙總與連續買超邏輯 (動態讀取代碼) ---");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- 1. 確保並獲取總結分頁，並從中讀取股票代碼清單 ---
    const summarySheetName = "個股法人買賣超總計";
    let summarySheet = ss.getSheetByName(summarySheetName);

    if (!summarySheet) {
        // 如果分頁不存在，則先創建它
        summarySheet = ss.insertSheet(summarySheetName);
    }
    
    // 🎯 確保總結分頁有正確的標頭
    const summaryHeader = [
        "代碼", "名稱", 
        "外資總買賣超", "投信總買賣超", "自營商總買賣超", "三大法人總買賣超合計",
        "外資連續買超天數", "投信連續買超天數", "自營商連續買超天數"
    ];
    // 如果 A1 不是正確的標頭，則清除並設置標頭
    if (summarySheet.getRange('A1').getValue() !== summaryHeader[0]) {
         summarySheet.clear();
         summarySheet.appendRow(summaryHeader);
    }

    // 🎯 動態讀取股票代碼清單 (這是新的 STOCK_CODES 來源)
    const CODES_TO_FETCH = getStockCodesFromSheet(summarySheet, STOCK_CODES_DEFAULT); 

    // --- 2. 處理歷史紀錄分頁 ---
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
    
    // 🎯 總結數據結構：用於累積淨額 *並* 儲存每日淨額 (為計算連續天數準備)
    const summaryMap = {}; 

    for (const code of CODES_TO_FETCH) {
        Logger.log(`---> 正在抓取 ${code} 的法人原始數據...`);
        const rawData = fetchFinMindInstitutionalHistory(code);

        // 初始化該股票的總計數據
        summaryMap[code] = {
            foreignNet: 0,
            trustNet: 0,
            dealerNet: 0,
            totalNet: 0,
            dailyNets: [] // 🎯 新增：儲存每日淨額，用於計算連續天數
        };

        if (!rawData || rawData.length === 0) {
            Logger.log(`[注意] ${code} 未抓到數據或數據為空。`);
            continue;
        }

        // --- 核心彙總邏輯：將 Long Format 轉為 Wide Format (每日一筆) ---
        
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
        
        // Step 2: Convert grouped data to rows, calculate daily net totals, and accumulate for summary
        for (const dateKey in dailyData) {
            const d = dailyData[dateKey];
            
            // a. 計算自營商總計 (自行買賣 + 避險)
            const dealerBuy = d.dealerSelf.buy + d.dealerHedging.buy;
            const dealerSell = d.dealerSelf.sell + d.dealerHedging.sell;
            const dealerNet = dealerBuy - dealerSell;
            
            // b. 計算其他法人淨額
            const foreignNet = d.foreign.buy - d.foreign.sell;
            const trustNet = d.trust.buy - d.trust.sell;
            
            // c. 計算三大法人總淨額
            const totalNet = foreignNet + trustNet + dealerNet;

            // 🎯 累積到總結數據
            summaryMap[code].foreignNet += foreignNet;
            summaryMap[code].trustNet += trustNet;
            summaryMap[code].dealerNet += dealerNet;
            summaryMap[code].totalNet += totalNet;

            // 🎯 儲存每日淨額，用於計算連續天數
            summaryMap[code].dailyNets.push({
                date: d.date,
                foreign: foreignNet,
                trust: trustNet,
                dealer: dealerNet
            });


            // d. 加入歷史紀錄行
            allRows.push([
                d.date,         // 0: 日期
                d.stock_id,     // 1: 證券代碼
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
        // 🎯 排序邏輯：主要依證券代碼 (索引 1)，次要依日期 (索引 0)
        allRows.sort((a, b) => {
            // Primary sort: stock_id (index 1)
            const codeA = a[1];
            const codeB = b[1];
            if (codeA < codeB) return -1;
            if (codeA > codeB) return 1;

            // Secondary sort: date (index 0)
            return new Date(a[0]) - new Date(b[0]);
        });
        
        historicalSheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
        Logger.log(`歷史紀錄更新完成！共寫入 ${allRows.length} 筆紀錄，已依代碼和日期排序。`);
    } else {
        Logger.log("未抓取到任何有效的法人買賣超歷史數據。");
    }

    // --- 4. 處理總結分頁 (動態寫入結果) ---
    
    const summaryRows = [];
    
    // 儲存現有的名稱欄位數據 (B欄)，以便在覆寫時保留
    // 從 A2 開始讀取整個範圍
    const existingRange = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, summarySheet.getLastColumn()).getValues();
    const existingNames = {};
    
    // 將現有的名稱 (Column B) 映射到代碼 (Column A)
    existingRange.forEach(row => {
        const code = String(row[0]).trim();
        const name = String(row[1]).trim();
        if (code && name) {
            existingNames[code] = name;
        }
    });

    for (const code of CODES_TO_FETCH) {
        const data = summaryMap[code];
        
        // 🎯 計算連續買超天數
        const foreignStreak = calculateConsecutiveBuyDays(data.dailyNets, 'foreign');
        const trustStreak = calculateConsecutiveBuyDays(data.dailyNets, 'trust');
        const dealerStreak = calculateConsecutiveBuyDays(data.dailyNets, 'dealer');

        summaryRows.push([
            code, // 代碼 (A 欄)
            existingNames[code] || "", // 名稱 (保留現有名稱，否則留空) (B 欄)
            data.foreignNet, // 外資總買賣超
            data.trustNet, // 投信總買賣超
            data.dealerNet, // 自營商總買賣超
            data.totalNet,  // 總買賣超合計
            foreignStreak, // 外資連續買超天數 (Column G)
            trustStreak,   // 投信連續買超天數 (Column H)
            dealerStreak   // 自營商連續買超天數 (Column I)
        ]);
    }
    
    // 寫入總結數據 (先清除舊的數據範圍，再寫入新的)
    if (summaryRows.length > 0) {
        // 清除 A2 開始的所有舊數據 (只清資料，不影響標頭)
        summarySheet.getRange(2, 1, summarySheet.getMaxRows() - 1, summarySheet.getMaxColumns()).clearContent();
        
        const dataRange = summarySheet.getRange(2, 1, summaryRows.length, summaryRows[0].length);
        dataRange.setValues(summaryRows);
        Logger.log(`總計分頁更新完成！共寫入 ${summaryRows.length} 筆總結。`);

        // 🎯 重新應用：連續買超天數的條件格式設定
        const streakRange = summarySheet.getRange(2, 7, summaryRows.length, 3); 
        const greenColor = "#b6d7a8"; // 淺綠色

        // 1. 建立條件格式規則：數值 >= 2 則標註為綠色
        const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThanOrEqualTo(2)
            .setBackground(greenColor)
            .setRanges([streakRange])
            .build();

        // 2. 清除舊規則並應用新規則
        summarySheet.setConditionalFormatRules([rule]);
        
        Logger.log(`[格式] 已為連續買超天數欄位 (G:I) 設定條件格式 (>= 2 為綠色)。`);

    } else {
        Logger.log("未抓取到任何有效的法人買賣超總結數據。");
    }

    // 🎯 關鍵新增行：在數據更新後，執行名稱填充
    // =================================================================
    updateStockNames();


    Browser.msgBox(`數據更新完成！已更新歷史紀錄 (${historicalSheetName}) 和總結 (${summarySheetName}) 兩個分頁。\n\n股票代碼已從「個股法人買賣超總計」分頁的 A 欄讀取。`);
    Logger.log("--- 程式執行結束 ---");
}