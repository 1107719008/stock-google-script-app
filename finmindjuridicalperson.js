// =================================================================
// 台灣股市三大法人買賣超數據抓取 (FinMind API v4)
// =================================================================

// ⚠️ 設定區：請務必填寫您的 FinMind API 金鑰
// 如果您沒有金鑰，請留空 ("")，但數據穩定性會降低。
const FINMIND_TOKEN = getMyUserToken();

// 🎯 設定區：要查詢的股票代號 (上市 TWSE) - 此為首次運行或清單為空時的預設值
const STOCK_CODES_DEFAULT = ["2330"];

// 🎯 設定區：查詢過去多少天的數據
const DAYS_TO_FETCH = 7; 

// 🎯 新增設定區：用於動量分析的買超比例
const BUY_RATIO = 1.2;

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
 * 應用動量分析條件格式：如果三大法人總買超量 >= 昨天總買超量 * BUY_RATIO (且兩日均為買超)
 */
function applyMomentumFormatting() {
    Logger.log("--- 程式開始執行：動量分析與格式化 ---");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheetName = "個股法人買賣超總計";
    const historicalSheetName = "TWSE 法人買賣超歷史紀錄";
    const summarySheet = ss.getSheetByName(summarySheetName);
    const historicalSheet = ss.getSheetByName(historicalSheetName);

    if (!summarySheet || !historicalSheet) {
        Logger.log("錯誤：找不到總結或歷史紀錄分頁。無法進行動量分析。");
        return;
    }

    // 1. 獲取歷史數據 (Column L 是三大法人淨買賣超合計)
    const lastHistoricalRow = historicalSheet.getLastRow();
    if (lastHistoricalRow < 2) {
        Logger.log("歷史數據為空，跳過動量格式化。");
        return;
    }
    // 範圍從 A2 開始，共 12 欄 (到 L 欄)
    const historicalData = historicalSheet.getRange(2, 1, lastHistoricalRow - 1, 12).getValues();

    // 2. 按股票代碼分組，並找出最新的兩筆淨買賣超數據
    const momentumCodes = [];
    const groupedData = {};

    historicalData.forEach(row => {
        const code = String(row[1]).trim();
        const date = new Date(row[0]);
        const totalNet = row[11]; // 三大法人淨買賣超合計

        if (!groupedData[code]) {
            groupedData[code] = [];
        }
        
        // 儲存為物件 {date: Date, net: number}
        groupedData[code].push({ date: date, net: totalNet });
    });

    // 3. 遍歷並分析動量
    for (const code in groupedData) {
        const dailyRecords = groupedData[code];
        
        // 按日期降序排列 (最新的在前)
        dailyRecords.sort((a, b) => b.date.getTime() - a.date.getTime());

        // 需要至少 2 個交易日
        if (dailyRecords.length < 2) continue;

        const todayNet = dailyRecords[0].net;
        const yesterdayNet = dailyRecords[1].net;

        // 條件檢查：
        // 1. 今天必須是買超 ( > 0)
        // 2. 昨天也必須是買超 ( > 0)
        // 3. 今天買超量 >= 昨天買超量 * BUY_RATIO
        if (todayNet > 0 && yesterdayNet > 0) {
            if (todayNet >= yesterdayNet * BUY_RATIO) {
                momentumCodes.push(code);
            }
        }
    }

    // 4. 應用條件格式到總結分頁 (A 欄)
    const lastSummaryRow = summarySheet.getLastRow();
    if (lastSummaryRow < 2) {
        Logger.log("總結分頁無數據，無法應用動量格式。");
        return;
    }
    
    // 獲取現有規則
    let rules = summarySheet.getConditionalFormatRules();
    
    // 過濾掉任何舊的動量紫色規則 (A 欄)
    rules = rules.filter(rule => {
        const ranges = rule.getRanges();
        // 如果規則是作用於 A 欄 (Column 1) 且不是連買天數的綠色規則，則移除
        if (ranges.length === 1 && ranges[0].getColumn() === 1) {
            return false; 
        }
        return true; 
    });

    const purpleColor = "#8e7cc3"; // 紫色
    const currentCodes = summarySheet.getRange(2, 1, lastSummaryRow - 1, 1).getValues().flat().map(String).map(s => s.trim());
    const rangesToHighlight = [];
    
    // 找出所有需要標註紫色的儲存格範圍
    currentCodes.forEach((code, index) => {
        if (momentumCodes.includes(code)) {
            // index 是 0-based，儲存格行號是 index + 2
            rangesToHighlight.push(summarySheet.getRange(index + 2, 1));
        }
    });

    if (rangesToHighlight.length > 0) {
        // 創建並加入新的動量規則
        const momentumRule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied("=TRUE") // 因為範圍是手動選的，所以使用 TRUE
            .setBackground(purpleColor)
            .setRanges(rangesToHighlight)
            .build();
        
        rules.push(momentumRule);
        Logger.log(`[格式] 已為 ${momentumCodes.length} 支股票代碼 (A 欄) 設定動量紫色標註。`);
    }
    
    // 應用所有規則 (保留了其他規則，如綠色連買天數)
    summarySheet.setConditionalFormatRules(rules);
    Logger.log("--- 程式執行結束：動量格式化 ---");
}

// -----------------------------------------------------------------
// 動量分析與總結分頁顏色標註 (根據 RATIO 標註 C, D, E 欄)
// -----------------------------------------------------------------

/**
 * 應用動量分析條件格式：如果三大法人任一法人今天的買超量 >= 昨天的買超量 * BUY_RATIO (且兩日均為買超)，
 * 則標註總結分頁對應的欄位。
 */
function applyMomentumAndSummaryColoring() {
    Logger.log("--- 程式開始執行：動量分析與總結分頁顏色標註 ---");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheetName = "個股法人買賣超總計";
    const historicalSheetName = "TWSE 法人買賣超歷史紀錄";
    const summarySheet = ss.getSheetByName(summarySheetName);
    const historicalSheet = ss.getSheetByName(historicalSheetName);

    if (!summarySheet || !historicalSheet) {
        Logger.log("錯誤：找不到總結或歷史紀錄分頁。無法進行動量分析。");
        return;
    }

    // 1. 獲取歷史數據 (用於計算動量)
    const lastHistoricalRow = historicalSheet.getLastRow();
    if (lastHistoricalRow < 2) {
        Logger.log("歷史數據為空，跳過動量格式化。");
        return;
    }
    // 範圍從 A2 開始，共 12 欄 (到 L 欄)
    const historicalData = historicalSheet.getRange(2, 1, lastHistoricalRow - 1, 12).getValues();

    // 2. 按股票代碼分組，並找出最新的兩筆淨買賣超數據
    const momentumFlags = {}; // 儲存符合動量條件的法人類型 { '2330': ['foreign', 'trust'], ... }
    const groupedData = {};

    historicalData.forEach(row => {
        const code = String(row[1]).trim(); // 股票代碼
        const date = new Date(row[0]); // 日期

        // 淨買賣超數據 (E, H, K 欄)
        const foreignNet = row[4]; 
        const trustNet = row[7]; 
        const dealerNet = row[10]; 

        if (!groupedData[code]) {
            groupedData[code] = [];
        }
        
        groupedData[code].push({ 
            date: date, 
            foreign: foreignNet, 
            trust: trustNet, 
            dealer: dealerNet 
        });
    });

    // 3. 遍歷並分析個別法人動量 (使用淨買超金額和 BUY_RATIO)
    for (const code in groupedData) {
        const dailyRecords = groupedData[code];
        
        // 按日期降序排列 (最新的在前)
        dailyRecords.sort((a, b) => b.date.getTime() - a.date.getTime());

        // 需要至少 2 個交易日
        if (dailyRecords.length < 2) continue;

        const today = dailyRecords[0];
        const yesterday = dailyRecords[1];

        const triggeredInstitutions = [];
        const institutions = ['foreign', 'trust', 'dealer'];

        for (const institution of institutions) {
            const todayNet = today[institution];
            const yesterdayNet = yesterday[institution];
            
            // 條件檢查：
            // 1. 今天必須是買超 ( > 0)
            // 2. 昨天也必須是買超 ( > 0)
            // 3. 今天買超量 >= 昨天買超量 * BUY_RATIO
            if (todayNet > 0 && yesterdayNet > 0) {
                if (todayNet >= yesterdayNet * BUY_RATIO) {
                    triggeredInstitutions.push(institution);
                }
            }
        }
        
        if (triggeredInstitutions.length > 0) {
            momentumFlags[code] = triggeredInstitutions;
        }
    }

    // 4. 應用條件格式到總結分頁 (A, C, D, E 欄)
    const lastSummaryRow = summarySheet.getLastRow();
    if (lastSummaryRow < 2) return;
    
    // 定義顏色
    const purpleColor = "#8e7cc3";    // A 欄 (代碼)
    const foreignColor = "#4a86e8";   // C 欄 (外資 - 藍色)
    const trustColor = "#f1c232";     // D 欄 (投信 - 黃色)
    const dealerColor = "#ea9999";    // E 欄 (自營商 - 紅色)
    
    // 總結頁的欄位索引 (A, C, D, E)
    const momentumCols = [1, 3, 4, 5]; 
    const colMap = { 'foreign': 3, 'trust': 4, 'dealer': 5 }; // 法人對應的欄位
    const colorMap = { 'foreign': foreignColor, 'trust': trustColor, 'dealer': dealerColor };

    // 獲取現有規則
    let rules = summarySheet.getConditionalFormatRules();
    
    // 過濾掉所有舊的 A, C, D, E 欄規則，以便重新應用
    rules = rules.filter(rule => {
        const ranges = rule.getRanges();
        // 如果規則只作用於一個範圍，且該範圍是 A, C, D, 或 E 欄中的任一列，則移除
        if (ranges.length === 1) {
            const col = ranges[0].getColumn();
            if (momentumCols.includes(col)) {
                return false;
            }
        }
        // 保留非 A, C, D, E 欄的規則 (例如 G:I 的連續買超綠色規則)
        return true; 
    });


    const currentCodes = summarySheet.getRange(2, 1, lastSummaryRow - 1, 1).getValues().flat().map(String).map(s => s.trim());
    
    // 找出所有需要標註的儲存格範圍並應用規則
    currentCodes.forEach((code, index) => {
        const row = index + 2; // 資料從第 2 行開始
        const flags = momentumFlags[code]; 

        if (flags) {
            // 🎯 A Column (整體動量 - 紫色)
            rules.push(SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied("=TRUE")
                .setBackground(purpleColor)
                .setRanges([summarySheet.getRange(row, 1)]) // A 欄
                .build());

            // 🎯 C, D, E Columns (個別法人動量)
            flags.forEach(institution => {
                const col = colMap[institution];
                const color = colorMap[institution];
                
                rules.push(SpreadsheetApp.newConditionalFormatRule()
                    .whenFormulaSatisfied("=TRUE")
                    .setBackground(color)
                    .setRanges([summarySheet.getRange(row, col)])
                    .build());
            });
            
             Logger.log(`[格式] 股票代碼 ${code} 觸發了法人動量，並已設定顏色。`);
        }
    });

    // 重新應用所有規則
    summarySheet.setConditionalFormatRules(rules);
    
    Logger.log("--- 程式執行結束：動量分析與總結分頁顏色標註 ---");
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

/**
 * 🎯 新增函數：抓取 FinMind 過去 DAYS_TO_FETCH 天的股票價格數據 (開盤、收盤、高點、低點)。
 * @param {string} code 股票代號 (e.g., "2330")
 * @returns {object} { 'YYYY-MM-DD': { open: number, close: number, high: number, low: number }, ... }
 */
function fetchFinMindStockPrice(code) {
    const startDate = getPastDate(DAYS_TO_FETCH); 
    const endDate = getTodayDate(); 
    
    // 使用 TaiwanStockPrice 資料集
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=${code}&start_date=${startDate}&end_date=${endDate}`;
    
    Logger.log(`[查詢] 股價數據 - 股票代碼: ${code}, 查詢區間: ${startDate} 至 ${endDate}`);

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
            Logger.log(`[失敗] FinMind API 呼叫股價數據失敗(${code})，狀態碼: ${res.getResponseCode()}`);
            return {};
        }

        const json = JSON.parse(res.getContentText());
        
        if (json.status !== 200) {
             Logger.log(`[失敗] 股價數據 API 回傳非 200 狀態，代碼: ${json.status}，訊息: ${json.msg}`);
             return {};
        }

        const priceMap = {};
        if (json.data) {
            json.data.forEach(item => {
                // FinMind 使用 max/min
                if (item.date && item.open && item.max && item.min && item.close) {
                    priceMap[item.date] = {
                        open: item.open,
                        close: item.close,
                        high: item.max, 
                        low: item.min,  
                    };
                }
            });
        }
        return priceMap;
        
    } catch (e) {
        Logger.log(`[錯誤] 抓取 ${code} 股價數據時發生例外: ${e.message}`);
        return {};
    }
}

// -----------------------------------------------------------------
// 歷史紀錄標頭更新函數 (M-P 欄位)
// -----------------------------------------------------------------

/**
 * 🎯 確保歷史紀錄分頁的標頭有 M, N, O, P 欄位。
 */
function updateHistoricalHeaderColumns() {
    Logger.log("--- 程式開始執行：更新歷史紀錄分頁標頭 ---");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historicalSheet = ss.getSheetByName("TWSE 法人買賣超歷史紀錄");

    if (!historicalSheet) {
        Logger.log("錯誤：找不到歷史紀錄分頁。");
        return;
    }

    const priceHeaderMtoP = [ "開盤價", "收盤價", "最高價", "最低價" ];
    
    // 檢查 M1 欄位是否為 "開盤價" (即檢查 M-P 是否已存在)
    const m1Value = historicalSheet.getRange('M1').getValue();

    if (m1Value !== priceHeaderMtoP[0]) {
        Logger.log("偵測到 M-P 欄位標頭缺失或不匹配，正在更新...");
        
        // 假設 A-L 欄已經存在 (共 12 欄)，從第 13 欄 (M) 開始寫入
        const startCol = 13; 
        const startRow = 1;
        
        historicalSheet.getRange(startRow, startCol, 1, priceHeaderMtoP.length).setValues([priceHeaderMtoP]);
        Logger.log("歷史紀錄分頁 M-P 欄位標頭已更新。");
    } else {
        Logger.log("歷史紀錄分頁 M-P 欄位標頭已存在，無須更新。");
    }
    
    Logger.log("--- 程式執行結束：更新歷史紀錄分頁標頭 ---");
}

/**
 * 🎯 新增函數：獨立抓取歷史股價數據並更新 M, N, O, P 欄位。
 * 此函數不影響 A-L 的法人數據。
 */
function updateHistoricalPrices() {
    Logger.log("--- 程式開始執行：更新歷史紀錄分頁價格數據 (M-P) ---");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historicalSheet = ss.getSheetByName("TWSE 法人買賣超歷史紀錄");
    const summarySheet = ss.getSheetByName("個股法人買賣超總計"); // 需要此頁面來獲取代碼清單

    if (!historicalSheet || !summarySheet) {
        Logger.log("錯誤：找不到總結或歷史紀錄分頁。無法更新價格數據。");
        Browser.msgBox("錯誤：請確保「個股法人買賣超總計」與「TWSE 法人買賣超歷史紀錄」分頁存在。");
        return;
    }
    
    // 1. 確保 M-P 欄位標頭存在
    updateHistoricalHeaderColumns();

    // 2. 獲取所有需要查詢的股票代碼
    const CODES_TO_FETCH = getStockCodesFromSheet(summarySheet, STOCK_CODES_DEFAULT); 
    
    // 3. 抓取所有股票的價格數據
    const allPriceData = {}; // { '2330': { '2025-01-01': { open: 1, close: 2, ... } } }
    
    for (const code of CODES_TO_FETCH) {
        Logger.log(`---> 正在抓取 ${code} 的股價數據...`);
        allPriceData[code] = fetchFinMindStockPrice(code);
        Utilities.sleep(500); // 避免 API 頻率限制
    }
    
    // 4. 讀取歷史紀錄分頁的資料 (只需要 A, B 欄和 M-P 欄位)
    const lastHistoricalRow = historicalSheet.getLastRow();
    if (lastHistoricalRow < 2) {
        Logger.log("歷史數據為空，無法更新價格。");
        return;
    }
    
    // 讀取 A:B (日期, 代碼) 和 M:P (價格) 欄位的數據
    // 讀取範圍為 16 欄，確保讀取到 M-P 欄位
    const historicalData = historicalSheet.getRange(2, 1, lastHistoricalRow - 1, 16).getValues(); 

    let updateCount = 0;
    
    // 5. 遍歷歷史數據，將價格數據覆蓋到 M, N, O, P 欄位
    historicalData.forEach(row => {
        // A 欄 (索引 0): 日期; B 欄 (索引 1): 股票代碼
        const dateKey = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        const code = String(row[1]).trim();
        
        const stockPrices = allPriceData[code];
        
        if (stockPrices) {
            const price = stockPrices[dateKey];
            
            if (price) {
                // M, N, O, P 欄位的索引分別是 12, 13, 14, 15
                // 檢查 M 欄 (索引 12) 是否已更新，避免重複寫入
                if (row[12] === "" || row[12] === null || row[12] !== price.open) { 
                    row[12] = price.open;  // M: 開盤價
                    row[13] = price.close; // N: 收盤價
                    row[14] = price.high;  // O: 最高價
                    row[15] = price.low;   // P: 最低價
                    updateCount++;
                }
            } else {
                 // 該交易日沒有價格數據，填入空值（如果之前有價格，則清除）
                 row[12] = ""; 
                 row[13] = ""; 
                 row[14] = ""; 
                 row[15] = ""; 
            }
        }
    });

    // 6. 一次性寫回更新後的價格數據 (從 A2 開始，寫入 16 欄)
    if (historicalData.length > 0) {
        // 寫入範圍為從 A2 到 P[LastRow]
        historicalSheet.getRange(2, 1, historicalData.length, historicalData[0].length).setValues(historicalData);
        Logger.log(`價格數據更新完成！共更新/檢查 ${updateCount} 筆交易日的價格資訊。`);
    }

    Browser.msgBox("歷史股價數據 (M-P 欄位) 更新完成！");
    Logger.log("--- 程式執行結束：更新歷史紀錄分頁價格數據 ---");
}


// =================================================================
// 主控函數：將結果寫入 Google Sheet (新增連續買超邏輯)
// =================================================================

function runAllCode(){
    Logger.log("------runnning code start")
    updateInstitutionalDataSheet();

    updateHistoricalPrices();

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

    // 執行動量分析和格式化
    // =================================================================
    applyMomentumFormatting();
    // =================================================================
    applyMomentumAndSummaryColoring();


    Browser.msgBox(`數據更新完成！已更新歷史紀錄 (${historicalSheetName}) 和總結 (${summarySheetName}) 兩個分頁。\n\n股票代碼已從「個股法人買賣超總計」分頁的 A 欄讀取。`);
    Logger.log("--- 程式執行結束 ---");
}