// ====================================================================
// Web App Functions - 網頁應用程式相關函式
// ====================================================================

/**
 * 當使用者打開網頁應用程式 URL 時，執行此函式。
 * @param {Object} e - 事件物件。
 * @returns {HtmlOutput} - 顯示給使用者的 HTML 頁面。
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('SettingsPage')
    .setTitle('儀器校正通知系統設定')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * [優化] 從「設定」工作表讀取所有設定，並傳送到網頁前端。
 * @returns {Object} - 包含所有設定資料的物件。
 */
function getSettingsForWebApp() {
  try {
    // 呼叫新的全域函式，並回傳給 WebApp 需要的 raw 格式
    const config = getGlobalConfig();
    return {
      globalCcEmails: config.raw.globalCcEmails,
      exclusionList: config.raw.exclusionList.map(item => item[0]), // 轉成一維陣列給前端
      notificationRules: config.raw.notificationRules,
      personnelMap: config.raw.personnelMap,
      groupMap: config.raw.groupMap,
      additionalCcMap: config.raw.additionalCcMap
    };
  } catch (e) {
    Logger.log(e);
    throw new Error('讀取設定時發生錯誤: ' + e.message);
  }
}

/**
 * 接收從網頁前端傳來的設定資料，並寫回「設定」工作表。
 * (此函式維持原樣以配合前端的儲存邏輯)
 * @param {Object} settings - 從前端傳來的設定物件。
 * @returns {String} - 成功訊息。
 */
function saveSettings(settings) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
    if (!sheet) throw new Error("找不到名為 '設定' 的工作表。");

    // 寫入通用副本
    sheet.getRange("B2").setValue(settings.globalCcEmails.join(', '));

    // 寫入各個表格
    writeDataToSheet(sheet, "D2:D", settings.exclusionList.map(item => [item]));
    writeDataToSheet(sheet, "F2:I", settings.notificationRules);
    writeDataToSheet(sheet, "K2:L", settings.personnelMap);
    writeDataToSheet(sheet, "N2:O", settings.groupMap);
    writeDataToSheet(sheet, "Q2:R", settings.additionalCcMap);

    return "所有設定已成功儲存！";
  } catch (e) {
    Logger.log(e);
    throw new Error('儲存設定時發生錯誤: ' + e.message);
  }
}

/**
 * 將資料寫入指定範圍的輔助函式。
 * (此函式維持原樣)
 */
function writeDataToSheet(sheet, rangeA1, data) {
  const startColLetter = rangeA1.match(/^[A-Z]+/)[0];
  const startColNum = sheet.getRange(startColLetter + "1").getColumn();

  const endColMatch = rangeA1.match(/[A-Z]+$/);
  const endColLetter = endColMatch ? endColMatch[0] : startColLetter;
  const endColNum = sheet.getRange(endColLetter + "1").getColumn();
  const numColumns = endColNum - startColNum + 1;

  // 先清除舊有資料範圍
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, startColNum, sheet.getLastRow() - 1, numColumns).clearContent();
  }

  // 再寫入新資料
  if (data && data.length > 0 && data[0].length > 0) {
    sheet.getRange(2, startColNum, data.length, data[0].length).setValues(data);
  }
}

// ====================================================================
// Notification System Core - 核心通知系統
// ====================================================================

/**
 * [優化] 讀取設定檔給後端通知系統使用。
 */
function getConfig() {
  try {
    // 呼叫新的全域函式，並回傳 parsed 好的設定物件
    return getGlobalConfig().parsed;
  } catch (e) {
    Logger.log(e);
    return null; // 維持原有的錯誤處理方式
  }
}

/**
 * 【最終架構版 - 核心大腦函式】
 * 全權負責所有通知判斷：計算日期、比對規則、檢查狀態、檢查排除、組合收件人與內容。
 * @param {object} instrumentInfo - 包含 id, name, custodian, group 的儀器物件。
 * @param {Date} expectedDate - 預計校正日期物件。
 * @param {string} currentStatus - 該儀器在表單上的目前狀態。
 * @param {object} config - 已解析過的完整設定物件。
 * @returns {object|null} 如果需要發送通知，則回傳包含細節的物件；否則回傳 null。
 */
function generateNotificationDetails(instrumentInfo, expectedDate, currentStatus, config) {
  const { id, name, custodian, group } = instrumentInfo;

  // 1. 計算日期差異
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffDays = Math.floor((expectedDate.getTime() - today.getTime()) / (1000 * 3600 * 24));

  // 2. 尋找規則
  const bestRule = config.notificationRules.find(rule => diffDays === rule.triggerDays);
  if (!bestRule) {
    return null; // 今天沒有符合的規則
  }

  // 3. 檢查狀態，避免重複發送
  if (currentStatus === bestRule.statusFlag) {
    return null; // 狀態已是目標狀態，不需重複發送
  }

  // 4. 檢查排除清單
  if (config.exclusionList.some(prefix => id && id.startsWith(prefix))) {
    return null; // 在排除清單中
  }

  // 5. 組合收件人
  const keeperInfo = config.personnelMap[custodian] || {};
  const toEmail = keeperInfo.email || null;
  if (!toEmail) {
    Logger.log(`在 generateNotificationDetails 中找不到保管人 '${custodian}' 的 Email。`);
    return null;
  }
  const ccEmails = new Set(config.globalCcEmails);
  const groupEmail = config.groupMap[group] || null;
  if (groupEmail) { ccEmails.add(groupEmail); }
  for (const condition in config.additionalCcMap) {
    if ((id && id.startsWith(condition)) || (group && group === condition)) {
      config.additionalCcMap[condition].forEach(email => ccEmails.add(email));
    }
  }

  // 6. 替換信件內容中的變數
  const placeholders = {
    '{年度}': today.getFullYear().toString(), '{設備編號}': id, '{設備名稱}': name,
    '{校正日期}': expectedDate.toLocaleDateString("zh-TW"), '{保管人}': custodian,
    '{稱謂}': generateSalutation(custodian), '{組別}': group,
  };
  let finalSubject = bestRule.subjectTemplate;
  let finalBody = bestRule.htmlBodyTemplate;
  for (const key in placeholders) {
    const regex = new RegExp(key.replace(/\{/g, '\\{').replace(/\}/g, '\\}'), 'g');
    finalSubject = finalSubject.replace(regex, placeholders[key]);
    finalBody = finalBody.replace(regex, placeholders[key]);
  }

  // 7. 組合並回傳最終結果
  return {
    to: toEmail, cc: [...ccEmails].join(','), subject: finalSubject,
    body: finalBody, rawBody: bestRule.htmlBodyTemplate, statusFlag: bestRule.statusFlag,
    triggerDays: bestRule.triggerDays, diffDays: diffDays // 也回傳天數供 Log 使用
  };
}

/**
 * [最終架構版] 核心通知函式，職責被大幅簡化。
 * 主要負責遍歷資料，並將判斷工作完全交給 generateNotificationDetails。
 */
function notifyKeeperCalibrationReminder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  if (!config) return;

  const deviceMap = getDeviceMapping();
  const logSheet = ss.getSheetByName("通知Log") || ss.insertSheet("通知Log");
  const currentYearSheet = ss.getSheetByName(String(new Date().getFullYear()));
  if (!currentYearSheet) {
    Logger.log("找不到今年的工作表：" + new Date().getFullYear());
    return;
  }
  const dataRange = currentYearSheet.getRange('A1:K' + currentYearSheet.getLastRow());
  const data = dataRange.getValues();
  const statuses = dataRange.getValues();
  const logEntries = [];

  for (let i = 1; i < data.length; i++) {
    // a. 準備資料
    let rowData = data[i];
    let deviceId = rowData[0], deviceName = rowData[1], keeperName = rowData[2];
    let expectedDateRaw = rowData[5], actualDate = rowData[6], status = rowData[10];
    if (actualDate) continue;
    if (!deviceId && !deviceName) { const mergeRowIndex = findMergedContent(data, i, 0); deviceId = data[mergeRowIndex][0]; deviceName = data[mergeRowIndex][1]; }
    if (!keeperName) { const keeperRowIndex = findMergedContent(data, i, 2); keeperName = data[keeperRowIndex][2]; }
    keeperName = keeperName ? String(keeperName).trim() : '';
    const expectedDate = parseDate(expectedDateRaw);
    if (!expectedDate || !deviceId || !keeperName) continue;

    const deviceInfo = deviceMap[deviceId] || {};
    const instrumentInfo = {
      id: deviceId, name: deviceName, custodian: keeperName,
      group: deviceInfo.group || "未知組別"
    };

    // b. 呼叫唯一的「大腦」函式，傳入所需的一切
    const details = generateNotificationDetails(instrumentInfo, expectedDate, status, config);

    // c. 如果大腦說要發信，就執行動作
    if (details) {
      const htmlBody = parseSimpleMarkup(details.body);
      MailApp.sendEmail({ to: details.to, cc: details.cc, subject: details.subject, htmlBody: htmlBody });

      statuses[i][10] = details.statusFlag;
      logEntries.push([
        new Date(), new Date().getFullYear().toString(), deviceId, deviceName, keeperName,
        expectedDate, details.diffDays, details.to, details.cc, details.statusFlag
      ]);
    }
  }

  // d. 批次寫入結果
  if (logEntries.length > 0) {
    dataRange.setValues(statuses);
    logNotificationBatch(logSheet, logEntries);
    Logger.log(`批次處理完成。共更新 ${logEntries.length} 筆狀態並寫入 Log。`);
  } else {
    Logger.log("沒有需要通知的項目。");
  }
}

/**
 * [最終決定版] 執行後端驅動的通知模擬。
 * 修正了先前版本中所有已知的錯誤：
 * 1. 修正了呼叫 generateNotificationDetails 時錯誤的參數。
 * 2. 修正了回傳給前端時，錯誤地使用了未處理的 `rawBody` 而不是已處理的 `body`。
 */
function simulateNotificationLogic(currentSettings, testData) {
  try {
    // 1. 準備 config 物件 (此部分正確)
    const config = { globalCcEmails: [], exclusionList: [], notificationRules: [], personnelMap: {}, groupMap: {}, additionalCcMap: {} };
    config.globalCcEmails = currentSettings.globalCcEmails || [];
    (currentSettings.exclusionList || []).forEach(item => config.exclusionList.push(item));
    (currentSettings.notificationRules || []).forEach(rule => {
      if (rule[0] !== "" && !isNaN(Number(rule[0]))) {
        config.notificationRules.push({ triggerDays: Number(rule[0]), statusFlag: rule[1], subjectTemplate: rule[2], htmlBodyTemplate: rule[3] });
      }
    });
    (currentSettings.personnelMap || []).forEach(row => { if (row[0] && row[1]) config.personnelMap[row[0]] = { email: row[1] }; });
    (currentSettings.groupMap || []).forEach(row => { if (row[0] && row[1]) config.groupMap[row[0]] = row[1]; });
    (currentSettings.additionalCcMap || []).forEach(row => { if (row[0] && row[1]) config.additionalCcMap[row[0]] = row[1].toString().split(',').map(e => e.trim()); });

    // 2. 準備要給「大腦」的資料 (修正了參數)
    const instrumentInfo = {
      id: testData.instrumentId,
      name: testData.deviceName,
      custodian: testData.custodianName,
      group: testData.groupName
    };
    const today = new Date();
    const expectedDate = new Date();
    expectedDate.setDate(today.getDate() + testData.daysUntilDue);

    // 3. 呼叫「大腦」函式，傳入正確的參數
    // 模擬時 status 傳入 null，因為我們想看的是觸發結果，而非是否要跳過
    const details = generateNotificationDetails(instrumentInfo, expectedDate, null, config);

    // 4. 根據「大腦」的回傳結果，組合正確的資料給前端
    if (details) {
      return {
        status: 'success',
        result: {
          triggeredRule: { days: details.triggerDays, flag: details.statusFlag },
          recipients: { to: details.to, cc: details.cc.split(',').filter(Boolean) },
          // ▼▼▼▼▼ 【最終修正！】 確保使用的是 details.body (已處理)▼▼▼▼▼
          content: { subject: details.subject, body: details.body }
        }
      };
    } else {
      return { status: 'no_rule', message: '根據輸入條件，此儀器不需發送通知。' };
    }
  } catch (e) {
    Logger.log('模擬時發生錯誤: ' + e.toString());
    return { status: 'error', message: '後端執行模擬時發生未知錯誤：' + e.toString() };
  }
}

// ====================================================================
// Helper Functions - 其他輔助函式
// ====================================================================

/**
 * [新增] 統一的設定讀取函式，作為所有設定的唯一來源。
 * @returns {Object} 包含 raw (給前端) 和 parsed (給後端) 兩種格式的設定。
 */
function getGlobalConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  if (!configSheet) throw new Error("找不到 '設定' 工作表。");

  const data = configSheet.getDataRange().getValues();
  const lastRow = data.length;

  const config = {
    raw: { globalCcEmails: [], exclusionList: [], notificationRules: [], personnelMap: [], groupMap: [], additionalCcMap: [] },
    parsed: { globalCcEmails: [], exclusionList: [], notificationRules: [], personnelMap: {}, groupMap: {}, additionalCcMap: {} }
  };

  if (lastRow < 2) return config;

  const globalCcsString = data[1][1] || '';
  config.raw.globalCcEmails = globalCcsString.toString().split(',').map(s => s.trim()).filter(Boolean);
  config.parsed.globalCcEmails = config.raw.globalCcEmails;

  for (let i = 1; i < lastRow; i++) {
    const row = data[i];
    if (row[3]) {
      config.raw.exclusionList.push([row[3]]);
      config.parsed.exclusionList.push(row[3]);
    }
    if (row[5] !== "" || row[6] || row[7] || row[8]) {
      const ruleRow = [row[5], row[6], row[7], row[8]];
      config.raw.notificationRules.push(ruleRow);
      if (ruleRow[0] !== "" && !isNaN(Number(ruleRow[0]))) {
        config.parsed.notificationRules.push({ triggerDays: Number(ruleRow[0]), statusFlag: ruleRow[1], subjectTemplate: ruleRow[2], htmlBodyTemplate: ruleRow[3] });
      }
    }
    if (row[10] || row[11]) {
      config.raw.personnelMap.push([row[10], row[11]]);
      if (row[10] && row[11]) config.parsed.personnelMap[row[10]] = { email: row[11] };
    }
    if (row[13] || row[14]) {
      config.raw.groupMap.push([row[13], row[14]]);
      if (row[13] && row[14]) config.parsed.groupMap[row[13]] = row[14];
    }
    if (row[16] || row[17]) {
      config.raw.additionalCcMap.push([row[16], row[17]]);
      if (row[16] && row[17]) config.parsed.additionalCcMap[row[16]] = row[17].toString().split(',').map(e => e.trim());
    }
  }
  return config;
}

/**
 * [優化] 從指定的外部試算表讀取儀器與組別的對應關係。
 * ID 從 Script Properties 讀取，增加彈性。
 */
function getDeviceMapping() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('DEVICE_MAPPING_SHEET_ID');
  if (!sheetId) {
    Logger.log("錯誤：找不到裝置對應表的試算表 ID。請先執行 setScriptProperties() 進行設定。");
    return {};
  }
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const deviceId = data[i][1],
        group = data[i][5];
      if (deviceId) { map[deviceId] = { group: group }; }
    }
    return map;
  } catch (e) {
    Logger.log(`讀取裝置對應表時發生錯誤 (ID: ${sheetId}): ${e.message}`);
    return {};
  }
}


/**
 * [升級版] 專為 WebApp 設計的函式。
 * 新增讀取「設備名稱」
 */
function getInstrumentsForWebApp() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'WEBAPP_INSTRUMENT_DATA_V4'; // 使用新的快取鍵以強制更新

  const cachedData = cache.get(CACHE_KEY);
  if (cachedData) {
    Logger.log('從快取中成功讀取組合後的儀器資料 (WebApp)。');
    return JSON.parse(cachedData);
  }

  Logger.log('WebApp 快取中無資料，正在從多個來源組合資料...');

  try {
    const deviceGroupMap = getDeviceMapping();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentYear = new Date().getFullYear().toString();
    const yearlySheet = ss.getSheetByName(currentYear);

    if (!yearlySheet) {
      Logger.log(`找不到今年度的工作表: ${currentYear}`);
      return [];
    }

    const lastRow = yearlySheet.getLastRow();
    if (lastRow < 2) return [];

    // 【請確認】根據您的 notify... 邏輯，A欄是儀器編號(0), B欄是設備名稱(1), C欄是保管人(2)
    const ID_COL_INDEX = 0;
    const DEVICE_NAME_COL_INDEX = 1; // <--- 新增讀取設備名稱
    const KEEPER_COL_INDEX = 2;

    const data = yearlySheet.getRange(2, 1, lastRow - 1, KEEPER_COL_INDEX + 1).getValues();
    const instrumentMap = new Map();

    data.forEach(row => {
      const deviceId = row[ID_COL_INDEX] ? String(row[ID_COL_INDEX]).trim() : '';
      if (deviceId && !instrumentMap.has(deviceId)) {
        instrumentMap.set(deviceId, {
          deviceName: row[DEVICE_NAME_COL_INDEX] ? String(row[DEVICE_NAME_COL_INDEX]).trim() : '',
          custodian: row[KEEPER_COL_INDEX] ? String(row[KEEPER_COL_INDEX]).trim() : ''
        });
      }
    });

    const finalInstrumentList = [];
    for (const [id, details] of instrumentMap.entries()) {
      const groupInfo = deviceGroupMap[id] || {};
      finalInstrumentList.push({
        id: id,
        deviceName: details.deviceName, // <--- 新增 deviceName
        custodian: details.custodian,
        group: groupInfo.group || '未知組別'
      });
    }

    cache.put(CACHE_KEY, JSON.stringify(finalInstrumentList), 21600);
    Logger.log(`成功組合並快取了 ${finalInstrumentList.length} 筆儀器資料給 WebApp。`);

    // ▼▼▼▼▼ 請在這裡插入以下這行 ▼▼▼▼▼
    Logger.log(JSON.stringify(finalInstrumentList));
    // ▲▲▲▲▲ 請在這裡插入以上這行 ▲▲▲▲▲
    return finalInstrumentList;

  } catch (e) {
    Logger.log('組合儀器資料給 WebApp 時發生錯誤: ' + e.toString());
    return [];
  }
}

/**
 * [說明] 一次性設定函式，用來儲存外部試算表的ID。
 * 請在編輯器中手動選擇此函式並執行一次即可。
 */
function setScriptProperties() {
  // *** 請將 'YOUR_SPREADSHEET_ID_HERE' 替換成你儀器對應表的實際 ID ***
  PropertiesService.getScriptProperties().setProperty('DEVICE_MAPPING_SHEET_ID', '1pnyTBCcfe6A7E0S2XZhZiSNzE6Er1qEUQPTMjJTWQtI');
  Logger.log("指令碼屬性已成功設定！");
}

/**
 * [優化] 批次寫入 Log。
 * 修正了標頭，增加了「距離校正天數」欄位。
 */
function logNotificationBatch(logSheet, logEntries) {
  // ▼▼▼▼▼ [修正] 檢查標頭，若不存在則寫入包含10個欄位的完整標頭 ▼▼▼▼▼
  if (logSheet.getRange(1, 1).getValue() !== "通知時間") {
    logSheet.getRange("A1:J1").setValues([ // 範圍擴大到 J1
      ["通知時間", "年度", "設備編號", "設備名稱", "保管人", "預計校正日期", "距離校正天數", "主要收件人(To)", "副本(CC)", "通知類型"]
    ]);
  }

  // 這行程式碼不需修改，它會動態寫入正確的欄數
  logSheet.getRange(logSheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
}

// --- 以下是其他未變動的輔助函式 ---
function generateSalutation(fullName) {
  if (!fullName) return "保管人";
  if (fullName.length > 2) return fullName.substring(1);
  return fullName;
}

function parseSimpleMarkup(text) {
  if (!text) return '';
  return text.replace(/\*\*(.*?)\*\*/g, '<b>$1</b>').replace(/\r\n/g, '\n').split('\n\n').map(p => '<p>' + p.replace(/\n/g, '<br>') + '</p>').join('');
}

function sendNotification(recipient, rule, placeholders, ccList) {
  let subject = rule.subjectTemplate;
  let htmlBody = parseSimpleMarkup(rule.htmlBodyTemplate);
  for (const key in placeholders) {
    const regex = new RegExp(key, 'g');
    subject = subject.replace(regex, placeholders[key]);
    htmlBody = htmlBody.replace(regex, placeholders[key]);
  }
  MailApp.sendEmail({ to: recipient, cc: ccList, subject: subject, htmlBody: htmlBody });
}

function parseDate(dateInput) {
  if (dateInput instanceof Date && !isNaN(dateInput)) {
    const d = new Date(dateInput);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  return null;
}

function findMergedContent(data, currentRow, colIndex) {
  for (let r = currentRow; r >= 0; r--) {
    if (data[r][colIndex]) return r;
  }
  return currentRow;
}

/**
 * 手動清除 WebApp 使用的快取。
 * 在編輯器中手動執行此函式以強制更新資料。
 */
function clearWebAppCache() {
  const cache = CacheService.getScriptCache();
  // 使用我們在 getInstrumentsForWebApp 中定義的快取鍵
  const CACHE_KEY = 'WEBAPP_INSTRUMENT_DATA_V4';
  cache.remove(CACHE_KEY);
  Logger.log(`快取 (Key: ${CACHE_KEY}) 已被成功清除！`);
}
