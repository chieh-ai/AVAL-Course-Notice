/**
 * 暐固科技 Apple Sales Training 課程公告系統
 * 後端 Apps Script 程式碼 (RESTful JSON API 版本)
 */

// 全域設定：Google Drive 資料夾 ID
const UPLOAD_FOLDER_ID = "1HaGeckqP47Pd82WTh9RQy3XmpGOJWSJF";
const RESOURCE_FOLDER_ID = "1VS-NjTzA8AYIm7lMcwRlN5Wn-haeKT0B";
const SPREADSHEET_ID = "1WM8PAYjBADe8hJpRjzABzM5S62a6kpzYdi2zV0fWPDs";

/**
 * 取得試算表實例（支援容器綁定與獨立指令碼）
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
}

/**
 * API GET 請求入口
 * 處理網頁資料讀取與資源清單查詢
 */
function doGet(e) {
  var action = e.parameter.action;
  var responseData;
  
  try {
    // 確保資料表已初始化
    initDatabase();
    
    if (action === "getPageData") {
      var isAdmin = e.parameter.admin === 'true';
      var selectedMonth = e.parameter.month;
      responseData = getPageData(isAdmin, selectedMonth);
    } 
    else if (action === "getResources") {
      responseData = getSharedResources();
    } 
    else if (action === "debugConfig") {
      var ss = getSpreadsheet();
      var sheet = ss.getSheetByName("WebConfig");
      responseData = { success: true, rows: sheet.getDataRange().getValues() };
    }
    else {
      responseData = { 
        success: false, 
        error: "未知的 GET 動作 (Action)。請指定 ?action=getPageData 或 ?action=getResources" 
      };
    }
  } catch (error) {
    responseData = { success: false, error: error.toString() };
  }
  
  // 以 JSON 格式輸出，GAS 會自動附加對應的 CORS 標頭與處理轉址
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * API POST 請求入口
 * 處理學生作業上傳及所有後端管理功能
 */
function doPost(e) {
  var responseData;
  
  try {
    // 確保資料表已初始化
    initDatabase();
    
    // 解析 POST 傳入的 JSON 內容
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    
    if (action === "uploadFile") {
      var studentName = postData.studentName;
      var fileName = postData.fileName;
      var base64Data = postData.base64Data;
      var mimeType = postData.mimeType;
      
      responseData = uploadFile(studentName, fileName, base64Data, mimeType);
    } 
    else if (action === "verifyAdmin") {
      var isValid = checkAdminAuth(postData.password);
      responseData = { success: isValid };
    }
    else if (action === "saveConfig") {
      if (!checkAdminAuth(postData.password)) {
        responseData = { success: false, error: "密碼驗證失敗，無管理權限。" };
      } else {
        responseData = saveWebConfig(postData.configData);
      }
    }
    else if (action === "saveEvents") {
      if (!checkAdminAuth(postData.password)) {
        responseData = { success: false, error: "密碼驗證失敗，無管理權限。" };
      } else {
        responseData = saveCalendarEvents(postData.yearMonth, postData.eventsData);
      }
    }
    else if (action === "saveCalendars") {
      if (!checkAdminAuth(postData.password)) {
        responseData = { success: false, error: "密碼驗證失敗，無管理權限。" };
      } else {
        responseData = saveCalendarsStatus(postData.calendarsData);
      }
    }
    else if (action === "getUploadLogs") {
      if (!checkAdminAuth(postData.password)) {
        responseData = { success: false, error: "密碼驗證失敗，無管理權限。" };
      } else {
        responseData = getUploadLogs();
      }
    }
    else {
      responseData = { success: false, error: "未知的 POST 動作 (Action)" };
    }
  } catch (error) {
    responseData = { success: false, error: "POST 請求處理錯誤: " + error.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 載入網頁所需的全部資料
 */
function getPageData(isAdmin, selectedMonth) {
  var ss = getSpreadsheet();
  
  // 1. 讀取 WebConfig 設定 (安全過濾敏感資訊)
  var configSheet = ss.getSheetByName("WebConfig");
  var configRows = configSheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < configRows.length; i++) {
    var key = configRows[i][0];
    if (key !== "admin_password") {
      var val = configRows[i][1];
      if (val instanceof Date) {
        config[key] = formatDateString(val);
      } else {
        config[key] = val;
      }
    }
  }
  
  // 2. 獲取行事曆月份列表 (動態由週期區間與資料庫事件計算得出)
  var startDate = formatDateString(config.calendar_start_date) || "2026-06-29";
  var endDate = formatDateString(config.calendar_end_date) || "2026-07-03";
  var startMonth = startDate.substring(0, 7);
  var endMonth = endDate.substring(0, 7);
  
  var configuredMonths = getMonthsInRange(startMonth, endMonth);
  var calendars = [];
  
  var eventSheet = ss.getSheetByName("CalendarEvents");
  var eventRows = eventSheet.getDataRange().getValues();
  
  if (isAdmin) {
    // 管理員可看到：所有已建立事件的月份 + 設定區間內的月份
    var allMonthsMap = {};
    configuredMonths.forEach(function(m) {
      allMonthsMap[m] = { yearMonth: m, status: "Active" };
    });
    for (var i = 1; i < eventRows.length; i++) {
      var ym = formatYearMonth(eventRows[i][0]);
      if (ym) {
        if (!allMonthsMap.hasOwnProperty(ym)) {
          allMonthsMap[ym] = { yearMonth: ym, status: "Archived" };
        }
      }
    }
    var sortedYM = Object.keys(allMonthsMap).sort();
    calendars = sortedYM.map(function(ym) {
      return allMonthsMap[ym];
    });
  } else {
    // 一般用戶：只可看到設定區間內的月份
    calendars = configuredMonths.map(function(m) {
      return { yearMonth: m, status: "Active" };
    });
  }
  
  // 計算預設顯示的月份 (activeMonth)
  var activeMonth = "";
  if (selectedMonth) {
    var exists = calendars.some(function(c) { return c.yearMonth === selectedMonth; });
    if (exists) activeMonth = selectedMonth;
  }
  
  if (!activeMonth) {
    var today = new Date();
    var curYM = today.getFullYear() + "-" + ((today.getMonth() + 1) < 10 ? "0" + (today.getMonth() + 1) : (today.getMonth() + 1));
    var hasCurrent = calendars.some(function(c) { return c.yearMonth === curYM; });
    if (hasCurrent) {
      activeMonth = curYM;
    } else if (calendars.length > 0) {
      activeMonth = calendars[0].yearMonth;
    }
  }
  
  var targetMonth = activeMonth;
  
  // 3. 讀取選定月份的課程事件 (CalendarEvents)
  var events = [];
  var isRangeMode = !selectedMonth; // 若前台加載且未選取月份，則為完整日期區間模式
  
  for (var i = 1; i < eventRows.length; i++) {
    var eventStart = formatDateString(eventRows[i][1]);
    var eventEnd = formatDateString(eventRows[i][2]);
    var eventMonth = formatYearMonth(eventRows[i][0]);
    
    var match = false;
    if (isRangeMode) {
      // 區間模式：只要事件與前台顯示週期有交集即可
      if (eventStart <= endDate && eventEnd >= startDate) {
        match = true;
      }
    } else {
      // 編輯模式：指定月份相符的事件
      if (eventMonth === targetMonth) {
        match = true;
      }
    }
    
    if (match) {
      if (!isAdmin) {
        // 一般使用者進行事件區間裁切以符合顯示週期
        if (eventStart < startDate) {
          eventStart = startDate;
        }
        if (eventEnd > endDate) {
          eventEnd = endDate;
        }
      }
      
      events.push({
        startDate: eventStart,
        endDate: eventEnd,
        name: eventRows[i][3],
        color: eventRows[i][4]
      });
    }
  }
  
  return {
    success: true,
    config: config,
    calendars: calendars,
    activeMonth: targetMonth,
    events: events
  };
}

/**
 * 讀取共享資源 (Resource) 資料夾中的檔案清單
 */
function getSharedResources() {
  var folder = DriveApp.getFolderById(RESOURCE_FOLDER_ID);
  var files = folder.getFiles();
  var list = [];
  
  while (files.hasNext()) {
    var file = files.next();
    list.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId(),
      size: formatBytes(file.getSize()),
      mimeType: file.getMimeType(),
      updated: Utilities.formatDate(file.getLastUpdated(), "GMT+8", "yyyy/MM/dd")
    });
  }
  
  return { success: true, files: list };
}

/**
 * 處理學生作業上傳
 */
function uploadFile(studentName, fileName, base64Data, mimeType) {
  if (!studentName || studentName.trim() === "") {
    throw new Error("請輸入您的中文全名");
  }
  if (!fileName || !base64Data) {
    throw new Error("未選取上傳檔案或檔案為空");
  }
  
  studentName = studentName.trim();
  var parentFolder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
  
  // 搜尋有無同名的學生資料夾
  var studentFolder;
  var subFolders = parentFolder.getFoldersByName(studentName);
  if (subFolders.hasNext()) {
    studentFolder = subFolders.next();
  } else {
    // 建立學生專屬資料夾
    studentFolder = parentFolder.createFolder(studentName);
  }
  
  // 解碼並建立檔案
  var decodedData = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decodedData, mimeType, fileName);
  var file = studentFolder.createFile(blob);
  
  // 將紀錄寫入 Google Sheet (UploadLogs)
  var ss = getSpreadsheet();
  var logSheet = ss.getSheetByName("UploadLogs");
  logSheet.appendRow([
    new Date(),
    studentName,
    fileName,
    file.getUrl(),
    studentFolder.getUrl()
  ]);
  
  return {
    success: true,
    fileName: fileName,
    fileUrl: file.getUrl()
  };
}

/**
 * 初始化資料庫表格及預設資料
 */
function initDatabase() {
  var ss = getSpreadsheet();
  
  // 1. 建立 WebConfig (網頁文字內容設定)
  var configSheet = ss.getSheetByName("WebConfig");
  if (!configSheet) {
    configSheet = ss.insertSheet("WebConfig");
    configSheet.appendRow(["Key", "Value", "Description"]);
    
    var defaultConfig = [
      ["course_code", "AF01", "課程代碼"],
      ["course_name", "AVAL Level 0 for Sales Training", "課程名稱"],
      ["course_outline", "• 以 Apple 通路、設計理念，及客戶認知連結銷售語言\n• 從 iPhone 產品特色及 iOS 帶出銷售\n• iPhone 周邊配件 及 MagSafe 產品特色\n• AppleCare+ 特色與銷售語言\n• AirPods 與 HomePod 家族產品特色\n• iCloud 服務、生態圈整合應用與銷售語言\n• 從 Apple Watch 產品特色及 watchOS 帶出銷售\n• 從 iPad 定位及 iPadOS 帶出銷售，iPad 周邊配件\n• 從 Mac 產品特色與 macOS 帶出銷售，並做出與 PC 的差異", "課程大綱 (每一列一個項目)"],
      ["class_date", "06/29（一）至 07/03（五）", "上課日期"],
      ["class_time", "09:30 至 17:00（第一天 10:00 開始）", "上課時間"],
      ["participants", "富鴻 王宥菁\n一七二 陳瑞滿\n優仕 李睿彬、鍾家杅", "參加名單 (姓名以換行或空格分隔)"],
      ["sales_drill", "◎", "銷售演練 (◎ 代表有，或填入說明)"],
      ["homework", "◎", "課後作業 (◎ 代表有，或填入說明)"],
      ["remarks", "① 需填寫課程報名表\n② 需提前完成並繳交課前作業\n③ 需自備 iPhone ，並升級至最新版本（iOS 26）\n④ 建議自備其他 iOS 裝置（iPad、Apple Pencil、Apple Watch），並升級至最新版本（iPadOS 26、watchOS 26）\n⑤ 建議自備 MacBook 筆電，並升級至最新版本 macOS 26", "備註事項 (每一列一項)"],
      ["location", "新北市新店區中正路 535 號 9 樓，906 會議室（進遠東園區後右側電梯上樓）", "上課地點 (將自動生成地圖)"],
      ["registration_text", "您已成功報名此課程或尚未完成登記？", "報名說明文字"],
      ["registration_url", "https://forms.gle/zFQcrKQNS3q7uqnK9", "Google Form 報名連結"],
      ["reg_l1_l3", "由各經銷業務窗口於課程前一個月 20 日前統一報名", "L1-L3 報名說明"],
      ["notices", "1. 上課過程若有發燒、咳嗽...等不適狀況請務必即時通報公司主管（窗口）及課程講師。\n2. 響應環保，若有飲水需求，請自備水壺、水杯...等日常用品。\n3. 北部近日天氣不穩定，請自行評估是否攜帶雨具及保暖衣物以備不時之需。", "其他重要注意事項 (每一列一項)"],
      ["contact_info", "暐固科技 業務行銷處 吳俊頡 副理 | 02-22186188 | chieh.wu@senao.com.tw", "聯絡資訊 (文字格式：公司 部門 聯絡人 | 電話 | 信箱)"]
    ];
    
    for (var i = 0; i < defaultConfig.length; i++) {
      configSheet.appendRow(defaultConfig[i]);
    }
    configSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
    configSheet.autoResizeColumns(1, 3);
  }
  
  // 2. 建立 Calendars (月曆列表)
  var calListSheet = ss.getSheetByName("Calendars");
  if (!calListSheet) {
    calListSheet = ss.insertSheet("Calendars");
    calListSheet.appendRow(["YearMonth", "Status"]);
    calListSheet.appendRow(["2026-06", "Active"]);
    calListSheet.appendRow(["2026-05", "Archived"]);
    calListSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");
    calListSheet.autoResizeColumns(1, 2);
  }
  
  // 3. 建立 CalendarEvents (月曆事件/課程區間)
  var calEventsSheet = ss.getSheetByName("CalendarEvents");
  if (!calEventsSheet) {
    calEventsSheet = ss.insertSheet("CalendarEvents");
    calEventsSheet.appendRow(["YearMonth", "StartDate", "EndDate", "EventName", "Color"]);
    
    var defaultEvents = [
      ["2026-06", "2026-06-01", "2026-06-05", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-08", "2026-06-12", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-15", "2026-06-19", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-22", "2026-06-26", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-29", "2026-07-03", "AVAL Level 0 for Sales Training (AF01)", "#ffcccc"],
      ["2026-05", "2026-05-11", "2026-05-15", "Apple Sales Training (Archived Demo)", "#e1e1e1"]
    ];
    
    for (var i = 0; i < defaultEvents.length; i++) {
      calEventsSheet.appendRow(defaultEvents[i]);
    }
    calEventsSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");
    calEventsSheet.autoResizeColumns(1, 5);
  }
  
  // 4. 建立 UploadLogs (學生上傳紀錄)
  var uploadLogsSheet = ss.getSheetByName("UploadLogs");
  if (!uploadLogsSheet) {
    uploadLogsSheet = ss.insertSheet("UploadLogs");
    uploadLogsSheet.appendRow(["Timestamp", "StudentName", "FileName", "FileUrl", "FolderUrl"]);
    uploadLogsSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#f3f3f3");
    uploadLogsSheet.autoResizeColumns(1, 5);
  }
  
  // 5. 確保 admin_password, calendar_start_month, calendar_end_month, calendar_start_date, calendar_end_date, location_remarks 鍵存在於 WebConfig
  var hasPasswordKey = false;
  var hasStartKey = false;
  var hasEndKey = false;
  var hasStartDateKey = false;
  var hasEndDateKey = false;
  var hasLocRemarksKey = false;
  var hasLocClassroomKey = false;
  var data = configSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    if (key === "admin_password") hasPasswordKey = true;
    if (key === "calendar_start_month") hasStartKey = true;
    if (key === "calendar_end_month") hasEndKey = true;
    if (key === "calendar_start_date") hasStartDateKey = true;
    if (key === "calendar_end_date") hasEndDateKey = true;
    if (key === "location_remarks") hasLocRemarksKey = true;
    if (key === "location_classroom") hasLocClassroomKey = true;
  }
  
  var needsResize = false;
  if (!hasPasswordKey) {
    configSheet.appendRow(["admin_password", "admin123", "系統管理員登入密碼"]);
    needsResize = true;
  }
  if (!hasStartKey) {
    configSheet.appendRow(["calendar_start_month", "2026-06", "行事曆顯示週期開始月份"]);
    needsResize = true;
  }
  if (!hasEndKey) {
    configSheet.appendRow(["calendar_end_month", "2026-07", "行事曆顯示週期結束月份"]);
    needsResize = true;
  }
  if (!hasStartDateKey) {
    configSheet.appendRow(["calendar_start_date", "2026-06-29", "行事曆顯示週期開始日期"]);
    needsResize = true;
  }
  if (!hasEndDateKey) {
    configSheet.appendRow(["calendar_end_date", "2026-07-03", "行事曆顯示週期結束日期"]);
    needsResize = true;
  }
  if (!hasLocRemarksKey) {
    configSheet.appendRow(["location_remarks", "進遠東園區後右側電梯上樓", "上課地點備註"]);
    needsResize = true;
  }
  if (!hasLocClassroomKey) {
    configSheet.appendRow(["location_classroom", "906 會議室", "上課地點教室"]);
    needsResize = true;
  }
  
  if (needsResize) {
    configSheet.autoResizeColumns(1, 3);
  }
}

/**
 * 格式化檔案大小
 */
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/**
 * 輔助函數：日期轉成字串格式 (YYYY-MM-DD)
 */
function formatDateString(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, "GMT+8", "yyyy-MM-dd");
  }
  if (!date) return "";
  var str = date.toString();
  if (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1) {
    return Utilities.formatDate(new Date(str), "GMT+8", "yyyy-MM-dd");
  }
  return str.split(" ")[0];
}

/**
 * 輔助函數：日期轉成 YYYY-MM 格式字串，容錯處理 Date 與 String
 */
function formatYearMonth(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, "GMT+8", "yyyy-MM");
  }
  if (!val) return "";
  var str = val.toString().trim();
  if (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1) {
    return Utilities.formatDate(new Date(str), "GMT+8", "yyyy-MM");
  }
  if (str.length >= 7) {
    return str.substring(0, 7);
  }
  return str;
}

/**
 * 驗證管理員密碼
 */
function checkAdminAuth(password) {
  if (!password) return false;
  var ss = getSpreadsheet();
  var configSheet = ss.getSheetByName("WebConfig");
  if (!configSheet) return false;
  var configRows = configSheet.getDataRange().getValues();
  for (var i = 1; i < configRows.length; i++) {
    if (configRows[i][0] === "admin_password") {
      return configRows[i][1].toString().trim() === password.trim();
    }
  }
  return false;
}

/**
 * 儲存網頁文字設定值
 */
function saveWebConfig(configData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("WebConfig");
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    for (var i = 1; i < values.length; i++) {
      var key = values[i][0];
      if (configData.hasOwnProperty(key)) {
        values[i][1] = configData[key];
      }
    }
    range.setValues(values);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 覆寫特定月份的日曆行程
 */
function saveCalendarEvents(yearMonth, eventsData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("CalendarEvents");
    var range = sheet.getDataRange();
    var values = range.getValues();
    var newValues = [values[0]]; // 保留表頭
    
    // 保留非本月事件，剔除本月事件以進行覆寫
    for (var i = 1; i < values.length; i++) {
      if (formatYearMonth(values[i][0]) !== yearMonth) {
        newValues.push(values[i]);
      }
    }
    
    // 加入新的事件列
    for (var i = 0; i < eventsData.length; i++) {
      var evt = eventsData[i];
      newValues.push([
        yearMonth,
        evt.startDate,
        evt.endDate,
        evt.name,
        evt.color || "#e1e1e1"
      ]);
    }
    
    sheet.clearContents();
    sheet.getRange(1, 1, newValues.length, 5).setValues(newValues);
    sheet.autoResizeColumns(1, 5);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 調整月份啟用/封存清單
 */
function saveCalendarsStatus(calendarsData) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Calendars");
    sheet.clearContents();
    
    var rows = [["YearMonth", "Status"]];
    for (var i = 0; i < calendarsData.length; i++) {
      rows.push([
        calendarsData[i].yearMonth,
        calendarsData[i].status
      ]);
    }
    sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    sheet.autoResizeColumns(1, 2);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 取得學員作業上傳紀錄
 */
function getUploadLogs() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("UploadLogs");
    if (!sheet) return { success: true, logs: [] };
    
    var rows = sheet.getDataRange().getValues();
    var logs = [];
    // 由新到舊 (倒序) 讀取
    for (var i = rows.length - 1; i >= 1; i--) {
      logs.push({
        timestamp: formatDateString(rows[i][0]) + " " + formatTimeString(rows[i][0]),
        studentName: rows[i][1],
        fileName: rows[i][2],
        fileUrl: rows[i][3],
        folderUrl: rows[i][4]
      });
    }
    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 輔助函數：時間轉成字串格式 (HH:mm:ss)
 */
function formatTimeString(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, "GMT+8", "HH:mm:ss");
  }
  if (!date) return "";
  var str = date.toString();
  if (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1) {
    return Utilities.formatDate(new Date(str), "GMT+8", "HH:mm:ss");
  }
  var parts = str.split(" ");
  if (parts.length > 1) return parts[1];
  return "";
}

/**
 * 計算兩個年份月份字串 (YYYY-MM) 之間所有的月份列表 (包含首尾)
 */
function getMonthsInRange(startStr, endStr) {
  var months = [];
  try {
    var startParts = startStr.split("-");
    var endParts = endStr.split("-");
    var startYear = parseInt(startParts[0]);
    var startMonth = parseInt(startParts[1]);
    var endYear = parseInt(endParts[0]);
    var endMonth = parseInt(endParts[1]);
    
    var currYear = startYear;
    var currMonth = startMonth;
    
    while (currYear < endYear || (currYear === endYear && currMonth <= endMonth)) {
      var mStr = currMonth < 10 ? "0" + currMonth : "" + currMonth;
      months.push(currYear + "-" + mStr);
      currMonth++;
      if (currMonth > 12) {
        currMonth = 1;
        currYear++;
      }
    }
  } catch (e) {
    // 容錯回傳起點月份
    return [startStr];
  }
  return months;
}
