/**
 * 暐固科技 Apple Sales Training 課程公告系統
 * 後端 Apps Script 程式碼
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
 * 網頁進入點
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  // 檢查是否為管理員模式 (網址加入 ?admin=true)
  var isAdmin = e.parameter.admin === 'true';
  template.isAdmin = isAdmin;
  
  return template.evaluate()
    .setTitle("暐固科技 - Apple Sales Training 課程公告")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 載入網頁所需的全部資料
 * @param {boolean} isAdmin 是否為管理員模式
 * @param {string} selectedMonth 選擇的年月 (格式: YYYY-MM)，選填
 */
function getPageData(isAdmin, selectedMonth) {
  try {
    // 確保資料表已初始化
    initDatabase();
    
    var ss = getSpreadsheet();
    
    // 1. 讀取 WebConfig 設定
    var configSheet = ss.getSheetByName("WebConfig");
    var configRows = configSheet.getDataRange().getValues();
    var config = {};
    for (var i = 1; i < configRows.length; i++) {
      config[configRows[i][0]] = configRows[i][1];
    }
    
    // 2. 讀取月曆列表 (Calendars)
    var calSheet = ss.getSheetByName("Calendars");
    var calRows = calSheet.getDataRange().getValues();
    var calendars = [];
    var activeMonth = "";
    
    for (var i = 1; i < calRows.length; i++) {
      var ym = calRows[i][0];
      var status = calRows[i][1];
      
      // 轉換為前端顯示格式：如果是管理員或是啟用狀態才加入
      if (status === "Active" || isAdmin) {
        calendars.push({
          yearMonth: ym,
          status: status
        });
      }
      
      if (status === "Active" && !selectedMonth) {
        activeMonth = ym;
      }
    }
    
    // 決定要查詢的月份
    var targetMonth = selectedMonth || activeMonth;
    if (!targetMonth && calendars.length > 0) {
      targetMonth = calendars[0].yearMonth;
    }
    
    // 3. 讀取選定月份的課程事件 (CalendarEvents)
    var eventSheet = ss.getSheetByName("CalendarEvents");
    var eventRows = eventSheet.getDataRange().getValues();
    var events = [];
    
    for (var i = 1; i < eventRows.length; i++) {
      if (eventRows[i][0] === targetMonth) {
        events.push({
          startDate: formatDateString(eventRows[i][1]),
          endDate: formatDateString(eventRows[i][2]),
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
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 讀取共享資源 (Resource) 資料夾中的檔案清單
 */
function getSharedResources() {
  try {
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
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * 處理學生作業上傳
 * @param {string} studentName 學生中文全名
 * @param {string} fileName 檔案名稱
 * @param {string} base64Data Base64 編碼的檔案內容
 * @param {string} mimeType 檔案的 Mime Type
 */
function uploadFile(studentName, fileName, base64Data, mimeType) {
  try {
    // 驗證輸入
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
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
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
    
    // 預設資料：2026 年 6 月設為 Active (啟用)，2026 年 5 月設為 Archived (封存) 以供測試
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
    
    // 預設事件：June 2026 的工作日
    var defaultEvents = [
      ["2026-06", "2026-06-01", "2026-06-05", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-08", "2026-06-12", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-15", "2026-06-19", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-22", "2026-06-26", "Apple Sales Training", "#e1e1e1"],
      ["2026-06", "2026-06-29", "2026-07-03", "AVAL Level 0 for Sales Training (AF01)", "#ffcccc"],
      // 封存的 5 月事件 (測試用)
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
  
  // 如果是 ISO 字串或物件字串
  var str = date.toString();
  if (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1) {
    return Utilities.formatDate(new Date(str), "GMT+8", "yyyy-MM-dd");
  }
  return str.split(" ")[0];
}
