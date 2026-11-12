function doGet(e) {
  var params = e ? e.parameter : {};
  
  if (params.i) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var detailSheet = ss.getSheetByName("記帳明細");
    var now = new Date();
    var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/M/d HH:mm:ss");
    
    var item = params.i;
    var amt = Number(params.a) || 0;
    var cat = params.c || "必要";
    var pay = params.p || "現金";
    
    detailSheet.appendRow([formattedTime, item, amt, cat, pay]);
    
    if (cat === "想要") {
      sendImpulseAlert_(ss, formattedTime, item, amt);
    }
    
    return HtmlService.createHtmlOutput(
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<div style="text-align:center; margin-top:60px; font-family:-apple-system,sans-serif; padding:24px;">' +
      '<div style="display:inline-block; background:#e8f5e9; color:#1b4d3e; padding:8px 16px; border-radius:99px; font-weight:bold; font-size:13px; margin-bottom:12px;">SCAN TO LOG · 掃碼入帳完成</div>' +
      '<h2 style="color:#111; margin:8px 0;">已記錄：' + item + '</h2>' +
      '<p style="font-size:24px; font-weight:bold; color:#1b4d3e; margin:8px 0;">$' + amt + ' <span style="font-size:14px; color:#666; font-weight:normal;">(' + cat + ' · ' + pay + ')</span></p>' +
      '<br><a href="' + ScriptApp.getService().getUrl() + '" style="display:inline-block; padding:12px 28px; background:#1b4d3e; color:white; text-decoration:none; border-radius:10px; font-weight:bold;">查看生活費看板</a>' +
      '</div>'
    );
  }
  
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('生活費與消費監控看板')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 取得或初始化「常用快捷」分頁
function getOrInitPresets_(ss) {
  var sheet = ss.getSheetByName("常用快捷");
  if (!sheet) {
    sheet = ss.insertSheet("常用快捷");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["項目名稱", "金額", "分類", "支付方式"]);
    sheet.appendRow(["早餐豆漿", 25, "必要", "現金"]);
    sheet.appendRow(["瓶裝水", 35, "必要", "現金"]);
    sheet.appendRow(["手搖飲", 75, "想要", "行動支付"]);
    sheet.appendRow(["三明治", 45, "必要", "現金"]);
  }
  
  var data = sheet.getDataRange().getValues();
  var presets = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      presets.push({
        rowIndex: i + 1,
        item: String(data[i][0]),
        amount: Number(data[i][1]) || 0,
        category: String(data[i][2] || "必要"),
        payMethod: String(data[i][3] || "現金")
      });
    }
  }
  return presets;
}

// 讀取本月生活費預算（基本 15000 + 本月加碼金額）
function getBudgetInfo_(ss, yearMonthKey) {
  var sheet = ss.getSheetByName("預算設定");
  if (!sheet) {
    sheet = ss.insertSheet("預算設定");
  }
  
  var data = sheet.getDataRange().getValues();
  var baseBudget = null;
  var extraBudget = 0;
  
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][0] || "");
    var val = Number(data[i][1]) || 0;
    if (label === "基本生活費") {
      baseBudget = val;
    } else if (label === "加碼_" + yearMonthKey) {
      extraBudget += val;
    }
  }
  
  if (baseBudget === null) {
    baseBudget = 15000;
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["預算項目", "金額", "備註"]);
    }
    sheet.appendRow(["基本生活費", 15000, "每月固定生活費"]);
  }
  
  return {
    baseBudget: baseBudget,
    extraBudget: extraBudget,
    totalBudget: baseBudget + extraBudget
  };
}

// 取得儀表板完整資料
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var detailSheet = ss.getSheetByName("記帳明細");
  var now = new Date();
  var currentMonth = now.getMonth();
  var currentYear = now.getFullYear();
  var tz = Session.getScriptTimeZone();
  var yearMonthKey = Utilities.formatDate(now, tz, "yyyy-MM");
  
  var budgetInfo = getBudgetInfo_(ss, yearMonthKey);
  var presets = getOrInitPresets_(ss);
  
  var data = detailSheet.getDataRange().getValues();
  var monthTotal = 0;
  var monthWantTotal = 0;
  var monthNeedTotal = 0;
  var recentRecords = [];
  
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    var rowDate = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    var rowCat = String(data[i][3] || "");
    var rowAmt = Number(data[i][2]) || 0;
    
    if (!isNaN(rowDate.getTime()) && rowDate.getFullYear() === currentYear && rowDate.getMonth() === currentMonth) {
      monthTotal += rowAmt;
      if (rowCat === "想要") {
        monthWantTotal += rowAmt;
      } else {
        monthNeedTotal += rowAmt;
      }
    }
  }
  
  for (var r = data.length - 1; r >= 1 && recentRecords.length < 5; r--) {
    var d = data[r][0];
    var timeStr = (d instanceof Date) ? Utilities.formatDate(d, tz, "MM/dd HH:mm") : String(d).slice(5, 16);
    recentRecords.push({
      rowIndex: r + 1,
      time: timeStr,
      item: String(data[r][1] || ""),
      amount: Number(data[r][2]) || 0,
      category: String(data[r][3] || ""),
      payMethod: String(data[r][4] || "")
    });
  }
  
  var lastDay = new Date(currentYear, currentMonth + 1, 0).getDate();
  var daysLeft = Math.max(1, lastDay - now.getDate() + 1);
  var remaining = budgetInfo.totalBudget - monthTotal;
  var dailyAllowance = Math.max(0, Math.floor(remaining / daysLeft));
  
  return {
    baseBudget: budgetInfo.baseBudget,
    extraBudget: budgetInfo.extraBudget,
    totalBudget: budgetInfo.totalBudget,
    monthTotal: monthTotal,
    monthWantTotal: monthWantTotal,
    monthNeedTotal: monthNeedTotal,
    remaining: remaining,
    daysLeft: daysLeft,
    dailyAllowance: dailyAllowance,
    presets: presets,
    recentRecords: recentRecords
  };
}

// 新增消費紀錄
function addExpenseRecord(item, amount, category, payMethod) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var detailSheet = ss.getSheetByName("記帳明細");
  var now = new Date();
  var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/M/d HH:mm:ss");
  
  detailSheet.appendRow([formattedTime, item, Number(amount), category, payMethod]);
  
  if (category === "想要") {
    sendImpulseAlert_(ss, formattedTime, item, amount);
  }
  
  var dashboard = getDashboardData();
  dashboard.message = "已記錄：" + item + " ($" + amount + ")";
  return dashboard;
}

// 寄送衝動消費提醒信
function sendImpulseAlert_(ss, formattedTime, item, amount) {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return;
    var dashboard = getDashboardData();
    var subject = "【生活費警示】新增一筆「想要」消費：$" + amount;
    var lines = [
      "哈囉！",
      "",
      "剛剛記錄了一筆「想要」類別的消費：",
      "- 時間：" + formattedTime,
      "- 項目：" + item,
      "- 金額：$" + amount,
      "- 本月生活費剩餘：$" + dashboard.remaining + " / 總額度 $" + dashboard.totalBudget,
      "- 本月「想要」累計：$" + dashboard.monthWantTotal,
      "- 距離月底剩 " + dashboard.daysLeft + " 天，平均每天可用：$" + dashboard.dailyAllowance,
      "",
      "留意小額高頻支出，守住本月生活費！"
    ];
    MailApp.sendEmail(userEmail, subject, lines.join("\n"));
  } catch (e) {
    // 略過郵件權限例外
  }
}

// 補充額外生活費
function addExtraBudget(amount, note) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("預算設定");
  var now = new Date();
  var yearMonthKey = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");
  
  sheet.appendRow(["加碼_" + yearMonthKey, Number(amount), note || "額外補充"]);
  
  var dashboard = getDashboardData();
  dashboard.message = "已補充生活費 +$" + amount;
  return dashboard;
}

// 新增自訂快速按鈕
function addQuickPreset(item, amount, category, payMethod) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("常用快捷");
  if (!sheet) {
    getOrInitPresets_(ss);
    sheet = ss.getSheetByName("常用快捷");
  }
  sheet.appendRow([item, Number(amount), category, payMethod]);
  
  var dashboard = getDashboardData();
  dashboard.message = "已新增快捷鍵：" + item;
  return dashboard;
}

// 刪除自訂快速按鈕
function deleteQuickPreset(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("常用快捷");
  var idx = Number(rowIndex);
  if (sheet && idx >= 2 && idx <= sheet.getLastRow()) {
    sheet.deleteRow(idx);
  }
  var dashboard = getDashboardData();
  dashboard.message = "已移除該快捷按鈕";
  return dashboard;
}

// 撤銷消費紀錄
function deleteExpenseRecord(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var detailSheet = ss.getSheetByName("記帳明細");
  var idx = Number(rowIndex);
  if (idx >= 2 && idx <= detailSheet.getLastRow()) {
    detailSheet.deleteRow(idx);
  }
  var dashboard = getDashboardData();
  dashboard.message = "已撤銷該筆消費紀錄";
  return dashboard;
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
