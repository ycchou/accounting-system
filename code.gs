/**
 * --- Code.gs : 後端核心與設定 (完整修復版 - 含年報與下載功能) ---
 */

const CONFIG = {
  // 🔴【請填寫】您的 Google 試算表 ID
  SPREADSHEET_ID: "1EEut01ck5yRp-Hk0vV5SBgGZ4Sczap6nvnsd6iWjUnE", 
  
  // ✅【已填寫】您的 Google Drive 資料夾 ID
  ROOT_FOLDER_ID: "1RmQqAAdjEZCJeWW2UpxxNGpi2oQDZ5n6", 
  
  SHEET_NAMES: { USERS: "Users", DB: "Database", SETTINGS: "Settings" }
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('帳務系統 Pro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

/** --- 驗證與使用者管理 --- */

function verifyToken(token) {
  if (!token) return { valid: false, message: "無 Token" };
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === token) {
        if (data[i][7] && new Date() > new Date(data[i][7])) {
          return { valid: false, message: "登入逾時" };
        }
        return { valid: true, username: data[i][1], name: data[i][2], role: data[i][5], uid: data[i][0] };
      }
    }
    return { valid: false, message: "無效的 Token" };
  } catch (e) {
    return { valid: false, message: "驗證錯誤: " + e.message };
  }
}

function loginUser(email, pass) {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email) {
        if (generateHash(pass, data[i][4]) === data[i][3]) {
          if (data[i][5] === 'Pending') return { success: false, message: "帳號審核中" };
          const token = Utilities.getUuid();
          sheet.getRange(i + 1, 7).setValue(token);
          sheet.getRange(i + 1, 8).setValue(new Date(Date.now() + 86400000));
          return { success: true, token: token, role: data[i][5], username: email, name: data[i][2] };
        }
      }
    }
    return { success: false, message: "帳號或密碼錯誤" };
  } catch (e) {
    return { success: false, message: "系統錯誤: " + e.message };
  }
}

function handleRegister(email, pass, name) {
  try {
    const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    if (data.length > 1 && data.slice(1).some(r => r[1] === email)) {
      return { success: false, message: "此 Email 已存在" };
    }
    const salt = generateSalt(10);
    sheet.appendRow([Utilities.getUuid(), email, name, generateHash(pass, salt), salt, 'Pending', '', '', new Date()]);
    return { success: true, message: "申請已送出" };
  } catch (e) {
    return { success: false, message: "註冊錯誤: " + e.message };
  }
}

function getAllUsers(token) {
  const user = verifyToken(token);
  if (!user.valid || user.role !== 'Admin') throw new Error("權限不足");
  return getSheet(CONFIG.SHEET_NAMES.USERS).getDataRange().getValues().slice(1).map(r => ({ id: r[0], username: r[1], name: r[2], role: r[5] }));
}

function adminUpdateUser(token, targetUid, action, newRole) {
  const user = verifyToken(token);
  if (!user.valid || user.role !== 'Admin') throw new Error("權限不足");
  if (targetUid === user.uid && action === 'delete') throw new Error("不能刪除自己");

  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetUid) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { success: false, message: "找不到使用者" };

  if (action === 'delete') {
    sheet.deleteRow(rowIndex);
    return { success: true, message: "已刪除" };
  } else {
    sheet.getRange(rowIndex, 6).setValue(newRole);
    return { success: true, message: "權限已更新" };
  }
}

/** --- 交易資料管理 --- */

function getSettingsData(token) {
  const check = verifyToken(token);
  if (!check.valid) throw new Error(check.message);
  const d = getSheet(CONFIG.SHEET_NAMES.SETTINGS).getDataRange().getValues();
  return { types: getCol(d, 0), categories: getCol(d, 1), payments: getCol(d, 2) };
}

// 1. 新增交易
function saveTransaction(token, form) {
  try {
    const user = verifyToken(token);
    if (!user.valid) return { success: false, message: "驗證失敗: " + user.message };
    if (user.role === 'Viewer') return { success: false, message: "權限不足" };

    let fileInfo = { url: "", id: "" };
    
    if (form.fileData) {
      try {
        fileInfo = uploadFile(form.fileData, form.fileName, form.mimeType, form.date);
      } catch (e) {
        return { success: false, message: "圖片上傳失敗，請檢查資料夾權限或 ID。錯誤: " + e.message };
      }
    }

    const sheet = getSheet(CONFIG.SHEET_NAMES.DB);
    sheet.appendRow([
      Utilities.getUuid(), form.date, form.type, form.category, form.subCategory||"", 
      form.amount, form.payment, form.memo, fileInfo.url, fileInfo.id, user.username, new Date()
    ]);
    return { success: true, message: "✅ 記帳成功！" };

  } catch (e) {
    return { success: false, message: "寫入失敗: " + e.message };
  }
}

// 2. 更新交易
function updateTransaction(token, id, form) {
  try {
    const user = verifyToken(token);
    if (!user.valid) return { success: false, message: "驗證失敗" };
    if (user.role === 'Viewer') return { success: false, message: "權限不足" };

    const sheet = getSheet(CONFIG.SHEET_NAMES.DB);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === id) { rowIndex = i + 1; break; }
    }
    if(rowIndex === -1) return { success: false, message: "找不到該筆資料" };

    let fileUrl = data[rowIndex-1][8];
    let fileId = data[rowIndex-1][9];

    if (form.fileData) {
      try {
        const newFile = uploadFile(form.fileData, form.fileName, form.mimeType, form.date);
        fileUrl = newFile.url;
        fileId = newFile.id;
      } catch(e) {
         return { success: false, message: "新圖片上傳失敗: " + e.message };
      }
    }

    const rowRange = sheet.getRange(rowIndex, 2, 1, 9); 
    rowRange.setValues([[
      form.date, form.type, form.category, form.subCategory||"", 
      form.amount, form.payment, form.memo, fileUrl, fileId
    ]]);

    return { success: true, message: "更新成功" };
  } catch(e) {
    return { success: false, message: "更新失敗: " + e.message };
  }
}

function deleteTransaction(token, id) {
  try {
    const user = verifyToken(token);
    if (!user.valid || user.role === 'Viewer') throw new Error("無權限");
    
    const sheet = getSheet(CONFIG.SHEET_NAMES.DB);
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === id) {
        if(data[i][9]) { try { DriveApp.getFileById(data[i][9]).setTrashed(true); } catch(e){} }
        sheet.deleteRow(i+1);
        return { success: true, message: "已刪除" };
      }
    }
    return { success: false, message: "找不到資料" };
  } catch(e) {
    return { success: false, message: "刪除失敗: " + e.message };
  }
}

// 修改: 支援 "ALL" 作為 monthStr 以取得整年資料
function getTransactionsByMonth(token, yearStr, monthStr) {
  const check = verifyToken(token);
  if (!check.valid) throw new Error(check.message);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.DB);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const filtered = data.slice(1).filter(r => {
    const d = new Date(r[1]);
    const isYearMatch = d.getFullYear() == yearStr;
    
    if (monthStr === 'ALL') {
      return isYearMatch;
    } else {
      return isYearMatch && (d.getMonth() + 1) == monthStr;
    }
  });

  return filtered.reverse().map(r => ({
    id: r[0], date: formatDate(r[1]), type: r[2], category: r[3],
    subCategory: r[4], amount: r[5], payment: r[6], memo: r[7], 
    fileUrl: r[8]
  }));
}

// 修改: 新增收入分類統計
function getReportData(token, yearStr, monthStr) {
  const txs = getTransactionsByMonth(token, yearStr, monthStr);
  let income = 0, expense = 0;
  let expMap = {};
  let incMap = {};

  txs.forEach(t => {
    const amt = Number(t.amount);
    if (t.type === '收入') {
      income += amt;
      if (!incMap[t.category]) incMap[t.category] = 0;
      incMap[t.category] += amt;
    } else if (t.type === '支出') {
      expense += amt;
      if (!expMap[t.category]) expMap[t.category] = 0;
      expMap[t.category] += amt;
    }
  });

  const expStats = Object.keys(expMap).map(k => ({ name: k, value: expMap[k] })).sort((a, b) => b.value - a.value);
  const incStats = Object.keys(incMap).map(k => ({ name: k, value: incMap[k] })).sort((a, b) => b.value - a.value);

  return { 
    income, 
    expense, 
    balance: income - expense, 
    categories: expStats,       // 支出分類
    incomeCategories: incStats  // 收入分類 (新增)
  };
}

// 新增: 產生並下載 Excel
function downloadReportExcel(token, yearStr, monthStr) {
  const user = verifyToken(token);
  if (!user.valid) throw new Error("權限不足");

  const data = getReportData(token, yearStr, monthStr);
  const title = `${yearStr}年${monthStr === 'ALL' ? '全年度' : monthStr + '月'}報表`;
  
  // 建立暫存試算表
  const tempSS = SpreadsheetApp.create("Temp_" + Date.now());
  const sheet = tempSS.getSheets()[0];
  
  // 寫入摘要
  sheet.getRange("A1").setValue(title).setFontSize(14).setFontWeight("bold");
  sheet.getRange("A2:B2").setValues([["項目", "金額"]]).setFontWeight("bold").setBackground("#efefef");
  sheet.getRange("A3:B5").setValues([
    ["總收入", data.income],
    ["總支出", data.expense],
    ["結餘", data.balance]
  ]);

  let row = 7;
  // 寫入收入細項
  sheet.getRange(row, 1).setValue("【收入分類統計】").setFontWeight("bold").setFontColor("#198754");
  row++;
  if (data.incomeCategories.length > 0) {
    data.incomeCategories.forEach(c => {
      sheet.getRange(row, 1, 1, 2).setValues([[c.name, c.value]]);
      row++;
    });
  } else {
    sheet.getRange(row, 1).setValue("(無收入資料)");
    row++;
  }

  // 寫入支出細項
  row++;
  sheet.getRange(row, 1).setValue("【支出分類統計】").setFontWeight("bold").setFontColor("#dc3545");
  row++;
  if (data.categories.length > 0) {
    data.categories.forEach(c => {
      sheet.getRange(row, 1, 1, 2).setValues([[c.name, c.value]]);
      row++;
    });
  } else {
    sheet.getRange(row, 1).setValue("(無支出資料)");
    row++;
  }

  // 匯出為 XLSX
  SpreadsheetApp.flush();
  const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
  const options = {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  const blob = response.getBlob().setName(title + ".xlsx");
  
  // 刪除暫存檔
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);

  // 回傳 Base64 供前端下載
  return { 
    filename: title + ".xlsx", 
    base64: Utilities.base64Encode(blob.getBytes()) 
  };
}

/** --- Helpers --- */
function getSheet(name) { 
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`找不到分頁: ${name}`);
    return sheet;
  } catch(e) {
    throw new Error("連接資料庫失敗: " + e.message);
  }
}
function getCol(data, idx) { return data.slice(1).map(r => r[idx]).filter(String); }
function formatDate(d) { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function generateHash(input, salt) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input + salt).map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join(''); }
function generateSalt(len) { let s="";const c="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";for(let i=0;i<len;i++)s+=c.charAt(Math.floor(Math.random()*c.length));return s;}

function uploadFile(base64, name, mime, dateStr) {
  try {
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const folder = getDateFolder(root, dateStr);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(',')[1]), mime, name);
    const ext = name.split('.').pop();
    const newName = `${dateStr.replace(/-/g,"")}_${Date.now().toString().slice(-6)}.${ext}`;
    blob.setName(newName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { url: file.getUrl(), id: file.getId() };
  } catch(e) {
    throw new Error("資料夾存取失敗: " + e.message);
  }
}

function getDateFolder(rootFolder, dateStr) {
  const d = new Date(dateStr);
  const y = d.getFullYear().toString();
  const m = (d.getMonth()+1).toString().padStart(2,'0');
  
  let yF;
  const yFolders = rootFolder.getFoldersByName(y);
  yF = yFolders.hasNext() ? yFolders.next() : rootFolder.createFolder(y);
  
  let mF;
  const mFolders = yF.getFoldersByName(m);
  mF = mFolders.hasNext() ? mFolders.next() : yF.createFolder(m);
  
  return mF;
}

// --- 請貼在 Code.gs 最下方 ---

function forceAuth() {
  // 這個函式的唯一目的是強迫系統跳出授權視窗
  // 隨便抓取一個網站，觸發 script.external_request 權限
  UrlFetchApp.fetch("https://www.google.com");
  Logger.log("✅ 授權成功！現在請去建立新版部署！");
}
