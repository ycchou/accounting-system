/**
 * --- Code.gs : 後端核心與設定 (完整修復版) ---
 */

const CONFIG = {
  // 🔴【請填寫】您的 Google 試算表 ID (網址 /d/ 後面那一長串)
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

/** --- 交易資料管理 (已加入防呆與錯誤捕捉) --- */

function getSettingsData(token) {
  const check = verifyToken(token);
  if (!check.valid) throw new Error(check.message);
  const d = getSheet(CONFIG.SHEET_NAMES.SETTINGS).getDataRange().getValues();
  return { types: getCol(d, 0), categories: getCol(d, 1), payments: getCol(d, 2) };
}

// 1. 新增交易 (安全版)
function saveTransaction(token, form) {
  try {
    const user = verifyToken(token);
    if (!user.valid) return { success: false, message: "驗證失敗: " + user.message };
    if (user.role === 'Viewer') return { success: false, message: "權限不足" };

    let fileInfo = { url: "", id: "" };
    
    // 處理檔案上傳
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

// 2. 更新交易 (安全版)
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

function getTransactionsByMonth(token, yearStr, monthStr) {
  const check = verifyToken(token);
  if (!check.valid) throw new Error(check.message);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.DB);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const filtered = data.slice(1).filter(r => {
    const d = new Date(r[1]);
    return d.getFullYear() == yearStr && (d.getMonth() + 1) == monthStr;
  });

  return filtered.reverse().map(r => ({
    id: r[0], date: formatDate(r[1]), type: r[2], category: r[3],
    subCategory: r[4], amount: r[5], payment: r[6], memo: r[7], 
    fileUrl: r[8]
  }));
}

function getReportData(token, yearStr, monthStr) {
  const txs = getTransactionsByMonth(token, yearStr, monthStr);
  let income = 0, expense = 0, catMap = {};

  txs.forEach(t => {
    const amt = Number(t.amount);
    if (t.type === '收入') income += amt;
    else if (t.type === '支出') {
      expense += amt;
      if (!catMap[t.category]) catMap[t.category] = 0;
      catMap[t.category] += amt;
    }
  });

  const catStats = Object.keys(catMap).map(k => ({ name: k, value: catMap[k] })).sort((a, b) => b.value - a.value);
  return { income, expense, balance: income - expense, categories: catStats };
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

// 3. 上傳檔案邏輯 (自動分類日期資料夾)
function uploadFile(base64, name, mime, dateStr) {
  try {
    // 取得根目錄
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    
    // 取得日期資料夾
    const folder = getDateFolder(root, dateStr);
    
    // 解碼 Base64
    const blob = Utilities.newBlob(Utilities.base64Decode(base64.split(',')[1]), mime, name);
    
    // 重新命名: YYYYMMDD_Timestamp.ext
    const ext = name.split('.').pop();
    const newName = `${dateStr.replace(/-/g,"")}_${Date.now().toString().slice(-6)}.${ext}`;
    blob.setName(newName);
    
    // 建立檔案
    const file = folder.createFile(blob);
    
    // 設定權限 (選擇性，設為知道連結者可檢視，避免破圖)
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

/** --- 手動建立 Admin 工具 --- */
function createAdminAccount() {
  const adminEmail = "admin@example.com"; 
  const adminPassword = "password123";    
  const adminName = "超級管理員";           

  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === adminEmail) {
      Logger.log("❌ 帳號已存在");
      return;
    }
  }

  const salt = generateSalt(10);
  const hash = generateHash(adminPassword, salt);
  const uuid = Utilities.getUuid();
  
  sheet.appendRow([uuid, adminEmail, adminName, hash, salt, 'Admin', '', '', new Date()]);
  Logger.log("✅ Admin 建立成功: " + adminEmail);
}
