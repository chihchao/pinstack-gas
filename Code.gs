const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// =========================================================================
// 移除原本寫死的 FOLDER_ID 和 ADMIN_PASSWORD 常數，改用動態讀取專案設定
// =========================================================================

// 初始化與網頁載入
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  // 支援網址帶入 ?page=看板id，前端載入後會自動開啟該看板
  template.initialPageId = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  // 提供正式的 Web App 網址，供前端組合「複製分享連結」使用
  template.baseAppUrl = ScriptApp.getService().getUrl();

  return template.evaluate()
      .setTitle('Pinstack 極簡看板平台')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 驗證管理員密碼
function verifyAdminPassword(password) {
  // 從專案設定（指令碼屬性）撈取最新密碼
  const adminPassword = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return password === adminPassword;
}

// 安全轉換工作表資料
function getSheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("找不到工作表: " + sheetName);
      return []; 
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    const headers = data[0];
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = val.toISOString();
        }
        obj[header] = val;
      });
      return obj;
    });
  } catch(e) {
    Logger.log("getSheetData 發生錯誤: " + e.toString());
    return []; 
  }
}

function appendRowData(sheetName, rowObject) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(header => rowObject[header] !== undefined ? rowObject[header] : "");
  sheet.appendRow(row);
}

function generateId() {
  return Utilities.getUuid();
}

// ==========================================
// 看板 (Pages) 相關 API
// ==========================================
function getPages() {
  return getSheetData('pages');
}

function savePage(pageData, adminPassword) {
  if (!verifyAdminPassword(adminPassword)) throw new Error("權限不足，管理員密碼錯誤");
  
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('pages');
  if (!pageData.id) {
    pageData.id = generateId();
    pageData.datetime = new Date();
    pageData.closed = "false";
    appendRowData('pages', pageData);
  } else {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == pageData.id) {
        sheet.getRange(i + 1, 2, 1, 4).setValues([[pageData.title, pageData.description, new Date(), pageData.closed]]);
        break;
      }
    }
  }
}

function deletePage(pageId, adminPassword) {
  if (!verifyAdminPassword(adminPassword)) throw new Error("權限不足，管理員密碼錯誤");
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const pageSheet = ss.getSheetByName('pages');
  const pageData = pageSheet.getDataRange().getValues();
  for (let i = 1; i < pageData.length; i++) {
    if (pageData[i][0] == pageId) { pageSheet.deleteRow(i + 1); break; }
  }
}

// ==========================================
// 貼文 (Posts) 相關 API
// ==========================================
function getPosts(pageId) {
  const posts = getSheetData('posts');
  return posts.filter(p => p.page_id == pageId);
}

function savePost(postData, nickname, adminPassword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const isAdmin = verifyAdminPassword(adminPassword);
  
  const pages = getSheetData('pages');
  const parentPage = pages.find(p => p.id == postData.page_id);
  if (parentPage && parentPage.closed === "true" && !isAdmin) {
    throw new Error("此看板已關閉，無法張貼或編輯貼文");
  }

  const sheet = ss.getSheetByName('posts');

  // 前端可能已預先產生 id（例如上傳圖片時需以貼文 id 命名檔案），
  // 因此新舊貼文的判斷改以「該 id 是否已存在於工作表」為準，而非 id 是否為空
  const data = sheet.getDataRange().getValues();
  let existingRow = -1;
  if (postData.id) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == postData.id) { existingRow = i; break; }
    }
  }

  if (existingRow === -1) {
    if (!postData.id) postData.id = generateId();
    postData.poster = nickname || "匿名使用者";
    postData.datetime = new Date();
    postData.likes = 0;
    appendRowData('posts', postData);

    const pageSheet = ss.getSheetByName('pages');
    const pageData = pageSheet.getDataRange().getValues();
    for (let i = 1; i < pageData.length; i++) {
      if (pageData[i][0] == postData.page_id) {
        pageSheet.getRange(i + 1, 4).setValue(new Date());
        break;
      }
    }
  } else {
    if (data[existingRow][5] !== nickname && !isAdmin) throw new Error("您無權修改此貼文");
    sheet.getRange(existingRow + 1, 3, 1, 3).setValues([[postData.content, postData.url, postData.image]]);
    sheet.getRange(existingRow + 1, 7).setValue(new Date());
  }
}

function deletePost(postId, nickname, adminPassword) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('posts');
  const data = sheet.getDataRange().getValues();
  const isAdmin = verifyAdminPassword(adminPassword);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == postId) {
      if (data[i][5] !== nickname && !isAdmin) throw new Error("您無權刪除此貼文");
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function likePost(postId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('posts');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == postId) {
      let currentLikes = Number(data[i][7]) || 0;
      sheet.getRange(i + 1, 8).setValue(currentLikes + 1);
      return currentLikes + 1;
    }
  }
}

// ==========================================
// 留言 (Messages) 相關 API
// ==========================================
function getMessages(postId) {
  const messages = getSheetData('messages');
  return messages.filter(m => m.post_id == postId);
}

function addMessage(msgData, nickname) {
  msgData.id = generateId();
  msgData.messager = nickname || "匿名使用者";
  msgData.datetime = new Date();
  appendRowData('messages', msgData);
  return msgData;
}

// ==========================================
// 圖片上傳處理 (Base64)
// ==========================================
function uploadImage(base64Data, fileName) {
  try {
    // 從專案設定（指令碼屬性）撈取最新雲端硬碟資料夾 ID
    const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    if (!folderId) throw new Error("未設定雲端硬碟資料夾 ID (FOLDER_ID)");

    const folder = DriveApp.getFolderById(folderId);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    throw new Error("圖片上傳失敗: " + e.toString());
  }
}