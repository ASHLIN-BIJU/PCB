/**
 * Google Apps Script for PCB Workshop Enrollment - REINFORCED
 */

function doGet(e) {
  return ContentService.createTextOutput("Backend is live. Please use POST for submissions.");
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Wait 10 seconds for lock
  
  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    
    // 1. Open the Spreadsheet
    var sheetId = "1UGoVwEuk21-_KOkyRgD-xD9ZAGpsAPsXmaVHqxplAps";
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    
    // 2. Handle File Upload
    var fileUrl = "No Image";
    if (data.screenshotBase64 && data.screenshotBase64.length > 0) {
      var folderName = "PCB Workshop Payments";
      var folders = DriveApp.getFoldersByName(folderName);
      var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      
      var blob = Utilities.newBlob(Utilities.base64Decode(data.screenshotBase64), data.screenshotMimeType, data.screenshotName);
      var file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }
    
    // 3. Append data
    sheet.appendRow([
      new Date(),
      data.name || "N/A",
      data.email || "N/A",
      data.phone || "N/A",
      data.motivation || "N/A",
      data.experience || "N/A",
      data.studentType || "N/A",
      fileUrl
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({ "result": "success", "rowAdded": sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
