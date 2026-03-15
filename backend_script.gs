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
    
    // 3. Process Smart Columns (Merge "Other" inputs)
    var motivation = data.motivation || "N/A";
    if (motivation === "Other" && data.motivationOther) {
      motivation = data.motivationOther;
    }

    var experience = data.experience || "N/A";
    if (experience === "Other" && data.experienceOther) {
      experience = data.experienceOther;
    }

    var studentStatus = data.studentType || "N/A";
    if (studentStatus === "Other" && data.otherCollege) {
      studentStatus = data.otherCollege;
    }

    // 4. Add header row if missing
    if (sheet.getRange(1, 1).getValue() !== "Timestamp") {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, 8).setValues([[
        "Timestamp", "Full Name", "Email", "Phone",
        "Motivation", "Experience", "Student Status / College", "Payment Screenshot"
      ]]);
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#d9ead3");
    }

    // 5. Append data (Condensed 8-column format)
    sheet.appendRow([
      new Date(),
      data.name    || "N/A",
      data.email   || "N/A",
      data.phone   || "N/A",
      motivation,
      experience,
      studentStatus,
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
