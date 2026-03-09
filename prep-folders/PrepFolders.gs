function createPrepFolder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PrepFolders");

  // --- Folder mapping (keyword → Google Drive folder ID) ---
  var parentFolders = {
    "BCB": "1XwhI59k2ZHXXexPBoJlqs3V_mvkvROCU",
    "Clarien": "1XgDrmSM57VZ-zxZfmP0ALF2GkFBx4VEe",
    "BNTB": "1qpTXZYpHJyh1YehrSCCzMOclyHGI26Rd",
    "HSBC": "1vjwPt5lXG2zC-YvLhRWkyOHPSGp5OHF9"
  };

  // --- Measurement template (Google Sheet ID) ---
  var templateId = "1KAJP0Zyd5JWsf0p1pl-A-C7OlUeiBUxqlj1vIw0sfZ8";

  // --- Private folder ---
  var privateFolderId = "1mpCJym7I-pwdJhkAykxS265v5h_QCh2d";

  // --- Get values from sheet ---
  var c2 = sheet.getRange("C2").getValue(); 
  var d2 = sheet.getRange("D2").getValue(); 
  var a2 = sheet.getRange("A2").getValue(); // date
  var b2 = sheet.getRange("B2").getValue(); 
  var keyword = sheet.getRange("E2").getValue(); // parent folder keyword or "Private"
  var tempFolderName = sheet.getRange("F2").getValue(); // temporary folder name for Private
  var wordFileLink = sheet.getRange("G2").getValue(); // existing Word file link

  // Format date as YYMMDD
  var formattedDate = Utilities.formatDate(new Date(a2), Session.getScriptTimeZone(), "yyMMdd");

  // Build main folder name
  var folderName = c2 + d2 + "-" + formattedDate + ". " + b2;

  var parentFolder;

  if (keyword === "Private") {
    var privateFolder = DriveApp.getFolderById(privateFolderId);
    parentFolder = privateFolder.createFolder(tempFolderName);
  } else {
    if (!parentFolders[keyword]) {
      SpreadsheetApp.getUi().alert("Invalid keyword in E2. Use one of: " + Object.keys(parentFolders).join(", ") + " or 'Private'");
      return;
    }
    parentFolder = DriveApp.getFolderById(parentFolders[keyword]);
  }

  // Step 3: create main folder inside the determined parent folder
  var newFolder = parentFolder.createFolder(folderName);

  // Step 4: create Photos subfolder
  var photosFolder = newFolder.createFolder("Photos");

  // Step 5: copy measurement template into new folder
  var templateFile = DriveApp.getFileById(templateId);
  var copiedTemplate = templateFile.makeCopy("Measurements-" + b2, newFolder);

  // Step 6: copy Word file (if link exists) into new folder
  if (wordFileLink) {
    try {
      // Extract file ID from link
      var fileIdMatch = wordFileLink.match(/[-\w]{25,}/);
      if (fileIdMatch) {
        var wordFileId = fileIdMatch[0];
        var wordFile = DriveApp.getFileById(wordFileId);
        wordFile.makeCopy("Appraisal Report - " + b2, newFolder);
      }
    } catch (e) {
      SpreadsheetApp.getUi().alert("Error copying Word file: " + e.message);
    }
  }

  // Step 7: save new folder link in H2
  sheet.getRange("H2").setValue("https://drive.google.com/drive/folders/" + newFolder.getId());

// --- Display a clickable link in a dialog with sans-serif font ---
var htmlOutput = HtmlService.createHtmlOutput(
  "<div style='font-family: Arial, Helvetica, sans-serif; line-height:1.5;'>" +
    "<p><b>Folder created:</b> " + newFolder.getName() + "</p>" +
    "<p><b>Subfolder created:</b> " + photosFolder.getName() + "</p>" +
    "<p><b>Copied template file:</b> " + copiedTemplate.getName() + "</p>" +
    "<p><b>Word file copied:</b> " + (wordFileLink ? "Yes" : "No") + "</p>" +
    "<p><b>Folder link:</b> <a href='https://drive.google.com/drive/folders/" + newFolder.getId() + "' target='_blank'>Open Folder</a></p>" +
  "</div>"
)
.setWidth(400)
.setHeight(250);

SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Folder Creation Complete");

}
