function listWordFiles() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the URL of the parent folder in Google Drive:');
  
  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('No URL entered. Exiting script.');
    return;
  }
  
  var folderUrl = response.getResponseText();
  var folderIdMatch = folderUrl.match(/[-\w]{25,}/);
  if (!folderIdMatch) {
    ui.alert('Invalid folder URL. Please check and try again.');
    return;
  }
  var parentFolderId = folderIdMatch[0];
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ReportsList");
  if (!sheet) {
    sheet = ss.insertSheet("ReportsList");
    sheet.appendRow(["File Name", "Folder Name", "Property Type", "Zoning", "Parish", "Parent Folder", "File URL"]);
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.appendRow(["File Name", "Folder Name", "Property Type", "Zoning", "Parish", "Parent Folder", "File URL"]);
    lastRow = 1;
  }
  
  // Get existing file URLs to avoid duplicates
  var existingUrls = [];
  if (lastRow > 1) {
    existingUrls = sheet.getRange(2, 7, lastRow - 1, 1).getValues().flat();
  }
  
  var newRows = [];
  
  // Map for property types
  var propertyTypeMap = {
    "APT": "Apartment",
    "H": "House",
    "MU": "Multiunit",
    "DU": "Duplex",
    "CN": "Condo",
    "VL": "Vacant Lot",
    "CM": "Commercial",
    "IND": "Industrial"
  };
  
  // List of Bermuda parishes (normalized)
  var parishes = ["Hamilton", "Pembroke", "Southampton", "Devonshire", "Warwick", "Paget", "Smiths", "St Georges", "Sandys"];
  
  function normalizeText(text) {
    return text.toLowerCase().replace(/[\.'’]/g, '').trim();
  }
  
  function processFolder(folder, parentFolderObj) {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (name.endsWith(".doc") || name.endsWith(".docx")) {
        var folderName = folder.getName();
        var propertyType = "";
        var zoning = "";
        var parish = "";
        var parentName = parentFolderObj ? parentFolderObj.getName() : "";
        var fileUrl = file.getUrl();
        
        // Skip if file already exists
        if (existingUrls.includes(fileUrl)) continue;
        
        // Extract property type and zoning from folder name
        var typeMatch = folderName.match(/(APT|H|MU|DU|CN|VL|CM|IND)(\d*)/i);
        if (typeMatch) {
          propertyType = propertyTypeMap[typeMatch[1].toUpperCase()] || "";
          if (typeMatch[2]) {
            zoning = "Residential " + typeMatch[2];
          }
        }
        
        // Find parish
        var folderNormalized = normalizeText(folderName);
        for (var i = 0; i < parishes.length; i++) {
          var parishNormalized = normalizeText(parishes[i]);
          if (folderNormalized.includes(parishNormalized)) {
            parish = parishes[i];
            break;
          }
        }
        
        newRows.push([name, folderName, propertyType, zoning, parish, parentName, fileUrl]);
        existingUrls.push(fileUrl); // Avoid duplicates in the same run
      }
    }
    
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      processFolder(subfolders.next(), folder);
    }
  }
  
  processFolder(parentFolder, null);
  
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);
  }
  
  SpreadsheetApp.flush();
  ui.alert("Listing complete! " + newRows.length + " new files added.");
}
