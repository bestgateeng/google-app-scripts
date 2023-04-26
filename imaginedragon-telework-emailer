// Telework Emailer - package telework logs into a ZIP file and email them.
// 
// This script is configured to run on the 1st of each month. It opens a set of Google Spreadsheets that 
// contain telework journal/log details. For each Spreadsheet, it opens the prior month's sheet and exports
// a PDF file of the sheet. It then zips all all of the PDFs and emails them. 
// Some assumptions:
//  - All Spreadsheets in `FOLDER_ID` are related to this application (i.e., they are all telework logs)
//  - All Spreadsheets' individual Sheets will be named "yyyy.MM", e.g. "2023.04")

function zipAndEmailMonthlyTabs() {
  // Replace the following with the appropriate folder ID and recipient email
  const folderId = 'FOLDER_ID';
  const recipientEmail = 'RECIPIENT_EMAIL';
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

  const today = new Date();
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const sheetName = Utilities.formatDate(lastMonth, Session.getScriptTimeZone(), 'yyyy.MM');

  const zipBlob = createZip(files, sheetName);
  if (zipBlob) {
    sendZipByEmail(recipientEmail, zipBlob, sheetName);
  }
}

function createZip(files, sheetName) {
  const pdfBlobs = [];

  while (files.hasNext()) {
    const file = files.next();
    const spreadsheet = SpreadsheetApp.open(file);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (sheet) {
      // Find the last row with content
      const lastRow = sheet.getLastRow();
      
      // Hide columns after column D
      const columnsToHide = sheet.getRange(1, 5, lastRow, sheet.getLastColumn() - 4);
      columnsToHide.activate();
      sheet.hideColumns(columnsToHide.getColumn(), columnsToHide.getNumColumns());

      // Hide empty rows
      const rowsToHide = sheet.getRange(lastRow + 1, 1, sheet.getMaxRows() - lastRow, 4);
      rowsToHide.activate();
      sheet.hideRows(rowsToHide.getRow(), rowsToHide.getNumRows());

      // Export the sheet as a PDF
      const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?';
      const exportOptions = 'exportFormat=pdf&format=pdf' +
        '&size=letter' +
        '&portrait=true' +
        '&fitw=true' +
        '&sheetnames=false&printtitle=false' +
        '&pagenumbers=false&gridlines=false' +
        '&fzr=true' +
        '&gid=' + sheet.getSheetId();
      const pdfBlob = UrlFetchApp.fetch(url + exportOptions, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      }).getBlob();
      pdfBlob.setName(file.getName() + '-' + sheetName + '.pdf');
      pdfBlobs.push(pdfBlob);

      // Unhide the hidden columns and rows
      sheet.unhideColumn(columnsToHide);
      sheet.unhideRow(rowsToHide);
    }
  }

  if (pdfBlobs.length > 0) {
    return Utilities.zip(pdfBlobs, 'Employee_Monthly_Logs_' + sheetName + '.zip');
  } else {
    return null;
  }
}

function sendZipByEmail(recipient, zipBlob, sheetName) {
  const subject = 'Employee Monthly Logs - ' + sheetName;
  const body = 'Please find the attached zip file containing the employee logs for the month ' + sheetName + '.';
  GmailApp.sendEmail(recipient, subject, body, {
    attachments: [zipBlob]
  });
}

// Set up a time-based trigger to run on the 1st of each month
function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('zipAndEmailMonthlyTabs')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
}
