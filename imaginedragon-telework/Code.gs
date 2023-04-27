const EMAIL_RECIPIENT = "keith.wegner@bestgateeng.com";
const FOLDER_ID = "";


function sendSpreadsheetsByEmail() {
  var folderId = FOLDER_ID;
  var emailAddress = EMAIL_RECIPIENT;
  var subject = 'Bestgate Engineering monthly telework logs';
  var message = 'Please find the attached spreadsheets for last month.';


  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var attachments = [];

  var previousMonth = getPreviousMonth();

  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    var sheetId = getSheetIdByName(fileId, previousMonth);

    if (sheetId) {
      var exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx&gid=' + sheetId;
      var fileBlob = UrlFetchApp.fetch(exportUrl, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      }).getBlob().setName(file.getName() + ' - ' + previousMonth + '.xlsx');
      attachments.push(fileBlob);
    }
  }

  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    body: message,
    attachments: attachments
  });
}

function getPreviousMonth() {
  var currentDate = new Date();
  var previousMonthDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, 1);
  var monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  return monthNames[previousMonthDate.getMonth()];
}

function getSheetIdByName(spreadsheetId, sheetName) {
  var sheets = Sheets.Spreadsheets.get(spreadsheetId).sheets;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].properties.title == sheetName) {
      return sheets[i].properties.sheetId;
    }
  }
  return null;
}
