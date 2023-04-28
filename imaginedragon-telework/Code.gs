const EMAIL_RECIPIENT = "keith.wegner@bestgateeng.com";
const FOLDER_ID = "";

function sendSpreadsheetsByEmail() {
  var folderId = FOLDER_ID;
  var emailAddress = EMAIL_RECIPIENT;
  var subject = 'Bestgate Engineering Telework Logs';
  var message = 'Please find the attached spreadsheets for this month for all of Bestgate Engineering.';

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var attachments = [];

  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();
    var exportUrl = 'https://docs.google.com/spreadsheets/export?id=' + fileId + '&exportFormat=xlsx';
    var fileBlob = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    }).getBlob().setName(file.getName() + '.xlsx');
    attachments.push(fileBlob);
  }

  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    body: message,
    attachments: attachments
  });
}
