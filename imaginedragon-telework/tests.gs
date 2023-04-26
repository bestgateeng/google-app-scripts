// In order to use GASUnit for testing, you'll first need to import the library into your Google Apps Script project. 
// To do this, follow these steps:
//
// 1. Go to your Google Apps Script project.
// 2. Click on "Extensions" in the left sidebar.
// 3. Click on the "+" button next to "Libraries" to add a new library.
// 4. Enter the GASUnit library ID 1tXPhnU5i5UIGd5z5q6b5Z6gQujK_S-QOC4id4HlGZ1bKc1N23yJhS0P9 and click "Look up."
// 5. Choose the latest version and click "Add."
//
// To run the tests, select runTests from the dropdown menu in the Script Editor and click the play button. 
// This will run both testCreateZip and testSendZipByEmail functions. Be sure to replace the placeholders 
// for the folder ID, sheet name, and your email address before running the tests.
// Please note that running the testSendZipByEmail function will send an actual email with the attached zip file 
// to the specified email address.

// Import GASUnit
const GASUnit = typeof require === 'function' ? require('GASUnit') : this.GASUnit;
const assert = GASUnit.assert;
const assertEquals = GASUnit.assertEquals;

// Test function for createZip
function testCreateZip() {
  // Replace with the appropriate folder ID
  const folderId = 'FOLDER_ID';
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  // Replace with a valid sheet name for testing
  const sheetName = '2023-03';

  const zipBlob = createZip(files, sheetName);
  assert(zipBlob !== null, 'Zip file should not be null');
  assertEquals(zipBlob.getName(), 'Employee_Monthly_Logs_' + sheetName + '.zip', 'Zip file should have the correct name');
}

// Test function for sendZipByEmail
function testSendZipByEmail() {
  // Replace with the appropriate folder ID
  const folderId = 'FOLDER_ID';
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  // Replace with a valid sheet name for testing and your email for testing
  const sheetName = '2023-03';
  const testEmail = 'YOUR_EMAIL@example.com';

  const zipBlob = createZip(files, sheetName);
  if (zipBlob) {
    sendZipByEmail(testEmail, zipBlob, sheetName);
    console.log('Test email sent to ' + testEmail);
  } else {
    console.log('No zip file created');
  }
}

// Run tests
function runTests() {
  testCreateZip();
  testSendZipByEmail();
}
