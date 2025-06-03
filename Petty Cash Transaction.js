var FOLDER_ID = 'YOUR_FOLDER_ID';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('PettyCashTransaction')
    .setTitle('Petty Cash Usage Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserInfo() {
  const email = Session.getActiveUser().getEmail();
  return getUserName(email);
}

function getUserName(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const emails = sheet.getRange('C5:C').getValues().flat();
  const names = sheet.getRange('B5:B').getValues().flat();
  
  const index = emails.findIndex(e => e.toLowerCase() === email.toLowerCase());
  return index !== -1 ? names[index] : 'Unknown User';
}

function getTransactionTypes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  return sheet.getRange('H5:H').getValues().flat().filter(String);
}

function submitForm(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash');
  const lastRow = sheet.getLastRow();
  
  const timestamp = new Date();
  const user = form.user;
  const purpose = form.purpose;
  const transactionType = form.transactionType;
  const amount = parseFloat(form.amount);
  const remarks = form.remarks;
  
  
  let fileUrl = '';
  if (form.additionalDocs) {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const blob = form.additionalDocs;
    const file = folder.createFile(blob);
    fileUrl = file.getUrl();
  }
  
  
  sheet.getRange(lastRow + 1, 1, 1, 9).setValues([[
    '', 
    timestamp,
    '',
    user,
    purpose,
    transactionType,
    amount,
    fileUrl,
    remarks
  ]]);
  
  return 'Form submitted successfully';
}