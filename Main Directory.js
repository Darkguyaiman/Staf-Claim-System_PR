function doGet() {
  const userEmail = Session.getActiveUser().getEmail();
  const username = getUsernameFromEmail(userEmail);
  const isAuthorized = checkAuthorization(userEmail);
  
  const template = HtmlService.createTemplateFromFile('MainDirectory');
  template.username = username;
  template.isAuthorized = isAuthorized;
  
  return template.evaluate()
    .setTitle('Main Directory')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUsernameFromEmail(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const range1 = sheet.getRange('C5:C');
  const range2 = sheet.getRange('J5:J');
  const values1 = range1.getValues();
  const values2 = range2.getValues();
  
  for (let i = 0; i < values1.length; i++) {
    if (values1[i][0] === email) {
      return sheet.getRange(`B${i+5}`).getValue();
    }
  }
  
  for (let i = 0; i < values2.length; i++) {
    if (values2[i][0] === email) {
      return sheet.getRange(`I${i+5}`).getValue();
    }
  }
  
  return 'Unknown User';
}

function checkAuthorization(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const range1 = sheet.getRange('C5:C');
  const range2 = sheet.getRange('J5:J');
  const values1 = range1.getValues();
  const values2 = range2.getValues();
  
  for (let i = 0; i < values1.length; i++) {
    if (values1[i][0] === email) {
      return true;
    }
  }
  
  for (let i = 0; i < values2.length; i++) {
    if (values2[i][0] === email) {
      return true;
    }
  }
  
  return false;
}