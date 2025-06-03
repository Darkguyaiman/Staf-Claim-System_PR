function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDAdvancePurchasesManagement')
    .setTitle('Advance Purchases Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

function checkUserAccess(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Settings');
    const emails = sheet.getRange('C5:C').getValues().flat();
    return emails.indexOf(email) !== -1 ? email : null;
  } catch (e) {
    return null;
  }
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const [advanceSheet, settingsSheet, eventSheet] = [
      ss.getSheetByName('Advance Purchases'),
      ss.getSheetByName('Settings'),
      ss.getSheetByName('Event Creation')
    ];

    const lastRow = advanceSheet.getLastRow();
    if (lastRow < 4) {
      return JSON.stringify({
        purchases: [],
        statusOptions: [],
        eventOptions: [],
        grouped: {},
        allData: []
      });
    }

    const [purchaseData, statusData, eventData] = [
      advanceSheet.getRange('A4:K' + lastRow).getValues().filter(row => row.some(cell => cell !== '')),
      settingsSheet.getRange('F5:F' + settingsSheet.getLastRow()).getValues().flat().filter(String),
      eventSheet.getRange('D4:D' + eventSheet.getLastRow()).getValues().flat().filter(String)
    ];

    const grouped = {};
    const allData = [];

    purchaseData.forEach((row, index) => {
      const eventId = row[9] || 'No Event ID';
      const eventName = row[3] || 'Uncategorized';
      const groupKey = `${eventId} - ${eventName}`;
      
      const purchaseItem = {
        index: index + 4,
        data: row,
        groupKey,
        eventId,
        eventName,
        status: row[0] || '',
        item: row[4] || '',
        submitter: row[2] || ''
      };

      if (!grouped[groupKey]) grouped[groupKey] = [];
      grouped[groupKey].push(purchaseItem);
      allData.push(purchaseItem);
    });

    return JSON.stringify({
      statusOptions: statusData,
      eventOptions: eventData,
      grouped,
      allData
    });
  } catch (e) {
    return JSON.stringify({
      statusOptions: [],
      eventOptions: [],
      grouped: {},
      allData: []
    });
  }
}

function batchUpdateAdvancePurchases(updates, userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Advance Purchases');
    
    updates.forEach(({ rowIndex, columnIndex, newValue }) => {
      sheet.getRange(rowIndex, columnIndex).setValue(newValue);
      sheet.getRange(rowIndex, 11).setValue(userEmail);
    });
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deletePurchase(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Advance Purchases');
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function uploadInvoiceFile(fileData, fileName, rowIndex, userEmail) {
  try {
    const folderId = 'YOUR_FOLDER_ID';
    const folder = DriveApp.getFolderById(folderId);
    
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.split(',')[1]),
      fileData.split(';')[0].split(':')[1],
      fileName
    );
    
    const file = folder.createFile(blob);
    const fileUrl = file.getUrl();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Advance Purchases');
    sheet.getRange(rowIndex, 9).setValue(fileUrl);
    sheet.getRange(rowIndex, 11).setValue(userEmail);
    
    return { 
      success: true, 
      fileUrl: fileUrl,
      fileName: fileName 
    };
  } catch (e) {
    return { 
      success: false, 
      message: e.toString() 
    };
  }
}

function isValidUrl(string) {
  try {
    new URL(string);
    return true;
  } catch (_) {
    return false;
  }
}