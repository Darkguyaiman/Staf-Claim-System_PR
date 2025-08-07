function doGet() {
  return HtmlService.createTemplateFromFile('CRUDBudget&Expenses')
    .evaluate()
    .setTitle('Event Budget & Expenses Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAllEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const eventSheet = ss.getSheetByName('Event Creation');
  const settingsSheet = ss.getSheetByName('Settings');
  
  const lastEventRow = eventSheet.getLastRow();
  if (lastEventRow < 4) return { events: [], eventTypes: [] };
  
  const lastSettingsRow = settingsSheet.getLastRow();
  
  const eventRange = eventSheet.getRange(4, 4, lastEventRow - 3, 14).getDisplayValues();
  
  const eventData = [];
  const len = eventRange.length;
  for (let i = 0; i < len; i++) {
    const row = eventRange[i];
    if (row[13] && row[12]) {
      eventData.push({ id: row[13], name: row[0] });
    }
  }

  const eventTypes = lastSettingsRow > 4 
    ? settingsSheet.getRange(5, 11, lastSettingsRow - 4, 1).getDisplayValues().map(r => r[0]).filter(v => v)
    : [];

  return { events: eventData, eventTypes: eventTypes };
}

function getPreEventBudgets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName('Pre-Event Budget');
  const lastRow = budgetSheet.getLastRow();
  
  if (lastRow < 4) return [];
  
  const budgetRange = budgetSheet.getRange(4, 1, lastRow - 3, 4).getValues();
  const budgetData = [];
  const len = budgetRange.length;
  
  for (let i = 0; i < len; i++) {
    const row = budgetRange[i];
    if (row[0]) {
      budgetData.push({
        eventId: row[0],
        type: row[1], 
        amount: row[2],
        lastModified: row[3]
      });
    }
  }

  return budgetData;
}

function getPostEventExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName('Post-Event Expenses');
  const lastRow = expensesSheet.getLastRow();
  
  if (lastRow < 4) return [];
  
  const expensesRange = expensesSheet.getRange(4, 1, lastRow - 3, 4).getValues();
  const expensesData = [];
  const len = expensesRange.length;
  
  for (let i = 0; i < len; i++) {
    const row = expensesRange[i];
    if (row[0]) {
      expensesData.push({
        eventId: row[0],
        type: row[1],
        amount: row[2], 
        lastModified: row[3]
      });
    }
  }

  return expensesData;
}

function updatePreEventBudgets(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pre-Event Budget');
  const userEmail = Session.getActiveUser().getEmail();
  
  
  const dataRange = sheet.getRange('A4:C' + sheet.getLastRow());
  const data = dataRange.getValues();
  
  
  const updateMap = new Map();
  updates.forEach(update => {
    const key = `${update.eventId}_${update.oldType}`;
    updateMap.set(key, update);
  });
  
  
  const cellsToUpdate = [];
  
  for (let i = 0; i < data.length; i++) {
    const eventId = data[i][0];
    const type = data[i][1];
    const key = `${eventId}_${type}`;
    
    if (updateMap.has(key)) {
      const update = updateMap.get(key);
      
      
      cellsToUpdate.push({
        row: i + 4,
        col: 2,
        value: update.newType
      });
      
      cellsToUpdate.push({
        row: i + 4,
        col: 3,
        value: update.newAmount
      });
      
      cellsToUpdate.push({
        row: i + 4,
        col: 4,
        value: userEmail
      });
    }
  }
  
  
  cellsToUpdate.forEach(cell => {
    sheet.getRange(cell.row, cell.col).setValue(cell.value);
  });
  
  return true;
}

function updatePostEventExpenses(updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Post-Event Expenses');
  const userEmail = Session.getActiveUser().getEmail();
  
  
  const dataRange = sheet.getRange('A4:C' + sheet.getLastRow());
  const data = dataRange.getValues();
  
  
  const updateMap = new Map();
  updates.forEach(update => {
    const key = `${update.eventId}_${update.oldType}`;
    updateMap.set(key, update);
  });
  
  
  const cellsToUpdate = [];
  
  for (let i = 0; i < data.length; i++) {
    const eventId = data[i][0];
    const type = data[i][1];
    const key = `${eventId}_${type}`;
    
    if (updateMap.has(key)) {
      const update = updateMap.get(key);
      
      
      cellsToUpdate.push({
        row: i + 4,
        col: 2,
        value: update.newType
      });
      
      cellsToUpdate.push({
        row: i + 4,
        col: 3,
        value: update.newAmount
      });
      
      cellsToUpdate.push({
        row: i + 4,
        col: 4,
        value: userEmail
      });
    }
  }
  
  
  cellsToUpdate.forEach(cell => {
    sheet.getRange(cell.row, cell.col).setValue(cell.value);
  });
  
  return true;
}

function deletePreEventBudgetItem(eventId, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pre-Event Budget');
  
  
  const dataRange = sheet.getRange('A4:D' + sheet.getLastRow());
  const data = dataRange.getValues();
  
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === eventId && data[i][1] === type) {
      
      sheet.deleteRow(i + 4);
      return true;
    }
  }
  
  return false;
}

function deletePostEventExpenseItem(eventId, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Post-Event Expenses');
  
  
  const dataRange = sheet.getRange('A4:D' + sheet.getLastRow());
  const data = dataRange.getValues();
  
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === eventId && data[i][1] === type) {
      
      sheet.deleteRow(i + 4);
      return true;
    }
  }
  
  return false;
}

function addNewBudgetItem(newItem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pre-Event Budget');
  const lastRow = sheet.getLastRow();
  const userEmail = Session.getActiveUser().getEmail();
  
  
  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([
    [newItem.eventId, newItem.type, newItem.amount, userEmail]
  ]);
  
  return true;
}

function addNewExpenseItem(newItem) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Post-Event Expenses');
  const lastRow = sheet.getLastRow();
  const userEmail = Session.getActiveUser().getEmail();
  
  
  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([
    [newItem.eventId, newItem.type, newItem.amount, userEmail]
  ]);
  
  return true;
}
