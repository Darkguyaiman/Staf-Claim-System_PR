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
  
  
  const eventRange = eventSheet.getRange('D4:Q' + eventSheet.getLastRow()).getValues();
  const eventTypes = settingsSheet.getRange('K5:K' + settingsSheet.getLastRow()).getValues().flat().filter(Boolean);
  
  
  const eventData = eventRange
    .map(row => {
      const eventId = row[13]; 
      const qualifier = row[12]; 
      const eventName = row[0]; 
      
      
      return (eventId && qualifier) ? { id: eventId, name: eventName } : null;
    })
    .filter(item => item !== null);

  return {
    events: eventData,
    eventTypes: eventTypes
  };
}

function getPreEventBudgets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName('Pre-Event Budget');
  
  const budgetRange = budgetSheet.getRange('A4:D' + budgetSheet.getLastRow()).getValues();
  
  const budgetData = budgetRange
    .map(row => row[0] ? { eventId: row[0], type: row[1], amount: row[2], lastModified: row[3] } : null)
    .filter(item => item !== null);

  return budgetData;
}

function getPostEventExpenses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName('Post-Event Expenses');
  
  const expensesRange = expensesSheet.getRange('A4:D' + expensesSheet.getLastRow()).getValues();
  
  const expensesData = expensesRange
    .map(row => row[0] ? { eventId: row[0], type: row[1], amount: row[2], lastModified: row[3] } : null)
    .filter(item => item !== null);

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