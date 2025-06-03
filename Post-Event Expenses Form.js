function doGet() {
  return HtmlService.createHtmlOutputFromFile('Post-EventExpensesForm')
    .setTitle('Post-Event Expenses Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAllEventData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('Event Creation');
  const expenseSheet = ss.getSheetByName('Post-Event Expenses');
  const budgetSheet = ss.getSheetByName('Pre-Event Budget');
  
  const statusRange = eventSheet.getRange('A4:A' + eventSheet.getLastRow()).getValues();
  const eventNameRange = eventSheet.getRange('D4:D' + eventSheet.getLastRow()).getValues();
  const eventIdRange = eventSheet.getRange('Q4:Q' + eventSheet.getLastRow()).getValues();

  const expenseEventIds = new Set(
    expenseSheet.getRange('A4:A' + expenseSheet.getLastRow()).getValues().flat()
  );

  const budgetData = {
    ids: budgetSheet.getRange('A4:A' + budgetSheet.getLastRow()).getValues(),
    types: budgetSheet.getRange('B4:B' + budgetSheet.getLastRow()).getValues(),
    amounts: budgetSheet.getRange('C4:C' + budgetSheet.getLastRow()).getValues()
  };
  
  const approvedEvents = [];
  for (let i = 0; i < statusRange.length; i++) {
    if (statusRange[i][0] === 'Approved' && !expenseEventIds.has(eventIdRange[i][0])) {
      approvedEvents.push({
        name: eventNameRange[i][0],
        id: eventIdRange[i][0]
      });
    }
  }
  
  const budgetItems = {};
  for (let i = 0; i < budgetData.ids.length; i++) {
    const eventId = budgetData.ids[i][0];
    if (!budgetItems[eventId]) {
      budgetItems[eventId] = [];
    }
    
    budgetItems[eventId].push({
      type: budgetData.types[i][0],
      amount: budgetData.amounts[i][0]
    });
  }
  
  return {
    approvedEvents: approvedEvents,
    budgetItems: budgetItems
  };
}

function submitExpenses(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expensesSheet = ss.getSheetByName('Post-Event Expenses');
  
  let lastRow = expensesSheet.getLastRow();
  if (lastRow < 4) lastRow = 3;
  
  formData.expenses.forEach(expense => {
    lastRow++;
    expensesSheet.getRange(`A${lastRow}`).setValue(formData.eventId);
    expensesSheet.getRange(`B${lastRow}`).setValue(expense.type);
    expensesSheet.getRange(`C${lastRow}`).setValue(expense.actualAmount);
  });
  
  return { success: true, message: 'Expenses submitted successfully!' };
}