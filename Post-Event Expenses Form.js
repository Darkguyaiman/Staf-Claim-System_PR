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

function uploadFile(base64Data, mimeType, fileName, eventId) {
  try {
    const folderId = 'YOUR_FOLDER_ID';
    const folder = DriveApp.getFolderById(folderId);
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setName(fileName);
    
    const fileUrl = file.getUrl();
    
    updateEventSheetWithFileUrl(eventId, fileUrl);
    
    return {
      success: true,
      fileUrl: fileUrl,
      fileName: fileName
    };
  } catch (error) {
    console.error('Error uploading file:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}


function updateEventSheetWithFileUrl(eventId, fileUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventSheet = ss.getSheetByName('Event Creation');
    

    const eventIdRange = eventSheet.getRange('Q4:Q' + eventSheet.getLastRow()).getValues();
    
    for (let i = 0; i < eventIdRange.length; i++) {
      if (eventIdRange[i][0] === eventId) {
        const rowNumber = i + 4; 
        eventSheet.getRange(`U${rowNumber}`).setValue(fileUrl);
        break;
      }
    }
  } catch (error) {
    console.error('Error updating event sheet with file URL:', error);
    throw error;
  }
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
