function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDPettyCashManagement')
    .setTitle('Petty Cash Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getPettyCashData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash');
    if (!sheet) {
      throw new Error('Petty Cash sheet not found');
    }
    
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 4) {
      return []; 
    }
    
    
    const range = sheet.getRange(4, 2, lastRow - 3, 9); 
    const values = range.getValues();
    
    
    const processedData = values.map((row, index) => {
      return {
        rowNumber: index + 4,
        timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss') : '',
        reviewedBy: row[1] || '',
        personOfSubmission: row[2] || '',
        purpose: row[3] || '',
        transactionType: row[4] || '',
        amount: row[5] ? parseFloat(row[5]).toFixed(2) : '',
        additionalDocs: row[6] || '',
        remarks: row[7] || '',
        status: row[8] || 'Submitted' 
      };
    }).filter(row => row.timestamp); 
    
    return processedData;
  } catch (error) {
    console.error('Error getting petty cash data:', error);
    throw error;
  }
}

function getUsersBalanceData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash Users Balance');
    if (!sheet) {
      throw new Error('Petty Cash Users Balance sheet not found');
    }
    
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 4) {
      return []; 
    }
    
    
    const range = sheet.getRange(4, 2, lastRow - 3, 3); 
    const values = range.getValues();
    
    
    const processedData = values.map((row, index) => {
      const username = row[0] || '';
      const currentBalance = parseFloat(row[1]) || 0;
      const minimumBalance = parseFloat(row[2]) || 0;
      
      if (!username) return null; 
      
      
      const difference = currentBalance - minimumBalance;
      let status = '';
      let statusClass = '';
      
      if (difference < 0) {
        status = 'Overdrawn';
        statusClass = 'overdrawn';
      } else if (difference === 0) {
        status = 'Low';
        statusClass = 'low';
      } else if (difference <= 10) {
        status = 'Critical';
        statusClass = 'critical';
      } else if (difference >= 50) {
        status = 'Sufficient';
        statusClass = 'sufficient';
      } else {
        status = 'Critical';
        statusClass = 'critical';
      }
      
      return {
        username: username,
        currentBalance: currentBalance.toFixed(2),
        minimumBalance: minimumBalance.toFixed(2),
        status: status,
        statusClass: statusClass,
        difference: difference,
        rowNumber: index + 4 
      };
    }).filter(row => row !== null); 
    
    return processedData;
  } catch (error) {
    console.error('Error getting users balance data:', error);
    throw error;
  }
}

function getStatusOptions() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sheet) {
      throw new Error('Settings sheet not found');
    }
    
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) {
      return ['Submitted', 'Approved', 'Rejected']; 
    }
    
    const range = sheet.getRange(5, 6, lastRow - 4, 1); 
    const values = range.getValues();
    
    const options = values.map(row => row[0]).filter(value => value && value.toString().trim());
    
    return options.length > 0 ? options : ['Submitted', 'Approved', 'Rejected'];
  } catch (error) {
    console.error('Error getting status options:', error);
    return ['Submitted', 'Approved', 'Rejected']; 
  }
}

function getUsernames() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sheet) {
      throw new Error('Settings sheet not found');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) {
      return [];
    }
    
    
    const rangeB = sheet.getRange(5, 2, lastRow - 4, 1); 
    const rangeI = sheet.getRange(5, 9, lastRow - 4, 1); 
    
    const valuesB = rangeB.getValues().map(row => row[0]).filter(value => value && value.toString().trim());
    const valuesI = rangeI.getValues().map(row => row[0]).filter(value => value && value.toString().trim());
    
    
    const allUsernames = [...valuesB, ...valuesI];
    const uniqueUsernames = [...new Set(allUsernames)];
    
    return uniqueUsernames.sort();
  } catch (error) {
    console.error('Error getting usernames:', error);
    throw error;
  }
}


function getCurrentUsername() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    
    if (!sheet) {
      return userEmail; 
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) {
      return userEmail;
    }
    
    
    const rangeC = sheet.getRange(5, 3, lastRow - 4, 1);
    const rangeB = sheet.getRange(5, 2, lastRow - 4, 1);
    const emailsC = rangeC.getValues();
    const usernamesB = rangeB.getValues();
    
    for (let i = 0; i < emailsC.length; i++) {
      if (emailsC[i][0] === userEmail) {
        return usernamesB[i][0] || userEmail;
      }
    }
    
    
    const rangeJ = sheet.getRange(5, 10, lastRow - 4, 1);
    const rangeI = sheet.getRange(5, 9, lastRow - 4, 1);
    const emailsJ = rangeJ.getValues();
    const usernamesI = rangeI.getValues();
    
    for (let i = 0; i < emailsJ.length; i++) {
      if (emailsJ[i][0] === userEmail) {
        return usernamesI[i][0] || userEmail;
      }
    }
    
    return userEmail; 
  } catch (error) {
    console.error('Error getting current username:', error);
    return Session.getActiveUser().getEmail();
  }
}

function updateRecord(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash');
    const username = getCurrentUsername();
    
    sheet.getRange(data.rowNumber, 5).setValue(data.purpose);
    sheet.getRange(data.rowNumber, 7).setValue(data.amount);
    sheet.getRange(data.rowNumber, 9).setValue(data.remarks);
    sheet.getRange(data.rowNumber, 10).setValue(data.status);
    sheet.getRange(data.rowNumber, 3).setValue(username);
    
    return true;
  } catch (error) {
    console.error('Error updating record:', error);
    throw error;
  }
}

function deleteRecord(rowNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash');
    sheet.deleteRow(rowNumber);
    return true;
  } catch (error) {
    console.error('Error deleting record:', error);
    throw error;
  }
}

function processTopup(username, amount) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash');
    const currentUsername = getCurrentUsername();
    const timestamp = new Date();
    
    
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    
    sheet.getRange(newRow, 2).setValue(timestamp); 
    sheet.getRange(newRow, 3).setValue(currentUsername); 
    sheet.getRange(newRow, 4).setValue(username); 
    sheet.getRange(newRow, 5).setValue(`Filling up petty cash for ${username}`); 
    sheet.getRange(newRow, 6).setValue('Top Up'); 
    sheet.getRange(newRow, 7).setValue(amount); 
    sheet.getRange(newRow, 8).setValue(''); 
    sheet.getRange(newRow, 9).setValue(''); 
    sheet.getRange(newRow, 10).setValue('Approved'); 
    
    return true;
  } catch (error) {
    console.error('Error processing top-up:', error);
    throw error;
  }
}

function updateMinimumBalance(username, newMinimum) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash Users Balance');
    if (!sheet) {
      throw new Error('Petty Cash Users Balance sheet not found');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      throw new Error('No user data found');
    }
    
    
    const range = sheet.getRange(4, 2, lastRow - 3, 1); 
    const usernames = range.getValues();
    
    for (let i = 0; i < usernames.length; i++) {
      if (usernames[i][0] === username) {
        const rowNumber = i + 4;
        sheet.getRange(rowNumber, 4).setValue(newMinimum); 
        return true;
      }
    }
    
    throw new Error('User not found');
  } catch (error) {
    console.error('Error updating minimum balance:', error);
    throw error;
  }
}
