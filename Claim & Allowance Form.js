function doGet() {
  return HtmlService.createHtmlOutputFromFile('Claim&AllowanceForm')
    .setTitle('Claims & Allowances Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getUserInfo() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const emailRange = settingsSheet.getRange('C5:C').getValues();
  const usernameRange = settingsSheet.getRange('B5:B').getValues();
  
  for (let i = 0; i < emailRange.length; i++) {
    if (emailRange[i][0] === userEmail) {
      return {
        email: userEmail,
        username: usernameRange[i][0]
      };
    }
  }
  
  return {
    email: userEmail,
    username: 'User not found'
  };
}


function generateApplicationNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const claimSheet = ss.getSheetByName('Claim&Allowance');
  const existingNumbers = claimSheet.getRange('C4:C').getValues().flat().filter(String);
  
  let newNumber;
  do {
    
    const randomNum = Math.floor(1000000 + Math.random() * 9000000);
    newNumber = 'HRM' + randomNum;
  } while (existingNumbers.includes(newNumber));
  
  return newNumber;
}


function getAvailableEvents(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('Event Creation');
  
  const eventData = eventSheet.getRange('A4:M').getValues();
  const availableEvents = [];
  
  for (let i = 0; i < eventData.length; i++) {
    if (!eventData[i][3]) continue; 
    
    const eventId = eventData[i][0]; 
    const eventName = eventData[i][3]; 
    const eventType = eventData[i][4]; 
    const eventLocation = eventData[i][5]; 
    const locationStatus = eventData[i][6]; 
    const timeSlots = eventData[i][9]; 
    const assignedUser = eventData[i][2]; 
    const multipleUsers = eventData[i][12]; 
    
    
    if (assignedUser === username || 
        (multipleUsers && multipleUsers.toString().split(',').some(u => u.trim() === username))) {
      availableEvents.push({
        id: eventId,
        name: eventName,
        type: eventType,
        location: eventLocation,
        locationStatus: locationStatus,
        timeSlots: timeSlots ? timeSlots.toString().split(',').map(slot => slot.trim()) : []
      });
    }
  }
  
  return availableEvents;
}




function getCompanyCars() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  return settingsSheet.getRange('N5:N').getValues().flat().filter(String);
}


function processFileUpload(fileObject, fileName, folderId) {
  try {
    const decodedBytes = Utilities.base64Decode(fileObject.bytes);
    const blob = Utilities.newBlob(decodedBytes, fileObject.mimeType, fileObject.fileName);

    
    if (blob.getBytes().length > 5 * 1024 * 1024) {
      return {
        success: false,
        error: 'File size exceeds 5MB limit'
      };
    }

    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setName(fileName);

    return {
      success: true,
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}



function submitFormData(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();
    
    
    const mainSheet = ss.getSheetByName('Claim&Allowance');
    const mainLastRow = mainSheet.getLastRow();
    
    mainSheet.getRange(mainLastRow + 1, 1).setValue(timestamp); 
    mainSheet.getRange(mainLastRow + 1, 2).setValue('Submitted'); 
    mainSheet.getRange(mainLastRow + 1, 3).setValue(formData.applicationNumber); 
    mainSheet.getRange(mainLastRow + 1, 4).setValue(formData.username); 
    mainSheet.getRange(mainLastRow + 1, 5).setValue(formData.eventName); 
    mainSheet.getRange(mainLastRow + 1, 6).setValue(formData.timeSlots.join(', ')); 
    mainSheet.getRange(mainLastRow + 1, 7).setValue(formData.mealAllowance.type); 
    mainSheet.getRange(mainLastRow + 1, 8).setValue(formData.otherAllowances.type); 
    mainSheet.getRange(mainLastRow + 1, 9).setValue(formData.totalClaim); 
    mainSheet.getRange(mainLastRow + 1, 10).setValue(formData.remarks); 
    mainSheet.getRange(mainLastRow + 1, 11).setValue(formData.eventId); 
    
    
    if (formData.transportation && formData.transportation.length > 0) {
      const transportSheet = ss.getSheetByName('Claim&Allowance Transportation');
      
      formData.transportation.forEach(function(item) {
        const transportLastRow = transportSheet.getLastRow();
        
        transportSheet.getRange(transportLastRow + 1, 1).setValue(formData.applicationNumber); 
        transportSheet.getRange(transportLastRow + 1, 2).setValue(item.type); 
        
        if (item.type === 'car') {
          transportSheet.getRange(transportLastRow + 1, 3).setValue(item.carType); 
          
          if (item.carType === 'company') {
            transportSheet.getRange(transportLastRow + 1, 4).setValue(item.companyCar); 
          } else if (item.carType === 'personal') {
            transportSheet.getRange(transportLastRow + 1, 5).setValue(item.kilometers); 
            transportSheet.getRange(transportLastRow + 1, 6).setValue(item.tngAmount); 
            if (item.receiptUrl) {
              transportSheet.getRange(transportLastRow + 1, 8).setValue(item.receiptUrl); 
            }
          } else if (item.carType === 'rental') {
            transportSheet.getRange(transportLastRow + 1, 7).setValue(item.rentalAmount); 
            if (item.receiptUrl) {
              transportSheet.getRange(transportLastRow + 1, 8).setValue(item.receiptUrl); 
            }
          }
        } else if (item.type === 'air' || item.type === 'train') {
          transportSheet.getRange(transportLastRow + 1, 7).setValue(item.ticketCost); 
          if (item.receiptUrl) {
            transportSheet.getRange(transportLastRow + 1, 8).setValue(item.receiptUrl); 
          }
        }
      });
    }
    
    
    if (formData.stay && formData.stay.length > 0) {
      const staySheet = ss.getSheetByName('Claim&Allowance Stay');
      
      formData.stay.forEach(function(item) {
        if (item.type === 'hotel' || item.type === 'homestay' || item.type === 'airport' || item.type === 'none') {
          const stayLastRow = staySheet.getLastRow();
          
          staySheet.getRange(stayLastRow + 1, 1).setValue(formData.applicationNumber); 
          staySheet.getRange(stayLastRow + 1, 2).setValue(item.type); 
          
          if (item.type === 'hotel' || item.type === 'homestay') {
            staySheet.getRange(stayLastRow + 1, 3).setValue(item.amount); 
            if (item.receiptUrl) {
              staySheet.getRange(stayLastRow + 1, 4).setValue(item.receiptUrl); 
            }
          }
        }
      });
    }
    
    
    if (formData.extraClaims && formData.extraClaims.length > 0) {
      const claimsSheet = ss.getSheetByName('Claim&Allowance Extra Claims');
      
      formData.extraClaims.forEach(function(item) {
        const claimsLastRow = claimsSheet.getLastRow();
        
        claimsSheet.getRange(claimsLastRow + 1, 1).setValue(formData.applicationNumber); 
        claimsSheet.getRange(claimsLastRow + 1, 2).setValue(item.description); 
        claimsSheet.getRange(claimsLastRow + 1, 3).setValue(item.amount); 
        if (item.documentUrl) {
          claimsSheet.getRange(claimsLastRow + 1, 4).setValue(item.documentUrl); 
        }
      });
    }
    
    return {
      success: true,
      message: 'Form submitted successfully',
      applicationNumber: formData.applicationNumber
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}