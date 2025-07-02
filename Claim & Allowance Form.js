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
    const mainLastRow = mainSheet.getLastRow() + 1;

    const totalHours = calculateTotalHours(formData.timeSlots);

    const mainRow = [
      timestamp,
      'Submitted',
      formData.applicationNumber,
      formData.username,
      formData.eventName,
      formData.timeSlots.join(', '),
      formData.mealAllowance.type,
      formData.otherAllowances.type,
      formData.totalClaim,
      formData.remarks,
      formData.eventId,
      '',
      totalHours
    ];

    mainSheet.getRange(mainLastRow, 1, 1, mainRow.length).setValues([mainRow]);

    if (formData.transportation?.length > 0) {
      const transportSheet = ss.getSheetByName('Claim&Allowance Transportation');
      const transportRows = formData.transportation.map(item => {
        const row = [
          formData.applicationNumber,
          item.type,
          '', '', '', '', '', '', ''
        ];

        if (item.type === 'car') {
          row[2] = item.carType;
          if (item.carType === 'company') {
            row[3] = item.companyCar || '';
          } else if (item.carType === 'personal') {
            row[4] = item.kilometers || '';
            row[5] = item.tngAmount || '';
            row[7] = item.receiptUrl || '';
          } else if (item.carType === 'rental') {
            row[6] = item.rentalAmount || '';
            row[7] = item.receiptUrl || '';
          }
        } else if (item.type === 'air' || item.type === 'train') {
          row[6] = item.ticketCost || '';
          row[7] = item.receiptUrl || '';
        }

        return row;
      });

      if (transportRows.length > 0) {
        transportSheet.getRange(transportSheet.getLastRow() + 1, 1, transportRows.length, 9).setValues(transportRows);
      }
    }

    if (formData.stay?.length > 0) {
      const staySheet = ss.getSheetByName('Claim&Allowance Stay');
      const stayRows = formData.stay.map(item => {
        if (['hotel', 'homestay', 'airport', 'none'].includes(item.type)) {
          return [
            formData.applicationNumber,
            item.type,
            (item.type === 'hotel' || item.type === 'homestay') ? item.amount || '' : '',
            (item.type === 'hotel' || item.type === 'homestay') ? item.receiptUrl || '' : ''
          ];
        }
      }).filter(row => row);

      if (stayRows.length > 0) {
        staySheet.getRange(staySheet.getLastRow() + 1, 1, stayRows.length, 4).setValues(stayRows);
      }
    }

    if (formData.extraClaims?.length > 0) {
      const claimsSheet = ss.getSheetByName('Claim&Allowance Extra Claims');
      const claimRows = formData.extraClaims.map(item => [
        formData.applicationNumber,
        item.description,
        item.amount,
        item.documentUrl || ''
      ]);

      if (claimRows.length > 0) {
        claimsSheet.getRange(claimsSheet.getLastRow() + 1, 1, claimRows.length, 4).setValues(claimRows);
      }
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

function calculateTotalHours(timeSlots) {
  let total = 0;
  timeSlots.forEach(slot => {
    const parts = slot.trim().split(' ');
    if (parts.length !== 2) return;
    const [startTime, endTime] = parts[1].split('-');
    const date = parts[0];

    const start = new Date(`${date}T${startTime}:00`);
    const end = new Date(`${date}T${endTime}:00`);
    const diffMs = end - start;

    if (!isNaN(diffMs)) {
      total += diffMs / (1000 * 60 * 60);
    }
  });

  return Math.round(total * 100) / 100;
}
