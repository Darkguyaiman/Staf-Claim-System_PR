function doGet() {
  return HtmlService.createHtmlOutputFromFile('Claim&AllowanceForm') 
    .setTitle('Claims & Allowances Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialFormData() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = {
    settings: ss.getSheetByName('Settings'),
    claim: ss.getSheetByName('Claim&Allowance'),
    event: ss.getSheetByName('Event Creation') 
  };

  const [emailRange, usernameRange, existingNumbers] = [
    sheets.settings.getRange('C5:C').getValues(),
    sheets.settings.getRange('B5:B').getValues(),
    sheets.claim.getRange('C4:C').getValues().flat().filter(String)
  ];

  let username = 'User not found';
  for (let i = 0; i < emailRange.length; i++) {
    if (emailRange[i][0] === userEmail) {
      username = usernameRange[i][0];
      break;
    }
  }

  let newNumber;
  const existingSet = new Set(existingNumbers);
  do {
    const randomNum = Math.floor(1000000 + Math.random() * 9000000);
    newNumber = 'HRM' + randomNum;
  } while (existingSet.has(newNumber));

  const userInfo = {
    email: userEmail,
    username: username
  };

  const eventData = sheets.event.getRange('A4:Q').getValues();
  const availableEvents = [];

  for (const row of eventData) {
    if (!row[3]) continue; 

    const assignedUser = row[2];
    const multipleUsers = row[12];

    if (assignedUser !== username &&
        !(multipleUsers && multipleUsers.toString().includes(username))) {
      continue; 
    }

    if (multipleUsers && assignedUser !== username) {
      const userList = multipleUsers.toString().split(',').map(u => u.trim());
      if (!userList.some(u => u === username)) {
        continue; 
      }
    }

    availableEvents.push({
      id: row[16],
      name: row[3],
      type: row[4],
      location: row[5],
      locationStatus: row[6],
      timeSlots: row[9] ? row[9].toString().split(',').map(slot => slot.trim()) : []
    });
  }

  return {
    userInfo: userInfo,
    applicationNumber: newNumber,
    availableEvents: availableEvents
  };
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
    const transportSheet = ss.getSheetByName('Claim&Allowance Transportation');
    const staySheet = ss.getSheetByName('Claim&Allowance Stay');
    const claimsSheet = ss.getSheetByName('Claim&Allowance Extra Claims');

    const totalHours = calculateTotalHours(formData.timeSlots);

    const mainRow = [
      timestamp,                                    
      'Submitted',                                  
      formData.applicationNumber,                   
      formData.username,                           
      formData.eventName,                          
      formData.timeSlots.join(', '),               
      JSON.stringify(formData.mealAllowance),      
      JSON.stringify(formData.otherAllowances),    
      formData.totalClaim,                         
      formData.remarks,                            
      formData.eventId,                            
      '',                                          
      totalHours,                                  
      formData.locationStatus                      
    ];

    const mainLastRow = mainSheet.getLastRow() + 1;
    mainSheet.getRange(mainLastRow, 1, 1, mainRow.length).setValues([mainRow]);

    if (formData.transportation?.length) {
      const transportStartRow = transportSheet.getLastRow() + 1;
      const transportRows = formData.transportation.map(item => {
        const row = [
          formData.applicationNumber,
          item.type,
          '', '', '', '', '', '', ''
        ];

        if (item.type === 'car') {
          row[2] = item.carType;
          if (item.carType === 'company') row[3] = item.companyCar || '';
          if (item.carType === 'personal') {
            row[4] = item.kilometers || '';
            row[5] = item.tngAmount || '';
            row[7] = item.receiptUrl || '';
          }
          if (item.carType === 'rental') {
            row[6] = item.rentalAmount || '';
            row[7] = item.receiptUrl || '';
          }
        } else if (item.type === 'air' || item.type === 'train') {
          row[6] = item.ticketCost || '';
          row[7] = item.receiptUrl || '';
        }
        return row;
      });
      transportSheet.getRange(transportStartRow, 1, transportRows.length, 9).setValues(transportRows);
    }

    if (formData.stay?.length) {
      const stayStartRow = staySheet.getLastRow() + 1;
      const stayRows = formData.stay
        .filter(item => ['hotel', 'homestay', 'airport', 'none'].includes(item.type))
        .map(item => [
          formData.applicationNumber,
          item.type,
          ['hotel', 'homestay'].includes(item.type) ? item.amount || '' : '',
          ['hotel', 'homestay'].includes(item.type) ? item.receiptUrl || '' : ''
        ]);
      staySheet.getRange(stayStartRow, 1, stayRows.length, 4).setValues(stayRows);
    }

    if (formData.extraClaims?.length) {
      const claimStartRow = claimsSheet.getLastRow() + 1;
      const claimRows = formData.extraClaims.map(item => [
        formData.applicationNumber,
        item.description,
        item.amount,
        item.documentUrl || ''
      ]);
      claimsSheet.getRange(claimStartRow, 1, claimRows.length, 4).setValues(claimRows);
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
