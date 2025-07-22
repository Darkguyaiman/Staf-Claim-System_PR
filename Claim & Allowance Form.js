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
const testEvents = [
  {
    id: "EVT001",
    name: "Sales Event - Local (KL)",
    type: "Sales",
    location: "Kuala Lumpur",
    locationStatus: "Local",
    timeSlots: ["2025-05-18 09:00-12:00", "2025-05-19 13:00-17:00"]
  },
  {
    id: "EVT002",
    name: "Marketing Event - Outstation",
    type: "Marketing",
    location: "Penang",
    locationStatus: "outstation",
    timeSlots: ["2025-05-19 09:00-17:00"]
  },
  {
    id: "EVT003",
    name: "Exhibition - Overseas",
    type: "Exhibition",
    location: "Singapore",
    locationStatus: "overseas",
    timeSlots: ["2025-05-20 09:00-17:00", "2025-05-21 09:00-17:00"]
  },
  {
    id: "EVT004",
    name: "Demo Session - Selangor",
    type: "Demo Session",
    location: "Selangor",
    locationStatus: "local",
    timeSlots: ["2025-05-21 09:00-12:00", "2025-05-22 13:00-17:00", "2025-05-23 18:00-21:00"]
  },
  {
    id: "EVT005",
    name: "PL3D Treatment - Outstation",
    type: "PL3D",
    location: "Johor Bahru",
    locationStatus: "outstation",
    timeSlots: ["2025-05-22 09:00-17:00"]
  },
  {
    id: "EVT006",
    name: "RF Treatment - Local",
    type: "RF",
    location: "Kuala Lumpur",
    locationStatus: "local",
    timeSlots: ["2025-05-23 09:00-12:00"]
  },
  {
    id: "EVT007",
    name: "Cryo Session - Outstation",
    type: "Cryo",
    location: "Ipoh",
    locationStatus: "outstation",
    timeSlots: ["2025-05-24 13:00-17:00"]
  },
  {
    id: "EVT008",
    name: "Training - Local",
    type: "Training",
    location: "Negeri Sembilan",
    locationStatus: "local",
    timeSlots: ["2025-05-25 09:00-17:00", "2025-05-26 09:00-17:00"]
  },
  {
    id: "EVT009",
    name: "Training - Outstation",
    type: "Training",
    location: "Terengganu",
    locationStatus: "outstation",
    timeSlots: ["2025-05-26 09:00-17:00"]
  },
  {
    id: "EVT010",
    name: "Training - Sabah",
    type: "Training",
    location: "Sabah",
    locationStatus: "outstation",
    timeSlots: ["2025-05-27 09:00-17:00", "2025-05-28 09:00-17:00"]
  },
  {
    id: "EVT011",
    name: "Marketing Event - Sarawak",
    type: "Marketing",
    location: "Sarawak",
    locationStatus: "outstation",
    timeSlots: ["2025-05-29 09:00-17:00"]
  },
  {
    id: "EVT012",
    name: "CUB Treatment - Outstation",
    type: "CUB",
    location: "Melaka",
    locationStatus: "outstation",
    timeSlots: ["2025-05-30 09:00-12:00", "2025-05-31 13:00-17:00"]
  },
  {
    id: "EVT013",
    name: "PL3D-RF Combo - Outstation",
    type: "PL3D-RF",
    location: "Kedah",
    locationStatus: "outstation",
    timeSlots: ["2025-06-01 09:00-17:00"]
  },
  {
    id: "EVT014",
    name: "Sales Event - Negeri Sembilan",
    type: "Sales",
    location: "Negeri Sembilan",
    locationStatus: "local",
    timeSlots: ["2025-06-02 09:00-12:00", "2025-06-03 13:00-17:00"]
  },
  {
    id: "EVT015",
    name: "Exhibition - Overseas (Tokyo)",
    type: "Exhibition",
    location: "Tokyo, Japan",
    locationStatus: "overseas",
    timeSlots: ["2025-06-04 09:00-17:00", "2025-06-05 09:00-17:00", "2025-06-06 09:00-17:00"]
  }
];

  return testEvents;
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
