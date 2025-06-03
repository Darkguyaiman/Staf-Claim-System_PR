function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDClaim&Allowance')
    .setTitle('Claim Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function parseAllowance(value) {
  if (!value || typeof value !== 'string') return 0;
  if (value.toLowerCase().includes('select')) return 0;

  const match = value.match(/RM\s*([\d,.]+)/i);
  return match ? parseFloat(match[1].replace(/,/g, '')) : 0;
}

function getAllClaims() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    
    const mainSheet = ss.getSheetByName('Claim&Allowance');
    if (!mainSheet) {
      throw new Error('Claim&Allowance sheet not found');
    }
    
    const lastRow = mainSheet.getLastRow();
    if (lastRow < 4) {
      return JSON.stringify([]); 
    }
    
    const mainData = mainSheet.getRange(4, 1, lastRow - 3, 12).getValues();
    
    
    const staySheet = ss.getSheetByName('Claim&Allowance Stay');
    let stayData = [];
    if (staySheet && staySheet.getLastRow() >= 4) {
      stayData = staySheet.getRange(4, 1, staySheet.getLastRow() - 3, 4).getValues();
    }
    
    
    const extraSheet = ss.getSheetByName('Claim&Allowance Extra Claims');
    let extraData = [];
    if (extraSheet && extraSheet.getLastRow() >= 4) {
      extraData = extraSheet.getRange(4, 1, extraSheet.getLastRow() - 3, 4).getValues();
    }
    
    
    const transportSheet = ss.getSheetByName('Claim&Allowance Transportation');
    let transportData = [];
    if (transportSheet && transportSheet.getLastRow() >= 4) {
      transportData = transportSheet.getRange(4, 1, transportSheet.getLastRow() - 3, 8).getValues();
    }
    
    
    const claims = mainData.map(row => {
      const applicationNumber = row[2]; 
      
      const claim = {
        createdDate: row[0],          
        eventStatus: row[1],          
        applicationNumber: row[2],    
        personOfSubmission: row[3],
        eventName: row[4],
        timeSlot: row[5],
        mealAllowance: {
          label: row[6],
          amount: parseAllowance(row[6])
        },
        otherAllowances: {
          label: row[7],
          amount: parseAllowance(row[7])
        },

        totalClaim: row[8],
        remarks: row[9],
        eventId: row[10],
        notesOperations: row[11],
        stays: null,
        extraClaims: [],
        transportation: []
      };

      
      const stayRecords = stayData.filter(stayRow => stayRow[0] === applicationNumber);
      claim.stays = stayRecords.map(stayRow => ({
        stayType: stayRow[1],
        stayCost: stayRow[2],
        stayReceipt: stayRow[3]
      }));

      
      
      const extraClaims = extraData.filter(extraRow => extraRow[0] === applicationNumber);
      claim.extraClaims = extraClaims.map(extraRow => ({
        claimDescription: extraRow[1],
        claimAmount: extraRow[2],
        claimReceipt: extraRow[3]
      }));
      
      
      const transportClaims = transportData.filter(transportRow => transportRow[0] === applicationNumber);
      claim.transportation = transportClaims.map(transportRow => ({
        transportationType: transportRow[1],
        carType: transportRow[2],
        companyCar: transportRow[3],
        distanceTravelled: transportRow[4],
        touchNGoCost: transportRow[5],
        transportCost: transportRow[6],
        receipt: transportRow[7]
      }));
      
      return claim;
    }).filter(claim => claim.applicationNumber); 
    
    const jsonOutput = JSON.stringify(claims, null, 2);
    Logger.log(jsonOutput); 
    return jsonOutput;

  } catch (error) {
    console.error('Error in getAllClaims:', error);
    throw new Error('Failed to fetch claims data: ' + error.message);
  }
}

function getClaimDetails(applicationNumber) {
  try {
    const claims = getAllClaims();
    const claim = claims.find(c => c.applicationNumber === applicationNumber);
    
    if (!claim) {
      throw new Error('Claim not found');
    }
    
    return claim;
    
  } catch (error) {
    console.error('Error in getClaimDetails:', error);
    throw new Error('Failed to fetch claim details: ' + error.message);
  }
}


function updateClaimStatus(applicationNumber, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName('Claim&Allowance');
    
    if (!mainSheet) {
      throw new Error('Claim&Allowance sheet not found');
    }
    
    const lastRow = mainSheet.getLastRow();
    if (lastRow < 4) {
      throw new Error('No data found in sheet');
    }
    
    
    const appNumbers = mainSheet.getRange(4, 3, lastRow - 3, 1).getValues(); 
    let rowIndex = -1;
    
    for (let i = 0; i < appNumbers.length; i++) {
      if (appNumbers[i][0] === applicationNumber) {
        rowIndex = i + 4; 
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('Application number not found');
    }
    
    
    mainSheet.getRange(rowIndex, 2).setValue(newStatus);
    
    return {
      success: true,
      message: 'Status updated successfully'
    };
    
  } catch (error) {
    console.error('Error in updateClaimStatus:', error);
    return {
      success: false,
      message: 'Failed to update status: ' + error.message
    };
  }
}