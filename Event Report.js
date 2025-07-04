function doGet(e) {
  var eventId = e.parameter.id;
  
  if (!eventId) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Event Report Generator</title>
          <base target="_top">
          <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
          <style>
            :root {
              --color-primary: #2D0C31;
              --color-secondary: #004D4D;
              --color-accent: #00B894;
              --color-light: #B8B5FF;
              --color-lighter: #B8FFB5;
              --color-dark: #333;
              --color-gray: #6c757d;
              --color-light-gray: #e9ecef;
              --color-white: #fff;
              --color-danger: #dc3545;
              --color-success: #28a745;
              --border-radius: 10px;
              --box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
              --transition: all 0.3s ease;
            }
            
            body {
              font-family: Arial, sans-serif;
              display: flex;
              justify-content: center;
              align-items: center;
              height: 100vh;
              margin: 0;
              background: linear-gradient(135deg, var(--color-light) 0%, var(--color-lighter) 100%);
            }
            .container {
              text-align: center;
              background: var(--color-white);
              padding: 40px;
              border-radius: var(--border-radius);
              box-shadow: var(--box-shadow);
              max-width: 500px;
              width: 90%;
            }
            .error-icon {
              font-size: 48px;
              color: var(--color-danger);
              margin-bottom: 20px;
            }
            h2 {
              color: var(--color-primary);
              margin-bottom: 10px;
            }
            p {
              color: var(--color-gray);
              font-size: 16px;
              line-height: 1.6;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="error-icon"><i class="fas fa-exclamation-triangle"></i></div>
            <h2>No Event Selected</h2>
            <p>Please choose an event on the Event Creation Tab and come back to get your report made.</p>
          </div>
        </body>
      </html>
    `);
  }
  
  var eventData = getEventData(eventId);
  
  if (!eventData) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Event Report Generator</title>
          <base target="_top">
          <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
          <style>
            :root {
              --color-primary: #2D0C31;
              --color-secondary: #004D4D;
              --color-accent: #00B894;
              --color-light: #B8B5FF;
              --color-lighter: #B8FFB5;
              --color-dark: #333;
              --color-gray: #6c757d;
              --color-light-gray: #e9ecef;
              --color-white: #fff;
              --color-danger: #dc3545;
              --color-success: #28a745;
              --border-radius: 10px;
              --box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
              --transition: all 0.3s ease;
            }
            
            body {
              font-family: Arial, sans-serif;
              display: flex;
              justify-content: center;
              align-items: center;
              height: 100vh;
              margin: 0;
              background: linear-gradient(135deg, var(--color-light) 0%, var(--color-lighter) 100%);
            }
            .container {
              text-align: center;
              background: var(--color-white);
              padding: 40px;
              border-radius: var(--border-radius);
              box-shadow: var(--box-shadow);
              max-width: 500px;
              width: 90%;
            }
            .error-icon {
              font-size: 48px;
              color: var(--color-danger);
              margin-bottom: 20px;
            }
            h2 {
              color: var(--color-primary);
              margin-bottom: 10px;
            }
            p {
              color: var(--color-gray);
              font-size: 16px;
              line-height: 1.6;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="error-icon"><i class="fas fa-times-circle"></i></div>
            <h2>Event Not Found</h2>
            <p>The event with ID "${eventId}" could not be found.</p>
          </div>
        </body>
      </html>
    `);
  }
  
  var template = HtmlService.createTemplateFromFile('EventReport');
  template.eventData = JSON.stringify(eventData);
  template.eventId = eventId;
  
  return template.evaluate()
    .setTitle('Event Report')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getEventData(eventId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Event Creation');
    if (!sheet) {
      Logger.log('Event Creation sheet not found');
      return null;
    }
    
    var lastRow = sheet.getLastRow();
    var eventIds = sheet.getRange('Q4:Q' + lastRow).getValues();
    
    for (var i = 0; i < eventIds.length; i++) {
      if (eventIds[i][0] == eventId) {
        var rowIndex = i + 4;
        
        
        var eventRow = sheet.getRange(rowIndex + ':' + rowIndex).getValues()[0];
        
        var eventData = {
          eventId: eventId,
          eventLeader: eventRow[2] || '', 
          eventName: eventRow[3] || '', 
          eventType: eventRow[4] || '', 
          location: eventRow[5] || '', 
          locationStatus: eventRow[6] || '', 
          startDate: eventRow[7] || '', 
          endDate: eventRow[8] || '', 
          timeslots: eventRow[9] || '', 
          supportingDocument: eventRow[10] || '', 
          officeAssets: eventRow[11] || '', 
          colleagues: eventRow[12] || '', 
          affiliatedCompany: eventRow[13] || '', 
          totalBudget: eventRow[14] || 0, 
          quotationDocument: eventRow[15] || '', 
          totalExpense: eventRow[18] || 0, 
          hasBudgetData: !!(eventRow[15] && eventRow[15].toString().trim())
        };
        
        
        if (eventData.hasBudgetData) {
          eventData.budgetBreakdown = getBudgetBreakdown(eventId);
          eventData.expenseBreakdown = getExpenseBreakdown(eventId);
        }
        
        eventData.advancePurchases = getAdvancePurchases(eventId);
        eventData.claimsAllowances = getClaimsAllowances(eventId);
        
        return eventData;
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('Error getting event data: ' + error.toString());
    return null;
  }
}

function getBudgetBreakdown(eventId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pre-Event Budget');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:C' + lastRow).getValues();
    var budgetItems = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == eventId) {
        budgetItems.push({
          budgetType: data[i][1] || '',
          budgetAmount: data[i][2] || 0
        });
      }
    }
    
    return budgetItems;
  } catch (error) {
    Logger.log('Error getting budget breakdown: ' + error.toString());
    return [];
  }
}

function getExpenseBreakdown(eventId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Post-Event Expenses');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:C' + lastRow).getValues();
    var expenseItems = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == eventId) {
        expenseItems.push({
          expenseType: data[i][1] || '',
          expenseAmount: data[i][2] || 0
        });
      }
    }
    
    return expenseItems;
  } catch (error) {
    Logger.log('Error getting expense breakdown: ' + error.toString());
    return [];
  }
}

function getAdvancePurchases(eventId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advance Purchases');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:J' + lastRow).getValues();
    var purchases = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][9] == eventId) { 
        purchases.push({
          status: data[i][0] || '',
          person: data[i][2] || '', 
          item: data[i][4] || '', 
          price: data[i][5] || 0, 
          remarks: data[i][6] || '', 
          invoiceDocument: data[i][8] || '' 
        });
      }
    }
    
    return purchases;
  } catch (error) {
    Logger.log('Error getting advance purchases: ' + error.toString());
    return [];
  }
}

function getClaimsAllowances(eventId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:M' + lastRow).getValues();
    var claims = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][10] == eventId) { 
        var applicationNumber = data[i][2] || ''; 
        
        var claim = {
          applicationStatus: data[i][1] || '', 
          applicationNumber: applicationNumber,
          person: data[i][3] || '', 
          timeslot: data[i][5] || '', 
          mealAllowance: data[i][6] || '', 
          otherAllowances: data[i][7] || '', 
          totalClaim: data[i][8] || 0, 
          totalTimeOnDuty: data[i][12] || '', 
          remarks: data[i][9] || '', 
          stayDetails: getStayDetails(applicationNumber),
          transportationDetails: getTransportationDetails(applicationNumber),
          extraClaims: getExtraClaims(applicationNumber)
        };
        
        claims.push(claim);
      }
    }
    
    return claims;
  } catch (error) {
    Logger.log('Error getting claims allowances: ' + error.toString());
    return [];
  }
}

function getStayDetails(applicationNumber) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance Stay');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:D' + lastRow).getValues();
    var stayDetails = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == applicationNumber) { 
        stayDetails.push({
          stayType: data[i][1] || '', 
          stayCost: data[i][2] || 0, 
          stayReceipt: data[i][3] || '' 
        });
      }
    }
    
    return stayDetails;
  } catch (error) {
    Logger.log('Error getting stay details: ' + error.toString());
    return [];
  }
}

function getTransportationDetails(applicationNumber) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance Transportation');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:H' + lastRow).getValues();
    var transportDetails = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == applicationNumber) { 
        transportDetails.push({
          transportationType: data[i][1] || '', 
          carType: data[i][2] || '', 
          companyCar: data[i][3] || '', 
          distanceTravelled: data[i][4] || '', 
          touchNGoCost: data[i][5] || 0, 
          transportCost: data[i][6] || 0, 
          transportReceipt: data[i][7] || '' 
        });
      }
    }
    
    return transportDetails;
  } catch (error) {
    Logger.log('Error getting transportation details: ' + error.toString());
    return [];
  }
}

function getExtraClaims(applicationNumber) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance Extra Claims');
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    
    var data = sheet.getRange('A4:D' + lastRow).getValues();
    var extraClaims = [];
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] == applicationNumber) { 
        extraClaims.push({
          claimDescription: data[i][1] || '', 
          claimAmount: data[i][2] || 0, 
          claimReceipt: data[i][3] || '' 
        });
      }
    }
    
    return extraClaims;
  } catch (error) {
    Logger.log('Error getting extra claims: ' + error.toString());
    return [];
  }
}
