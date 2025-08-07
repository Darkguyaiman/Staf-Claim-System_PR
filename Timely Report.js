function doGet() {
  return HtmlService.createHtmlOutputFromFile('TimelyReport')
    .setTitle('Timely Report Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function generateReport(startDate, endDate) {
  try {
    const start = parseLocalDate(startDate);
    const end = parseLocalDate(endDate);

    const eventsData = getEventsData(start, end);
    const advancePurchasesData = getAdvancePurchasesData(start, end);
    const pettyCashData = getPettyCashData(start, end);
    const claimsData = getClaimsAndAllowancesData(start, end);

    const report = {
      dateRange: {
        start: startDate,
        end: endDate
      },
      events: eventsData,
      advancePurchases: advancePurchasesData,
      pettyCash: pettyCashData,
      claimsAndAllowances: claimsData
    };

    return {
      success: true,
      data: report
    };
  } catch (error) {
    console.error('Error generating report:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function parseLocalDate(dateString) {
  const [year, month, day] = dateString.split('-').map(Number);
  return new Date(year, month - 1, day);
}


function getEventsData(startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Event Creation');
    if (!sheet) {
      throw new Error('Event Creation sheet not found');
    }

    const statusRange = sheet.getRange('A4:A').getValues();
    const dateRange = sheet.getRange('B4:B').getValues();
    const budgetRange = sheet.getRange('O4:O').getValues();
    const expenseRange = sheet.getRange('S4:S').getValues();
    
    let eventCount = 0;
    let totalBudget = 0;
    let totalExpenses = 0;

    for (let i = 0; i < statusRange.length; i++) {
      const status = statusRange[i][0];
      const date = dateRange[i][0];
      const budget = budgetRange[i][0];
      const expense = expenseRange[i][0];

      if (!status || !date) continue;

      if (status.toString().toLowerCase() === 'approved' && 
          date instanceof Date && 
          date >= startDate && 
          date <= endDate) {
        
        eventCount++;

        if (typeof budget === 'number' && !isNaN(budget)) {
          totalBudget += budget;
        }

        if (typeof expense === 'number' && !isNaN(expense)) {
          totalExpenses += expense;
        }
      }
    }
    
    return {
      count: eventCount,
      totalBudget: totalBudget,
      totalExpenses: totalExpenses,
      difference: totalBudget - totalExpenses
    };
    
  } catch (error) {
    console.error('Error in getEventsData:', error);
    return {
      count: 0,
      totalBudget: 0,
      totalExpenses: 0,
      difference: 0,
      error: error.toString()
    };
  }
}

function getAdvancePurchasesData(startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advance Purchases');
    if (!sheet) {
      throw new Error('Advance Purchases sheet not found');
    }

    const statusRange = sheet.getRange('A4:A').getValues();
    const dateRange = sheet.getRange('B4:B').getValues();
    const priceRange = sheet.getRange('F4:F').getValues();
    
    let purchaseCount = 0;
    let totalPrice = 0;

    for (let i = 0; i < statusRange.length; i++) {
      const status = statusRange[i][0];
      const date = dateRange[i][0];
      const price = priceRange[i][0];

      if (!status || !date) continue;

      if (status.toString().toLowerCase() === 'approved' && 
          date instanceof Date && 
          date >= startDate && 
          date <= endDate) {
        
        purchaseCount++;

        if (typeof price === 'number' && !isNaN(price)) {
          totalPrice += price;
        }
      }
    }
    
    return {
      count: purchaseCount,
      totalPrice: totalPrice
    };
    
  } catch (error) {
    console.error('Error in getAdvancePurchasesData:', error);
    return {
      count: 0,
      totalPrice: 0,
      error: error.toString()
    };
  }
}

function getPettyCashData(startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Petty Cash');
    if (!sheet) {
      throw new Error('Petty Cash sheet not found');
    }

    const dateRange = sheet.getRange('B4:B').getValues();
    const typeRange = sheet.getRange('F4:F').getValues();
    const amountRange = sheet.getRange('G4:G').getValues();
    const statusRange = sheet.getRange('J4:J').getValues();
    
    let transactionCount = 0;
    let topUpCount = 0;
    let totalTransactionAmount = 0;
    let totalTopUpAmount = 0;

    for (let i = 0; i < dateRange.length; i++) {
      const date = dateRange[i][0];
      const type = typeRange[i][0];
      const amount = amountRange[i][0];
      const status = statusRange[i][0];

      if (!date || !status) continue;

      if (date instanceof Date && 
          date >= startDate && 
          date <= endDate &&
          status.toString().toLowerCase() === 'approved') {
        
        const typeStr = type ? type.toString().toLowerCase() : '';
        const amountNum = typeof amount === 'number' && !isNaN(amount) ? amount : 0;
        
        if (typeStr === 'top up') {

          topUpCount++;
          totalTopUpAmount += amountNum;
        } else {

          transactionCount++;
          totalTransactionAmount += amountNum;
        }
      }
    }
    
    return {
      transactionCount: transactionCount,
      topUpCount: topUpCount,
      totalTransactionAmount: totalTransactionAmount,
      totalTopUpAmount: totalTopUpAmount
    };
    
  } catch (error) {
    console.error('Error in getPettyCashData:', error);
    return {
      transactionCount: 0,
      topUpCount: 0,
      totalTransactionAmount: 0,
      totalTopUpAmount: 0,
      error: error.toString()
    };
  }
}

function getClaimsAndAllowancesData(startDate, endDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance');
    if (!sheet) {
      throw new Error('Claim&Allowance sheet not found');
    }

    const dateRange = sheet.getRange('A4:A').getValues();
    const statusRange = sheet.getRange('B4:B').getValues();
    const applicationNumberRange = sheet.getRange('C4:C').getValues();
    const claimAmountRange = sheet.getRange('I4:I').getValues();
    const mealAllowanceRange = sheet.getRange('G4:G').getValues();
    const otherAllowanceRange = sheet.getRange('H4:H').getValues();
    
    let applicationCount = 0;
    let totalClaimAmount = 0;
    let totalMealAllowance = 0;
    let totalOtherAllowance = 0;
    let validApplicationNumbers = [];

    for (let i = 0; i < dateRange.length; i++) {
      const date = dateRange[i][0];
      const status = statusRange[i][0];
      const applicationNumber = applicationNumberRange[i][0];
      const claimAmount = claimAmountRange[i][0];
      const mealAllowanceJson = mealAllowanceRange[i][0];
      const otherAllowanceJson = otherAllowanceRange[i][0];

      if (!date || !status) continue;

      if (date instanceof Date && 
          date >= startDate && 
          date <= endDate &&
          status.toString().toLowerCase() === 'approved') {
        
        applicationCount++;

        if (applicationNumber) {
          validApplicationNumbers.push(applicationNumber);
        }

        if (typeof claimAmount === 'number' && !isNaN(claimAmount)) {
          totalClaimAmount += claimAmount;
        }

        if (mealAllowanceJson) {
          const mealAmount = extractTotalAmountFromJson(mealAllowanceJson.toString());
          totalMealAllowance += mealAmount;
        }

        if (otherAllowanceJson) {
          const otherAmount = extractTotalAmountFromJson(otherAllowanceJson.toString());
          totalOtherAllowance += otherAmount;
        }
      }
    }

    const stayData = getStayAmountsData(validApplicationNumbers);
    const extraClaimsData = getExtraClaimsData(validApplicationNumbers);
    const transportationData = getTransportationSummaryData(validApplicationNumbers);
    
    return {
      applicationCount: applicationCount,
      totalClaimAmount: totalClaimAmount,
      totalMealAllowance: totalMealAllowance,
      totalOtherAllowance: totalOtherAllowance,
      stayData: stayData,
      extraClaimsAmount: extraClaimsData,
      transportationData: transportationData,
      validApplicationNumbers: validApplicationNumbers
    };
    
  } catch (error) {
    console.error('Error in getClaimsAndAllowancesData:', error);
    return {
      applicationCount: 0,
      totalClaimAmount: 0,
      totalMealAllowance: 0,
      totalOtherAllowance: 0,
      stayData: { totalAmount: 0, breakdown: [] },
      extraClaimsAmount: 0,
      transportationData: { totalAmount: 0, details: [] },
      validApplicationNumbers: [],
      error: error.toString()
    };
  }
}

function extractTotalAmountFromJson(jsonString) {
  try {

    const totalAmountMatch = jsonString.match(/"totalAmount"\s*:\s*(\d+(?:\.\d+)?)/);
    if (totalAmountMatch) {
      return parseFloat(totalAmountMatch[1]) || 0;
    }
    return 0;
  } catch (error) {
    console.error('Error parsing JSON:', error);
    return 0;
  }
}

function getStayAmountsData(validApplicationNumbers) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance Stay');
    if (!sheet || validApplicationNumbers.length === 0) {
      return { totalAmount: 0, breakdown: [] };
    }
    
    const applicationNumberRange = sheet.getRange('A4:A').getValues();
    const stayTypeRange = sheet.getRange('B4:B').getValues();
    const stayAmountRange = sheet.getRange('C4:C').getValues();
    
    let totalStayAmount = 0;
    let stayTypeBreakdown = {};

    for (let i = 0; i < applicationNumberRange.length; i++) {
      const applicationNumber = applicationNumberRange[i][0];
      const stayType = stayTypeRange[i][0];
      const stayAmount = stayAmountRange[i][0];

      if (validApplicationNumbers.includes(applicationNumber)) {
        const amount = typeof stayAmount === 'number' && !isNaN(stayAmount) ? stayAmount : 0;
        totalStayAmount += amount;

        if (stayType) {
          const type = stayType.toString();
          if (!stayTypeBreakdown[type]) {
            stayTypeBreakdown[type] = 0;
          }
          stayTypeBreakdown[type] += amount;
        }
      }
    }

    const breakdown = Object.keys(stayTypeBreakdown).map(type => ({
      type: type,
      amount: stayTypeBreakdown[type]
    }));
    
    return {
      totalAmount: totalStayAmount,
      breakdown: breakdown
    };
    
  } catch (error) {
    console.error('Error in getStayAmountsData:', error);
    return { totalAmount: 0, breakdown: [] };
  }
}

function getExtraClaimsData(validApplicationNumbers) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Claim&Allowance Extra Claims');
    if (!sheet || validApplicationNumbers.length === 0) {
      return 0;
    }
    
    const applicationNumberRange = sheet.getRange('A4:A').getValues();
    const extraClaimAmountRange = sheet.getRange('C4:C').getValues();
    
    let totalExtraClaims = 0;

    for (let i = 0; i < applicationNumberRange.length; i++) {
      const applicationNumber = applicationNumberRange[i][0];
      const extraClaimAmount = extraClaimAmountRange[i][0];

      if (validApplicationNumbers.includes(applicationNumber)) {
        const amount = typeof extraClaimAmount === 'number' && !isNaN(extraClaimAmount) ? extraClaimAmount : 0;
        totalExtraClaims += amount;
      }
    }
    
    return totalExtraClaims;
    
  } catch (error) {
    console.error('Error in getExtraClaimsData:', error);
    return 0;
  }
}

function getTransportationSummaryData(validApplicationNumbers) {
  try {
    let totalTransportationAmount = 0;
    let allTransportationDetails = [];
    
    for (let appNumber of validApplicationNumbers) {
      const transportDetails = getClaimTransportationDetails(appNumber);
      allTransportationDetails = allTransportationDetails.concat(transportDetails);

      for (let transport of transportDetails) {
        totalTransportationAmount += (transport.touchNGoCost || 0) + (transport.transportCost || 0);
      }
    }
    
    return {
      totalAmount: totalTransportationAmount,
      details: allTransportationDetails
    };
    
  } catch (error) {
    console.error('Error in getTransportationSummaryData:', error);
    return { totalAmount: 0, details: [] };
  }
}

function getClaimTransportationDetails(applicationNumber) {
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
