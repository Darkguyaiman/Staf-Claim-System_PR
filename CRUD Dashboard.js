function doGet() {
  return HtmlService.createTemplateFromFile('CRUDDashboard')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getDashboardData() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Dashboard");
    if (!sheet) {
      return { error: true, message: 'Dashboard sheet not found' };
    }

    const eventsOverview = {
      totalEvents: sheet.getRange('C4').getValue() || 0,
      todaysEvents: sheet.getRange('C5').getValue() || 0,
      averageDevices: sheet.getRange('C6').getValue() || 0,
      budgetVariance: sheet.getRange('C7').getValue() || 0
    };

    const financialSummary = {
      totalApplications: sheet.getRange('F4').getValue() || 0,
      transportationCost: sheet.getRange('F5').getValue() || 0,
      stayCost: sheet.getRange('F6').getValue() || 0,
      mealAllowances: sheet.getRange('F7').getValue() || 0,
      extraClaims: sheet.getRange('F8').getValue() || 0,
      otherAllowances: sheet.getRange('F9').getValue() || 0,
      averageAdvanceCost: sheet.getRange('F10').getValue() || 0,
      averagePettyCash: sheet.getRange('F11').getValue() || 0
    };

    const statusRange = sheet.getRange('I4:I').getValues();
    const countRange = sheet.getRange('J4:J').getValues();

    const statusBreakdown = [];
    for (let i = 0; i < statusRange.length; i++) {
      if (statusRange[i][0] && statusRange[i][0] !== '') {
        statusBreakdown.push({
          status: statusRange[i][0],
          count: countRange[i][0] || 0
        });
      }
    }

    const daysRange = sheet.getRange('B15:B45').getValues();
    const eventsRange = sheet.getRange('C15:C45').getValues();
    const applicationsRange = sheet.getRange('D15:D45').getValues();

    const chartData = {
      days: [],
      events: [],
      applications: []
    };

    for (let i = 0; i < daysRange.length; i++) {
      if (daysRange[i][0] && daysRange[i][0] !== '') {
        let dayValue = daysRange[i][0];
        if (dayValue instanceof Date) {
          dayValue = Utilities.formatDate(dayValue, Session.getScriptTimeZone(), 'MMM d');
        }
        
        chartData.days.push(dayValue);
        chartData.events.push(eventsRange[i][0] || 0);
        chartData.applications.push(applicationsRange[i][0] || 0);
      }
    }

    
    const recentEvents = getRecentEvents();

    return {
      eventsOverview: eventsOverview,
      financialSummary: financialSummary,
      statusBreakdown: statusBreakdown,
      chartData: chartData,
      recentEvents: recentEvents
    };

  } catch (error) {
    return {
      error: true,
      message: 'Failed to retrieve dashboard data'
    };
  }
}

function getRecentEvents() {
  try {
    const eventSheet = SpreadsheetApp.getActive().getSheetByName("Event Creation");
    if (!eventSheet) {
      return { error: true, message: 'Event Creation sheet not found' };
    }

    
    const lastRow = eventSheet.getLastRow();
    
    
    if (lastRow < 4) {
      return [];
    }

    
    const statusRange = eventSheet.getRange('A4:A' + lastRow).getValues();
    const createdByRange = eventSheet.getRange('C4:C' + lastRow).getValues();
    const eventNameRange = eventSheet.getRange('D4:D' + lastRow).getValues();
    const eventTypeRange = eventSheet.getRange('E4:E' + lastRow).getValues();
    const locationRange = eventSheet.getRange('F4:F' + lastRow).getValues();
    const startDateRange = eventSheet.getRange('H4:H' + lastRow).getValues();
    const endDateRange = eventSheet.getRange('I4:I' + lastRow).getValues();

    const events = [];
    
    
    for (let i = 0; i < statusRange.length; i++) {
      
      if (!eventNameRange[i][0] || eventNameRange[i][0] === '') {
        continue;
      }

      const startDate = startDateRange[i][0];
      const endDate = endDateRange[i][0];
      
      
      let dateRange = '';
      let shortDateRange = '';
      if (startDate && endDate) {
        const startFormatted = startDate instanceof Date ? 
          Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMM d, yyyy') : 
          startDate.toString();
        const endFormatted = endDate instanceof Date ? 
          Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'MMM d, yyyy') : 
          endDate.toString();
        const startShort = startDate instanceof Date ? 
          Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMM d') : 
          startDate.toString();
        const endShort = endDate instanceof Date ? 
          Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'MMM d') : 
          endDate.toString();
        dateRange = startFormatted + ' - ' + endFormatted;
        shortDateRange = startShort + ' - ' + endShort;
      } else if (startDate) {
        dateRange = startDate instanceof Date ? 
          Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMM d, yyyy') : 
          startDate.toString();
        shortDateRange = startDate instanceof Date ? 
          Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MMM d') : 
          startDate.toString();
      }

      events.push({
        status: statusRange[i][0] || 'N/A',
        createdBy: createdByRange[i][0] || 'N/A',
        eventName: eventNameRange[i][0] || 'N/A',
        eventType: eventTypeRange[i][0] || 'N/A',
        location: locationRange[i][0] || 'N/A',
        dateRange: dateRange || 'N/A',
        shortDateRange: shortDateRange || 'N/A',
        rowIndex: i + 4 
      });
    }

    
    events.sort((a, b) => b.rowIndex - a.rowIndex);

    
    return events.slice(0, 10);

  } catch (error) {
    return { error: true, message: 'Failed to retrieve recent events data' };
  }
}