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

    const values = sheet.getRange('C4:C7').getValues().flat();
    const financialValues = sheet.getRange('F4:F11').getValues().flat();
    const statusRange = sheet.getRange('I4:J').getValues();
    const chartValues = sheet.getRange('B15:D45').getValues();

    const eventsOverview = {
      totalEvents: values[0] || 0,
      todaysEvents: values[1] || 0,
      averageDevices: values[2] || 0,
      budgetVariance: values[3] || 0
    };

    const financialSummary = {
      totalApplications: financialValues[0] || 0,
      transportationCost: financialValues[1] || 0,
      stayCost: financialValues[2] || 0,
      mealAllowances: financialValues[3] || 0,
      extraClaims: financialValues[4] || 0,
      otherAllowances: financialValues[5] || 0,
      averageAdvanceCost: financialValues[6] || 0,
      averagePettyCash: financialValues[7] || 0
    };

    const statusBreakdown = statusRange
      .filter(row => row[0])
      .map(row => ({ status: row[0], count: row[1] || 0 }));

    const chartData = {
      days: [],
      events: [],
      applications: []
    };

    const timeZone = Session.getScriptTimeZone();
    for (const row of chartValues) {
      const [day, eventCount, applicationCount] = row;
      if (!day) continue;
      const formattedDay = day instanceof Date
        ? Utilities.formatDate(day, timeZone, 'MMM d')
        : day.toString();
      chartData.days.push(formattedDay);
      chartData.events.push(eventCount || 0);
      chartData.applications.push(applicationCount || 0);
    }

    const recentEvents = getRecentEvents();

    return {
      eventsOverview,
      financialSummary,
      statusBreakdown,
      chartData,
      recentEvents
    };

  } catch (error) {
    console.error('Dashboard error:', error);
    return {
      error: true,
      message: 'Failed to retrieve dashboard data'
    };
  }
}

function getRecentEvents() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Event Creation");
    if (!sheet) {
      return { error: true, message: 'Event Creation sheet not found' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];

    const range = sheet.getRange(`A4:I${lastRow}`).getValues();
    const events = [];

    const timeZone = Session.getScriptTimeZone();

    for (let i = 0; i < range.length; i++) {
      const row = range[i];
      const [status, , createdBy, eventName, eventType, location, , startDate, endDate] = row;

      if (!eventName) continue;

      const formatDate = (date, format) =>
        date instanceof Date
          ? Utilities.formatDate(date, timeZone, format)
          : date?.toString() || '';

      const fullStart = formatDate(startDate, 'MMM d, yyyy');
      const fullEnd = formatDate(endDate, 'MMM d, yyyy');
      const shortStart = formatDate(startDate, 'MMM d');
      const shortEnd = formatDate(endDate, 'MMM d');

      let dateRange = fullStart;
      let shortDateRange = shortStart;

      if (startDate && endDate) {
        dateRange += ' - ' + fullEnd;
        shortDateRange += ' - ' + shortEnd;
      }

      events.push({
        status: status || 'N/A',
        createdBy: createdBy || 'N/A',
        eventName: eventName || 'N/A',
        eventType: eventType || 'N/A',
        location: location || 'N/A',
        dateRange: dateRange || 'N/A',
        shortDateRange: shortDateRange || 'N/A',
        rowIndex: i + 4
      });
    }

    return events
      .sort((a, b) => b.rowIndex - a.rowIndex)
      .slice(0, 10);

  } catch (error) {
    console.error('Recent events error:', error);
    return { error: true, message: 'Failed to retrieve recent events data' };
  }
}
