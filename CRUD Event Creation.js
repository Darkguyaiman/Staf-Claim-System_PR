function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDEventCreation')
    .setTitle('Event Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
}

function getEventData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventSheet = ss.getSheetByName('Event Creation');
    const deviceSheet = ss.getSheetByName('Devices For Events');
    
    if (!eventSheet || !deviceSheet) {
      throw new Error('Required sheets not found');
    }
    
    
    const eventData = eventSheet.getDataRange().getValues();
    const deviceData = deviceSheet.getDataRange().getValues();
    
    const events = [];
    
    
    for (let i = 3; i < eventData.length; i++) {
      const row = eventData[i];
      
      
      if (!row[16] || row[16] === '') continue; 
      
      const eventId = row[16];
      
      
      const eventDevices = [];
      for (let j = 3; j < deviceData.length; j++) {
        const deviceRow = deviceData[j];
        if (deviceRow[0] === eventId) { 
          eventDevices.push({
            demoLaser: deviceRow[1] || '',
            accessories: deviceRow[2] || '',
            rowIndex: j + 1 
          });
        }
      }
      
      const event = {
        eventStatus: row[0] || '',
        createdDate: row[1] ? formatDateForDisplay(row[1]) : '',
        createdBy: row[2] || '',
        eventName: row[3] || '',
        eventType: row[4] || '',
        location: row[5] || '',
        locationStatus: row[6] || '',
        startDate: row[7] ? formatDateForDisplay(row[7]) : '',
        endDate: row[8] ? formatDateForDisplay(row[8]) : '',
        timeSlots: row[9] || '',
        supportingDocument: row[10] || '',
        officeAssets: row[11] || '',
        accompanyingColleagues: row[12] || '',
        affiliatedCompany: row[13] || '',
        totalBudget: row[14] || '',
        quotationDocument: row[15] || '',
        eventId: eventId,
        lastModifiedBy: row[17] || '',
        totalExpenses: row[18] || '',
        devices: eventDevices,
        rowIndex: i + 1 
      };
      
      events.push(event);
    }
    
    return events;
    
  } catch (error) {
    console.error('Error getting event data:', error);
    return [];
  }
}

function getDropdownData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      throw new Error('Settings sheet not found');
    }
    
    const data = settingsSheet.getDataRange().getValues();
    
    
    const eventTypes = [];
    for (let i = 4; i < data.length; i++) {
      if (data[i][3] && data[i][3] !== '') {
        eventTypes.push(data[i][3]);
      }
    }
    
    
    const colleagues = [];
    for (let i = 4; i < data.length; i++) {
      if (data[i][1] && data[i][1] !== '') {
        colleagues.push(data[i][1]);
      }
    }
    
    
    const officeAssets = [];
    for (let i = 4; i < data.length; i++) {
      if (data[i][6] && data[i][6] !== '') {
        officeAssets.push(data[i][6]);
      }
    }
    
    
    const devices = [];
    const deviceAccessories = {};
    for (let i = 4; i < data.length; i++) {
      if (data[i][14] && data[i][14] !== '') { 
        const deviceName = data[i][14];
        devices.push(deviceName);
        
        
        const accessories = data[i][15] ? data[i][15].toString().split(',').map(acc => acc.trim()).filter(acc => acc !== '') : [];
        deviceAccessories[deviceName] = accessories;
      }
    }
    
    return {
      eventTypes: eventTypes,
      colleagues: colleagues,
      officeAssets: officeAssets,
      devices: devices,
      deviceAccessories: deviceAccessories,
      eventStatuses: ['Submitted', 'In Review', 'Approved', 'Rejected'],
      affiliatedCompanies: ['Example 2', 'Example 1']
    };
    
  } catch (error) {
    console.error('Error getting dropdown data:', error);
    return {
      eventTypes: [],
      colleagues: [],
      officeAssets: [],
      devices: [],
      deviceAccessories: {},
      eventStatuses: ['Submitted', 'In Review', 'Approved', 'Rejected'],
      affiliatedCompanies: ['Example 2', 'Example 1']
    };
  }
}

function updateEvent(eventData, supportingDocumentFile, quotationDocumentFile) {
  try {
    console.log('updateEvent called with:', {
      eventData: eventData,
      supportingDocumentFile: supportingDocumentFile ? 'File provided' : 'No file',
      quotationDocumentFile: quotationDocumentFile ? 'File provided' : 'No file'
    });
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventSheet = ss.getSheetByName('Event Creation');
    const deviceSheet = ss.getSheetByName('Devices For Events');
    
    if (!eventSheet || !deviceSheet) {
      throw new Error('Required sheets not found');
    }
    
    const rowIndex = eventData.rowIndex;
    let supportingDocumentUrl = eventData.supportingDocument;
    let quotationDocumentUrl = eventData.quotationDocument;
    
    
    if (supportingDocumentFile && supportingDocumentFile.data) {
      console.log('Uploading supporting document:', supportingDocumentFile.name);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(supportingDocumentFile.data),
        supportingDocumentFile.mimeType,
        supportingDocumentFile.name
      );
      const uploadResult = uploadFile(blob, supportingDocumentFile.name, 'supporting');
      if (uploadResult.success) {
        supportingDocumentUrl = uploadResult.url;
        console.log('Supporting document uploaded successfully:', uploadResult.url);
      } else {
        throw new Error('Failed to upload supporting document: ' + uploadResult.message);
      }
    }
    
    
    if (quotationDocumentFile && quotationDocumentFile.data) {
      console.log('Uploading quotation document:', quotationDocumentFile.name);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(quotationDocumentFile.data),
        quotationDocumentFile.mimeType,
        quotationDocumentFile.name
      );
      const uploadResult = uploadFile(blob, quotationDocumentFile.name, 'quotation');
      if (uploadResult.success) {
        quotationDocumentUrl = uploadResult.url;
        console.log('Quotation document uploaded successfully:', uploadResult.url);
      } else {
        throw new Error('Failed to upload quotation document: ' + uploadResult.message);
      }
    }
    
    
    const startDate = eventData.startDate ? parseDate(eventData.startDate) : '';
    const endDate = eventData.endDate ? parseDate(eventData.endDate) : '';
    
    
    eventSheet.getRange(rowIndex, 1).setValue(eventData.eventStatus); 
    eventSheet.getRange(rowIndex, 4).setValue(eventData.eventName); 
    eventSheet.getRange(rowIndex, 5).setValue(eventData.eventType); 
    eventSheet.getRange(rowIndex, 6).setValue(eventData.location); 
    eventSheet.getRange(rowIndex, 7).setValue(eventData.locationStatus); 
    eventSheet.getRange(rowIndex, 8).setValue(startDate); 
    eventSheet.getRange(rowIndex, 9).setValue(endDate); 
    eventSheet.getRange(rowIndex, 10).setValue(eventData.timeSlots); 
    eventSheet.getRange(rowIndex, 11).setValue(supportingDocumentUrl); 
    eventSheet.getRange(rowIndex, 12).setValue(eventData.officeAssets); 
    eventSheet.getRange(rowIndex, 13).setValue(eventData.accompanyingColleagues); 
    eventSheet.getRange(rowIndex, 14).setValue(eventData.affiliatedCompany); 
    eventSheet.getRange(rowIndex, 16).setValue(quotationDocumentUrl); 
    eventSheet.getRange(rowIndex, 18).setValue(Session.getActiveUser().getEmail()); 
    
    
    updateEventDevices(eventData.eventId, eventData.devices || []);
    
    return { 
      success: true, 
      message: 'Event updated successfully',
      updatedEvent: {
        ...eventData,
        startDate: formatDateForDisplay(startDate),
        endDate: formatDateForDisplay(endDate),
        supportingDocument: supportingDocumentUrl,
        quotationDocument: quotationDocumentUrl,
        lastModifiedBy: Session.getActiveUser().getEmail()
      }
    };
    
  } catch (error) {
    console.error('Error updating event:', error);
    return { success: false, message: 'Error updating event: ' + error.toString() };
  }
}

function updateEventDevices(eventId, devices) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deviceSheet = ss.getSheetByName('Devices For Events');
    
    if (!deviceSheet) {
      throw new Error('Devices For Events sheet not found');
    }
    
    
    const data = deviceSheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    for (let i = data.length - 1; i >= 3; i--) { 
      if (data[i][0] === eventId) {
        rowsToDelete.push(i + 1); 
      }
    }
    
    
    rowsToDelete.forEach(rowIndex => {
      deviceSheet.deleteRow(rowIndex);
    });
    
    
    if (devices && devices.length > 0) {
      const newRows = devices.map(device => [
        eventId,
        device.demoLaser || '',
        device.accessories || ''
      ]);
      
      
      const lastRow = deviceSheet.getLastRow();
      const startRow = lastRow + 1;
      
      
      const range = deviceSheet.getRange(startRow, 1, newRows.length, 3);
      range.setValues(newRows);
    }
    
    console.log('Devices updated successfully for event:', eventId);
    
  } catch (error) {
    console.error('Error updating devices:', error);
    throw error;
  }
}

function uploadFile(fileBlob, fileName, fileType) {
  try {
    console.log('uploadFile called with:', fileName, fileType);
    
    const folderId = '1YOUR_FOLDER_ID';
    const folder = DriveApp.getFolderById(folderId);
    
    
    if (fileBlob.getBytes().length > 5 * 1024 * 1024) {
      return { success: false, message: 'File size exceeds 5MB limit' };
    }
    
    
    if (fileType === 'quotation' && !fileName.toLowerCase().endsWith('.xlsx')) {
      return { success: false, message: 'Quotation documents must be in XLSX format' };
    }
    
    
    const file = folder.createFile(fileBlob.setName(fileName));
    
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = file.getUrl();
    console.log('File uploaded successfully:', fileUrl);
    
    return { success: true, url: fileUrl, message: 'File uploaded successfully' };
    
  } catch (error) {
    console.error('Error uploading file:', error);
    return { success: false, message: 'Error uploading file: ' + error.toString() };
  }
}

function calculateLocationStatus(location) {
  try {
    
    const fixedLocation = "3 Towers, 349, Jln Ampang, Kampung Berembang, 55000 Kuala Lumpur, Malaysia";
    
    
    const maps = Maps.newGeocoder();
    
    
    const fixedLocationResult = maps.geocode(fixedLocation);
    const targetLocationResult = maps.geocode(location);
    
    if (fixedLocationResult.status !== 'OK' || targetLocationResult.status !== 'OK') {
      return 'Unknown';
    }
    
    const fixedCoords = fixedLocationResult.results[0].geometry.location;
    const targetCoords = targetLocationResult.results[0].geometry.location;
    const targetCountry = getCountryFromResult(targetLocationResult.results[0]);
    
    
    if (targetCountry && !['Malaysia', 'Singapore'].includes(targetCountry)) {
      return 'Overseas';
    }
    
    
    const distance = calculateDistance(
      fixedCoords.lat, fixedCoords.lng,
      targetCoords.lat, targetCoords.lng
    );
    
    
    if (distance <= 85) {
      return 'Local';
    } else {
      return 'Outstation';
    }
    
  } catch (error) {
    console.error('Error calculating location status:', error);
    return 'Unknown';
  }
}

function getCountryFromResult(result) {
  for (let component of result.address_components) {
    if (component.types.includes('country')) {
      return component.long_name;
    }
  }
  return null;
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371; 
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  const distance = R * c;
  return distance;
}

function formatDateForDisplay(date) {
  if (!date) return '';
  if (date instanceof Date) {
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return date.toString();
}

function parseDate(dateString) {
  if (!dateString) return '';
  
  
  const parts = dateString.split('-');
  if (parts.length === 3) {
    const year = parseInt(parts[0]);
    const month = parseInt(parts[1]) - 1; 
    const day = parseInt(parts[2]);
    return new Date(year, month, day);
  }
  
  return new Date(dateString);
}

function searchEvents(searchTerm) {
  const events = getEventData();
  if (!searchTerm || searchTerm.trim() === '') {
    return events;
  }
  
  const term = searchTerm.toLowerCase();
  return events.filter(event => 
    event.eventName.toLowerCase().includes(term) ||
    event.eventType.toLowerCase().includes(term) ||
    event.location.toLowerCase().includes(term) ||
    event.createdBy.toLowerCase().includes(term) ||
    event.eventStatus.toLowerCase().includes(term)
  );
}

function getEventById(eventId) {
  const events = getEventData();
  return events.find(event => event.eventId === eventId);
}

function generateAssetMovementForm(eventData) {
  try {
    console.log('generateAssetMovementForm called with:', eventData);
    
    const eventName = eventData.eventName || 'Event';
    const userName = eventData.createdBy || 'Unknown User';
    const affiliatedCompany = eventData.affiliatedCompany || 'PMS';
    const devices = eventData.devices || [];
    
    
    const lasers = devices.map(device => device.demoLaser || 'N/A');
    const accessories = devices.map(device => 
        device.accessories ? device.accessories.split(',').map(acc => acc.trim()) : []
    );
    
    
    const logoBase64 = getLogoAsBase64(affiliatedCompany);
    
    
    const formHtml = generateFormHtml(eventName, userName, affiliatedCompany, lasers, accessories, logoBase64);
    
    return {
      success: true,
      html: formHtml
    };
    
  } catch (error) {
    console.error('Error generating asset movement form:', error);
    return {
      success: false,
      message: 'Error generating form: ' + error.toString()
    };
  }
}

function getLogoAsBase64(affiliatedCompany) {
  try {
    const pmsLogoId = '18WKmAt3S4XkWVz3Z4mdN-Vkd13mama4f';
    const qssLogoId = '1yzzbZNGEQ6ov5iEmNYgNphZfQsydim8y';
    
    const isQSS = affiliatedCompany === 'Example 1';
    const logoId = isQSS ? qssLogoId : pmsLogoId;
    
    
    const file = DriveApp.getFileById(logoId);
    const blob = file.getBlob();
    
    
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    
    return `data:${mimeType};base64,${base64}`;
    
  } catch (error) {
    console.error('Error getting logo as base64:', error);
    return ''; 
  }
}

function generateFormHtml(eventName, userName, affiliatedCompany, lasers, accessories, logoBase64) {
  
  const isQSS = affiliatedCompany === 'Example 1';
  const companyFullName = isQSS ? 'Example 1 Ltd Pte' : 'Example 2 Ltd Pte.';
  
  return `
  <!DOCTYPE html>
  <html>
  <head>
      <meta charset="utf-8">
      <title>Asset Movement Form for ${eventName}</title>
      <style>
          @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
          
          @media print {
              body { margin: 0; padding: 0; }
              .no-print { display: none; }
              .form-container { box-shadow: none; border: none; padding: 0; }
              @page { margin: 0.25in; size: A4 portrait; }
          }
          
          * {
              margin: 0;
              padding: 0;
              box-sizing: border-box;
              font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          }
          
          body {
              margin: 0;
              padding: 0;
              line-height: 1.4;
              color: #333;
              font-size: 12px;
          }
          
          .form-container {
              width: 100%;
              max-width: 8.5in;
              margin: 0 auto;
              padding: 10px;
              background: white;
              position: relative;
              page-break-after: avoid;
          }
          
          .form-container::before {
              content: '';
              position: absolute;
              top: 0;
              left: 0;
              right: 0;
              height: 3px;
              background: #2c3e50;
          }
          
          .header {
              text-align: center;
              margin-bottom: 10px;
              padding-bottom: 5px;
              border-bottom: 1px solid #e1e8ed;
          }
          
          .header img {
              height: 40px;
              margin-bottom: 5px;
          }
          
          .company-name {
              font-size: 12px;
              font-weight: 600;
              color: #555;
              margin-top: 2px;
          }
          
          .title {
              color: #2c3e50;
              text-align: center;
              margin: 10px 0;
              font-size: 18px;
              font-weight: 600;
              position: relative;
              padding-bottom: 5px;
          }
          
          .title::after {
              content: '';
              position: absolute;
              bottom: 0;
              left: 50%;
              transform: translateX(-50%);
              width: 50px;
              height: 2px;
              background: #2c3e50;
          }
          
          .event-info {
              background: #f8f9fa;
              padding: 10px;
              border-radius: 3px;
              margin: 10px 0;
              border-left: 3px solid #2c3e50;
              font-size: 11px;
          }
          
          .event-info h4 {
              color: #2c3e50;
              font-weight: 600;
              margin-bottom: 5px;
              font-size: 11px;
              text-transform: uppercase;
          }
          
          .event-info p {
              color: #555;
              font-weight: 400;
              margin: 4px 0;
          }
          
          .devices-table {
              border-collapse: collapse;
              width: 100%;
              margin: 10px 0;
              border: 1px solid #ddd;
              font-size: 11px;
          }
          
          .devices-table th {
              background: #2c3e50;
              color: white;
              padding: 6px 8px;
              text-align: left;
              font-weight: 500;
              font-size: 11px;
              text-transform: uppercase;
          }
          
          .devices-table td {
              padding: 6px 8px;
              border-bottom: 1px solid #e1e8ed;
              background: white;
          }
          
          .devices-table tr:nth-child(even) td {
              background: #f8f9fa;
          }
          
          .devices-table td:first-child {
              font-weight: 500;
              color: #2c3e50;
          }
          
          .acknowledgment {
              background: #f8f9fa;
              border: 1px solid #ddd;
              padding: 10px;
              margin: 10px 0;
              border-radius: 3px;
              font-size: 11px;
          }
          
          .acknowledgment h3 {
              color: #2c3e50;
              margin: 0 0 8px;
              font-weight: 600;
              font-size: 11px;
              text-transform: uppercase;
              padding-bottom: 3px;
              border-bottom: 1px solid #ddd;
          }
          
          .acknowledgment p {
              color: #333;
              font-weight: 400;
              line-height: 1.4;
              text-align: justify;
          }
          
          .signature-table {
              border-collapse: collapse;
              width: 100%;
              margin: 10px 0;
              border: 1px solid #ddd;
              font-size: 11px;
          }
          
          .signature-table th {
              background: #2c3e50;
              color: white;
              padding: 6px 8px;
              font-weight: 500;
              font-size: 11px;
              text-transform: uppercase;
              width: 50%;
          }
          
          .signature-table td {
              padding: 10px;
              background: white;
              border-right: 1px solid #e1e8ed;
              vertical-align: top;
              height: 80px;
          }
          
          .signature-table td:last-child {
              border-right: none;
          }
          
          .signature-table p {
              margin: 6px 0;
              color: #333;
              font-weight: 400;
          }
          
          .signature-section {
              margin-top: 15px;
              padding-top: 10px;
              border-top: 1px solid #e1e8ed;
              font-size: 11px;
          }
          
          .signature-section h3 {
              color: #2c3e50;
              margin-bottom: 8px;
              font-weight: 600;
              font-size: 11px;
              text-transform: uppercase;
          }
          
          .signature-section p {
              margin: 4px 0;
              color: #555;
              font-weight: 400;
          }
          
          .signature-section strong {
              color: #2c3e50;
              font-weight: 600;
          }
          
          .signature-line {
              margin: 8px 0 3px;
              border-bottom: 1px solid #333;
              width: 200px;
              height: 15px;
              display: inline-block;
          }
          
          .print-button {
              background: #2c3e50;
              color: white;
              border: none;
              padding: 6px 12px;
              border-radius: 3px;
              cursor: pointer;
              margin: 10px 5px 10px 0;
              font-size: 12px;
              font-weight: 500;
              transition: all 0.2s ease;
          }
          
          .print-button:hover {
              background: #1a2634;
          }
          
          .close-button {
              background: #555;
          }
          
          .close-button:hover {
              background: #333;
          }
          
          .form-footer {
              margin-top: 15px;
              text-align: center;
              padding-top: 10px;
              border-top: 1px solid #e1e8ed;
              color: #777;
              font-size: 10px;
              font-weight: 400;
          }
          
          .no-devices {
              text-align: center;
              font-style: italic;
              color: #777;
              padding: 10px;
              background: #f8f9fa;
              border-radius: 3px;
              border: 1px dashed #ddd;
              font-size: 11px;
          }
      </style>
  </head>
  <body>
      <div class="no-print">
          <button class="print-button" onclick="window.print()">
              Print / Save as PDF
          </button>
          <button class="print-button close-button" onclick="window.close()">
              Close
          </button>
      </div>
      
      <div class="form-container">
          <div class="header">
              ${logoBase64 ? `<img src="${logoBase64}" alt="${affiliatedCompany} Logo">` : ''}
              <div class="company-name">${companyFullName}</div>
          </div>
          
          <h1 class="title">Asset Movement Form</h1>
          
          <div class="event-info">
              <h4>Event Details</h4>
              <p><strong>Event Name:</strong> ${eventName}</p>
              <p><strong>Responsible Person:</strong> ${userName}</p>
              <p><strong>Company:</strong> ${companyFullName}</p>
              <p><strong>Form Generated:</strong> ${new Date().toLocaleDateString('en-GB', { 
                  year: 'numeric', 
                  month: 'short', 
                  day: 'numeric' 
              })}</p>
          </div>
          
          <table class="devices-table">
              <thead>
                  <tr>
                      <th>Demo Laser Equipment</th>
                      <th>Accessories & Components</th>
                  </tr>
              </thead>
              <tbody>
                  ${lasers.length > 0 ? lasers.map((laser, index) => `
                      <tr>
                          <td>${laser}</td>
                          <td>${(accessories[index] || []).join(', ') || 'No accessories specified'}</td>
                      </tr>
                  `).join('') : `
                      <tr>
                          <td colspan="2" class="no-devices">
                              No devices have been assigned to this event
                          </td>
                      </tr>
                  `}
              </tbody>
          </table>

          <div class="acknowledgment">
              <h3>Acknowledgment and Declaration by Employee</h3>
              <p>
                  I, <strong>${userName}</strong>, acknowledge receipt of the above assets in good condition. 
                  I understand these assets belong to <strong>${companyFullName}</strong> and are under my 
                  possession for work purposes. I will be responsible for their safe handling and timely return.
              </p>
          </div>

          <table class="signature-table">
              <thead>
                  <tr>
                      <th>Asset Collection</th>
                      <th>Asset Return</th>
                  </tr>
              </thead>
              <tbody>
                  <tr>
                      <td>
                          <p><strong>Received by:</strong> _______________________</p>
                          <p><strong>Date:</strong> ________________________</p>
                          <p><strong>Time:</strong> ________________________</p>
                          <p><strong>Condition:</strong> _____________________</p>
                      </td>
                      <td>
                          <p><strong>Returned by:</strong> _______________________</p>
                          <p><strong>Date:</strong> ________________________</p>
                          <p><strong>Time:</strong> ________________________</p>
                          <p><strong>Condition:</strong> _____________________</p>
                      </td>
                  </tr>
              </tbody>
          </table>

          <div class="signature-section">
              <h3>Form Submission Details</h3>
              <p><strong>Submitted by:</strong> ${userName}</p>
              <p><strong>Event:</strong> ${eventName}</p>
              <p><strong>Company:</strong> ${companyFullName}</p>
              <p><strong>Signature:</strong><span class="signature-line"></span></p>
              <p><strong>Date:</strong> ${new Date().toLocaleDateString('en-GB')}</p>
          </div>
          
          <div class="form-footer">
              This document was automatically generated by the Event Management System<br>
              ${companyFullName} - Asset Management Division
          </div>
      </div>
  </body>
  </html>`;
}
