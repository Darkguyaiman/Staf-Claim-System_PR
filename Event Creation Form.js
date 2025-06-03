const FIXED_LOCATION = "YOUR_FIXED_LOCATION";
const LOCAL_DISTANCE_THRESHOLD = 85; 
const MAPS_API_KEY = "YOUR_MAPS_API_KEY";
const DOCUMENT_FOLDER_ID = "YOUR_FOLDER_ID"; 

function doGet() {
  const email = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const jColumn = sheet.getRange('J5:J' + sheet.getLastRow()).getValues().flat();
  const cColumn = sheet.getRange('C5:C' + sheet.getLastRow()).getValues().flat();

  const isAuthorized = jColumn.includes(email) || cColumn.includes(email);

  if (isAuthorized) {
    return HtmlService.createTemplateFromFile('EventCreationForm')
      .evaluate()
      .setTitle('Event Creation Form')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('Unauthorized')
      .setTitle('Unauthorized Access')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getMapsApiKey() {
  return MAPS_API_KEY;
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
    username: 'User Not Found'
  };
}

function getFormOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  
  
  const eventTypes = settingsSheet.getRange('D5:D')
    .getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
  
  
  const officeAssets = settingsSheet.getRange('G5:G')
    .getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
  
  
  const budgetTypes = settingsSheet.getRange('K5:K')
    .getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
  
  
  const devices = settingsSheet.getRange('O5:O')
    .getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
  
  
  const currentUser = Session.getActiveUser().getEmail();
  const allUsernames = settingsSheet.getRange('B5:C')
    .getValues()
    .filter(row => row[0] !== '' && row[1] !== currentUser)
    .map(row => row[0]);
  
  return {
    eventTypes: eventTypes,
    officeAssets: officeAssets,
    budgetTypes: budgetTypes,
    devices: devices,
    colleagues: allUsernames
  };
}

function calculateDistance(destination) {
  try {
    const mapsUrl = `https://maps.googleapis.com/maps/api/directions/json?origin=${encodeURIComponent(FIXED_LOCATION)}&destination=${encodeURIComponent(destination)}&key=${MAPS_API_KEY}`;
    
    const response = UrlFetchApp.fetch(mapsUrl);
    const data = JSON.parse(response.getContentText());
    
    
    const geocodeUrl = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(destination)}&key=${MAPS_API_KEY}`;
    const geocodeResponse = UrlFetchApp.fetch(geocodeUrl);
    const geocodeData = JSON.parse(geocodeResponse.getContentText());
    
    let country = "";
    if (geocodeData.status === "OK" && geocodeData.results.length > 0) {
      const addressComponents = geocodeData.results[0].address_components;
      const countryComponent = addressComponents.find(component => 
        component.types.includes("country"));
      
      if (countryComponent) {
        country = countryComponent.long_name;
        
        
        if (country !== "Malaysia" && country !== "Singapore") {
          return { distance: null, status: "Overseas", country: country };
        }
      }
    }
    
    
    if (data.status === "OK" && data.routes.length > 0) {
      
      const distanceInMeters = data.routes[0].legs[0].distance.value;
      const distanceInKm = distanceInMeters / 1000;
      
      if (distanceInKm <= LOCAL_DISTANCE_THRESHOLD) {
        return { distance: distanceInKm, status: "Local", country: country };
      } else {
        return { distance: distanceInKm, status: "Outstation", country: country };
      }
    } else {
      
      if (country === "Malaysia" || country === "Singapore") {
        return { distance: null, status: "Outstation", country: country };
      }
      
      
      return { distance: null, status: "Unknown", country: country || "Unknown" };
    }
  } catch (error) {
    Logger.log("Error calculating distance: " + error);
    return { distance: null, status: "Error", country: "Unknown" };
  }
}

function generateEventId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('Event Creation');
  
  
  const existingIds = eventSheet.getRange('Q4:Q')
    .getValues()
    .filter(row => row[0] !== '')
    .map(row => row[0]);
  
  let newId;
  let isUnique = false;
  
  
  while (!isUnique) {
    
    const randomDigits = Math.floor(100000 + Math.random() * 900000).toString();
    newId = 'E' + randomDigits;
    
    
    isUnique = !existingIds.includes(newId);
  }
  
  return newId;
}

function uploadFileToDrive(fileBlob, fileName) {
  try {
    const folder = DriveApp.getFolderById(DOCUMENT_FOLDER_ID);
    
    let blob;

    
    if (fileBlob && fileBlob.data && fileBlob.mimeType) {
      blob = Utilities.newBlob(
        Utilities.base64Decode(fileBlob.data),
        fileBlob.mimeType,
        fileName
      );
    } else if (typeof fileBlob.getBytes === 'function') {
      
      blob = fileBlob;
    } else {
      throw new Error("Invalid fileBlob format. Must be a Blob or an object with base64 'data' and 'mimeType'.");
    }

    const file = folder.createFile(blob);
    file.setName(fileName);

    return file.getUrl();
  } catch (error) {
    Logger.log("Error uploading file: " + error);
    throw new Error("Failed to upload file: " + error.toString());
  }
}

function processForm(formData, documentBlob, quotationBlob) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eventSheet = ss.getSheetByName('Event Creation');
    const budgetSheet = ss.getSheetByName('Pre-Event Budget');
    const devicesSheet = ss.getSheetByName('Devices For Events');
    
    
    const eventId = generateEventId();
    
    
    let documentUrl = "";
    let quotationUrl = "";
    
    if (documentBlob) {
      documentUrl = uploadFileToDrive(documentBlob, `${eventId}_SupportingDoc_${formData.documentName}`);
    }
    
    if (quotationBlob) {
      quotationUrl = uploadFileToDrive(quotationBlob, `${eventId}_Quotation_${formData.quotationName}`);
    }
    
    
    const eventIds = eventSheet.getRange("Q4:Q").getValues(); 
    let nextRow = 4;

    for (let i = 0; i < eventIds.length; i++) {
      if (!eventIds[i][0]) {
        nextRow = i + 4; 
        break;
      }
    }
    
    
    const timeSlots = formData.timeSlots.map(slot => 
      `${slot.date} ${slot.startTime}-${slot.endTime}`
    ).join(", ");
    
    
    const officeAssets = formData.officeAssets.join(", ");
    const colleagues = formData.colleagues.join(", ");  
    const cleanStatus = formData.locationStatus.match(/\b(Local|Outstation|Overseas)\b/)?.[0] || "Unknown";
    
    
    eventSheet.getRange(nextRow, 1).setValue("Submitted"); 
    eventSheet.getRange(nextRow, 2).setValue(new Date()); 
    eventSheet.getRange(nextRow, 3).setValue(formData.username); 
    eventSheet.getRange(nextRow, 4).setValue(formData.eventName); 
    eventSheet.getRange(nextRow, 5).setValue(formData.eventType); 
    eventSheet.getRange(nextRow, 6).setValue(formData.location); 
    eventSheet.getRange(nextRow, 7).setValue(cleanStatus);
    eventSheet.getRange(nextRow, 8).setValue(formData.startDate); 
    eventSheet.getRange(nextRow, 9).setValue(formData.endDate); 
    eventSheet.getRange(nextRow, 10).setValue(timeSlots); 
    eventSheet.getRange(nextRow, 11).setValue(documentUrl); 
    eventSheet.getRange(nextRow, 12).setValue(officeAssets); 
    eventSheet.getRange(nextRow, 13).setValue(colleagues); 
    eventSheet.getRange(nextRow, 14).setValue(formData.company); 
    if (eventSheet.getRange("O4").getFormula() === "") {
      eventSheet.getRange("O4").setFormula('=ARRAYFORMULA(IF(Q4:Q = "", "", SUMIF(\'Pre-Event Budget\'!A4:A, Q4:Q, \'Pre-Event Budget\'!C4:C)))');
    }
    eventSheet.getRange(nextRow, 16).setValue(quotationUrl); 
    eventSheet.getRange(nextRow, 17).setValue(eventId); 

    
    
    if (formData.budgetItems.length > 0) {
      const budgetLastRow = budgetSheet.getLastRow();
      let budgetNextRow = budgetLastRow + 1;
      
      formData.budgetItems.forEach(item => {
        budgetSheet.getRange(budgetNextRow, 1).setValue(eventId); 
        budgetSheet.getRange(budgetNextRow, 2).setValue(item.type); 
        budgetSheet.getRange(budgetNextRow, 3).setValue(parseFloat(item.amount)); 
        budgetNextRow++;
      });
    }
    
    
    if (formData.devices && formData.devices.length > 0) {
      const devicesLastRow = devicesSheet.getLastRow();
      let devicesNextRow = devicesLastRow + 1;
      
      formData.devices.forEach(device => {
        devicesSheet.getRange(devicesNextRow, 1).setValue(eventId); 
        devicesSheet.getRange(devicesNextRow, 2).setValue(device); 
        devicesNextRow++;
      });
    }
    
    return {
      success: true,
      message: "Event created successfully!",
      eventId: eventId
    };
  } catch (error) {
    Logger.log("Error processing form: " + error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}