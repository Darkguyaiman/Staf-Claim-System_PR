function doGet() {
  return HtmlService.createTemplateFromFile('CRUDSettings')
    .evaluate()
    .setTitle('Settings')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function getSettingsData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sheet) {
      throw new Error('Settings sheet not found');
    }
    
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 5) {
      return {
        users: [],
        admins: [],
        eventTypes: [],
        pettyCashTypes: [],
        expensesBudgetTypes: [],
        officeAssets: [],
        companyCars: [],
        devices: [],
        accessories: []
      };
    }
    
    
    const users = [];
    const usernames = sheet.getRange('B5:B' + lastRow).getValues().flat();
    const userEmails = sheet.getRange('C5:C' + lastRow).getValues().flat();
    
    for (let i = 0; i < usernames.length; i++) {
      if (usernames[i] && userEmails[i]) {
        users.push({
          username: usernames[i],
          email: userEmails[i]
        });
      }
    }
    
    
    const admins = [];
    const adminUsernames = sheet.getRange('I5:I' + lastRow).getValues().flat();
    const adminEmails = sheet.getRange('J5:J' + lastRow).getValues().flat();
    
    for (let i = 0; i < adminUsernames.length; i++) {
      if (adminUsernames[i] && adminEmails[i]) {
        admins.push({
          username: adminUsernames[i],
          email: adminEmails[i]
        });
      }
    }
    
    
    const eventTypes = sheet.getRange('D5:D' + lastRow).getValues().flat().filter(item => item);
    
    
    const pettyCashTypes = sheet.getRange('H5:H' + lastRow).getValues().flat().filter(item => item);
    
    
    const expensesBudgetTypes = sheet.getRange('K5:K' + lastRow).getValues().flat().filter(item => item);
    
    
    const officeAssets = sheet.getRange('G5:G' + lastRow).getValues().flat().filter(item => item);
    
    
    const companyCars = sheet.getRange('N5:N' + lastRow).getValues().flat().filter(item => item);
    
    
    const devices = [];
    const deviceNames = sheet.getRange('O5:O' + lastRow).getValues().flat();
    const deviceAccessories = sheet.getRange('P5:P' + lastRow).getValues().flat();
    
    for (let i = 0; i < deviceNames.length; i++) {
      if (deviceNames[i]) {
        const accessoryList = deviceAccessories[i] ? 
          deviceAccessories[i].toString().split(',').map(acc => acc.trim()).filter(acc => acc) : [];
        
        devices.push({
          name: deviceNames[i],
          accessories: accessoryList
        });
      }
    }
    
    
    const accessories = sheet.getRange('Q5:Q' + lastRow).getValues().flat().filter(item => item);
    
    return {
      users: users,
      admins: admins,
      eventTypes: eventTypes,
      pettyCashTypes: pettyCashTypes,
      expensesBudgetTypes: expensesBudgetTypes,
      officeAssets: officeAssets,
      companyCars: companyCars,
      devices: devices,
      accessories: accessories
    };
    
  } catch (error) {
    console.error('Error getting settings data:', error);
    return {
      error: error.toString(),
      users: [],
      admins: [],
      eventTypes: [],
      pettyCashTypes: [],
      expensesBudgetTypes: [],
      officeAssets: [],
      companyCars: [],
      devices: [],
      accessories: []
    };
  }
}

function updateSettingsData(type, action, data, index) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
    if (!sheet) {
      throw new Error('Settings sheet not found');
    }
    
    const columnMappings = {
      'user': { username: 'B', email: 'C' },
      'admin': { username: 'I', email: 'J' },
      'eventType': 'D',
      'pettyCashType': 'H',
      'expensesBudgetType': 'K',
      'officeAsset': 'G',
      'companyCar': 'N',
      'device': { name: 'O', accessories: 'P' },
      'accessory': 'Q'
    };
    
    const mapping = columnMappings[type];
    if (!mapping) {
      throw new Error('Unknown data type: ' + type);
    }
    
    if (action === 'add') {
      addItem(sheet, type, mapping, data);
    } else if (action === 'edit') {
      editItem(sheet, type, mapping, data, index);
    } else if (action === 'delete') {
      deleteItem(sheet, type, mapping, index);
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error updating settings data:', error);
    throw error;
  }
}

function addItem(sheet, type, mapping, data) {
  if (type === 'user' || type === 'admin') {
    
    const usernameCol = mapping.username;
    const emailCol = mapping.email;
    
    let row = 5;
    while (sheet.getRange(usernameCol + row).getValue() !== '') {
      row++;
    }
    
    sheet.getRange(usernameCol + row).setValue(data.username);
    sheet.getRange(emailCol + row).setValue(data.email);
    
  } else if (type === 'device') {
    
    let row = 5;
    while (sheet.getRange(mapping.name + row).getValue() !== '') {
      row++;
    }
    
    sheet.getRange(mapping.name + row).setValue(data.name);
    if (data.accessories && data.accessories.length > 0) {
      sheet.getRange(mapping.accessories + row).setValue(data.accessories.join(', '));
    }
    
  } else {
    
    let row = 5;
    while (sheet.getRange(mapping + row).getValue() !== '') {
      row++;
    }
    
    sheet.getRange(mapping + row).setValue(data);
  }
}

function editItem(sheet, type, mapping, data, index) {
  const row = 5 + index;
  
  if (type === 'user' || type === 'admin') {
    sheet.getRange(mapping.username + row).setValue(data.username);
    sheet.getRange(mapping.email + row).setValue(data.email);
    
  } else if (type === 'device') {
    sheet.getRange(mapping.name + row).setValue(data.name);
    const accessoriesValue = data.accessories && data.accessories.length > 0 ? 
      data.accessories.join(', ') : '';
    sheet.getRange(mapping.accessories + row).setValue(accessoriesValue);
    
  } else {
    sheet.getRange(mapping + row).setValue(data);
  }
}

function deleteItem(sheet, type, mapping, index) {
  const row = 5 + index;
  
  if (type === 'user' || type === 'admin') {
    
    const usernameCol = mapping.username;
    const emailCol = mapping.email;
    
    
    const lastRow = sheet.getLastRow();
    if (row < lastRow) {
      const usernameRange = sheet.getRange(usernameCol + (row + 1) + ':' + usernameCol + lastRow);
      const emailRange = sheet.getRange(emailCol + (row + 1) + ':' + emailCol + lastRow);
      
      const usernameValues = usernameRange.getValues();
      const emailValues = emailRange.getValues();
      
      
      sheet.getRange(usernameCol + row + ':' + usernameCol + (lastRow - 1)).setValues(usernameValues);
      sheet.getRange(emailCol + row + ':' + emailCol + (lastRow - 1)).setValues(emailValues);
    }
    
    
    sheet.getRange(usernameCol + lastRow).clearContent();
    sheet.getRange(emailCol + lastRow).clearContent();
    
  } else if (type === 'device') {
    
    const nameCol = mapping.name;
    const accessoriesCol = mapping.accessories;
    
    const lastRow = sheet.getLastRow();
    if (row < lastRow) {
      const nameRange = sheet.getRange(nameCol + (row + 1) + ':' + nameCol + lastRow);
      const accessoriesRange = sheet.getRange(accessoriesCol + (row + 1) + ':' + accessoriesCol + lastRow);
      
      const nameValues = nameRange.getValues();
      const accessoriesValues = accessoriesRange.getValues();
      
      
      sheet.getRange(nameCol + row + ':' + nameCol + (lastRow - 1)).setValues(nameValues);
      sheet.getRange(accessoriesCol + row + ':' + accessoriesCol + (lastRow - 1)).setValues(accessoriesValues);
    }
    
    
    sheet.getRange(nameCol + lastRow).clearContent();
    sheet.getRange(accessoriesCol + lastRow).clearContent();
    
  } else {
    
    const lastRow = sheet.getLastRow();
    if (row < lastRow) {
      const range = sheet.getRange(mapping + (row + 1) + ':' + mapping + lastRow);
      const values = range.getValues();
      
      
      sheet.getRange(mapping + row + ':' + mapping + (lastRow - 1)).setValues(values);
    }
    
    
    sheet.getRange(mapping + lastRow).clearContent();
  }
}