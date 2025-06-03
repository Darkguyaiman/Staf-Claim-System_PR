function doGet() {
  return HtmlService.createHtmlOutputFromFile('AdvancePurchase')
    .setTitle('Advance Purchase Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getUserName() {
  var email = Session.getActiveUser().getEmail();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var data = sheet.getRange('B5:C').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] === email) {
      return data[i][0];
    }
  }
  return 'Unknown User';
}

function getEventOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Event Creation');
  var data = sheet.getRange('A4:Q').getValues(); 
  var username = getUserName();
  var options = [];
  var today = new Date();
  today.setHours(0, 0, 0, 0); 
  var twoDaysLater = new Date(today.getTime() + (2 * 24 * 60 * 60 * 1000));

  for (var i = 0; i < data.length; i++) {
    if (data[i][3]) { 
      var eventDate = new Date(data[i][7]); 
      eventDate.setHours(0, 0, 0, 0); 
      var isApproved = data[i][0] === 'Approved';
      var isUserInvolved = data[i][2] === username || (data[i][12] && data[i][12].split(',').map(s => s.trim()).includes(username));

      if (isApproved && isUserInvolved && eventDate >= twoDaysLater) {
        options.push({
          name: data[i][3],
          id: data[i][16] || '' 
        });
      }
    }
  }

  return options;
}

function submitForm(formData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advance Purchases');
  var lastRow = sheet.getLastRow();
  var transactionId = generateUniqueId();
  
  formData.items.forEach(function(item, index) {
    lastRow++;
    sheet.getRange(lastRow, 1, 1, 10).setValues([[
      "Submitted",           
      new Date(),            
      formData.username,     
      formData.eventName,    
      item.name,             
      item.price,            
      formData.remarks,      
      transactionId,         
      "",                    
      formData.eventId       
    ]]);
  });
  
  sendConfirmationEmail(formData, transactionId);
  
  return transactionId;
}

function generateUniqueId() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advance Purchases');
  var existingIds = sheet.getRange('H4:H').getValues().flat();
  
  while (true) {
    var id = 'AP' + Math.random().toString().substr(2, 6);
    if (!existingIds.includes(id)) {
      return id;
    }
  }
}

function sendConfirmationEmail(formData, transactionId) {
  var emailAddress = Session.getActiveUser().getEmail();
  var userName = getUserName(); 
  var subject = 'Advance Purchase Confirmation - ' + transactionId;

  var htmlBody = `
    <html>
      <body style="font-family: Arial, sans-serif; margin: 0; padding: 0; background-color: #f5f7fa; color: #333;">
        <div style="max-width: 600px; margin: 20px auto; background: #ffffff; border-radius: 12px; box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08); overflow: hidden;">
          <!-- Header -->
          <div style="background: linear-gradient(135deg, #2D0C31 0%, #004D4D 100%); color: #ffffff; text-align: center; padding: 30px 20px;">
            <h1 style="margin: 0; font-size: 26px;">Advance Purchase Confirmation</h1>
            <p style="margin: 8px 0 0; font-size: 14px;">We've received your submission</p>
          </div>

          <!-- Content -->
          <div style="padding: 30px;">
            <p style="font-size: 16px;">Dear ${userName},</p>
            <p style="font-size: 15px; color: #555;">We've received your submission and are processing it. Below are the details of your purchase:
            </p>

            <div style="background: #f9f9ff; border-radius: 8px; padding: 1px; margin-bottom: 25px; border: 1px solid #B8B5FF;">
              <table style="width: 100%; font-size: 14px;">
                <tr>
                  <th style="padding: 16px; text-align: left; background: #B8B5FF; color: #2D0C31;">Advance Purchase ID</th>
                  <td style="padding: 16px;">${transactionId}</td>
                </tr>
                <tr>
                  <th style="padding: 16px; background: #f9f9ff; color: #2D0C31;">Event Name</th>
                  <td style="padding: 16px;">${formData.eventName}</td>
                </tr>
                <tr>
                  <th style="padding: 16px; background: #f9f9ff; color: #2D0C31;">Items</th>
                  <td style="padding: 16px;">
                    <ul style="padding-left: 20px; list-style-type: none; margin: 0;">
                      ${formData.items.map(item => `
                        <li style="margin-bottom: 8px; display: flex;">
                          <span style="flex: 1;">${item.name}</span>
                          <span style="margin-left: 20px; font-weight: 600; color: #2D0C31;">RM ${item.price.toFixed(2)}</span>
                        </li>
                      `).join('')}
                    </ul>
                  </td>
                </tr>
                <tr>
                  <th style="padding: 16px; background: #2D0C31; color: #ffffff;">Total Amount</th>
                  <td style="padding: 16px; background: #2D0C31; color: #ffffff;">RM ${formData.totalAmount.toFixed(2)}</td>
                </tr>
              </table>
            </div>

            <!-- Remarks -->
            <div style="background: #B8FFB5; border-left: 4px solid #00B894; padding: 12px 16px; border-radius: 4px;">
              <p style="margin: 0;"><strong style="color: #2D0C31;">Remarks:</strong> ${formData.remarks}</p>
            </div>

            <div style="text-align: center; margin: 30px 0 20px;">
              <p style="font-size: 15px; color: #555;">Need help with your submission?</p>
              <a href="mailto:mohamedaiman103@gmail.com" style="background: #00B894; color: #ffffff; padding: 12px 24px; border-radius: 6px; text-decoration: none;">Contact Our Support Team</a>
            </div>
          </div>

          <!-- Footer -->
          <div style="background: linear-gradient(135deg, #B8FFB5 0%, #B8B5FF 100%); text-align: center; padding: 20px; font-size: 13px; color: #2D0C31;">
            <p style="margin: 0;">Thank you</p>
            <p style="margin: 0;">© QSS/PMS</p>
          </div>
        </div>
      </body>
    </html>
  `;

  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: htmlBody
  });
}


function getInvoiceRows(transactionId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advance Purchases');
  var data = sheet.getRange('B4:I').getValues();
  var rows = [];
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][6] === transactionId && data[i][7] === '') {
      rows.push({
        rowIndex: i + 4,
        item: data[i][3],
        price: data[i][4]
      });
    }
  }
  
  return rows;
}

function uploadInvoice(rowIndex, file) {
  var folder = DriveApp.getFolderById('YOUR_FOLDER_ID');
  var blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType, file.filename);
  var driveFile = folder.createFile(blob);
  var fileUrl = driveFile.getUrl();
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Advance Purchases');
  sheet.getRange(rowIndex, 9).setValue(fileUrl);
  
  return fileUrl;
}

function submitInvoices(invoices) {
  var results = [];
  for (var i = 0; i < invoices.length; i++) {
    var invoice = invoices[i];
    try {
      var fileUrl = uploadInvoice(invoice.rowIndex, invoice.file);
      results.push({
        success: true,
        message: 'Invoice uploaded successfully for row ' + invoice.rowIndex,
        url: fileUrl
      });
    } catch (error) {
      results.push({
        success: false,
        message: 'Error uploading invoice for row ' + invoice.rowIndex + ': ' + error.toString()
      });
    }
  }
  return results;
}