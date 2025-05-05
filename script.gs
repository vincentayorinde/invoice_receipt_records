function saveToRecords() {
    var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Are you sure you want to save this record?", ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return; // Cancel if user clicks No
  }
  

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceSheet = spreadsheet.getSheetByName("Invoice");
  var receiptSheet = spreadsheet.getSheetByName("Receipt");
  var recordsSheet = spreadsheet.getSheetByName("Records");
  
  // Find the last row in Records sheet
  var lastRow = recordsSheet.getLastRow() + 1;
  
  // Get values from Invoice and Receipt
  var invoiceNumber = invoiceSheet.getRange("B3").getValue();
  var date = invoiceSheet.getRange("B4").getValue();
  var customerName = invoiceSheet.getRange("B5").getValue();
  var subtotal = invoiceSheet.getRange("D14").getValue();
  var tax = invoiceSheet.getRange("D15").getValue();
  var total = invoiceSheet.getRange("D16").getValue();
  var paymentMethod = receiptSheet.getRange("B8").getValue();
  
  // Append to Records
  recordsSheet.getRange(lastRow, 1).setValue(invoiceNumber);
  recordsSheet.getRange(lastRow, 2).setValue(date);
  recordsSheet.getRange(lastRow, 3).setValue(customerName);
  recordsSheet.getRange(lastRow, 4).setValue(subtotal);
  recordsSheet.getRange(lastRow, 5).setValue(tax);
  recordsSheet.getRange(lastRow, 6).setValue(total);
  recordsSheet.getRange(lastRow, 7).setValue(paymentMethod);
  
  SpreadsheetApp.getUi().alert("Record saved successfully!");
}

function clearForm() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceSheet = spreadsheet.getSheetByName("Invoice");
  var receiptSheet = spreadsheet.getSheetByName("Receipt");
  
  // Clear input fields
  invoiceSheet.getRange("B5:B7").clearContent();
  invoiceSheet.getRange("A10:C12").clearContent();
  receiptSheet.getRange("B8").clearContent();
  
  SpreadsheetApp.getUi().alert("Invoice and Receipt cleared!");
}

function getSheetAsPDF(sheet, fileName) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId();
  var sheetId = sheet.getSheetId();

 // Set print range to exclude buttons
  var printRange;
  if (sheet.getName() === "Invoice") {
    printRange = "A1:D28"; // Adjust to include content up to row 29, excluding buttons at A29+
  } else if (sheet.getName() === "Receipt") {
    printRange = "A1:C25"; // Adjust to include content up to row 24, excluding any potential buttons
  }
  sheet.setActiveRange(sheet.getRange(printRange));
  
  var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?" +
            "format=pdf" +
            "&gid=" + sheetId +
            "&portrait=true" +
            "&fitw=true" +
            "&sheetnames=false" +
            "&printtitle=false" +
            "&pagenum=false" +
            "&gridlines=false" +
            "&top_margin=0.5" +
            "&bottom_margin=0.5" +
            "&left_margin=0.5" +
            "&right_margin=0.5" +
            "&range=" + encodeURIComponent(printRange); // Properly encode the range
  
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    }
  });

  
  return response.getBlob().setName(fileName + ".pdf");
}

function sendEmailViaBrevo(to, subject, body, attachmentBlob, fromEmail, fromName) {
  var apiKey = "key"; // Replace with your Brevo API key (e.g., xkeysib-...)
  var apiUrl = "https://api.brevo.com/v3/smtp/email";
  
  var payload = {
    "sender": {
      "name": fromName,
      "email": fromEmail
    },
    "to": [
      {
        "email": to
      }
    ],
    "subject": subject,
    "textContent": body,
    "attachment": [
      {
        "name": attachmentBlob.getName(),
        "content": Utilities.base64Encode(attachmentBlob.getBytes())
      }
    ]
  };
  
  var options = {
    "method": "post",
    "headers": {
      "accept": "application/json",
      "api-key": apiKey,
      "content-type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseCode = response.getResponseCode();
    if (responseCode === 201) {
      return true;
    } else {
      throw new Error("Brevo API responded with status " + responseCode + ": " + response.getContentText());
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error sending email: " + e.message);
    return false;
  }
}

function sendInvoiceEmail() {
    var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Are you sure you want to send the invoice email?", ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return; // Cancel if user clicks No
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var invoiceSheet = spreadsheet.getSheetByName("Invoice");
  
  var customerEmail = invoiceSheet.getRange("B7").getValue();
  var customerName = invoiceSheet.getRange("B5").getValue();
  var invoiceNumber = invoiceSheet.getRange("B3").getValue();
  var totalAmount = invoiceSheet.getRange("D16").getValue();
  
  if (!customerEmail) {
    SpreadsheetApp.getUi().alert("Please enter a customer email in B7.");
    return;
  }
  
  var pdfBlob = getSheetAsPDF(invoiceSheet, "Invoice_" + invoiceNumber);
  var subject = "Invoice #" + invoiceNumber + " from Company";
  var body = "Dear " + customerName + ",\n\n" +
             "I hope this message finds you well. Thank you for choosing Company for your needs.\n\n" +
             "We have prepared your invoice #" + invoiceNumber + " for the amount of $" + totalAmount.toFixed(2) + ". " +
             "Please find the detailed invoice attached for your records.\n\n" +
             "Should you have any questions or require further assistance, please don’t hesitate to reach out to us at email@Company.com or call us at (000) 000-0000. " +
             "We’re here to assist you!\n\n" +
             "We appreciate your business and look forward to serving you again.\n\n" +
             "Warm regards,\n" +
             "email Team\n" +
             "Company\n" +
             "email@Company.com | (000) 000-0000";


  if (sendEmailViaBrevo(customerEmail, subject, body, pdfBlob, "email@Company.com", "Company")) {
    SpreadsheetApp.getUi().alert("Invoice emailed successfully to " + customerEmail + "!");
  }
}

function sendReceiptEmail() {
   var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Are you sure you want to send the receipt email?", ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    return; // Cancel if user clicks No
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var receiptSheet = spreadsheet.getSheetByName("Receipt");
  
  var customerEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Invoice").getRange("B7").getValue();
  var customerName = receiptSheet.getRange("B6").getValue();
  var receiptNumber = receiptSheet.getRange("B3").getValue();
  var amountPaid = receiptSheet.getRange("B7").getValue();
  
  if (!customerEmail) {
    SpreadsheetApp.getUi().alert("Please enter a customer email in Invoice!B7.");
    return;
  }
  
  var pdfBlob = getSheetAsPDF(receiptSheet, "Receipt_" + receiptNumber);
  var subject = "Receipt #" + receiptNumber + " from Company";
  var body = "Dear " + customerName + ",\n\n" +
             "I hope you’re doing well. Thank you for your recent payment to Company.\n\n" +
             "We’re pleased to confirm that we have received your payment of $" + amountPaid.toFixed(2) + ". " +
             "Please find attached your receipt #" + receiptNumber + " for your records.\n\n" +
             "If you have any questions or need further assistance, feel free to contact us at email@Company.com or (000) 000-0000. " +
             "We’re always happy to help!\n\n" +
             "Thank you for your trust in us. We look forward to serving you again in the future.\n\n" +
             "Warm regards,\n" +
             "email Team\n" +
             "Company\n" +
             "email@Company.com | (000) 000-0000";
  
  if (sendEmailViaBrevo(customerEmail, subject, body, pdfBlob, "email@Company.com", "Company")) {
    SpreadsheetApp.getUi().alert("Receipt emailed successfully to " + customerEmail + "!");
  }
}
