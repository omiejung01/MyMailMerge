function myFunction() {
  //Logger.log('Hello World4');
  readSheetRowByRow("ITX3004", "Midterm and Quiz score ITX3004 1/2025");
}

function sendMyEmail(recipient, subject, body) {
  GmailApp.sendEmail(recipient, subject, body);
}

function readSheetRowByRow(sheetname,subject) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  const sheet = ss.getSheetByName(sheetname); // Get the sheet by name

  // Get all data in the sheet, assuming data starts from A1
  const range = sheet.getDataRange(); 
  const values = range.getValues(); // Get the data as a 2D array

  var fields = [];
  // Iterate through each row
  for (let i = 0; i < values.length; i++) {
    const row = values[i]; // Get the current row (which is an array of cell values)

    // first row
    if (i ==0) {    
      for (let j = 0; j < row.length; j++) {
        fields[j] = row[j];
        //Logger.log(`Row ${i + 1}, Column ${j + 1}: ${row[j]}`);
      }
    } else {
      var message_body = 'Dear student\n\n';
      var recipient = ''; // 2nd column is the email
      for (let j = 0; j < row.length; j++) {  
        message_body = message_body.concat(`${fields[j]}: ${row[j]}`);
        message_body = message_body.concat("\n");
        Logger.log(`${fields[j]}: ${row[j]}`);

        if (j == 1) {
          recipient = `${row[j]}`;
          Logger.log(`${row[j]}`);
        }
      }
      message_body = message_body.concat("\nThank you");
      Logger.log("\n");

      sendMyEmail(recipient, subject, message_body) 
    }  
  }
}