function doPost(e) {
    const data = JSON.parse(e.postData.contents);
    const userEmail = data.userEmail;
    const userOTP = data.userOTP || '';

    const userOtp = parseInt(userOTP);

    let userExist = isUserExist(userEmail);

    if (userExist && userOTP === '') {
        generateOTP(userEmail);
        const response = sendDatEmail(userEmail);
        return ContentService.createTextOutput(response);

    } else if (userExist && userOtp > 0) {
        const response = checkUserOtp(userEmail, userOtp);
        return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);

    } else {
        return ContentService.createTextOutput(JSON.stringify({ status: 500, message: 'No user found' }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}


function doGet(e) {
    let summary = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs').getRange('B1').getValue();
    return ContentService.createTextOutput(summary)
}

function isUserExist(userEmail) {
    let sheetEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');
    const lastRow = sheetEmail.getLastRow();
    if (lastRow < 2) {
        // If there are no rows with data after the header (assuming header is in the first row)
        Logger.log('No data found beyond the first row.');
        return;
    }

    // Get the range from the second row to the last row that has data in the specified column
    const values = sheetEmail.getRange(2, 1, lastRow - 1, 1).getValues();

    // Convert the 2D array to a 1D array with value at index 0 of each sub-array
    const allEmail = values.map(function (row) {
        return row[0];
    });

    return allEmail.includes(userEmail);
}

function generateOTP(userEmail) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');

  // Verify that the sheet exists
  if (!sheet) {
    Logger.log("The 'Email' sheet does not exist.");
    return null;
  }

  // Get all email addresses starting from row 2
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const emailValues = dataRange.getValues();

  // Find the row containing the user email
  let userRow = -1;
  for (let i = 0; i < emailValues.length; i++) {
    if (emailValues[i][0] === userEmail) {
      userRow = i + 2; // Adjust for zero-index and header row
      break;
    }
  }

  if (userRow === -1) {
    Logger.log('User email not found.');
    return null;
  }

  // Clear any existing trigger for deleting OTP
  clearPreviousTrigger(userEmail);

  // Generate a six-digit random OTP
  const sixDigitCode = Math.floor(100000 + Math.random() * 900000);

  // Attempt to store the OTP in the sheet
  try {
    sheet.getRange(userRow, 2).setValue(sixDigitCode); // Store the OTP in column 2
    // Verify that the OTP was set successfully
    const storedOTP = sheet.getRange(userRow, 2).getValue();
    if (storedOTP !== sixDigitCode) {
      throw new Error("OTP could not be stored in the sheet.");
    }
  } catch (error) {
    Logger.log(`Error while storing OTP: ${error.message}`);
    return null;
  }

  // Set a trigger to delete the OTP after 5 minutes
  ScriptApp.newTrigger('deleteCodeAfterTimeout')
    .timeBased()
    .after(5 * 60 * 1000) // 5 minutes in milliseconds
    .create();

  PropertiesService.getScriptProperties().setProperty('userEmailToDelete', userEmail);

  Logger.log(`OTP ${sixDigitCode} successfully stored for user: ${userEmail}`);
  return sixDigitCode;
}

// Function to clear previous trigger related to the user email
function clearPreviousTrigger(userEmail) {
  const triggers = ScriptApp.getProjectTriggers();
  const userEmailToDelete = PropertiesService.getScriptProperties().getProperty('userEmailToDelete');

  if (userEmailToDelete && userEmailToDelete === userEmail) {
    let triggerDeleted = false;
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'deleteCodeAfterTimeout') {
        ScriptApp.deleteTrigger(triggers[i]);
        Logger.log(`Previous trigger for user ${userEmail} deleted successfully.`);
        triggerDeleted = true;
        break; // Exit loop after deleting the relevant trigger
      }
    }
    if (!triggerDeleted) {
      Logger.log(`No existing trigger found for user ${userEmail}.`);
    }
  } else {
    Logger.log(`No trigger exists for user ${userEmail} to be deleted.`);
  }
}


function deleteCodeAfterTimeout() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');
  const userEmailToDelete = PropertiesService.getScriptProperties().getProperty('userEmailToDelete');

  if (!userEmailToDelete) {
    Logger.log('No user email to delete.');
    return;
  }

  if (!sheet) {
    Logger.log("The 'Email' sheet does not exist.");
    return;
  }

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1); // Get all emails starting from row 2
  const emailValues = dataRange.getValues();

  // Find the row containing the user email
  let userRow = -1;
  for (let i = 0; i < emailValues.length; i++) {
    if (emailValues[i][0] === userEmailToDelete) {
      userRow = i + 2; // Adjust for zero-index and header row
      break;
    }
  }

  if (userRow === -1) {
    Logger.log('User email not found in the sheet.');
    return;
  }

  // Clear the OTP in column 2
  sheet.getRange(userRow, 2).setValue('');

  // Delete the trigger after the OTP has been deleted
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'deleteCodeAfterTimeout') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log(`Trigger for deleting OTP of user ${userEmailToDelete} has been removed.`);
      break; // Exit after deleting the trigger
    }
  }

  // Clear the property for the email to delete
  PropertiesService.getScriptProperties().deleteProperty('userEmailToDelete');

  Logger.log(`OTP for user ${userEmailToDelete} has been deleted.`);
}

function sendDatEmail(userEmail) {
    //userEmail = 'pakit.tu@rsu.ac.th';
    let sheetEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');
    const lastRow = sheetEmail.getLastRow();
    if (lastRow < 2) {
        // If there are no rows with data after the header (assuming header is in the first row)
        Logger.log('No data found beyond the first row.');
        return;
    }
    let emailsAndOTP = sheetEmail.getRange(2, 1, lastRow - 1, 2).getValues();

    let genOTP = '';
    let attempts = 0;

    let response = '';

    while (genOTP.length === 0 && attempts < 5) {
        genOTP = getOtpForUser(userEmail, emailsAndOTP);
        if (genOTP && genOTP.length > 0) {
            break;
        }
        Utilities.sleep(5000); // Wait for 5 seconds
        attempts++;
    }

    if (genOTP.length === 0) {
        Logger.log('Unable to retrieve OTP after 5 attempts.');
        // Send error response to the client (can be adjusted as needed)
        GmailApp.sendEmail(userEmail, 'Error - OTP Generation', 'Failed to generate OTP. Please try again later.');
        response = 'Error - OTP Generation';
        return response;
    } else {
        const subject = "DentRSU Connect APP";
        let body = "Email OTP Verification: " + genOTP;
        
        // Create an HTML body with the OTP prominently displayed
        let htmlBody = `
            <div style="font-family: Arial, sans-serif; text-align: center;">
                <h2>Email OTP Verification</h2>
                <p style="font-size: 24px; font-weight: bold; background-color: #f0f0f0; padding: 10px; display: inline-block;">
                    ${genOTP}
                </p>
                <p>Please use this OTP to verify your email address.</p>
                <p>It is valid for 5 minutes.</p>
            </div>
        `;

        // Send the email with the HTML body
        GmailApp.sendEmail(userEmail, subject, body, { htmlBody: htmlBody });

        response = 'OTP via email';
        return response;
    }
}

function getOtpForUser(userEmail, dataArray) {
    for (let i = 0; i < dataArray.length; i++) {
        if (dataArray[i][0] === userEmail) {
            return dataArray[i][1]; // Return the OTP that matches the userEmail
        }
    }
    return null; // Return null if no match is found
}

function withSignatureLine() {
    let emails = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("emails").getValues()
    const signature = Gmail.Users.Settings.SendAs.list('me').sendAs.find(account => account.isDefault).signature
    emails.forEach((e) => {
        let body = e[2] + signature
        if (e[4]) {
            GmailApp.sendEmail(e[0], e[1], null, {
                htmlBody: body
            })
        }
    })

}

function checkUserOtp(userEmail, userOtp) {
    //userEmail = 'pakit.tu@rsu.ac.th';
    //userOtp = 123456;
    let sheetEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');
    const lastRow = sheetEmail.getLastRow();
    if (lastRow < 2) {
        // If there are no rows with data after the header (assuming header is in the first row)
        Logger.log('No data found beyond the first row.');
        recordLog(userEmail, 'failure');
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'No data found beyond the first row.' })).setMimeType(ContentService.MimeType.JSON);
    }
    let emailsAndOTP = sheetEmail.getRange(2, 1, lastRow - 1, 2).getValues();
    const otp = getOtpForUser(userEmail, emailsAndOTP);

    if (otp && otp === userOtp) {
        const jwt = createJwtToken(userEmail);
        recordLog(userEmail, 'success');
        return { status: 200, message: 'OTP verified successfully.', userEmail: userEmail, token: jwt };
        //return ContentService.createTextOutput(JSON.stringify({ status: 200, message: 'OTP verified successfully.', userEmail: userEmail, token: jwt })).setMimeType(ContentService.MimeType.JSON);
    } else {
        recordLog(userEmail, 'failure');
        return { status: 500, message: 'OTP verification failed. Incorrect OTP.' };
        //return ContentService.createTextOutput(JSON.stringify({ status: 500, message: 'OTP verification failed. Incorrect OTP.' })).setMimeType(ContentService.MimeType.JSON);
    }
}

function recordLog(userEmail, status) {
    //userEmail = 'test@rsu.ac.th';
    //status ='Day test';
    let sheetLogs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
    
    // Create the Logs sheet if it doesn't exist
    if (!sheetLogs) {
        sheetLogs = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Logs');
        sheetLogs.appendRow(['User Email', 'Timestamp', 'Status']);
    }

    // Insert a new row at the second position (row 2)
    sheetLogs.insertRowBefore(2);

    // Set the timestamp and add the new log at the second row
    const timestamp = new Date();
    sheetLogs.getRange(2, 1, 1, 3).setValues([[userEmail, timestamp, status]]);
}


function createJwtToken(userEmail) {
    const header = { alg: 'HS256', typ: 'JWT' };
    const payload = { email: userEmail, exp: Math.floor(Date.now() / 1000) + (60 * 60) }; // Expires in 1 hour
    const secret = ''; // Add your secret key here

    const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header));
    const encodedPayload = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
    const signature = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(`${encodedHeader}.${encodedPayload}`, secret));

    return `${encodedHeader}.${encodedPayload}.${signature}`;
}

function deleteExcessRows() {
  const sheetLogs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  if (!sheetLogs) {
    Logger.log("The 'Logs' sheet does not exist.");
    return;
  }

  const lastRow = sheetLogs.getLastRow();

  // Check if there are more than 10,000 rows
  if (lastRow > 10000) {
    // Calculate the number of rows to delete
    const rowsToDelete = lastRow - 10000;

    // Delete rows starting from row 10,001 onward
    sheetLogs.deleteRows(10001, rowsToDelete);
    Logger.log(`${rowsToDelete} rows deleted starting from row 10,001.`);
  } else {
    Logger.log("No rows to delete. Total rows are within the limit.");
  }
}

// Function to set a trigger that runs the deleteExcessRows function every week
function setWeeklyTrigger() {
  // First, clear existing triggers for the function to avoid duplicates
  deleteExistingTriggers();

  // Create a time-based trigger to run every week
  ScriptApp.newTrigger("deleteExcessRows")
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY) // Optionally choose a specific day of the week
    .atHour(1) // Optionally choose the time (1 AM in this example)
    .create();
}

// Function to delete existing triggers for the deleteExcessRows function
function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "deleteExcessRows") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

