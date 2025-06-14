const CHANNEL_ACCESS_TOKEN = '2007214665';
const SPREADSHEET_ID = '15p-c58qNFNNHSUe-H7_QNafGCFA4v6em7FscOx8hbzE';
const SHEET_NAME = 'PermittedUsers';

function doPost(e) {
  // Add error handling for missing postData
  if (!e || !e.postData || !e.postData.contents) {
    console.error("Invalid request format");
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Invalid request"})).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const jsonData = JSON.parse(e.postData.contents);
    
    // Verify the request contains events
    if (!jsonData.events || jsonData.events.length === 0) {
      console.error("No events in request");
      return ContentService.createTextOutput(JSON.stringify({success: false, message: "No events"})).setMimeType(ContentService.MimeType.JSON);
    }

    const event = jsonData.events[0];
    
    // Get permitted users from spreadsheet
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const permittedUserIds = data.map(row => row[0].toString()); // Convert to string for comparison
    
    // Check if user is permitted
    if (permittedUserIds.includes(event.source.userId)) {
      // User is permitted - send response
      const replyToken = event.replyToken;
      const message = {
        type: 'text',
        text: 'Welcome! You are a permitted user.'
      };
      
      replyMessage(replyToken, message);
      
      // Log interaction
      sheet.appendRow([event.source.userId, new Date(), 'Accessed']);
    }
    
    return ContentService.createTextOutput(JSON.stringify({success: true})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error("Error processing request:", error);
    return ContentService.createTextOutput(JSON.stringify({success: false, message: error.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function replyMessage(replyToken, message) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [message]
    })
  };
  UrlFetchApp.fetch(url, options);
}

// Automation Script for "Column" on the sheet of Permitted Users (Main Sheet)
function setStatusValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Permitted Users");
  const range = sheet.getRange("E2:E"); // Status column
  
  // Create validation rule
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Active", "Pending", "Suspended"], true)
    .setAllowInvalid(false) // Reject invalid entries
    .setHelpText("Must be: Active, Pending, or Suspended")
    .build();
  
  // Apply to range
  range.setDataValidation(rule);
}

// Automation Script for "UserID" on the sheet of Permitted Users (Main Sheet)
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const UserIDcolumn = 1; // Column A (UserID)
  const regexPattern = "^U[a-f0-9]{32}$";

  // Check if the edit was made in the UserID column of the "Permitted Users" sheet.
  if (e.range.getColumn() === UserIDcolumn && sheet.getName() === "Permitted Users") {
    const inputUser = e.value;
    
    if (!RegExp(regexPattern).test(inputUser)) {
      e.range.setValue(""); // Empty cells if invalid
      SpreadsheetApp.getUi().alert(
        "ERROR: LINE UserID format is invalid!\n" +
        "Correct example: U4d3f1... (32 hex characters after 'U')"
      );
    }
  }
}

// Automation Script to create/update the named range dynamically
function createPermittedUsersNamedRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Permitted Users");
  
  // Get all non-empty UserIDs (skips header and blanks)
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(`A2:A${lastRow}`);
  
  // Create or update named range
  const namedRanges = ss.getNamedRanges();
  let existingRange = namedRanges.find(nr => nr.getName() === "PermittedUsers");
  
  if (existingRange) {
    existingRange.setRange(range); // Update existing
  } else {
    ss.setNamedRange("PermittedUsers", range); // Create new
  }
  
  console.log(`Named range "PermittedUsers" set to: ${range.getA1Notation()}`);
}

// Automation Script to implement the INDIRECT reference for my "PermittedUsers" named range
function checkUserPermission(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Option 1: Direct named range access
  const permittedUsers = ss.getRangeByName("PermittedUsers").getValues().flat();
  
  // Option 2: Using INDIRECT equivalent
  const indirectUsers = ss.getRange("PermittedUsers!A2:A").getValues().flat();
  
  return permittedUsers.includes(userId);
}

// Automation Script to implement the protection rules (lock columns and enforce required fields)
function fullProtectionSetup() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Permitted Users');
  const adminEmails = ['admin@yourdomain.com'];
  
  // 1. Unlock only Notes column
  const notesRange = sheet.getRange('F:F');
  const protection = sheet.protect();
  protection.setUnprotectedRanges([notesRange]);
  
  // 2. Set UserID validation
  enforceRequiredFields(); // From previous script
  
  // 3. Extra protection for other columns
  const otherColumns = sheet.getRange('B:E');
  otherColumns.protect()
    .setDescription('Admin-only columns')
    .addEditors(adminEmails);
  
  // 4. Set sheet-wide warning
  protection.setWarningOnly(true)
    .setDescription('Only Notes column is editable without admin permission');
}