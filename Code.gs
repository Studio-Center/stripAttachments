// this is the id of the Auditions folder on google drive
// you can get this by running folderIds and then checking the logs
var FOLDER_ID = 'YOUR_FOLDER_ID';
var SPREADSHEET_ID = 'YOUR_SPEADSHEET_ID';
var SHEET_NAME = 'YOUR_SHEET_NAME';
var dropbox_token = 'YOUR_DROPBOX_Token';
var SLACK_URL = 'YOUR_SLACK_WEBHOOK_URL';

// use this to install the script
function Install() {
   ScriptApp.newTrigger('main')
   .timeBased().everyMinutes(5).create();
}

// this does all the work
function main() {
  // strip attachments
  var threads = GmailApp.getInboxThreads();
  var label = GmailApp.getUserLabelByName("_processed");
  try {
    for(j=0; j<threads.length; j++) {
      processThread(threads[j]);
    }
    // move message to _processed folder
    label.addToThreads(threads);
    GmailApp.moveThreadsToArchive(threads)
  }
  catch(error) {
    sendSlackMessage(error)
    console.error(error);
  }
}

// Process the Email Threads
function processThread(thread) {
  var messages = thread.getMessages();
  
  for(var j=0; j<messages.length; j++) {
    var message = messages[j];
    
    message.markRead();
    
    var attachments = message.getAttachments();
    
    for(var i=0; i<attachments.length; i++) {
      var attachment = attachments[i];
  
      saveAttachment(attachment);
      logMessage(message, attachment);
      
      Logger.log(attachment.getName());
    }
  }
}

// save attachments to gdrive
function saveAttachment(attachment) {
  
  var file = attachment.getName();
  var folder = DriveApp.getFolderById(Folder_ID);
  var check = folder.getFilesByName(file);
  var invalidExtensions = ['png', 'jpg', 'jpeg', 'pdf', 'doc', 'docx'];
  
  // strip the extension
  var extension = file.split('.');

  // check the filetype to make sure it isn't in the list of invalid types
  if(invalidExtensions.indexOf(extension[extension.length-1]) > -1) {
    
    Logger.log(file + ' is not audio.');
    return;
  }
  
  if(check.hasNext()) {
    Logger.log(file + ' already exists. File not overwritten.');
    return;
  }
  uploadGoogleFilesToDropbox(attachment); // Uplaod attachment to Drop Box
  folder.createFile(attachment).setName(file); // Save attachement to Google Drive
  Logger.log(file + ' saved.');
}

//  log attachments in Google Sheets
function logMessage(message, attachment) {
  var fileName = attachment.getName();
  var sender = message.getFrom();
  
  var matches = /(.*?)<(.*?)>/gi.exec(sender); //split senders name and email address apart
  
  var senderEmail = matches[2];
  var senderName = matches[1];
  var sentTime = message.getDate().toString();
  var messageBody = message.getPlainBody();
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var emailData= [];
  emailData[0] = [fileName, senderEmail, senderName, sentTime, messageBody];
  
  var nextRow = 2;
  var numRows = 1;
  var numCols = 5;
  sheet.insertRowsBefore(nextRow, 1); // insert a blank row at the top just before the Data
  sheet.getRange(nextRow,1,numRows,numCols).setValues(emailData); // write emailData to our new blank row
}

// Save attachements to Drop Box
function uploadGoogleFilesToDropbox(attachment) {
  
  var path = '/Auditions/' + attachment.getName();
  
  var headers = {
    "Content-Type": "application/octet-stream",
    'Authorization': 'Bearer ' + dropbox_token,
    "Dropbox-API-Arg": JSON.stringify({"path": path})
    
  };
  
  var options = {
    "method": "POST",
    "headers": headers,
    "payload": attachment.getBytes()
  };
  
  var apiUrl = "https://content.dropboxapi.com/2/files/upload";
  var response = JSON.parse(UrlFetchApp.fetch(apiUrl, options).getContentText());
  
  Logger.log("File uploaded successfully to Dropbox");
  
}

function testMessage(){
 sendSlackMessage("test Message");
  
  
}

// Send Slack Message if failure occurs
function sendSlackMessage(text, opt_channel) {
  var slackMessage = {
    text: text,
    icon_url:
        'https://www.gstatic.com/images/icons/material/product/1x/adwords_64dp.png',
    username: 'Auditions Bot',
    channel: opt_channel || '#auditions'
  };

  var options = {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(slackMessage)
  };
  UrlFetchApp.fetch(SLACK_URL, options);
}


// ----------------- Utlities -------------

// Log the name of every folder in the user's Drive.
function folderIds(){
  var folders = DriveApp.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    Logger.log(folder.getName() + '  -  ' + folder.getId());
  }
}