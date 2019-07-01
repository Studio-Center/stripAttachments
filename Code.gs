// this is the id of the Auditions folder on google drive
// you can get this by running folderIds and then checking the logs
var AUDITIONS_ID = '1tR82ZaGscjlPqAh_o8uKCuE6sYNt7XZT';

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
    
  for(j=0; j<threads.length; j++) {
      processThread(threads[j]);
    }
  // move message to _processed folder
  label.addToThreads(threads);
  GmailApp.moveThreadsToArchive(threads)
}

function processThread(thread) {
  var messages = thread.getMessages();
  
  for(var j=0; j<messages.length; j++) {
    
    var message = messages[j];
    
    message.markRead();
    
    var attachments = message.getAttachments();
    
    for(var i=0; i<attachments.length; i++) {
      var attachment = attachments[i];
      
      saveAttachment(attachment);
      
      Logger.log(attachment.getName());
    }
  }
}

// move attachments to gdrive
function saveAttachment(attachment) {
  
  var file = attachment.getName();
  var folder = DriveApp.getFolderById(AUDITIONS_ID);
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
  
  folder.createFile(attachment).setName(file);
  Logger.log(file + ' saved.');
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