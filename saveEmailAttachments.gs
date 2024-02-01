// from @davidlee in MacAdmins.org Slack
// https://macadmins.slack.com/archives/C08PZM3EX/p1705679666550589?thread_ts=1705593705.892889&cid=C08PZM3EX

function saveAttachmentsToDrive() {
  var folderId = 'FOLDER-ID'; // The ID of the target folder
  var folder = DriveApp.getFolderById(folderId);
  var threads = GmailApp.getInboxThreads();
  var label = getOrCreateLabel('InvoiceSavedToDrive'); // Retrieve or create the label

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (!message.isUnread()) continue; // Skip if the message is read

      var sender = message.getFrom();
      var domain = sender.match(/@([\w.-]+)/)[1]; // Extract domain from email address
      var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
      var prefix = domain + " - INVOICE - " + date + " - ";

      var attachments = message.getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var contentType = attachment.getContentType();

        // Check for specific file types
        if (contentType === 'application/pdf' || contentType === 'application/vnd.ms-excel' || 
            contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
            contentType === 'application/msword' || 
            contentType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
          
          var fileName = prefix + attachment.getName();
          folder.createFile(attachment).setName(fileName);
        }
      }
      label.addToThread(message.getThread()); // Apply the label to the thread
    }
  }
}

// Function to retrieve or create a label
function getOrCreateLabel(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}
