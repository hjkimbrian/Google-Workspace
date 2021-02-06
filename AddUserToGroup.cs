/*
Purpose of the Script:
Add Shared Drive Access Requestor to Google Group used for ACL
Shared Drive Link: https://drive.google.com/drive/u/0/folders/0ANkIXd3coZwTUk9PVA
Link to join group: https://groups.google.com/a/workspaceadmins.org/g/workspace-admins-community-comment
Workspace Admins: https://workspaceadmins.org
*/

function addtoGroup() {
  var interval = 5;    //  if the script runs every 5 minutes; change otherwise
  var emails = [];
  var timeFrom = Math.floor(Date.now()/1000) - 60 * interval;
  var group = GroupsApp.getGroupByEmail("workspace-admins-community-comment@workspaceadmins.org");
  var threads = GmailApp.search('subject: "Workspace Admins [Public] - Request for access" after:' + timeFrom);
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    if (messages.length == 1) {      // only deal with threads with one message
      var ReplyToEmail = messages[0].getReplyTo().match(/([^<]+@[^>]+)/)[1];
      emails.push(ReplyToEmail);
      messages[0].replyAll("We have added you to the Google Group, workspace-admins-community-comment@workspaceadmins.org. You should have access to the Shared Drive shortly.")
    }
  }
  for (i=0; i < emails.length; i++) {
    try {
      addMember (emails[i], group);
    }
    catch(e) {
      console.error(e);
      continue;
    }
  }
}

function addMember (email, group) {
  var hasMember = group.hasUser(email);
  Utilities.sleep(1000);

  if(!hasMember) {
    var newMember = {email: email,
    role: "MEMBER",
    delivery_settings: "NONE"};
    AdminDirectory.Members.insert(newMember, "workspace-admins-community-comment@workspaceadmins.org");
  }
}
