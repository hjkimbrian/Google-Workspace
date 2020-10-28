//Add Menu Item to Lock Google Docs Files
//Credit goes to @mike-bc on https://saasops.community
//https://better-it.slack.com/archives/C2DMQ0BGE/p1602277511365000?thread_ts=1602259544.344300&cid=C2DMQ0BGE

var fileId=DocumentApp.getActiveDocument();
function onOpen() {
  var ui = DocumentApp.getUi();
  // Or SpreadsheetApp or FormApp.
  ui.createMenu('Lock File')
      .addItem('Lock File', 'lockFile')
      .addToUi();
}
function lockFile(fileId) {
  var file = Drive.Files.get(fileId)
  var resource = 
      {
        "contentRestrictions": [
          {
            "readOnly": true
          }
        ]
      }
  Drive.Files.update(resource, fileId)
}
