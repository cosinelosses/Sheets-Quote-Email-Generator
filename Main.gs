function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      options = [{name: "Create Copy and Insert Above", functionName: "createRowCopyInsert"}, {name: "Generate Email", functionName: "createDraftPDFAttached"},
                 {name: "Fix Drive Links", functionName: "fix_links"}];
  ss.addMenu("Custom", options);
}

// CONSTANTS - Enter Columns Here
var documentColumn = 3; 
var clientEmailColumn = 5; 

function createRowCopyInsert() {
  // get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; 
  
  // get the active row nummber and chop (string)
  var activeRow = String(ss.getActiveRange().getA1Notation()).substring(1); 
  var activeRow = parseInt(activeRow); 
  
  // formula contains link and id 
  var formula = sheet.getRange(activeRow, documentColumn).getFormula(); 
  var formulaAsString = String(formula); 
  
  //check to make sure not already valid     
  var endOfFormulaString = formulaAsString.slice(-12);
  Logger.log(endOfFormulaString != ',\"Document\")');
  
  if (endOfFormulaString != ',\"Document\")') {
    var id = 'bad';
    SpreadsheetApp.getUi().alert('no file');
  }
  
  // split the url by slashes
  var parts = formulaAsString.split('/'); 
  
  if (parts.length > 7) {
      id = parts[7];  
    }  
  
  // insert the new row
  sheet.insertRowBefore(activeRow); 
  
  // copy the data from the old row to the new
  var newRowNum = activeRow; 
  var newDocLinkRange = sheet.getRange(newRowNum, documentColumn);

  // copy the values from above column (1 row, 8 columns over)
     // change to getMaxColumns 
  var firstRange = sheet.getRange(activeRow + 1, documentColumn + 1, 1, 9);
  var secondRange = sheet.getRange(activeRow, documentColumn + 1);
  firstRange.copyTo(secondRange);
  Logger.log(secondRange.getA1Notation()); 
     
  // get currently selected file link to copy with id 
  var oldFile = DriveApp.getFileById(id);
    
  // create copy of old file
  var newFile = oldFile.makeCopy(oldFile.getName().slice(0, -4) + ' - Copy'); 
  
  // get link and id of new file
  var newFileLink = newFile.getUrl();
  var newFileId = newFile.getId();
  
  // insert new copy 
  //check to make sure there is a link
  if (String(newFileLink) != 'undefined') {
     newDocLinkRange.setFormula("=hyperlink(\"" + String(newFileLink) + "\", \"Document\")"); 
  }
  else {
    newDocLinkRange.setValue('Not in Drive');
    newDocLinkRange.setBackground('red');
  }
 
}

function createDraftPDFAttached() {
  // get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; 
  
  // get active row number and chop off first letter 
  var activeRow = String(ss.getActiveRange().getA1Notation()).substring(1); 
  var activeRow = parseInt(activeRow); 
  
  // get formula of new doc cell
  var formula = ss.getActiveRange().getFormula(); 
  var formulaAsString = String(formula); 
  
  // extract id (8th item when string is split by '/')
  var parts = String(sheet.getRange(activeRow, documentColumn).getFormula()).split('/'); 
  var id = parts[7]; 
  
  // create pdf file to attach
  var file = DriveApp.getFileById(id);
  Logger.log(file.getName());
  
  // create new pdf file 
  var pdfFile = DriveApp.createFile(file.getAs('application/pdf'));
  Logger.log(pdfFile.getName()); 
  pdfFile.setName('Quote ' + file.getName() + '.pdf');                              
  
  // create new message object 
  var msg = new Message();
  
  // give message properties (may not need to convert to string, check)
  var recipient = String(sheet.getRange(activeRow, clientEmailColumn).getValue());
  msg.to = recipient;
  msg.subject = 'Quotation';
  msg.body = '<p>This is a quote email with pdf attached.</p>'; 
  
  //attach file 
  var fileAttach = DriveApp.getFileById(pdfFile.getId());
  fileAttach.setName(pdfFile)
  msg.attachments = [fileAttach];
  
  //messages will be saved as a draft to be opened with link
  msg.createDraft();
  
  // grab latest draft 
  var draftId = GmailApp.getDraftMessages()[0].getId();
  var linkToDraft = 'https://mail.google.com/mail/u/0/#drafts?compose=' + draftId;
  Logger.log(linkToDraft);
  
  // show confirmation dialog 
  var ui = HtmlService.createTemplateFromFile('linkDialog');
  ui.draftLink = linkToDraft; 
  ui.draftName = draftId
  ui.docEditLink = 'exmaple.com';
  var uiout = ui.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle("Draft Created").setHeight(75).setWidth(200);
  SpreadsheetApp.getUi().showDialog(uiout);
}
