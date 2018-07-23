function createDocument() {
  var newDocument = DocumentApp.create("TestDoc");
  newDocument.getBody().appendParagraph("Test Creation!");
  
  var newDocumentFile = DriveApp.getFileById(newDocument.getId());
  DriveApp.getFoldersByName('Google Script examples').next().addFile(newDocumentFile);
  DriveApp.getRootFolder().removeFile(newDocumentFile);
}

function editDocument() {
  var directory = DriveApp.getFoldersByName('Google Script examples').next();
  var docFile = directory.getFilesByName("Skills Card Copy").next();
  var doc = DocumentApp.openById(docFile.getId());
  
//  for (var paragraph in doc.getBody().getParagraphs())
//    Logger.log(paragraph.getText());
  
  var paragraphs = doc.getBody().getParagraphs();
  
//  Logger.log(paragraphs[0].getText());
//  Logger.log(paragraphs[1].getText());
  
  for (i = 0; i < paragraphs.length; i++) {
    //Logger.log(paragraphs[i].setText("Paragraph" + i));
    Logger.log(paragraphs[i].getText());
  }
}
