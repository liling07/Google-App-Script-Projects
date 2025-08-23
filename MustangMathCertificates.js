//essentially we had a folder of score report pdfs for individual students who can be identified by unique IDs, we then had a google slide template for certificates; 
//this program pulls the score report from the folder + generates a certificate and sends it to them :) 
/// was actually one of my first projects :D completed in september 2024! and.. took like 7-8 hours cuz i was slow and didn't know what i was doing


function autoEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Special'; 
  var competitorSheet = ss.getSheetByName(sheetName);
  var competitorData = competitorSheet.getDataRange().getValues();  
  var folderId = {redacted};  
  var folder = DriveApp.getFolderById(folderId);
  var pdfFiles = folder.getFiles();
  
  
  var pdfMap = {};
  
  while (pdfFiles.hasNext()) {
    var file = pdfFiles.next();
    if (file.getName().endsWith('.pdf')) {
      var fileName = file.getName().replace('.pdf', '').trim();  
      pdfMap[fileName] = file;  
    }
  }

  
  for (var i = 35; i < 127; i++) {
    var competitorName = competitorData[i][2];  
    var email = competitorData[i][4];            
    var compID = competitorData[i][0];           

    if (email && pdfMap["individual_" + compID]) {  
      // Export certificate as PDF
      var slideId = {redacted};   
      var slideDeck = SlidesApp.openById(slideId);
      
      
      if (i < slideDeck.getSlides().length) {
        var slide = slideDeck.getSlides()[i];
        
        
        var tempPresentation = SlidesApp.openById('{redacted}');
        var tempSlide = tempPresentation.appendSlide(slide);
        
        
        if (tempPresentation.getSlides().length > 1) {
          tempPresentation.getSlides()[0].remove();
        }
        
        
        tempPresentation.saveAndClose();

        
        var tempFile = DriveApp.getFileById(tempPresentation.getId());
        var pdfBlob = tempFile.getAs(MimeType.PDF).setName(competitorName + '_certificate.pdf');
        
        
        tempFile.setTrashed(true);
        
        
        var subject = 'Mustang Math Mania 2024 Score Report and Certificate';
       var message = 'Dear ' + competitorName +','+ 'message'

        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: message,
          attachments: [pdfBlob, pdfMap["individual_" + compID].getAs(MimeType.PDF)]
        });
        
        Logger.log("Email sent to " + email);
      } else {
        Logger.log("No slide available for index: " + i);
      }
    } else {
      Logger.log("No matching email or PDF found for " + competitorName);
    }
  }
}
