//august 2024
//scraped the scioly wiki to create a foundation for my fossil binder .. :) 
// this was kinda simple tbh 

function createFossilSlides() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Phylum'; 
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  var data = sheet.getDataRange().getValues();
  
  var presentationId = '{redacted}'; 
  var presentation = SlidesApp.openById(presentationId);
  
  const layoutSlide = presentation.getSlides()[1];
  
  for (let i = 1; i < data.length; i++) {
    const fossil = data[i];
    
    // Copy the layout slide for each fossil
    const newSlide = presentation.appendSlide(layoutSlide);

    // Replace placeholders with actual data
    const name = fossil[0]; // Name
    const description = fossil[1]; // Description
    const range = fossil[2];
    const adapt = fossil[3];
    const mol = fossil[4];
    const distribution = fossil[5];
    const etym = fossil[6];
    const add = fossil[7];

    newSlide.replaceAllText("<<Name>>", name, false);
    newSlide.replaceAllText("<<Etymology>>", etym, false);
    newSlide.replaceAllText("<<Range>>", range, false);
    newSlide.replaceAllText(" <<Mode of Life>>", mol, false);
    newSlide.replaceAllText(" <<Distribution>>", distribution, false);
    newSlide.replaceAllText(" <<Description>>", description, false);
    newSlide.replaceAllText(" <<Adaptation>>", adapt, false);
    newSlide.replaceAllText(" <<Additional Information>>", add, false);
  }
  
  presentation.saveAndClose();
}
