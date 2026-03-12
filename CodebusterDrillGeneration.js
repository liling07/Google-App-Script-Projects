/*template slides: https://docs.google.com/presentation/d/1LGYtXHjBr4_2GZpT-bKRLkXg38DRkvvFUVmQQle6vgE/edit?usp=sharing
I think making a copy of it works :)  
pretty simple program- but yay! */ 

function onOpen(e){
  SlidesApp.getUi()
    .createMenu('Generate Practice')
    .addItem('Baconian / Morse Fill In (2)', 'alphabetTranslation') 
    .addItem('Baconian / Morse Fill In (5)', 'alphabetTranslations')
    .addItem('Letter Replacement (2)', 'alphabetReplacement')
    .addItem('Letter Replacement (5)', 'alphabetReplacements')
    .addToUi();
}

function alphabetTranslation() {

  const presentation = SlidesApp.getActivePresentation();
  const template = presentation.getSlides()[0];

  for(let k = 0; k<2; k++){
    const slide = presentation.appendSlide(template);
    const tables = slide.getTables();
    for(let i =0; i< tables.length; i++){
      const table = tables[i];
      for(let j =0; j<table.getNumColumns()-1; j++){
        const cell = table.getCell(0, j);
        cell.getText().setText(generateRandomLetter());

        cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      }
    }

  }
  
}

function alphabetTranslations() {
  const presentation = SlidesApp.getActivePresentation();
  const template = presentation.getSlides()[0];

  for(let k = 0; k<5; k++){
    const slide = presentation.appendSlide(template);
    const tables = slide.getTables();
    for(let i =0; i< tables.length; i++){
      const table = tables[i];
      for(let j =0; j<table.getNumColumns()-1; j++){
        const cell = table.getCell(0, j);
        cell.getText().setText(generateRandomLetter());

        cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      }
    }

  }
  
}

function alphabetReplacement(){
  const presentation = SlidesApp.getActivePresentation();
  const template = presentation.getSlides()[1];
  for(let k = 0; k <2; k++){
    const slide = presentation.appendSlide(template);
    const tables = slide.getTables();
    for(let i = 0; i<tables.length; i++){
      table = tables[i];
      if(table.getNumRows() == 1){
        const j = Math.floor(Math.random() * (50)) + 200;
        let string = "";
        for(let k = 0; k<j; k++){
          string += generateRandomLetterSpace();
        }
        table.getCell(0,0).getText().setText(string);
      }
      else{
        const scrambledReplacement = scrambleString('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
        for(let j = 3; j<29; j++){
          const cell = table.getCell(1,j);
          cell.getText().setText(scrambledReplacement.charAt(j-3));
          cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
          cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        }
      }
    }
  }
  

}

function alphabetReplacements(){
  const presentation = SlidesApp.getActivePresentation();
  const template = presentation.getSlides()[1];
  for(let k = 0; k <5; k++){
    const slide = presentation.appendSlide(template);
    const tables = slide.getTables();
    for(let i = 0; i<tables.length; i++){
      table = tables[i];
      if(table.getNumRows() == 1){
        const j = Math.floor(Math.random() * (50)) + 200;
        let string = "";
        for(let k = 0; k<j; k++){
          string += generateRandomLetterSpace();
        }
        table.getCell(0,0).getText().setText(string);
      }
      else{
        const scrambledReplacement = scrambleString('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
        for(let j = 3; j<29; j++){
          const cell = table.getCell(1,j);
          cell.getText().setText(scrambledReplacement.charAt(j-3));
          cell.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
          cell.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        }
      }
    }
  }
  

}

function generateRandomLetter() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const randomIndex = Math.floor(Math.random() * characters.length);
  const randomLetter = characters.charAt(randomIndex);
  return randomLetter;
}

function generateRandomLetterSpace() {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ             ';
  const randomIndex = Math.floor(Math.random() * characters.length);
  const randomLetter = characters.charAt(randomIndex);
  return randomLetter;
}

function scrambleString(inputString) {
  let chars = inputString.split('');

  // Fisher-Yates shuffle algorithm
  for (let i = chars.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  const scrambledString = chars.join('');
  return scrambledString;
}

