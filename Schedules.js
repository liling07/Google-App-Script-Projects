//late august 2025
//also worked off of Stephen's code + logic from last year
//helping produce schedules based on a spreadsheet of data
//main changes: added "batches" + triggers to run after each batch bc I ran into run-time errors T-T
//added a page split every 2 schedules for easier printing
//changed data structure and used maps for filling in placeholders 

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AutoFill Docs')
    .addItem('Create New Docs', 'createNewGoogleDocs')
    .addToUi();
}

const BATCH_SIZE = 100; 

function createNewGoogleDocs() {
  const scriptProps = PropertiesService.getScriptProperties();
  const startIdx = Number(scriptProps.getProperty('cursor') || 0);

  const templateFile = DriveApp.getFileById('{redacted}');
  const destinationFolder = DriveApp.getFolderById('{redacted}');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  const rows = sheet.getDataRange().getValues();
  const groups = [];
  let current = '', bucket = [];
  for (let i = 1; i < rows.length; i++) { 
    const r = rows[i];
    if (!r[0]) continue;
    const name = String(r[0]).trim();
    if (name !== current) {
      if (bucket.length) groups.push(bucket);
      current = name;
      bucket = [];
    }
    bucket.push(r);
  }
  if (bucket.length) groups.push(bucket);

  let bigDocId = scriptProps.getProperty('bigDocId');
  let bigDoc, bigBody;
  if (!bigDocId) {
    bigDoc = DocumentApp.create('All Students Details');
    bigBody = bigDoc.getBody();
    bigDocId = bigDoc.getId();
    const bigDocFile = DriveApp.getFileById(bigDocId);
    destinationFolder.addFile(bigDocFile);
    try { DriveApp.getRootFolder().removeFile(bigDocFile); } catch (_) {}
    scriptProps.setProperty('bigDocId', bigDocId);
  } else {
    bigDoc = DocumentApp.openById(bigDocId);
    bigBody = bigDoc.getBody();
  }

  let sectionsOnPage = Number(scriptProps.getProperty('sectionsOnPage') || 0);
  const endIdx = Math.min(startIdx + BATCH_SIZE, groups.length);
  for (let idx = startIdx; idx < endIdx; idx++) {
    const studentRows = groups[idx];


    const tempId = templateFile.makeCopy('tmp-' + safeFileName(studentRows[0][0])).getId();
    const tempDoc = DocumentApp.openById(tempId);
    const tempBody = tempDoc.getBody();


    rep(tempBody, 'Name',     studentRows[0][0]);
    rep(tempBody, 'Grade',    studentRows[0][1]);
    rep(tempBody, 'Advisory', studentRows[0][2]);


    const perMap = {};
    studentRows.forEach((row, i) => {
      const perText = (row[3] || '').toString().trim();
      let key = null;
      if (/^A\b/i.test(perText)) key = 'A';
      const m = /^([1-8])\b/.exec(perText);
      if (!key && m) key = m[1];
      if (!key) key = (i + 1 === 9) ? 'A' : String(i + 1);
      perMap[key] = { Per: perText, Cls: row[4] || '', Desc: row[5] || '', T: row[6] || '' };
    });

    ['A','1','2','3','4','5','6','7','8'].forEach(s => {
      const v = perMap[s] || { Per:'', Cls:'', Desc:'', T:'' };
      rep(tempBody, `Per${s}`,         v.Per);
      rep(tempBody, `Clssrm${s}`,      v.Cls);
      rep(tempBody, `Description${s}`, v.Desc);
      rep(tempBody, `TName${s}`,       v.T);
    });

   
    const n = tempBody.getNumChildren();
    for (let i = 0; i < n; i++) {
      const child = tempBody.getChild(i).copy();
      const t = child.getType();
      if (t === DocumentApp.ElementType.PARAGRAPH)      bigBody.appendParagraph(child.asParagraph());
      else if (t === DocumentApp.ElementType.LIST_ITEM) bigBody.appendListItem(child.asListItem());
      else if (t === DocumentApp.ElementType.TABLE)     bigBody.appendTable(child.asTable());
      else if (t === DocumentApp.ElementType.PAGE_BREAK)bigBody.appendPageBreak();
      else try { bigBody.appendParagraph(child.asText().getText()); } catch (_) {}
    }

    // page break after every 2 students (except at very end)
    sectionsOnPage++;
    const isLastOverall = (idx === groups.length - 1);
    if (sectionsOnPage === 2 && !isLastOverall) {
      bigBody.appendPageBreak();
      sectionsOnPage = 0;
    }


    try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {}
  }

  bigDoc.saveAndClose();

  if (endIdx < groups.length) {
    scriptProps.setProperty('cursor', String(endIdx));
    scriptProps.setProperty('sectionsOnPage', String(sectionsOnPage));
    scheduleContinuation_(); // run again shortly
    Logger.log(`Progress: ${endIdx}/${groups.length} students…`);
  } else {
    scriptProps.deleteProperty('cursor');
    scriptProps.deleteProperty('sectionsOnPage');
    Logger.log(`Done: ${bigDoc.getUrl()}`);
  }
}

function scheduleContinuation_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'createNewGoogleDocs')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('createNewGoogleDocs')
    .timeBased()
    .after(60 * 1000) 
    .create();
}

function rep(body, key, val) {
  const pattern = '\\{\\{' + escapeRegex(key) + '\\}\\}';
  body.replaceText(pattern, (val ?? '').toString());
}
function escapeRegex(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function safeFileName(s) { return String(s).replace(/[\\/:*?"<>|]/g, '_').slice(0, 100); }

/*
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create New Docs', 'createNewGoogleDocs');
  menu.addToUi();
}

function createNewGoogleDocs() {
  const googleDocTemplate = DriveApp.getFileById('11IUtuslgICqilwVBHMn9utnIjnGxfAYdfG7cgvtlrc4');
  const destinationFolder = DriveApp.getFolderById('19Unyjuhh-8jQp_urk0Kklv6qSazaEGJq');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test');
  const rows = sheet.getDataRange().getValues();

  let currentStudent = '';
  let studentRows = [];

  const bigDoc = DocumentApp.create('All Students Details');
  const body = bigDoc.getBody(); 
  const bigDocFile = DriveApp.getFileById(bigDoc.getId());
  destinationFolder.addFile(bigDocFile);
  DriveApp.getRootFolder().removeFile(bigDocFile);

  rows.forEach(function(row, index) {
    if (index === 0 || !row[0]) return;

    if (row[0] !== currentStudent) {
      if (studentRows.length > 0) {

        processStudentRows(studentRows, googleDocTemplate, body);
      }

      currentStudent = row[0];
      studentRows = [];
    }
    
    studentRows.push(row);
  });


  if (studentRows.length > 0) {
    processStudentRows(studentRows, googleDocTemplate, body);
  }

  bigDoc.saveAndClose();
  const url = bigDoc.getUrl();
  Logger.log(url);
}

function processStudentRows(studentRows, template, body) {

  const tempDoc = template.makeCopy().getId();
  const tempBody = DocumentApp.openById(tempDoc).getBody();;

  tempBody.replaceText(`{{Name}}`, studentRows[0][0]);
  tempBody.replaceText(`{{Grade}}`, studentRows[0][1]);
  tempBody.replaceText(`{{Advisory}}`, studentRows[0][2]);


  studentRows.forEach((row, index) => {
    let periodIndex = index + 1;
    if (periodIndex === 9) periodIndex = "A"; // handle "A" period

    tempBody.replaceText(`{{Per${periodIndex}}}`, row[3] || '');
    tempBody.replaceText(`{{Clssrm${periodIndex}}}`, row[4] || '');
    tempBody.replaceText(`{{Description${periodIndex}}}`, row[5] || '');
    tempBody.replaceText(`{{TName${periodIndex}}}`, row[6] || '');
  });

  // append template to the big doc
  const numChildren = tempBody.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const element = tempBody.getChild(i).copy();
    
    // check the type of element and append it as such
    if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
      body.appendParagraph(element);
    } else if (element.getType() == DocumentApp.ElementType.TABLE) {
      body.appendTable(element);
    } else if (element.getType() == DocumentApp.ElementType.LIST_ITEM) {
      body.appendListItem(element);
    } else {
      body.appendParagraph(element.asText().getText());
    }
  }
  //DriveApp.getFileById(tempDoc).setTrashed(true);
}
*/
