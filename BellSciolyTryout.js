//september 2025
//essentially, we have a big google doc of questions (as of right now there are 89) so made a program to auto generate a google form :D
//had to search up a few things because i've never generate google forms before :D but was pretty fun to do over labor day weekend~
//also will comment soon.. as well as comment on my other projects :(

const DOC_ID = '{redacted}';
const DEFAULT_POINTS = 1;

function buildQuizFromDoc() {
  const lines = getDocLines(DOC_ID);

  const form = FormApp.getActiveForm();
  form.getItems().forEach(item => form.deleteItem(item));
  
  let currentSectionTitle = null;

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];

    if (!line) { i++; continue; }

    if (!line.endsWith('?') && !line.endsWith(':')) {
      const headerLines = [line];
      i++;
      while (i < lines.length && lines[i] && !lines[i].endsWith('?')) {
        headerLines.push(lines[i]);
        i++;
      }
      const header = headerLines.join(' ').trim();
      if (header) {
        form.addPageBreakItem().setTitle(header);
        currentSectionTitle = header;
      }
      continue;
    }

    const questionText = line;
    i++;

    const options = [];
    while (i < lines.length) {
      const optLine = lines[i];
      if (!optLine) { i++; continue; } 
      if (/^Answer\s*:/i.test(optLine)) break; 
      const m = optLine.match(/^([A-Z])\)\s*(.+)$/); 
      if (!m) break; 
      options.push({ letter: m[1], text: m[2].trim() });
      i++;
    }

    let answerLine = '';
    while (i < lines.length) {
      const maybeAns = lines[i];
      if (/^Answer\s*:/i.test(maybeAns)) {
        answerLine = maybeAns;
        i++;
        break;
      }
      if (maybeAns) break; 
      i++; 
    }

    let correctLetter = null;
    let correctText = null;

    if (answerLine) {
      const letterMatch = answerLine.match(/^Answer\s*:\s*([A-Z])\)/i);
      if (letterMatch) {
        correctLetter = letterMatch[1].toUpperCase();
      } else {
        const textPart = answerLine.replace(/^Answer\s*:\s*/i, '').trim();
        correctText = stripLeadingLetterParen(textPart).trim();
      }
    }

    let correctIndex = -1;
    if (correctLetter) {
      correctIndex = options.findIndex(o => o.letter === correctLetter);
    } else if (correctText) {
      const norm = s => s.replace(/\s+/g, ' ').trim().toLowerCase();
      const tgt = norm(correctText);
      correctIndex = options.findIndex(o => {
        const oText = norm(o.text);
        return oText === tgt || oText.includes(tgt) || tgt.includes(oText);
      });
    }

    if (options.length === 0) {
      form.addParagraphTextItem().setTitle(questionText).setPoints(DEFAULT_POINTS);
      continue;
    }

    const item = form.addMultipleChoiceItem().setTitle(questionText);
    const choices = options.map((opt, idx) =>
      item.createChoice(`${opt.text}`, idx === correctIndex)
    );
    item.setChoices(choices);
    item.setPoints(DEFAULT_POINTS);
  }
}

function getDocLines(docId) {
  const body = DocumentApp.openById(docId).getBody();
  const paras = body.getParagraphs();
  const lines = [];
  for (const p of paras) {
    lines.push((p.getText() || '').trim());
  }
  return lines;
}

function stripLeadingLetterParen(s) {
  return s.replace(/^[A-Z]\)\s*/, '');
}
