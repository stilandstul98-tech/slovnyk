// ============================================================
// Вставте цей код у Google Apps Script
// Розширення → Apps Script → замінити весь вміст → Зберегти
// Потім: Деплой → Новий деплой → Веб-застосунок
//   Виконувати як: Я
//   Доступ: Усі
// Скопіюйте URL і вставте у налаштування додатку
// ============================================================

const SHEET_NAME = 'Words';
const HEADERS = ['id','word','translation','partOfSpeech','forms','sentences',
                 'visualScene','grammarTip','addedAt','nextReview','intervalIndex','reviewCount'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function rowToWord(headers, row) {
  const o = {};
  headers.forEach((h, i) => {
    if (h === 'sentences') {
      try { o[h] = JSON.parse(row[i]); } catch { o[h] = row[i] ? [String(row[i])] : []; }
    } else if (h === 'intervalIndex' || h === 'reviewCount') {
      o[h] = Number(row[i]) || 0;
    } else {
      o[h] = String(row[i] || '');
    }
  });
  return o;
}

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'getAll';
    const data = (e.parameter && e.parameter.data) ? JSON.parse(e.parameter.data) : {};
    return handleAction(action, data);
  } catch(err) {
    return out({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return handleAction(body.action || 'getAll', body.data || body);
  } catch(err) {
    return out({ success: false, error: err.message });
  }
}

function handleAction(action, data) {
  const sh = getSheet();

  if (action === 'getAll') {
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return out({ success: true, words: [] });
    const headers = rows[0];
    const words = rows.slice(1)
      .map(r => rowToWord(headers, r))
      .filter(w => w.id);
    return out({ success: true, words });
  }

  if (action === 'add') {
    const w = data;
    sh.appendRow([
      w.id, w.word, w.translation, w.partOfSpeech,
      w.forms, JSON.stringify(w.sentences || []),
      w.visualScene || '', w.grammarTip || '',
      w.addedAt, w.nextReview,
      w.intervalIndex || 0, w.reviewCount || 0
    ]);
    return out({ success: true });
  }

  if (action === 'delete') {
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sh.deleteRow(i + 1);
        return out({ success: true });
      }
    }
    return out({ success: false, error: 'Слово не знайдено' });
  }

  if (action === 'update') {
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sh.getRange(i + 1, 10).setValue(data.nextReview);
        sh.getRange(i + 1, 11).setValue(data.intervalIndex);
        sh.getRange(i + 1, 12).setValue(data.reviewCount);
        return out({ success: true });
      }
    }
    return out({ success: false, error: 'Слово не знайдено' });
  }

  return out({ success: false, error: 'Невідома дія: ' + action });
}

function out(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
