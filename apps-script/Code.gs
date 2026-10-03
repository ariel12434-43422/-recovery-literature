/**
 * מאגר ספרות החלמה — חיבור חי בין תיקיית Drive לאתר
 *
 * מה הסקריפט עושה:
 *  • exportLibrary   — סורק את תיקיית "ספרות החלמה" ושומר תמונת מצב (JSON) בתוך התיקייה.
 *  • doGet           — מגיש את תמונת המצב לאתר (כשהסקריפט פרוס כ-Web App).
 *  • setupAutoUpdate — מתזמן סריקה אוטומטית כל שעה, כך שהאתר מתעדכן לבד.
 *
 * הפעלה חד-פעמית:
 *  1. מדביקים את הקובץ הזה במקום הקוד הקיים בפרויקט ב-script.google.com ושומרים.
 *  2. בוחרים בפונקציה setupAutoUpdate ולוחצים "הפעלה" (מאשרים הרשאות).
 *  3. פריסה ← פריסה חדשה ← סוג: אפליקציית אינטרנט
 *       הפעלה בתור: אני | מי יכול לגשת: כל אחד  ← פריסה.
 *  4. מעתיקים את כתובת ה-Web App (מסתיימת ב-/exec) ושמים אותה ב-index.html ב-LIVE_DATA_URL.
 *
 * מעכשיו: כל שינוי בדרייב (קובץ חדש, מחיקה, שינוי שם) יופיע באתר תוך שעה לכל היותר.
 * לעדכון מיידי — מריצים exportLibrary ידנית.
 */

const ROOT_FOLDER_ID = '1wHTMvZJxZ3tyQ56qLxEc__oBUu-ZTJqe';  // תיקיית "ספרות החלמה"
const OUTPUT_FILE_NAME = 'recovery-literature.json';
const UNCATEGORIZED = 'כללי';          // קבצים שיושבים ישירות בתיקייה הראשית
const CACHE_KEY = 'library-json';

// ───────────────────────── Web App ─────────────────────────

function doGet() {
  const json = readSnapshot_() || exportLibraryJson_();
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────── סטטיסטיקה ─────────────────────────
// האתר שולח לכאן אירוע אנונימי בכל פעולה (בלי שם, בלי IP, בלי חשבון).

const EVENT_NAMES = {
  visit: 'כניסה לאתר',
  view: 'תצוגה מקדימה',
  open: 'פתיחה בדרייב',
  download: 'הורדה',
  folder: 'כניסה לתיקייה',
  search: 'חיפוש'
};
const LOG_SHEET = 'פעולות';
const SUMMARY_SHEET = 'סיכום';

function doPost(e) {
  try {
    const ev = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const kind = EVENT_NAMES[ev.ev];
    if (!kind) return ok_();
    const cut = (s, n) => String(s == null ? '' : s).slice(0, n || 200);
    const ss = statsBook_();
    if (!ss) return ok_();
    const lock = LockService.getScriptLock();
    lock.tryLock(5000);
    ss.getSheetByName(LOG_SHEET).appendRow([
      new Date(), kind, cut(ev.folder), cut(ev.name), cut(ev.q, 100),
      ev.results === undefined || ev.results === '' ? '' : Number(ev.results) || 0,
      cut(ev.id, 60)
    ]);
    lock.releaseLock();
  } catch (err) { /* לא מפילים את האתר בגלל סטטיסטיקה */ }
  return ok_();
}

function ok_() {
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

function statsBook_() {
  const id = PropertiesService.getScriptProperties().getProperty('STATS_SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : null;
}

/** הפעלה חד-פעמית: יוצר את גיליון הסטטיסטיקה בדרייב ומדפיס את הקישור אליו. */
function setupStats() {
  let ss = statsBook_();
  if (!ss) {
    ss = SpreadsheetApp.create('מאגר ספרות החלמה — סטטיסטיקת שימוש');
    PropertiesService.getScriptProperties().setProperty('STATS_SHEET_ID', ss.getId());
  }
  ss.setSpreadsheetLocale('iw_IL');
  ss.setSpreadsheetTimeZone('Asia/Jerusalem');

  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.getSheets()[0];
    log.setName(LOG_SHEET);
  }
  log.setRightToLeft(true);
  log.getRange(1, 1, 1, 7).setValues([['זמן', 'פעולה', 'תיקייה', 'קובץ', 'חיפוש', 'תוצאות', 'מזהה קובץ']])
    .setFontWeight('bold').setBackground('#0f4f4c').setFontColor('#ffffff');
  log.setFrozenRows(1);
  log.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
  log.setColumnWidth(3, 260); log.setColumnWidth(4, 320); log.setColumnWidth(5, 180);

  let sum = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, 0);
  sum.clear();
  sum.setRightToLeft(true);
  const L = "'" + LOG_SHEET + "'!A2:G";
  const blocks = [
    ['A', 'הקבצים הכי נצפים', `=IFERROR(QUERY(${L},"select D, count(A) where (B='תצוגה מקדימה' or B='פתיחה בדרייב') and D<>'' group by D order by count(A) desc limit 30 label D 'קובץ', count(A) 'צפיות'",0),"עדיין אין נתונים")`],
    ['D', 'הקבצים הכי מורדים', `=IFERROR(QUERY(${L},"select D, count(A) where B='הורדה' and D<>'' group by D order by count(A) desc limit 30 label D 'קובץ', count(A) 'הורדות'",0),"עדיין אין נתונים")`],
    ['G', 'התיקיות הכי פופולריות', `=IFERROR(QUERY(${L},"select C, count(A) where B='כניסה לתיקייה' and C<>'' group by C order by count(A) desc limit 30 label C 'תיקייה', count(A) 'כניסות'",0),"עדיין אין נתונים")`],
    ['J', 'מה חיפשו', `=IFERROR(QUERY(${L},"select E, count(A) where B='חיפוש' and E<>'' group by E order by count(A) desc limit 30 label E 'חיפוש', count(A) 'פעמים'",0),"עדיין אין נתונים")`],
    ['M', 'חיפושים בלי תוצאות', `=IFERROR(QUERY(${L},"select E, count(A) where B='חיפוש' and F=0 and E<>'' group by E order by count(A) desc limit 30 label E 'חיפוש', count(A) 'פעמים'",0),"אין — כל החיפושים מצאו משהו")`],
    ['P', 'פעילות לפי יום', `=IFERROR(QUERY(${L},"select toDate(A), count(A) where A is not null group by toDate(A) order by toDate(A) desc limit 60 label toDate(A) 'יום', count(A) 'פעולות'",0),"עדיין אין נתונים")`]
  ];
  blocks.forEach(([col, title, formula]) => {
    sum.getRange(col + '4').setValue(title).setFontWeight('bold').setFontSize(12).setFontColor('#0f4f4c');
    sum.getRange(col + '5').setFormula(formula);
  });
  sum.getRange('A1').setValue('סיכום שימוש באתר').setFontSize(18).setFontWeight('bold');
  sum.getRange('A2').setFormula(`="כניסות: "&COUNTIF('${LOG_SHEET}'!B:B,"כניסה לאתר")&"   |   צפיות: "&(COUNTIF('${LOG_SHEET}'!B:B,"תצוגה מקדימה")+COUNTIF('${LOG_SHEET}'!B:B,"פתיחה בדרייב"))&"   |   הורדות: "&COUNTIF('${LOG_SHEET}'!B:B,"הורדה")&"   |   חיפושים: "&COUNTIF('${LOG_SHEET}'!B:B,"חיפוש")`);
  sum.getRange('P6:P').setNumberFormat('dd/MM/yyyy');
  ['A', 'D', 'G', 'J', 'M'].forEach(c => sum.setColumnWidth(sum.getRange(c + '1').getColumn(), 280));
  ss.setActiveSheet(sum);

  Logger.log('✔ גיליון הסטטיסטיקה מוכן: %s', ss.getUrl());
  return ss.getUrl();
}

// ───────────────────────── סריקה ─────────────────────────

function exportLibrary() {
  const json = exportLibraryJson_();
  const data = JSON.parse(json);
  Logger.log('✔ נסרקו %s קבצים ב-%s קטגוריות', data.stats.files, data.stats.categories);
  if (data.stats.notShared) {
    Logger.log('⚠ %s קבצים אינם פתוחים ל"כל מי שיש לו את הקישור" — מבקרים באתר לא יוכלו לפתוח אותם.', data.stats.notShared);
  }
}

function setupAutoUpdate() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'exportLibrary')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('exportLibrary').timeBased().everyHours(1).create();
  exportLibrary();
  Logger.log('✔ עדכון אוטומטי הופעל: הסריקה תרוץ כל שעה.');
}

function exportLibraryJson_() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const files = [];
  const categories = [];
  walk_(root, [], files, categories, isPublic_(root) === true);

  const data = {
    version: 2,
    generatedAt: new Date().toISOString(),
    root: { id: root.getId(), name: root.getName(), url: root.getUrl() },
    stats: {
      files: files.length,
      categories: categories.length,
      notShared: files.filter(f => f.shared === false).length
    },
    categories: categories,
    files: files
  };
  const json = JSON.stringify(data);
  saveSnapshot_(root, json);
  return json;
}

// ───────────────────────── פונקציות עזר ─────────────────────────

function walk_(folder, path, files, categories, inheritedPublic) {
  const fit = folder.getFiles();
  while (fit.hasNext()) {
    const f = fit.next();
    const n = f.getName();
    if (n === OUTPUT_FILE_NAME || n.startsWith('.') || n.startsWith('~$') || n === 'Icon\r') continue;
    files.push(describeFile_(f, path, inheritedPublic));
  }

  const subs = [];
  const dit = folder.getFolders();
  while (dit.hasNext()) {
    const d = dit.next();
    if (d.getName().startsWith('.') || d.getName() === '__MACOSX') continue;
    subs.push(d);
  }
  subs.sort((a, b) => a.getName().localeCompare(b.getName(), 'he'));

  subs.forEach(sub => {
    const subPath = path.concat(sub.getName());
    const before = files.length;
    const catIndex = categories.length;
    categories.push({ id: sub.getId(), name: sub.getName(), path: subPath, url: sub.getUrl(), fileCount: 0 });
    walk_(sub, subPath, files, categories, inheritedPublic || isPublic_(sub) === true);
    categories[catIndex].fileCount = files.length - before; // כולל תתי-קטגוריות
  });
}

function describeFile_(file, path, inheritedPublic) {
  const id = file.getId();
  const mime = file.getMimeType();
  const name = file.getName();
  const dot = name.lastIndexOf('.');
  const own = isPublic_(file);
  return {
    id: id,
    name: name,
    ext: dot > 0 ? name.slice(dot + 1).toLowerCase() : '',
    mimeType: mime,
    type: typeOf_(mime),
    path: path,
    category: path.length ? path.join(' / ') : UNCATEGORIZED,
    description: file.getDescription() || '',
    size: file.getSize(),
    updated: file.getLastUpdated().toISOString(),
    viewUrl: file.getUrl(),
    downloadUrl: downloadUrl_(id, mime),
    shared: own === true || inheritedPublic
  };
}

function typeOf_(mime) {
  if (mime === 'application/pdf') return 'pdf';
  if (mime === MimeType.GOOGLE_DOCS || /wordprocessingml|msword|rtf|text\/plain|opendocument\.text/.test(mime)) return 'doc';
  if (mime === MimeType.GOOGLE_SHEETS || /spreadsheetml|ms-excel|csv|opendocument\.spreadsheet/.test(mime)) return 'sheet';
  if (mime === MimeType.GOOGLE_SLIDES || /presentationml|ms-powerpoint|opendocument\.presentation/.test(mime)) return 'slides';
  if (/^image\//.test(mime)) return 'image';
  if (/^audio\//.test(mime)) return 'audio';
  if (/^video\//.test(mime)) return 'video';
  if (/zip|rar|7z|tar|gzip/.test(mime)) return 'archive';
  return 'other';
}

function downloadUrl_(id, mime) {
  switch (mime) {
    case MimeType.GOOGLE_DOCS:   return 'https://docs.google.com/document/d/' + id + '/export?format=pdf';
    case MimeType.GOOGLE_SHEETS: return 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
    case MimeType.GOOGLE_SLIDES: return 'https://docs.google.com/presentation/d/' + id + '/export/pdf';
    default:                     return 'https://drive.google.com/uc?export=download&id=' + id;
  }
}

function isPublic_(item) {
  try {
    const a = item.getSharingAccess();
    return a === DriveApp.Access.ANYONE || a === DriveApp.Access.ANYONE_WITH_LINK;
  } catch (e) {
    return null;
  }
}

// תמונת המצב נשמרת גם במטמון (לטעינה מהירה) וגם כקובץ בתיקייה (גיבוי)
function saveSnapshot_(root, json) {
  try {
    const cache = CacheService.getScriptCache();
    const chunks = json.match(/[\s\S]{1,90000}/g) || [];
    const map = { [CACHE_KEY + ':n']: String(chunks.length) };
    chunks.forEach((c, i) => map[CACHE_KEY + ':' + i] = c);
    cache.putAll(map, 6 * 60 * 60);
  } catch (e) { /* המטמון הוא שיפור ביצועים בלבד */ }

  const it = root.getFilesByName(OUTPUT_FILE_NAME);
  if (it.hasNext()) it.next().setContent(json);
  else root.createFile(OUTPUT_FILE_NAME, json, MimeType.PLAIN_TEXT);
}

function readSnapshot_() {
  try {
    const cache = CacheService.getScriptCache();
    const n = Number(cache.get(CACHE_KEY + ':n'));
    if (n) {
      const keys = Array.from({ length: n }, (_, i) => CACHE_KEY + ':' + i);
      const got = cache.getAll(keys);
      if (keys.every(k => got[k] != null)) return keys.map(k => got[k]).join('');
    }
  } catch (e) {}
  const it = DriveApp.getFolderById(ROOT_FOLDER_ID).getFilesByName(OUTPUT_FILE_NAME);
  return it.hasNext() ? it.next().getBlob().getDataAsString('UTF-8') : null;
}
