/*************************************************************************
 * DAILY WORK STATUS + ATTENDANCE AUTOMATION
 * STANDALONE (unbound) VERSION
 * -----------------------------------------------------------------------
 * WHAT THIS AUTOMATES
 *  1. "Present Status" (col C) of PAST dates becomes read-only — only
 *     today's row can be edited. (Full-access admins can always edit.)
 *  2. A professional daily email report — Total Present/Absent, who was
 *     present/absent, and what each present person worked on.
 *  3. Today's block of rows (Date/Name/KRA/etc.) is auto-created every
 *     day by cloning yesterday's employee list.
 *
 * NIGHT-SHIFT AWARE "DAY" BOUNDARY
 * Your shift runs 3 PM to 8 AM. This script uses a configurable "Shift
 * start hour" (default 15 = 3 PM) as the day boundary. Any real clock
 * time BEFORE that hour still belongs to the PREVIOUS calendar date's
 * block. The automated email and auto-absent logic will target that
 * logical "yesterday".
 *
 * -----------------------------------------------------------------------
 * ONLY CHANGE IN THIS VERSION (baaki sab pehle jaisa):
 *  - KRA column (col D) ab pichle din se copyTo() se clone hota hai, isliye
 *    MERGED cells + colour + font + alignment sab hu-ba-hu aate hain.
 *  - Naya helper fixTodayKra(): aaj ke already-bane (bina merge) block ko
 *    theek karne ke liye ek baar manually run karne wala function.
 *************************************************************************/

/* ============================ CONFIG ============================ */
const SHEET_ID = '1KVL-eTqZpW1Yp_RG3Mp45MEF_dHODAADiNNneF2WuI0';   // <-- REQUIRED

const DATE_COL = 1, NAME_COL = 2, STATUS_COL = 3, KRA_COL = 4, WORK_COL = 5, ADD_COL = 6;
const HEADERS = ['Date', 'Employee Names', 'Present Status', 'Key Responsibility Area (KRA)', 'Today Working ', 'Additional Work Done'];
const STATUS_LOCK_TAG = 'status-lock:no-manual-edit-past-dates';
const DRIVER_INTERVAL_MIN = 10;

const PRESENT_COLOR = '#34a853', PRESENT_TEXT = '#ffffff';
const ABSENT_COLOR  = '#ea4335', ABSENT_TEXT  = '#ffffff';
const HOLIDAY_COLOR = '#4285f4', HOLIDAY_TEXT = '#ffffff';

const FULL_ACCESS_EMAILS_DEFAULT = ['jeetuyogi.vefogix@gmail.com'];
const MAIL_RECIPIENTS_DEFAULT = ['jeetuyogi.vefogix@gmail.com'];

/* ============================ SETUP ============================= */
function setup() {
  initDefaults_();

  const mine = ['driver'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (mine.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('driver').timeBased().everyMinutes(DRIVER_INTERVAL_MIN).create();

  const sheet = getDataSheet_();
  const cfg = getConfig_();
  applyStatusFormatting_(sheet);
  ensureTodayBlock_(sheet, cfg);
  applyStatusLock_(sheet, cfg);

  Logger.log('Setup done. Deploy > New deployment > Web app to get your Admin Control Panel.');
}

/* ========================= WEB APP ENTRY ========================= */
function doGet(e) {
  return HtmlService.createHtmlOutput(buildPanelHtml_())
    .setTitle('Work Status Automation Panel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== PRIVATE CONFIG STORAGE ==================== */
function props_() { return PropertiesService.getScriptProperties(); }

function initDefaults_() {
  const p = props_();
  const d = {
    tabOverride: '',
    shiftStartHour: '15',
    autoAbsent: 'ON', autoAbsentHour: '9',
    autoEmail: 'ON', emailHour: '9', emailMin: '30',
    recipients: MAIL_RECIPIENTS_DEFAULT.join(', '),
    fullAccess: FULL_ACCESS_EMAILS_DEFAULT.join(', ')
  };
  Object.keys(d).forEach(function (k) { if (p.getProperty(k) === null) p.setProperty(k, d[k]); });
}

function uiGetConfig() {
  const p = props_();
  return {
    tabOverride: p.getProperty('tabOverride') || '',
    shiftStartHour: p.getProperty('shiftStartHour') || '15',
    autoAbsent: p.getProperty('autoAbsent') || 'ON',
    autoAbsentHour: p.getProperty('autoAbsentHour') || '9',
    autoEmail: p.getProperty('autoEmail') || 'ON',
    emailHour: p.getProperty('emailHour') || '9',
    emailMin: p.getProperty('emailMin') || '30',
    recipients: p.getProperty('recipients') || MAIL_RECIPIENTS_DEFAULT.join(', '),
    fullAccess: p.getProperty('fullAccess') || FULL_ACCESS_EMAILS_DEFAULT.join(', ')
  };
}

function uiSaveConfig(d) {
  props_().setProperties({
    tabOverride: d.tabOverride || '',
    shiftStartHour: String(d.shiftStartHour),
    autoAbsent: d.autoAbsent, autoAbsentHour: String(d.autoAbsentHour),
    autoEmail: d.autoEmail, emailHour: String(d.emailHour), emailMin: String(d.emailMin),
    recipients: d.recipients, fullAccess: d.fullAccess
  }, false);
  const sheet = getDataSheet_(), cfg = getConfig_();
  applyStatusLock_(sheet, cfg);
  return 'Settings saved.';
}
function uiSendTest()        { sendStatusEmail_(getDataSheet_(), getConfig_()); return 'Report sent.'; }
function uiCreateTodayBlock(){ const s = getDataSheet_(); ensureTodayBlock_(s, getConfig_()); applyStatusLock_(s, getConfig_()); return 'Today\'s rows are ready.'; }
function uiLockNow()         { applyStatusLock_(getDataSheet_(), getConfig_()); return 'Past dates locked.'; }

function splitList_(s, def) {
  const list = String(s || '').split(/[,\n;]+/).map(function (x) { return x.trim(); }).filter(Boolean);
  return list.length ? list : def;
}
function isOn_(v) { return /^(on|yes|true|1)$/i.test(String(v || '').trim()); }

function getConfig_() {
  const c = uiGetConfig();
  return {
    tabOverride: c.tabOverride,
    shiftStartHour: parseInt(c.shiftStartHour, 10) || 15,
    autoAbsent: isOn_(c.autoAbsent), autoAbsentHour: parseInt(c.autoAbsentHour, 10) || 9,
    autoEmail: isOn_(c.autoEmail), emailHour: parseInt(c.emailHour, 10) || 9, emailMin: parseInt(c.emailMin, 10) || 30,
    recipients: splitList_(c.recipients, MAIL_RECIPIENTS_DEFAULT),
    fullAccess: splitList_(c.fullAccess, FULL_ACCESS_EMAILS_DEFAULT)
  };
}

/* ==================== LOGICAL DATE / SHIFT ENGINE ================= */
function getLogicalDate_(cfg) {
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, tz_(), 'H'));
  if (hour < cfg.shiftStartHour) {
    return new Date(now.getTime() - 24 * 60 * 60 * 1000);
  }
  return now;
}

function getRelativeMinutes_(hour, min, shiftStartHour) {
  let h = hour < shiftStartHour ? hour + 24 : hour;
  return (h - shiftStartHour) * 60 + min;
}

/* =========================== PANEL HTML ========================== */
function buildPanelHtml_() {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Work Status Automation</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {
    --primary: #4f46e5; --primary-hover: #4338ca; --primary-light: #e0e7ff;
    --bg: #f8fafc; --surface: #ffffff;
    --text-main: #0f172a; --text-muted: #64748b;
    --border: #e2e8f0; --ring: rgba(79, 70, 229, 0.2);
    --success: #10b981; --error: #ef4444;
    --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
    --shadow-md: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', sans-serif; background: var(--bg); color: var(--text-main); line-height: 1.5; }
  .header { background: var(--surface); border-bottom: 1px solid var(--border); padding: 24px 20px; position: sticky; top: 0; z-index: 10; box-shadow: var(--shadow-sm); }
  .header-content { max-width: 600px; margin: 0 auto; display: flex; align-items: center; justify-content: space-between; }
  .header h2 { font-size: 20px; font-weight: 700; display: flex; align-items: center; gap: 8px; }
  .header .badge { background: var(--primary-light); color: var(--primary); font-size: 11px; padding: 4px 8px; border-radius: 12px; font-weight: 600; text-transform: uppercase; }
  .wrap { max-width: 600px; margin: 0 auto; padding: 32px 20px 60px; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 16px; padding: 24px; margin-bottom: 24px; box-shadow: var(--shadow-sm); transition: box-shadow 0.2s; }
  .card:hover { box-shadow: var(--shadow-md); }
  .card h3 { font-size: 14px; font-weight: 600; margin-bottom: 16px; display: flex; align-items: center; gap: 6px; }
  label { display: block; font-size: 13px; font-weight: 500; margin: 12px 0 6px; }
  input, select, textarea { width: 100%; padding: 10px 14px; border: 1px solid var(--border); border-radius: 10px; font-size: 14px; font-family: inherit; color: var(--text-main); }
  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 4px var(--ring); }
  textarea { min-height: 80px; resize: vertical; }
  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .flex-between { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
  .hint { font-size: 12px; color: var(--text-muted); margin-top: 6px; line-height: 1.6; }
  button { display: inline-flex; align-items: center; justify-content: center; gap: 8px; border: none; border-radius: 10px; padding: 12px 20px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; position: relative; }
  button:active { transform: scale(0.98); }
  button:disabled { opacity: 0.7; cursor: not-allowed; }
  .btn-primary { background: var(--primary); color: #fff; width: 100%; box-shadow: 0 4px 6px -1px rgba(79, 70, 229, 0.3); }
  .btn-primary:hover:not(:disabled) { background: var(--primary-hover); }
  .btn-secondary { background: #fff; border: 1px solid var(--border); color: var(--text-main); width: 100%; }
  .btn-secondary:hover:not(:disabled) { background: var(--bg); }
  .actions-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-top: 16px; }

  /* Spinner & Loading Fixes */
  .spinner { width: 16px; height: 16px; border-radius: 50%; animation: spin 0.8s linear infinite; display: none; position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); }
  .btn-primary .spinner { border: 2px solid rgba(255,255,255,0.3); border-top-color: #fff; }
  .btn-secondary .spinner { border: 2px solid rgba(79, 70, 229, 0.2); border-top-color: var(--primary); }
  .btn-loading .btn-text { opacity: 0; }
  .btn-loading .spinner { display: block; }

  @keyframes spin { to { transform: translate(-50%, -50%) rotate(360deg); } }
  #toast-container { position: fixed; bottom: 24px; right: 24px; display: flex; flex-direction: column; gap: 12px; z-index: 1000; }
  .toast { display: flex; align-items: center; gap: 10px; min-width: 280px; padding: 14px 20px; border-radius: 12px; background: #fff; font-size: 14px; font-weight: 500; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); border-left: 4px solid var(--primary); opacity: 0; transform: translateX(50px); animation: slideIn 0.3s forwards; }
  .toast.success { border-left-color: var(--success); }
  .toast.error { border-left-color: var(--error); }
  @keyframes slideIn { to { opacity: 1; transform: translateX(0); } }
  @keyframes fadeOut { to { opacity: 0; transform: translateY(20px); } }
  #app-loader { position: fixed; inset: 0; background: var(--bg); display: flex; align-items: center; justify-content: center; z-index: 9999; flex-direction: column; gap: 16px; transition: opacity 0.3s; }
  .app-spinner { width: 40px; height: 40px; border: 3px solid var(--border); border-top-color: var(--primary); border-radius: 50%; animation: spin-normal 1s linear infinite; }
  @keyframes spin-normal { to { transform: rotate(360deg); } }
</style>
</head>
<body>

<div id="app-loader"><div class="app-spinner"></div><div style="font-weight:500;color:var(--text-muted);">Loading workspace...</div></div>

<div class="header"><div class="header-content"><h2>⚙️ Control Panel</h2><span class="badge">Admin Only</span></div></div>

<div class="wrap">
  <div class="card">
    <h3>⏱️ Shift & Attendance Rules</h3>
    <div class="grid-2">
      <div><label>Shift start hour (0-23)</label><input id="shiftStartHour" type="number" min="0" max="23"></div>
      <div><label>Cutoff hour (0-23)</label><input id="autoAbsentHour" type="number" min="0" max="23"></div>
    </div>
    <div class="hint">Shift resets at the start hour. Data prior to this hour belongs to yesterday (e.g. 3 AM still maps to yesterday's date).</div>
    <div class="flex-between" style="margin-top:16px;padding-top:16px;border-top:1px solid var(--border);">
      <div style="font-size: 14px; font-weight: 500;">Auto-mark Absent</div>
      <select id="autoAbsent" style="width:auto;margin:0;"><option>ON</option><option>OFF</option></select>
    </div>
  </div>

  <div class="card">
    <h3>📧 Daily Report Settings</h3>
    <div class="flex-between">
      <div style="font-size: 14px; font-weight: 500;">Automated Emailing</div>
      <select id="autoEmail" style="width:auto;margin:0;"><option>ON</option><option>OFF</option></select>
    </div>
    <div class="grid-2">
      <div><label>Send hour (0-23)</label><input id="emailHour" type="number" min="0" max="23"></div>
      <div><label>Send minute</label><input id="emailMin" type="number" min="0" max="59"></div>
    </div>
    <label>Recipients (Comma separated)</label>
    <textarea id="recipients" placeholder="admin@company.com"></textarea>
  </div>

  <div class="card">
    <h3>👥 Access Management</h3>
    <label>Full Access Admins (Comma separated)</label>
    <textarea id="fullAccess"></textarea>
    <div class="hint">These users bypass date locks and can edit past attendance directly in the sheet.</div>
  </div>

  <div class="card">
    <h3>🛠️ Advanced Options</h3>
    <label>Force Specific Sheet Tab</label>
    <input id="tabOverride" placeholder="e.g. Sep Month 2026 (Leave blank for auto)">
  </div>

  <button class="btn-primary" onclick="save(this)"><span class="spinner"></span><span class="btn-text">💾 Save All Settings</span></button>
  <div class="actions-grid">
    <button class="btn-secondary" onclick="act('uiCreateTodayBlock', this, '🗓️ Preparing...')"><span class="spinner"></span><span class="btn-text">🗓️ Prep Today's Rows</span></button>
    <button class="btn-secondary" onclick="act('uiLockNow', this, '🔒 Locking...')"><span class="spinner"></span><span class="btn-text">🔒 Lock Past Dates</span></button>
    <button class="btn-secondary" onclick="act('uiSendTest', this, '📧 Sending...')"><span class="spinner"></span><span class="btn-text">📧 Test Email</span></button>
  </div>
</div>

<div id="toast-container"></div>

<script>
  const FIELDS = ['tabOverride','autoAbsent','autoAbsentHour','autoEmail','emailHour','emailMin','recipients','fullAccess'];
  google.script.run.withSuccessHandler(function(c) {
    FIELDS.forEach(k => { if(document.getElementById(k)) document.getElementById(k).value = c[k]; });
    const loader = document.getElementById('app-loader');
    loader.style.opacity = '0';
    setTimeout(() => loader.style.display = 'none', 300);
  }).uiGetConfig();

  function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = 'toast ' + type;
    toast.innerHTML = '<span class="toast-icon">' + (type === 'success' ? '✅' : '⚠️') + '</span><span>' + message + '</span>';
    container.appendChild(toast);
    setTimeout(() => { toast.style.animation = 'fadeOut 0.3s forwards'; setTimeout(() => toast.remove(), 300); }, 4000);
  }

  function setBtnLoading(btn, isLoading, loadingText = '') {
    const textSpan = btn.querySelector('.btn-text');
    if (isLoading) { btn.dataset.originalText = textSpan.innerText; btn.classList.add('btn-loading'); btn.disabled = true; if (loadingText) textSpan.innerText = loadingText; }
    else { btn.classList.remove('btn-loading'); btn.disabled = false; textSpan.innerText = btn.dataset.originalText; }
  }

  function save(btn) {
    setBtnLoading(btn, true, 'Saving...');
    let d = {}; FIELDS.forEach(k => d[k] = document.getElementById(k).value);
    google.script.run.withSuccessHandler(r => { setBtnLoading(btn, false); showToast(r, 'success'); }).withFailureHandler(e => { setBtnLoading(btn, false); showToast(e.message, 'error'); }).uiSaveConfig(d);
  }

  function act(fn, btn, loadingText) {
    setBtnLoading(btn, true, loadingText);
    google.script.run.withSuccessHandler(r => { setBtnLoading(btn, false); showToast(r, 'success'); }).withFailureHandler(e => { setBtnLoading(btn, false); showToast(e.message, 'error'); })[fn]();
  }
</script>
</body>
</html>`;
}

/* ==================== SHEET / TAB / TZ HELPERS ==================== */
var _ss = null;
function getSS_() { if (!_ss) _ss = SpreadsheetApp.openById(SHEET_ID); return _ss; }
function tz_()    { return getSS_().getSpreadsheetTimeZone(); }

function findMonthSheet_(ss, date) {
  const year = date.getFullYear();
  const idx = date.getMonth();
  const short = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][idx];
  const long  = ['January','February','March','April','May','June','July','August','September','October','November','December'][idx];
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName();
    if (n.indexOf(String(year)) === -1) continue;
    const low = n.toLowerCase();
    if (low.indexOf(short.toLowerCase()) === 0 || low.indexOf(long.toLowerCase()) === 0) return sheets[i];
  }
  return null;
}
function getDataSheet_() {
  const ss = getSS_(), cfg = getConfig_();
  if (cfg.tabOverride) {
    const s = ss.getSheetByName(cfg.tabOverride);
    if (s) return s;
  }
  const s = findMonthSheet_(ss, new Date());
  if (!s) throw new Error('No tab found for the current month. Create one named like "Sep Month 2026", or set a Tab override in the panel.');
  return s;
}

/* ====================== BLOCK PARSING (per day) ==================== */
function getAllBlocks_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  const colB = sheet.getRange(1, 2, lastRow, 1).getValues();
  const blocks = [];
  let r = 0;
  while (r < lastRow) {
    if (String(colA[r][0]).trim() === 'Date') {
      const headerRow = r + 1;
      let dr = r + 1;
      while (dr < lastRow && String(colB[dr][0]).trim() !== '' && String(colA[dr][0]).trim() !== 'Date') dr++;
      const dataStart = headerRow + 1, dataEnd = dr;
      if (dataEnd >= dataStart) {
        const dCol = sheet.getRange(dataStart, DATE_COL, dataEnd - dataStart + 1, 1).getValues();
        let dateStr = '';
        for (let k = 0; k < dCol.length; k++) {
          const dv = dCol[k][0];
          if (dv instanceof Date) {
            dateStr = Utilities.formatDate(dv, tz_(), 'yyyy-MM-dd');
            break;
          }
          if (dv) {                       // text date bhi accept
            const parsed = new Date(dv);
            if (!isNaN(parsed.getTime())) {
              dateStr = Utilities.formatDate(parsed, tz_(), 'yyyy-MM-dd');
              break;
            }
          }
        }
        blocks.push({ headerRow: headerRow, dataStart: dataStart, dataEnd: dataEnd, dateStr: dateStr });
      }
      r = dr;
    } else {
      r++;
    }
  }
  return blocks;
}

function getTodayBlock_(sheet, cfg) {
  const logicalToday = getLogicalDate_(cfg);
  const todayStr = Utilities.formatDate(logicalToday, tz_(), 'yyyy-MM-dd');
  const blocks = getAllBlocks_(sheet);
  for (let i = 0; i < blocks.length; i++) if (blocks[i].dateStr === todayStr) return blocks[i];
  return null;
}
function ensureRowCount_(sheet, neededRows) {
  const have = sheet.getMaxRows();
  if (have < neededRows) sheet.insertRowsAfter(have, neededRows - have);
}

function ensureTodayBlock_(sheet, cfg) {
  let today = getTodayBlock_(sheet, cfg);
  if (today) return today;

  const blocks = getAllBlocks_(sheet);
  if (!blocks.length) throw new Error('No existing day-block found to clone from. Add at least one day manually first.');
  const last = blocks[blocks.length - 1];
  const empCount = last.dataEnd - last.dataStart + 1;

  const blankRow = sheet.getLastRow() + 1;
  const headerRow = blankRow + 1;
  const dataStart = headerRow + 1;
  const dataEnd = dataStart + empCount - 1;
  ensureRowCount_(sheet, dataEnd);

  const hdrRange = sheet.getRange(headerRow, 1, 1, HEADERS.length);
  hdrRange.setValues([HEADERS]).setBackground('#ffff00').setFontWeight('bold');

  const srcNames = sheet.getRange(last.dataStart, NAME_COL, empCount, 1).getValues();

  const logicalToday = getLogicalDate_(cfg);
  const dateVals = [];
  for (let i = 0; i < empCount; i++) dateVals.push([logicalToday]);

  sheet.getRange(dataStart, DATE_COL, empCount, 1).setValues(dateVals).setNumberFormat('d-mmm-yy');
  sheet.getRange(dataStart, NAME_COL, empCount, 1).setValues(srcNames);
  sheet.getRange(dataStart, STATUS_COL, empCount, 1).clearContent();

  // === KRA COLUMN (col D): pichle din se hu-ba-hu clone ===
  // copyTo() text + colour + font + alignment + MERGED cells sab ek saath copy
  // kar deta hai, isliye jis user ke saamne jo KRA/merge tha wahi naye din me aayega.
  sheet.getRange(last.dataStart, KRA_COL, empCount, 1)
       .copyTo(sheet.getRange(dataStart, KRA_COL, empCount, 1));

  sheet.getRange(dataStart, WORK_COL, empCount, 2).clearContent();

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Present', 'Absent', 'Holiday'], true).setAllowInvalid(false).build();
  sheet.getRange(dataStart, STATUS_COL, empCount, 1).setDataValidation(rule);

  applyStatusFormatting_(sheet);
  return { headerRow: headerRow, dataStart: dataStart, dataEnd: dataEnd, dateStr: Utilities.formatDate(logicalToday, tz_(), 'yyyy-MM-dd') };
}

/* ========================== FORMATTING ============================ */
function applyStatusFormatting_(sheet) {
  const range = sheet.getRange(2, STATUS_COL, Math.max(sheet.getMaxRows() - 1, 1), 1);
  const rules = sheet.getConditionalFormatRules().filter(function (r) {
    return r.getRanges().every(function (rg) { return rg.getColumn() !== STATUS_COL; });
  });
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Present')
      .setBackground(PRESENT_COLOR).setFontColor(PRESENT_TEXT).setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Absent')
      .setBackground(ABSENT_COLOR).setFontColor(ABSENT_TEXT).setBold(true).setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Holiday')
      .setBackground(HOLIDAY_COLOR).setFontColor(HOLIDAY_TEXT).setBold(true).setRanges([range]).build()
  );
  sheet.setConditionalFormatRules(rules);
}

/* ========================== PROTECTION ============================ */
function applyStatusLock_(sheet, cfg) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === STATUS_LOCK_TAG) p.remove();
  });
  const today = ensureTodayBlock_(sheet, cfg);
  const lockRows = today.headerRow - 2;
  if (lockRows >= 1) {
    const rng = sheet.getRange(2, STATUS_COL, lockRows, 1);
    const p = rng.protect().setDescription(STATUS_LOCK_TAG);
    p.setWarningOnly(false);
    try { p.addEditors(cfg.fullAccess); } catch (e) { Logger.log('addEditors: ' + e); }
    pruneEditors_(p, cfg.fullAccess);
    try { if (p.canDomainEdit()) p.setDomainEdit(false); } catch (e) {}
  }
}
function pruneEditors_(protection, allow) {
  let owner = '';
  try { const o = getSS_().getOwner(); owner = o ? o.getEmail() : ''; } catch (e) {}
  protection.getEditors().forEach(function (u) {
    const em = u.getEmail();
    if (em && em !== owner && allow.indexOf(em) === -1) {
      try { protection.removeEditors([em]); } catch (e) { }
    }
  });
}

function autoAbsentSweep_(sheet, cfg) {
  if (!cfg.autoAbsent) return;
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, tz_(), 'H'));
  const min = Number(Utilities.formatDate(now, tz_(), 'm'));

  const currentMins = getRelativeMinutes_(hour, min, cfg.shiftStartHour);
  const cutoffMins = getRelativeMinutes_(cfg.autoAbsentHour, 0, cfg.shiftStartHour);
  if (currentMins < cutoffMins) return;

  const logicalTodayStr = Utilities.formatDate(getLogicalDate_(cfg), tz_(), 'yyyy-MM-dd');
  const p = props_();
  if (p.getProperty('AUTO_ABSENT_DATE') === logicalTodayStr) return;

  // Weekend (Sat/Sun) = office holiday -> blank cell ko 'Holiday' mark karo;
  // weekday -> blank cell ko 'Absent'. Jo pehle se bhari hai (Present/Holiday/Absent)
  // use kabhi overwrite nahi karte -- SIRF blank Present Status cell fill hoti hai.
  const dow = Number(Utilities.formatDate(getLogicalDate_(cfg), tz_(), 'u')); // 1=Mon ... 6=Sat, 7=Sun
  const blankFill = (dow === 6 || dow === 7) ? 'Holiday' : 'Absent';

  const block = ensureTodayBlock_(sheet, cfg);
  const count = block.dataEnd - block.dataStart + 1;
  const rng = sheet.getRange(block.dataStart, STATUS_COL, count, 1);
  const vals = rng.getValues();
  let changed = false;
  for (let i = 0; i < vals.length; i++) {
    if (!vals[i][0]) { vals[i][0] = blankFill; changed = true; }
  }
  if (changed) rng.setValues(vals);
  p.setProperty('AUTO_ABSENT_DATE', logicalTodayStr);
}

/* ======================= CONTROLLER (10 min) ======================= */
function driver() {
  const sheet = getDataSheet_();
  const cfg = getConfig_();
  ensureTodayBlock_(sheet, cfg);
  applyStatusLock_(sheet, cfg);
  autoAbsentSweep_(sheet, cfg);
  maybeSendEmail_(sheet, cfg);
}

function maybeSendEmail_(sheet, cfg) {
  if (!cfg.autoEmail) return;
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, tz_(), 'H'));
  const min = Number(Utilities.formatDate(now, tz_(), 'm'));

  const currentMins = getRelativeMinutes_(hour, min, cfg.shiftStartHour);
  const emailMins = getRelativeMinutes_(cfg.emailHour, cfg.emailMin, cfg.shiftStartHour);
  if (currentMins < emailMins) return;

  const logicalTodayStr = Utilities.formatDate(getLogicalDate_(cfg), tz_(), 'yyyy-MM-dd');
  const p = props_();
  if (p.getProperty('EMAIL_SENT_DATE') === logicalTodayStr) return;

  sendStatusEmail_(sheet, cfg);
  p.setProperty('EMAIL_SENT_DATE', logicalTodayStr);
}

/* ============================= REPORT DATA ========================= */
function getTodayReportData_(sheet, cfg) {
  const block = ensureTodayBlock_(sheet, cfg);
  const count = block.dataEnd - block.dataStart + 1;
  const data = sheet.getRange(block.dataStart, 1, count, 6).getValues();

  const logicalDate = getLogicalDate_(cfg);
  const dateStr = Utilities.formatDate(logicalDate, tz_(), 'EEE, dd-MMM-yyyy');

  const present = [], absent = [], holiday = [];
  data.forEach(function (r) {
    const name = String(r[NAME_COL - 1]).trim();
    if (!name) return;
    const status = String(r[STATUS_COL - 1]).trim();
    const kra = String(r[KRA_COL - 1] || '').trim();
    const work = String(r[WORK_COL - 1] || '').trim();
    const add = String(r[ADD_COL - 1] || '').trim();
    if (status === 'Present') present.push({ name: name, kra: kra, work: work, add: add });
    else if (status === 'Holiday') holiday.push({ name: name, kra: kra });
    else absent.push({ name: name, status: status || 'Not marked yet' });
  });
  // expected = kaam par aane wale (Holiday walon ko chhod ke) -- attendance % isi par nikalta hai
  const expected = present.length + absent.length;
  const total = present.length + absent.length + holiday.length;
  return { dateStr: dateStr, present: present, absent: absent, holiday: holiday, total: total, expected: expected };
}

/* ============================= EMAIL =============================== */
function sendStatusEmail_(sheet, cfg) {
  if (!cfg) cfg = getConfig_();
  if (!cfg.recipients.length) { Logger.log('No recipients.'); return; }

  const r = getTodayReportData_(sheet, cfg);
  const html = buildReportEmailHtml_(r);
  const pct = r.expected ? Math.round((r.present.length / r.expected) * 100) : 0;
  let subject = 'Daily Work Status — ' + r.dateStr + '  (Present ' + r.present.length + '/' + r.expected + ', ' + pct + '%';
  if (r.holiday.length) subject += ', Holiday ' + r.holiday.length;
  subject += ')';
  MailApp.sendEmail({ to: cfg.recipients.join(','), subject: subject, htmlBody: html });
}

function buildReportEmailHtml_(r) {
  const F = 'font-family:Arial,Helvetica,sans-serif;';
  const pct = r.expected ? Math.round((r.present.length / r.expected) * 100) : 0;

  function statCard(label, value, color) {
    return '<td style="padding:6px;"><div style="' + F + 'background:#f8f9fa;border:1px solid #e8eaed;border-radius:10px;padding:14px 10px;text-align:center;">' +
      '<div style="font-size:26px;font-weight:bold;color:' + color + ';">' + value + '</div>' +
      '<div style="font-size:12px;color:#5f6368;margin-top:4px;">' + label + '</div></div></td>';
  }

  let presentHtml = '';
  r.present.forEach(function (p) {
    presentHtml += '<div style="' + F + 'border:1px solid #e8eaed;border-radius:12px;padding:14px 16px;margin-top:10px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;">' +
      '<div style="font-weight:bold;font-size:14px;color:#202124;">' + esc_(p.name) + '</div>' +
      '<span style="background:#e6f4ea;color:#137333;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:bold;">Present</span></div>' +
      (p.kra ? '<div style="font-size:12px;color:#5f6368;margin-top:4px;"><b>KRA:</b> ' + esc_(p.kra) + '</div>' : '') +
      (p.work ? '<div style="font-size:13px;color:#3c4043;margin-top:8px;line-height:1.5;">' + esc_(p.work) + '</div>' : '<div style="font-size:12px;color:#9aa0a6;margin-top:8px;">No work update added yet.</div>') +
      (p.add ? '<div style="font-size:12px;color:#5f6368;margin-top:6px;"><b>Additional:</b> ' + esc_(p.add) + '</div>' : '') +
      '</div>';
  });

  let absentHtml = '';
  if (r.absent.length) {
    let li = '';
    r.absent.forEach(function (a) {
      li += '<li style="margin:3px 0;">' + esc_(a.name) +
        (a.status === 'Not marked yet' ? ' <span style="color:#c5221f;font-size:11px;">(not marked)</span>' : '') + '</li>';
    });
    absentHtml = '<div style="' + F + 'margin-top:22px;background:#fce8e6;border-left:4px solid #ea4335;border-radius:6px;padding:14px 16px;">' +
      '<div style="font-weight:bold;color:#c5221f;font-size:14px;">Absent today (' + r.absent.length + ')</div>' +
      '<ul style="margin:8px 0 2px;padding-left:20px;color:#5f6368;font-size:13px;">' + li + '</ul></div>';
  }

  let holidayHtml = '';
  if (r.holiday.length) {
    let lih = '';
    r.holiday.forEach(function (h) {
      lih += '<li style="margin:3px 0;">' + esc_(h.name) +
        (h.kra ? ' <span style="color:#9aa0a6;font-size:11px;">(' + esc_(h.kra) + ')</span>' : '') + '</li>';
    });
    holidayHtml = '<div style="' + F + 'margin-top:22px;background:#e8f0fe;border-left:4px solid #4285f4;border-radius:6px;padding:14px 16px;">' +
      '<div style="font-weight:bold;color:#1a56c4;font-size:14px;">On holiday today (' + r.holiday.length + ')</div>' +
      '<ul style="margin:8px 0 2px;padding-left:20px;color:#5f6368;font-size:13px;">' + lih + '</ul></div>';
  }

  return '<div style="' + F + 'background:#eef1f4;padding:24px 0;"><div style="max-width:640px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden;border:1px solid #e3e6ea;">' +
    '<div style="background:#4f46e5;padding:22px 24px;"><div style="' + F + 'color:#fff;font-size:20px;font-weight:bold;">📋 Daily Work Status Report</div>' +
    '<div style="' + F + 'color:#c7d2fe;font-size:13px;margin-top:4px;">' + esc_(r.dateStr) + '</div></div>' +
    '<div style="padding:20px 24px 4px;"><table width="100%" cellspacing="0" cellpadding="0"><tr>' +
      statCard('Total staff', r.total, '#202124') +
      statCard('Present', r.present.length, '#137333') +
      statCard('Absent', r.absent.length, '#c5221f') +
      statCard('Holiday', r.holiday.length, '#1a56c4') +
      statCard('Attendance', pct + '%', '#4f46e5') +
    '</tr></table></div>' +
    '<div style="padding:8px 24px 24px;">' +
      '<div style="' + F + 'font-size:13px;font-weight:bold;color:#5f6368;text-transform:uppercase;letter-spacing:.4px;margin-top:10px;">Today\'s work — Present staff</div>' +
      (presentHtml || '<div style="color:#9aa0a6;font-size:13px;margin-top:8px;">No one marked present yet.</div>') +
      absentHtml +
      holidayHtml +
    '</div>' +
    '<div style="padding:20px 24px;text-align:center;border-top:1px solid #f1f3f4;">' +
      '<a href="' + getSS_().getUrl() + '" style="' + F + 'display:inline-block;background:#4f46e5;color:#fff;text-decoration:none;padding:11px 24px;border-radius:8px;font-size:14px;font-weight:bold;">Open the sheet</a>' +
      '<div style="' + F + 'color:#9aa0a6;font-size:11px;margin-top:14px;">Automated report • Work Status Automation</div></div></div></div>';
}
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* ================= ONE-TIME FIX: aaj ka KRA merge repair ================
 * Aaj (jis din yeh chalao) ka block agar pehle hi bina merge ke ban chuka hai,
 * to yeh function pichle din ka KRA (text + colour + font + alignment + MERGES)
 * dobara aaj ke block par copy kar dega.
 * Editor me function dropdown se "fixTodayKra" select karke Run karo — bas ek baar.
 * Sirf KRA column (col D) ko chhoota hai; Status/Working/Date kuch nahi badalta.
 * ==================================================================== */
function fixTodayKra() {
  const sheet = getDataSheet_(), cfg = getConfig_();
  const blocks = getAllBlocks_(sheet);
  const todayStr = Utilities.formatDate(getLogicalDate_(cfg), tz_(), 'yyyy-MM-dd');

  let ti = -1;
  for (let i = 0; i < blocks.length; i++) if (blocks[i].dateStr === todayStr) ti = i;
  if (ti <= 0) throw new Error('Today\'s block or the previous block was not found.');

  const today = blocks[ti], prev = blocks[ti - 1];
  const count = Math.min(today.dataEnd - today.dataStart + 1,
                         prev.dataEnd  - prev.dataStart  + 1);

  // aaj ke KRA me jo bhi (galat) merge hai use pehle todo, phir clone karo
  sheet.getRange(today.dataStart, KRA_COL, count, 1).breakApart();
  sheet.getRange(prev.dataStart, KRA_COL, count, 1)
       .copyTo(sheet.getRange(today.dataStart, KRA_COL, count, 1));

  return 'Today\'s formatting (merge + color + alignment) was copied over from the previous day.';
}

/* ============ ONE-TIME: aaj ke block me Holiday option + colour ============
 * Naye code ke baad: naye din ke block me Holiday dropdown + colour apne aap
 * aate hain. Lekin AAJ ka block agar pehle hi ban chuka hai to usme purana
 * 2-option dropdown hai. Yeh function ek baar run karo to aaj ke block me bhi
 * Present/Absent/Holiday dropdown aur Holiday colour turant lag jaayega.
 * Sirf Present Status column (col C) ko chhoota hai; data kuch nahi badalta.
 * ==================================================================== */
function enableHolidayForToday() {
  const sheet = getDataSheet_(), cfg = getConfig_();
  const block = getTodayBlock_(sheet, cfg);
  if (!block) throw new Error('Today\'s block is not available. Play "Prep Today\'s Rouse" first.');
  const count = block.dataEnd - block.dataStart + 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Present', 'Absent', 'Holiday'], true).setAllowInvalid(false).build();
  sheet.getRange(block.dataStart, STATUS_COL, count, 1).setDataValidation(rule);
  applyStatusFormatting_(sheet);
  return 'Holiday options + colors ready in today\'s batch.';
}
