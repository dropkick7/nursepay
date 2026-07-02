// ─────────────────────────────────────────────────────────────────────────────
// NursePay — Push Approved Jobs to GitHub
// Add this below the existing combine pipeline code.
// After running runCombinePipeline() and checking off rows in the Review tab,
// run pushApprovedJobsToGitHub() to push to jobs.json and clear the Review tab.
// ─────────────────────────────────────────────────────────────────────────────

const GITHUB_TOKEN = "YOUR_GITHUB_TOKEN_HERE"; // paste your existing token
const GITHUB_OWNER = "dropkick7";
const GITHUB_REPO  = "nursepay";
const GITHUB_FILE  = "jobs.json";

function pushApprovedJobsToGitHub() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REVIEW_TAB); // reuses REVIEW_TAB from combine pipeline
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];

  // Build column index map
  const col = {};
  headers.forEach((h, i) => col[h.trim()] = i);

  // Filter approved rows (checkbox = true)
  const approved = data.slice(1).filter(row => row[col['Approve']] === true);

  if (approved.length === 0) {
    SpreadsheetApp.getUi().alert('No approved jobs found. Check off some rows first.');
    return;
  }

  // Map to jobs.json schema
  const jobs = approved.map(row => ({
    hospital:  row[col['Hospital System']] || '',
    role:      row[col['Role Type']]       || '',
    title:     row[col['Job Title']]       || '',
    city:      row[col['Location']]        || '',
    shift:     row[col['Shift Type']]      || '',
    length:    String(row[col['Shift Length']] || ''),
    workStyle: row[col['Work Style']]      || '',
    pay:       row[col['Pay Range']]       || '',
    url:       row[col['Job Link']]        || '',
    pulled:    row[col['Pulled Date']]     || '',
    rateKey:   getRateKey(row[col['Hospital System']], row[col['Role Type']], row[col['Job Title']])
  }));

  // Get current file SHA from GitHub (required for update)
  const apiUrl  = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_FILE}`;
  const getResp = UrlFetchApp.fetch(apiUrl, {
    headers: {
      'Authorization': `token ${GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json'
    },
    muteHttpExceptions: true
  });

  let sha = null;
  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  }

  // Push to GitHub
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const content = Utilities.base64Encode(JSON.stringify(jobs, null, 2));
  const payload = { message: `Update per diem jobs ${today}`, content, ...(sha ? { sha } : {}) };

  const putResp = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: {
      'Authorization': `token ${GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = putResp.getResponseCode();
  if (status === 200 || status === 201) {
    // Clear Review tab data rows, keep header
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      sheet.getRange(2, 1, lastRow - 1, 1).insertCheckboxes();
    }
    SpreadsheetApp.getUi().alert(`✅ Done! ${jobs.length} jobs pushed to jobs.json. Review tab cleared and ready for next week.`);
  } else {
    SpreadsheetApp.getUi().alert(`❌ GitHub push failed (${status}).\n\n${putResp.getContentText()}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Map hospital + role to contract rate key used by per-diem.html
// Handles APRN sub-roles by checking the job title
// ─────────────────────────────────────────────────────────────────────────────
function getRateKey(hospital, role, title) {
  const h = (hospital || '').trim();
  const r = (role    || '').trim().toUpperCase();
  const t = (title   || '').toLowerCase();

  if (h === 'UCSF') {
    if (r === 'RN') return 'RN';
    if (r === 'APRN') {
      if (t.includes('senior') && t.includes('anesthetist')) return 'CRNASr';
      if (t.includes('anesthetist') || t.includes('crna'))   return 'CRNA';
      if (t.includes('practitioner') || t.includes('np'))    return 'NP';
      return null;
    }
    return null;
  }
  if (h === 'Sutter') {
    if (r === 'RN') return 'RN';
    return null;
  }
  if (h === 'Kaiser') {
    if (r === 'RN')   return 'StaffNurseII';
    if (r === 'APRN') return 'NPII';
    return null;
  }
  return null;
}

// ─────────────────────────────────────────────────────────────────────────────
// MENU — replace your existing onOpen() with this one, or just add the
// second menu item to whichever onOpen() you already have.
// ─────────────────────────────────────────────────────────────────────────────
// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu('NursePay')
//     .addItem('1. Run Combine Pipeline', 'runCombinePipeline')
//     .addItem('2. Push Approved Jobs to GitHub', 'pushApprovedJobsToGitHub')
//     .addToUi();
// }
