/**
 * KunalMIS_Panel.gs  — v5
 * Pulls data directly from panel.parcelx.in (no Metabase)
 *
 * HOW TO GET COOKIES (do this after every login):
 *  1. Log in to panel.parcelx.in in Chrome (complete OTP)
 *  2. Open DevTools (F12) → Network tab
 *  3. Go to panel.parcelx.in/mis_report
 *  4. Click ANY request to panel.parcelx.in → Headers → copy the full "Cookie:" value
 *  5. Paste into Script Properties → PARCELX_COOKIES
 *
 * SETUP:
 *  1. Enable Drive API: Services (+) → Drive API → Add
 *  2. Set Script Properties: PARCELX_COOKIES (from step above)
 *  3. Run testConnection() to verify
 *  4. Run testFindExportUrl() to discover export endpoint
 *  5. Run setupTrigger() once for daily 9 AM IST
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────
const CFG = {
  sheetId   : '1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE',
  dataTab   : 'MIS Data',
  pivotTab  : 'Pivot',
  emailTo   : 'kunal.chauhan@parcelx.in',
  emailCc   : 'sudhanshu@parcelx.in,sparsh@parcelx.in',
  alertEmail: 'sudhanshu@parcelx.in',   // gets session-expired alerts
  panelBase : 'https://panel.parcelx.in',
  basicUser : 'pax',
  basicPass : 'cloud@4w5',
  salesPoc  : '26',
  daysBack  : 40
};

const COLS_WANTED = [
  'Placed Date (D-M-Y)', 'Placed Time', 'ParcelX Order ID',
  'Pickup Date (D-M-Y)', 'Pickup Time', 'Delivered Date',
  'Current User Status', 'Parcelx Tracking ID', 'Pre Generated WayBill',
  'RTO Waybill', 'Client Order ID / Invoice no', 'Channel Order Number',
  'User Email', 'RTS Date', 'RTO Marked Date',
  'First Attempt Date', 'Latest Attempt Date', 'Sales POC',
  'Item Name', 'Item Qty',
  'Item Length', 'Item Height', 'Item Width', 'Item Weight',
  'CN Weight', 'DN Weight', 'Charged Weight', 'Updated Actual Weight',
  'Invoice Value', 'COD', 'Order Type (Air / Surface)',
  'Courier Used', 'Courier Account', 'Warehouse Name',
  'Pickup Name', 'Pickup City', 'Pickup State',
  'RTO Reason', 'Attempt Counts',
  'NDR First Date', 'NDR First Remarks', 'NDR Last Date', 'NDR Last Remarks'
];

const STATUSES   = ['Booked','Manifested','Not Picked','Out For Pickup','Pickup Pending'];
const STATUS_CLR = {
  'Booked':'#3498db','Manifested':'#9b59b6','Not Picked':'#e74c3c',
  'Out For Pickup':'#e67e22','Pickup Pending':'#f1c40f'
};


// ── MAIN ──────────────────────────────────────────────────────────────────────
function runMIS() {
  Logger.log('▶ Kunal MIS starting...');
  try {
    const cookies = getStoredCookies_();
    const headers = buildHeaders_(cookies);
    const { fromStr, toStr } = getDateRange_();
    Logger.log('Range: ' + fromStr + ' → ' + toStr);

    triggerSearch_(headers, fromStr, toStr);

    const blob = downloadExcel_(headers, fromStr, toStr);
    const { allHeaders, allRows } = parseData_(blob);
    Logger.log('Export cols: ' + allHeaders.length + ', rows: ' + allRows.length);

    const { colIdx, finalHeaders } = mapColumns_(allHeaders);
    const finalRows = allRows.map(r =>
      finalHeaders.map((_, i) => colIdx[i] !== -1 ? r[colIdx[i]] : '')
    );

    writeSheet_(finalHeaders, finalRows);
    writePivot_(finalHeaders, finalRows);
    sendEmail_(finalHeaders, finalRows);
    Logger.log('✅ Done — ' + finalRows.length + ' rows');

  } catch (err) {
    Logger.log('❌ ' + err);
    handleError_(err);
    throw err;
  }
}


// ── ERROR HANDLER — sends alert email on session expiry ───────────────────────
function handleError_(err) {
  const msg = err.toString();
  const isSession = msg.includes('SESSION_EXPIRED') || msg.includes('OTP_REQUIRED');

  const subject = isSession
    ? '⚠️ ParcelX MIS — Session Expired (action needed)'
    : '❌ ParcelX MIS — Script Error';

  const body = isSession ? `
<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto">
  <div style="background:#e74c3c;padding:20px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff">⚠️ MIS Session Expired</h2>
  </div>
  <div style="background:#fff;padding:20px 24px;border:1px solid #ddd;border-top:none">
    <p>The ParcelX panel session has expired. The MIS report could not run.</p>
    <p><strong>To fix — takes 2 minutes:</strong></p>
    <ol>
      <li>Open Chrome → go to <a href="https://panel.parcelx.in">panel.parcelx.in</a></li>
      <li>Log in (complete OTP as usual)</li>
      <li>Open DevTools → <strong>Network</strong> tab</li>
      <li>Navigate to <strong>panel.parcelx.in/mis_report</strong></li>
      <li>Click any request → Headers → copy the full <strong>Cookie:</strong> value</li>
      <li>Open <a href="https://script.google.com">script.google.com</a> → your project → ⚙️ Project Settings → Script Properties</li>
      <li>Update <strong>PARCELX_COOKIES</strong> with the copied value</li>
      <li>Run <strong>runMIS()</strong> manually to confirm it works</li>
    </ol>
    <p style="font-size:12px;color:#999">Sessions typically last 7–14 days.</p>
  </div>
</div>` : `<pre>${msg}</pre>`;

  GmailApp.sendEmail(
    CFG.alertEmail,
    subject,
    isSession ? 'ParcelX session expired. Please refresh cookies.' : msg,
    { htmlBody: body, name: 'ParcelX MIS Alert' }
  );
  Logger.log('Alert email sent to ' + CFG.alertEmail);
}


// ── AUTH ──────────────────────────────────────────────────────────────────────
function getStoredCookies_() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty('PARCELX_COOKIES');
  if (!raw) throw new Error(
    'PARCELX_COOKIES not set in Script Properties.\n' +
    'See script header for instructions on how to get the cookie string.'
  );
  // Accept raw string OR JSON array
  try {
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) return arr.map(c => c.name + '=' + c.value).join('; ');
  } catch (e) {}
  return raw.trim();
}

function buildHeaders_(cookies) {
  return {
    'Authorization'   : 'Basic ' + Utilities.base64Encode(CFG.basicUser + ':' + CFG.basicPass),
    'Cookie'          : cookies,
    'X-Requested-With': 'XMLHttpRequest',
    'Accept'          : 'text/html,application/json,*/*',
    'User-Agent'      : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
  };
}

function isSessionExpired_(body) {
  const b = body.toLowerCase();
  if (b.includes('enter otp') || b.includes('input the otp'))
    throw new Error('OTP_REQUIRED: Session expired. Please log in and update PARCELX_COOKIES.');
  if ((b.includes('login') || b.includes('sign in')) && b.includes('password') && !b.includes('mis_report'))
    throw new Error('SESSION_EXPIRED: Please log in again and update PARCELX_COOKIES in Script Properties.');
}


// ── DATES ─────────────────────────────────────────────────────────────────────
function getDateRange_() {
  const to   = new Date();
  const from = new Date();
  from.setDate(from.getDate() - CFG.daysBack);
  return {
    fromStr: Utilities.formatDate(from, 'Asia/Kolkata', 'dd-MM-yyyy'),
    toStr  : Utilities.formatDate(to,   'Asia/Kolkata', 'dd-MM-yyyy')
  };
}


// ── SEARCH ────────────────────────────────────────────────────────────────────
function triggerSearch_(headers, fromStr, toStr) {
  const payload =
    'from_date='  + encodeURIComponent(fromStr) +
    '&to_date='   + encodeURIComponent(toStr)   +
    '&sales_poc=' + CFG.salesPoc                +
    '&box8=1&searchbtn=1';

  Logger.log('Triggering search...');
  const resp = UrlFetchApp.fetch(CFG.panelBase + '/mis_report', {
    method            : 'post',
    headers           : Object.assign({}, headers, { 'Content-Type': 'application/x-www-form-urlencoded' }),
    payload           : payload,
    muteHttpExceptions: true,
    followRedirects   : true
  });

  Logger.log('Search → ' + resp.getResponseCode());
  isSessionExpired_(resp.getContentText());
  Utilities.sleep(2500);
}


// ── DOWNLOAD ─────────────────────────────────────────────────────────────────
function downloadExcel_(headers, fromStr, toStr) {
  const props      = PropertiesService.getScriptProperties();
  const overrideUrl = props.getProperty('EXPORT_URL_OVERRIDE');

  const candidates = overrideUrl ? [overrideUrl] : [
    CFG.panelBase + '/mis_report/export_excel',
    CFG.panelBase + '/mis_report/exportexcel',
    CFG.panelBase + '/mis_report/export',
    CFG.panelBase + '/mis_report/download',
    CFG.panelBase + '/mis_report/excel',
    CFG.panelBase + '/mis_report/export_csv',
    CFG.panelBase + '/mis_report/csv',
    CFG.panelBase + '/export/mis',
    CFG.panelBase + '/mis_report?export=excel&from_date=' + fromStr + '&to_date=' + toStr + '&sales_poc=' + CFG.salesPoc + '&box8=1',
    CFG.panelBase + '/mis_report?type=excel&from_date=' + fromStr + '&to_date=' + toStr + '&sales_poc=' + CFG.salesPoc,
    CFG.panelBase + '/mis_report?export=csv&from_date='  + fromStr + '&to_date=' + toStr + '&sales_poc=' + CFG.salesPoc
  ];

  for (const url of candidates) {
    try {
      Logger.log('Trying: ' + url);
      const resp = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true, followRedirects: true });
      const code = resp.getResponseCode();
      const ct   = (resp.getHeaders()['Content-Type'] || resp.getHeaders()['content-type'] || '').toLowerCase();
      Logger.log('  ' + code + ' | ' + ct.substring(0, 70));

      if (code === 200 && isFileResponse_(ct)) {
        Logger.log('✅ Got file from: ' + url);
        props.setProperty('EXPORT_URL_OVERRIDE', url); // remember it
        return resp.getBlob().setContentType(ct.split(';')[0].trim());
      }
    } catch (e) { Logger.log('  error: ' + e); }
  }

  // POST fallback
  const postPayload = 'from_date=' + encodeURIComponent(fromStr) +
    '&to_date=' + encodeURIComponent(toStr) +
    '&sales_poc=' + CFG.salesPoc + '&box8=1&export=1';

  for (const url of [CFG.panelBase + '/mis_report', CFG.panelBase + '/mis_report/export']) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: Object.assign({}, headers, { 'Content-Type': 'application/x-www-form-urlencoded' }),
        payload: postPayload, muteHttpExceptions: true
      });
      const ct = (resp.getHeaders()['Content-Type'] || '').toLowerCase();
      if (resp.getResponseCode() === 200 && isFileResponse_(ct)) {
        Logger.log('✅ POST export worked: ' + url);
        return resp.getBlob().setContentType(ct.split(';')[0].trim());
      }
    } catch (e) { Logger.log('  POST error: ' + e); }
  }

  throw new Error(
    'Could not download data file.\n' +
    'Run testFindExportUrl() — it will log which URL returns the Excel/CSV.\n' +
    'Then call setExportUrl("the_url") to save it.'
  );
}

function isFileResponse_(ct) {
  return ct.includes('spreadsheetml') || ct.includes('excel') ||
         ct.includes('octet-stream')  || ct.includes('openxmlformats') ||
         ct.includes('ms-excel')      || ct.includes('csv');
}


// ── PARSE ─────────────────────────────────────────────────────────────────────
function parseData_(blob) {
  const ct = (blob.getContentType() || '').toLowerCase();

  if (ct.includes('csv') || ct.includes('text')) {
    const rows = Utilities.parseCsv(blob.getDataAsString('UTF-8'));
    return { allHeaders: rows[0] || [], allRows: rows.slice(1) };
  }

  // Excel → convert via Drive Advanced Service
  const tempName = 'kunal_mis_' + Date.now();
  const xlsBlob  = blob.setName(tempName + '.xlsx')
    .setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  let fileId = null;
  try {
    const f = Drive.Files.insert(
      { title: tempName, mimeType: 'application/vnd.google-apps.spreadsheet' },
      xlsBlob, { convert: true }
    );
    fileId = f.id;
    const raw = SpreadsheetApp.openById(fileId).getSheets()[0].getDataRange().getValues();
    return { allHeaders: (raw[0] || []).map(String), allRows: raw.slice(1) };
  } finally {
    if (fileId) try { Drive.Files.remove(fileId); } catch(e) {}
  }
}


// ── COLUMN MAP ────────────────────────────────────────────────────────────────
function mapColumns_(exportHeaders) {
  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const ne   = exportHeaders.map(norm);
  const colIdx = COLS_WANTED.map(w => {
    const nw = norm(w);
    let idx = ne.indexOf(nw);
    if (idx === -1) idx = ne.findIndex(h => h.includes(nw) || nw.includes(h));
    return idx;
  });
  const missed = COLS_WANTED.filter((_, i) => colIdx[i] === -1);
  if (missed.length) Logger.log('Blank cols (not found in export): ' + missed.join(', '));
  return { colIdx, finalHeaders: COLS_WANTED.slice() };
}


// ── SHEET ─────────────────────────────────────────────────────────────────────
function writeSheet_(headers, rows) {
  const ss    = SpreadsheetApp.openById(CFG.sheetId);
  let   sheet = ss.getSheetByName(CFG.dataTab) || ss.insertSheet(CFG.dataTab, 0);
  sheet.clearContents();
  const now = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  sheet.getRange(1,1).setValue('Kunal Chauhan MIS — Updated: ' + now);
  sheet.getRange(1,1,1,headers.length).merge().setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(2,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8ecf0');
  if (rows.length) sheet.getRange(3,1,rows.length,headers.length).setValues(rows);
  sheet.setFrozenRows(2);
  headers.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  Logger.log('✅ Sheet: ' + rows.length + ' rows');
}


// ── PIVOT ─────────────────────────────────────────────────────────────────────
function writePivot_(headers, rows) {
  const ss    = SpreadsheetApp.openById(CFG.sheetId);
  let   sheet = ss.getSheetByName(CFG.pivotTab) || ss.insertSheet(CFG.pivotTab);
  sheet.clearContents();
  const now  = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const sIdx = headers.indexOf('Current User Status');

  function section(col, label) {
    const gIdx = headers.indexOf(col);
    if (gIdx === -1) return [];
    const grp = {};
    rows.forEach(r => {
      const k = (r[gIdx]||'(blank)').toString().substring(0,40);
      const s = r[sIdx]||'';
      if (!grp[k]) grp[k] = {};
      grp[k][s] = (grp[k][s]||0)+1;
    });
    const out = [['── By '+label+' ──'],[label,...STATUSES,'Total']];
    Object.entries(grp)
      .map(([k,v])=>[k,v,STATUSES.reduce((s,x)=>s+(v[x]||0),0)])
      .sort((a,b)=>b[2]-a[2])
      .forEach(([k,v,tot])=>out.push([k,...STATUSES.map(s=>v[s]||0),tot]));
    out.push(['TOTAL',...STATUSES.map(s=>rows.filter(r=>r[sIdx]===s).length),rows.length]);
    out.push(['']);
    return out;
  }

  let out = [['Kunal Chauhan – Pivot | '+now],['']];
  out = out.concat(section('User Email','Seller'));
  out = out.concat(section('Order Type (Air / Surface)','Courier Type'));
  out = out.concat(section('Placed Date (D-M-Y)','Placed Date'));
  out = out.concat(section('Courier Used','Courier'));
  out = out.concat(section('Pickup City','Pickup City'));

  const maxC = Math.max(...out.map(r=>r.length||1));
  sheet.getRange(1,1,out.length,maxC).setValues(out);
  sheet.getRange(1,1,1,maxC).merge().setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
}


// ── EMAIL ─────────────────────────────────────────────────────────────────────
function sendEmail_(headers, rows) {
  const now   = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const sIdx  = headers.indexOf('Current User Status');
  const counts = {};
  STATUSES.forEach(s => { counts[s] = rows.filter(r => r[sIdx]===s).length; });

  const statusRows = STATUSES.map(s => `
    <tr>
      <td style="padding:9px 16px;border-bottom:1px solid #f0f0f0;font-size:14px">${s}</td>
      <td style="padding:9px 16px;border-bottom:1px solid #f0f0f0;text-align:center">
        <span style="background:${STATUS_CLR[s]};color:#fff;padding:3px 14px;border-radius:12px;font-weight:700">${counts[s]}</span>
      </td></tr>`).join('');

  const html = `
<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto">
  <div style="background:#1a1a2e;padding:24px 28px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff">📦 ParcelX MIS Report</h2>
    <p style="margin:6px 0 0;color:#aaa;font-size:13px">Kunal Chauhan &nbsp;·&nbsp; ${now} &nbsp;·&nbsp; Last ${CFG.daysBack} Days</p>
  </div>
  <div style="background:#fff;padding:22px 28px;border:1px solid #ddd;border-top:none">
    <p><strong>Total orders:</strong> ${rows.length}</p>
    <table style="width:100%;border-collapse:collapse">
      <tr style="background:#f0f2f5">
        <th style="padding:10px 16px;text-align:left">Status</th>
        <th style="padding:10px 16px;text-align:center">Count</th>
      </tr>${statusRows}
    </table>
    <div style="text-align:center;margin:20px 0 8px">
      <a href="https://docs.google.com/spreadsheets/d/${CFG.sheetId}"
         style="background:#34a853;color:#fff;padding:11px 28px;border-radius:6px;text-decoration:none;font-weight:700">
        📊 Open Google Sheet
      </a>
    </div>
    <p style="font-size:12px;color:#999;text-align:center">CSV attached — ${rows.length} rows, ${headers.length} columns</p>
  </div>
</div>`;

  const esc = v => { const s=String(v==null?'':v); return (s.includes(',')||s.includes('"')||s.includes('\n'))?'"'+s.replace(/"/g,'""')+'"':s; };
  const csv = [COLS_WANTED.map(esc).join(','), ...rows.map(r=>r.map(esc).join(','))].join('\n');
  const dateTag = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'ddMMMyyyy');

  GmailApp.sendEmail(CFG.emailTo,
    'ParcelX MIS | Kunal Chauhan | ' + now,
    rows.length + ' orders. CSV attached.',
    { htmlBody: html, cc: CFG.emailCc, name: 'ParcelX MIS',
      attachments: [Utilities.newBlob(csv, 'text/csv', 'Kunal_MIS_'+dateTag+'.csv')] }
  );
  Logger.log('✅ Email → ' + CFG.emailTo);
}


// ── TRIGGER ───────────────────────────────────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMIS')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runMIS').timeBased().everyDays(1).atHour(9).inTimezone('Asia/Kolkata').create();
  Logger.log('✅ Daily 9 AM IST trigger set');
}


// ── DEBUG ─────────────────────────────────────────────────────────────────────
/** Step 1: Run this first to check if cookies work */
function testConnection() {
  const cookies = getStoredCookies_();
  const headers = buildHeaders_(cookies);
  const resp = UrlFetchApp.fetch(CFG.panelBase + '/mis_report', { headers, muteHttpExceptions: true });
  const body = resp.getContentText();
  Logger.log('Response code: ' + resp.getResponseCode());
  Logger.log('Page has mis_report content: ' + body.includes('mis_report'));
  Logger.log('Session expired?: ' + (body.toLowerCase().includes('login') && !body.includes('mis_report')));
  Logger.log('OTP page?: ' + body.toLowerCase().includes('enter otp'));
  Logger.log('First 800 chars:\n' + body.substring(0, 800));
}

/** Step 2: Run after testConnection() passes — finds the export URL */
function testFindExportUrl() {
  const cookies = getStoredCookies_();
  const headers = buildHeaders_(cookies);
  const { fromStr, toStr } = getDateRange_();
  triggerSearch_(headers, fromStr, toStr);

  const paths = [
    '/mis_report/export_excel', '/mis_report/exportexcel', '/mis_report/export',
    '/mis_report/download', '/mis_report/excel', '/mis_report/export_csv', '/mis_report/csv',
    '/export/mis', '/mis_report?export=excel&from_date='+fromStr+'&to_date='+toStr+'&sales_poc='+CFG.salesPoc+'&box8=1',
    '/mis_report?export=csv&from_date='+fromStr+'&to_date='+toStr+'&sales_poc='+CFG.salesPoc
  ];

  paths.forEach(path => {
    try {
      const r = UrlFetchApp.fetch(CFG.panelBase + path, { headers, muteHttpExceptions: true });
      const ct = r.getHeaders()['Content-Type'] || '';
      const size = r.getBlob().getBytes().length;
      Logger.log(r.getResponseCode() + ' | ' + size + ' bytes | ' + ct.substring(0,60) + '  →  ' + path);
    } catch(e) { Logger.log('ERR | ' + path + ' | ' + e); }
  });
  Logger.log('\nLook for a 200 response with excel/spreadsheet content type and size > 1000 bytes');
}

/** After finding the URL from testFindExportUrl(), save it here */
function setExportUrl(url) {
  PropertiesService.getScriptProperties().setProperty('EXPORT_URL_OVERRIDE', url);
  Logger.log('✅ Export URL saved: ' + url);
}
