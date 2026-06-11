/**
 * KunalMIS_Panel.gs  — v6
 * Pulls data from panel.parcelx.in (no Metabase)
 * Export URL confirmed: /mis_report?export=excel&...  (returns HTML table @ 1.1MB)
 *
 * HOW TO REFRESH COOKIES (every ~7 days):
 *  1. Log in to panel.parcelx.in in Chrome (complete OTP)
 *  2. F12 → Network tab → navigate to panel.parcelx.in/mis_report
 *  3. Click the mis_report request → Request Headers → copy "Cookie:" value
 *  4. Script Properties → PARCELX_COOKIES → paste
 */

const CFG = {
  sheetId   : '1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE',
  dataTab   : 'MIS Data',
  pivotTab  : 'Pivot',
  emailTo   : 'kunal.chauhan@parcelx.in',
  emailCc   : 'sudhanshu@parcelx.in,sparsh@parcelx.in',
  alertEmail: 'sudhanshu@parcelx.in',
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

const STATUSES  = ['Booked','Manifested','Not Picked','Out For Pickup','Pickup Pending'];
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

    // 1. Trigger search (loads data server-side)
    triggerSearch_(headers, fromStr, toStr);

    // 2. Download export (HTML table @ ~1.1MB)
    const html = downloadExport_(headers, fromStr, toStr);

    // 3. Parse HTML table → headers + rows
    const { exportHeaders, exportRows } = parseHtmlTable_(html);
    Logger.log('Export cols: ' + exportHeaders.length + ', rows: ' + exportRows.length);

    if (exportRows.length === 0)
      throw new Error('No data rows found in export. Check if search returned results.');

    // 4. Filter to 43 wanted columns
    const { colIdx } = mapColumns_(exportHeaders);
    const finalRows = exportRows.map(r =>
      COLS_WANTED.map((_, i) => colIdx[i] !== -1 ? (r[colIdx[i]] || '') : '')
    );

    // 5. Write → Sheet, Pivot, Email
    writeSheet_(COLS_WANTED, finalRows);
    writePivot_(COLS_WANTED, finalRows);
    sendEmail_(COLS_WANTED, finalRows);
    Logger.log('✅ Done — ' + finalRows.length + ' rows');

  } catch (err) {
    Logger.log('❌ ' + err);
    handleError_(err);
    throw err;
  }
}


// ── AUTH ──────────────────────────────────────────────────────────────────────
function getStoredCookies_() {
  const raw = PropertiesService.getScriptProperties().getProperty('PARCELX_COOKIES');
  if (!raw || raw.length < 10)
    throw new Error('PARCELX_COOKIES not set. See script header for instructions.');
  try {
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) return arr.map(c => c.name + '=' + c.value).join('; ');
  } catch(e) {}
  return raw.trim();
}

function buildHeaders_(cookies) {
  return {
    'Authorization'   : 'Basic ' + Utilities.base64Encode(CFG.basicUser + ':' + CFG.basicPass),
    'Cookie'          : cookies,
    'X-Requested-With': 'XMLHttpRequest',
    'Accept'          : '*/*',
    'User-Agent'      : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
  };
}

function checkSession_(body) {
  const b = body.substring(0, 3000).toLowerCase();
  if (b.includes('enter otp') || b.includes('input the otp'))
    throw new Error('OTP_REQUIRED: Session expired. Log in again and update PARCELX_COOKIES.');
  if ((b.includes('password') || b.includes('sign in')) && !b.includes('mis_report') && !b.includes('searchbtn'))
    throw new Error('SESSION_EXPIRED: Session expired. Log in again and update PARCELX_COOKIES.');
}


// ── DATES ─────────────────────────────────────────────────────────────────────
function getDateRange_() {
  const to = new Date(), from = new Date();
  from.setDate(from.getDate() - CFG.daysBack);
  return {
    fromStr: Utilities.formatDate(from, 'Asia/Kolkata', 'dd-MM-yyyy'),
    toStr  : Utilities.formatDate(to,   'Asia/Kolkata', 'dd-MM-yyyy')
  };
}


// ── SEARCH ────────────────────────────────────────────────────────────────────
function triggerSearch_(headers, fromStr, toStr) {
  const payload = 'from_date=' + encodeURIComponent(fromStr) +
    '&to_date=' + encodeURIComponent(toStr) +
    '&sales_poc=' + CFG.salesPoc + '&box8=1&searchbtn=1';

  const resp = UrlFetchApp.fetch(CFG.panelBase + '/mis_report', {
    method: 'post',
    headers: Object.assign({}, headers, { 'Content-Type': 'application/x-www-form-urlencoded' }),
    payload, muteHttpExceptions: true, followRedirects: true
  });
  Logger.log('Search → ' + resp.getResponseCode());
  checkSession_(resp.getContentText());
  Utilities.sleep(2000);
}


// ── DOWNLOAD (confirmed URL) ───────────────────────────────────────────────────
function downloadExport_(headers, fromStr, toStr) {
  // Confirmed working URL from testFindExportUrl() — returns HTML table @ ~1.1MB
  const url = CFG.panelBase +
    '/mis_report?export=excel' +
    '&from_date=' + encodeURIComponent(fromStr) +
    '&to_date='   + encodeURIComponent(toStr) +
    '&sales_poc=' + CFG.salesPoc +
    '&box8=1';

  Logger.log('Downloading: ' + url);
  const resp = UrlFetchApp.fetch(url, { headers, muteHttpExceptions: true, followRedirects: true });
  const code = resp.getResponseCode();
  const size = resp.getBlob().getBytes().length;
  Logger.log('Download → ' + code + ' | ' + size + ' bytes');

  if (code !== 200)
    throw new Error('Export returned HTTP ' + code);

  const body = resp.getContentText('UTF-8');
  checkSession_(body);

  if (size < 5000)
    throw new Error('Export file too small (' + size + ' bytes) — no data or wrong URL.');

  return body;
}


// ── HTML TABLE PARSER ─────────────────────────────────────────────────────────
function parseHtmlTable_(html) {
  // Strip scripts, styles, comments to simplify parsing
  const cleaned = html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '');

  function cellText(cellHtml) {
    return cellHtml
      .replace(/<[^>]+>/g, ' ')          // strip all tags
      .replace(/&amp;/g,  '&')
      .replace(/&lt;/g,   '<')
      .replace(/&gt;/g,   '>')
      .replace(/&nbsp;/g, ' ')
      .replace(/&#160;/g, ' ')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g,  "'")
      .replace(/\s+/g,    ' ')
      .trim();
  }

  function parseRow(trHtml) {
    const cells = [];
    const re = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    let m;
    while ((m = re.exec(trHtml)) !== null) cells.push(cellText(m[1]));
    return cells;
  }

  // Find header row (first <tr> with <th> tags, or first <tr> inside <thead>)
  let exportHeaders = [];
  const theadMatch = cleaned.match(/<thead[\s\S]*?>([\s\S]*?)<\/thead>/i);
  if (theadMatch) {
    const trMatch = theadMatch[1].match(/<tr[^>]*>([\s\S]*?)<\/tr>/i);
    if (trMatch) exportHeaders = parseRow(trMatch[1]);
  }

  if (exportHeaders.length === 0) {
    // No thead — try first <tr> of table
    const firstTr = cleaned.match(/<tr[^>]*>([\s\S]*?)<\/tr>/i);
    if (firstTr) exportHeaders = parseRow(firstTr[1]);
  }

  // All data rows — find tbody or all <tr> after header
  const exportRows = [];
  const tbodyMatch = cleaned.match(/<tbody[\s\S]*?>([\s\S]*?)<\/tbody>/i);
  const bodyHtml   = tbodyMatch ? tbodyMatch[1] : cleaned;

  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let trMatch;
  while ((trMatch = trRe.exec(bodyHtml)) !== null) {
    const row = parseRow(trMatch[1]);
    if (row.length > 0 && row.some(c => c !== '')) exportRows.push(row);
  }

  Logger.log('Parsed headers: ' + exportHeaders.length + ', rows: ' + exportRows.length);
  if (exportHeaders.length > 0) Logger.log('First 5 headers: ' + exportHeaders.slice(0,5).join(' | '));
  if (exportRows.length > 0)    Logger.log('First row sample: ' + exportRows[0].slice(0,5).join(' | '));

  return { exportHeaders, exportRows };
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
  if (missed.length) Logger.log('Blank (not matched): ' + missed.join(', '));
  return { colIdx };
}


// ── WRITE SHEET ───────────────────────────────────────────────────────────────
function writeSheet_(headers, rows) {
  const ss    = SpreadsheetApp.openById(CFG.sheetId);
  let   sheet = ss.getSheetByName(CFG.dataTab) || ss.insertSheet(CFG.dataTab, 0);
  sheet.clearContents();
  const now = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  sheet.getRange(1,1).setValue('Kunal Chauhan MIS — Updated: ' + now);
  sheet.getRange(1,1,1,headers.length).merge()
       .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(2,1,1,headers.length).setValues([headers])
       .setFontWeight('bold').setBackground('#e8ecf0');
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
  sheet.getRange(1,1,1,maxC).merge()
       .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
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
        <span style="background:${STATUS_CLR[s]};color:#fff;padding:3px 14px;
              border-radius:12px;font-weight:700;font-size:13px">${counts[s]}</span>
      </td>
    </tr>`).join('');

  const html = `
<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;color:#333">
  <div style="background:#1a1a2e;padding:24px 28px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff;font-size:20px">📦 ParcelX MIS Report</h2>
    <p style="margin:6px 0 0;color:#aaa;font-size:13px">
      Kunal Chauhan &nbsp;·&nbsp; ${now} &nbsp;·&nbsp; Last ${CFG.daysBack} Days
    </p>
  </div>
  <div style="background:#fff;padding:22px 28px;border:1px solid #ddd;border-top:none">
    <p style="margin:0 0 14px;font-size:14px"><strong>Total orders:</strong> ${rows.length}</p>
    <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:6px">
      <tr style="background:#f0f2f5">
        <th style="padding:10px 16px;text-align:left;font-size:13px">Status</th>
        <th style="padding:10px 16px;text-align:center;font-size:13px">Count</th>
      </tr>${statusRows}
    </table>
    <div style="text-align:center;margin:20px 0 8px">
      <a href="https://docs.google.com/spreadsheets/d/${CFG.sheetId}"
         style="display:inline-block;background:#34a853;color:#fff;
                padding:11px 28px;border-radius:6px;text-decoration:none;font-weight:700;font-size:14px">
        📊 Open Google Sheet
      </a>
    </div>
    <p style="margin:12px 0 0;font-size:12px;color:#999;text-align:center">
      CSV attached — ${rows.length} rows · ${headers.length} columns
    </p>
  </div>
</div>`;

  const esc = v => {
    const s = String(v == null ? '' : v);
    return (s.includes(',') || s.includes('"') || s.includes('\n'))
      ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = [COLS_WANTED.map(esc).join(','),
               ...rows.map(r => r.map(esc).join(','))].join('\n');
  const tag = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'ddMMMyyyy');

  GmailApp.sendEmail(CFG.emailTo,
    'ParcelX MIS | Kunal Chauhan | ' + now,
    rows.length + ' orders. CSV attached.',
    { htmlBody: html, cc: CFG.emailCc, name: 'ParcelX MIS',
      attachments: [Utilities.newBlob(csv, 'text/csv', 'Kunal_MIS_' + tag + '.csv')] }
  );
  Logger.log('✅ Email → ' + CFG.emailTo);
}


// ── ERROR HANDLER ─────────────────────────────────────────────────────────────
function handleError_(err) {
  const msg = err.toString();
  const isSession = msg.includes('SESSION_EXPIRED') || msg.includes('OTP_REQUIRED');
  const subject = isSession ? '⚠️ ParcelX MIS — Session Expired' : '❌ ParcelX MIS — Error';
  const body = isSession ? `
<div style="font-family:Arial,sans-serif;max-width:580px">
  <div style="background:#e74c3c;padding:20px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff">⚠️ MIS Session Expired</h2>
  </div>
  <div style="background:#fff;padding:20px;border:1px solid #ddd;border-top:none">
    <p>The automation could not run because the ParcelX session expired.</p>
    <p><strong>Fix in 2 minutes:</strong></p>
    <ol>
      <li>Go to <a href="https://panel.parcelx.in">panel.parcelx.in</a> → log in (complete OTP)</li>
      <li>F12 → Network tab → open mis_report page</li>
      <li>Click the request → Request Headers → copy <strong>Cookie:</strong> value</li>
      <li><a href="https://script.google.com">script.google.com</a> → Project Settings → Script Properties → update <strong>PARCELX_COOKIES</strong></li>
    </ol>
  </div>
</div>` : '<pre>' + msg + '</pre>';
  try {
    GmailApp.sendEmail(CFG.alertEmail, subject, msg.substring(0,500), { htmlBody: body, name: 'ParcelX MIS Alert' });
  } catch(e) {}
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
function testConnection() {
  const cookies = getStoredCookies_();
  const resp = UrlFetchApp.fetch(CFG.panelBase + '/mis_report', {
    headers: buildHeaders_(cookies), muteHttpExceptions: true
  });
  const body = resp.getContentText();
  Logger.log('Code: ' + resp.getResponseCode());
  Logger.log('Has searchbtn: ' + body.includes('searchbtn'));
  Logger.log('Has mis_report: ' + body.includes('mis_report'));
  Logger.log('Session OK: ' + (!body.toLowerCase().includes('enter otp') && body.includes('searchbtn')));
}

function testParseOnly() {
  // Test just the download + parse without writing to sheet
  const cookies = getStoredCookies_();
  const headers = buildHeaders_(cookies);
  const { fromStr, toStr } = getDateRange_();
  triggerSearch_(headers, fromStr, toStr);
  const html = downloadExport_(headers, fromStr, toStr);
  const { exportHeaders, exportRows } = parseHtmlTable_(html);
  Logger.log('Headers found (' + exportHeaders.length + '): ' + exportHeaders.join(' | '));
  Logger.log('Row 1: ' + (exportRows[0] || []).join(' | '));
  Logger.log('Row 2: ' + (exportRows[1] || []).join(' | '));
  Logger.log('Total rows: ' + exportRows.length);
}
