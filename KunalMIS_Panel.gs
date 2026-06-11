/**
 * KunalMIS_Panel.gs  — v4
 * Pulls data directly from panel.parcelx.in (no Metabase)
 * Uses stored session cookies → searches MIS → downloads Excel → Sheets + Email
 *
 * SETUP (one-time):
 *  1. Log in to panel.parcelx.in in Chrome (complete OTP)
 *  2. Open Console (F12) → paste this → press Enter:
 *       copy(JSON.stringify([...document.cookie.split(';').map(c=>{const[n,...v]=c.trim().split('=');return{name:n,value:v.join('=')}})]))
 *  3. Extensions → Apps Script → Project Settings → Script Properties → Add:
 *       PARCELX_COOKIES  =  (paste from clipboard)
 *  4. Run setupTrigger() once → sets 9 AM IST daily run
 *  5. Run testConnection() first to verify everything works
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────
const CFG = {
  sheetId   : '1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE',
  dataTab   : 'MIS Data',
  pivotTab  : 'Pivot',
  emailTo   : 'kunal.chauhan@parcelx.in',
  emailCc   : 'sudhanshu@parcelx.in,sparsh@parcelx.in',
  panelBase : 'https://panel.parcelx.in',
  basicUser : 'pax',
  basicPass : 'cloud@4w5',
  salesPoc  : '26',          // Kunal Chauhan
  daysBack  : 40
};

// Exact 43 columns to keep (in this order)
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

const STATUSES    = ['Booked','Manifested','Not Picked','Out For Pickup','Pickup Pending'];
const STATUS_CLR  = {
  'Booked':'#3498db','Manifested':'#9b59b6','Not Picked':'#e74c3c',
  'Out For Pickup':'#e67e22','Pickup Pending':'#f1c40f'
};


// ── MAIN ──────────────────────────────────────────────────────────────────────
function runMIS() {
  Logger.log('▶ Kunal MIS starting...');

  const cookies = getStoredCookies_();
  const headers = buildHeaders_(cookies);

  // Date range
  const { fromStr, toStr } = getDateRange_();
  Logger.log('Range: ' + fromStr + ' → ' + toStr);

  // 1. Trigger search on the panel
  triggerSearch_(headers, fromStr, toStr);

  // 2. Download Excel
  const blob = downloadExcel_(headers, fromStr, toStr);

  // 3. Parse Excel → rows
  const { allHeaders, allRows } = parseExcel_(blob);
  Logger.log('Total cols in export: ' + allHeaders.length + ', rows: ' + allRows.length);

  // 4. Filter to wanted 43 columns
  const { colIdx, finalHeaders } = mapColumns_(allHeaders);
  const finalRows = allRows.map(r => finalHeaders.map((_, i) => {
    const src = colIdx[i];
    return src !== -1 ? r[src] : '';
  }));

  // 5. Write to sheet
  writeSheet_(finalHeaders, finalRows);

  // 6. Pivot
  writePivot_(finalHeaders, finalRows);

  // 7. Email
  sendEmail_(finalHeaders, finalRows);

  Logger.log('✅ Done — ' + finalRows.length + ' rows');
}


// ── AUTH & HEADERS ────────────────────────────────────────────────────────────
function getStoredCookies_() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty('PARCELX_COOKIES');
  if (!raw) throw new Error(
    'PARCELX_COOKIES not set.\n' +
    'Log in to panel.parcelx.in, run copy(JSON.stringify([...document.cookie.split(\';\').map(c=>{const[n,...v]=c.trim().split(\'=\');return{name:n,value:v.join(\'=\')}})])) in Console, then set Script Property.'
  );

  // Accept both raw string and JSON array
  let cookieStr = raw;
  try {
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) {
      cookieStr = arr.map(c => c.name + '=' + c.value).join('; ');
    }
  } catch (e) { /* already a raw string */ }

  return cookieStr;
}

function buildHeaders_(cookies) {
  return {
    'Authorization'   : 'Basic ' + Utilities.base64Encode(CFG.basicUser + ':' + CFG.basicPass),
    'Cookie'          : cookies,
    'X-Requested-With': 'XMLHttpRequest',
    'Accept'          : 'text/html,application/xhtml+xml,application/json,*/*',
    'User-Agent'      : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
  };
}


// ── DATES ─────────────────────────────────────────────────────────────────────
function getDateRange_() {
  const to   = new Date();
  const from = new Date();
  from.setDate(from.getDate() - CFG.daysBack);
  return {
    fromStr : Utilities.formatDate(from, 'Asia/Kolkata', 'dd-MM-yyyy'),
    toStr   : Utilities.formatDate(to,   'Asia/Kolkata', 'dd-MM-yyyy')
  };
}


// ── SEARCH ────────────────────────────────────────────────────────────────────
function triggerSearch_(headers, fromStr, toStr) {
  const payload =
    'from_date='  + encodeURIComponent(fromStr)   +
    '&to_date='   + encodeURIComponent(toStr)     +
    '&sales_poc=' + CFG.salesPoc                  +
    '&box8=1&searchbtn=1';

  Logger.log('POSTing search...');
  const resp = UrlFetchApp.fetch(CFG.panelBase + '/mis_report', {
    method          : 'post',
    headers         : Object.assign({}, headers, {'Content-Type': 'application/x-www-form-urlencoded'}),
    payload         : payload,
    muteHttpExceptions: true,
    followRedirects : true
  });

  const code = resp.getResponseCode();
  Logger.log('Search response: ' + code);

  if (code === 200) {
    const body = resp.getContentText();
    if (body.toLowerCase().includes('login') && body.toLowerCase().includes('password')) {
      throw new Error('SESSION_EXPIRED: Please log in to panel again and update PARCELX_COOKIES in Script Properties.');
    }
    if (body.toLowerCase().includes('otp')) {
      throw new Error('OTP_REQUIRED: Session has expired. Log in manually (complete OTP), then update PARCELX_COOKIES.');
    }
  }

  // Brief pause so server can generate the export
  Utilities.sleep(3000);
}


// ── DOWNLOAD EXCEL ────────────────────────────────────────────────────────────
function downloadExcel_(headers, fromStr, toStr) {
  // Try export URL candidates
  const candidates = [
    CFG.panelBase + '/mis_report/export_excel',
    CFG.panelBase + '/mis_report/exportexcel',
    CFG.panelBase + '/mis_report/export',
    CFG.panelBase + '/mis_report/download',
    CFG.panelBase + '/mis_report/excel',
    CFG.panelBase + '/export/mis',
    CFG.panelBase + '/mis_report?export=excel&from_date=' + fromStr + '&to_date=' + toStr + '&sales_poc=' + CFG.salesPoc + '&box8=1',
    CFG.panelBase + '/mis_report?type=excel&from_date=' + fromStr + '&to_date=' + toStr + '&sales_poc=' + CFG.salesPoc
  ];

  for (const url of candidates) {
    try {
      Logger.log('Trying: ' + url);
      const resp = UrlFetchApp.fetch(url, {
        headers: headers,
        muteHttpExceptions: true,
        followRedirects: true
      });

      const code = resp.getResponseCode();
      const ct   = (resp.getHeaders()['Content-Type'] || resp.getHeaders()['content-type'] || '').toLowerCase();
      Logger.log('  → ' + code + ' | ' + ct);

      if (code === 200 && (
        ct.includes('spreadsheetml') || ct.includes('excel') ||
        ct.includes('octet-stream')  || ct.includes('openxmlformats') ||
        ct.includes('ms-excel')
      )) {
        Logger.log('✅ Got Excel from: ' + url);
        return resp.getBlob();
      }

      // If CSV
      if (code === 200 && ct.includes('csv')) {
        Logger.log('✅ Got CSV from: ' + url);
        return resp.getBlob().setContentType('text/csv');
      }

    } catch (e) {
      Logger.log('  Failed: ' + e);
    }
  }

  // Last resort: try POST export
  const exportPayload =
    'from_date='  + encodeURIComponent(fromStr) +
    '&to_date='   + encodeURIComponent(toStr)   +
    '&sales_poc=' + CFG.salesPoc                +
    '&box8=1&export=1&type=excel';

  for (const url of [CFG.panelBase + '/mis_report', CFG.panelBase + '/mis_report/export']) {
    try {
      Logger.log('POST export try: ' + url);
      const resp = UrlFetchApp.fetch(url, {
        method  : 'post',
        headers : Object.assign({}, headers, {'Content-Type': 'application/x-www-form-urlencoded'}),
        payload : exportPayload,
        muteHttpExceptions: true
      });
      const ct = (resp.getHeaders()['Content-Type'] || '').toLowerCase();
      if (resp.getResponseCode() === 200 && (ct.includes('excel') || ct.includes('spreadsheetml') || ct.includes('octet-stream'))) {
        Logger.log('✅ POST export worked: ' + url);
        return resp.getBlob();
      }
    } catch (e) {
      Logger.log('  POST failed: ' + e);
    }
  }

  throw new Error(
    'Could not download Excel file.\n' +
    'Run testFindExportUrl() to discover the correct export endpoint, ' +
    'then update the candidates list in downloadExcel_().'
  );
}


// ── PARSE EXCEL ───────────────────────────────────────────────────────────────
function parseExcel_(blob) {
  // Convert Excel → Google Sheets via Drive API, read data, then delete
  const ct = blob.getContentType() || '';

  if (ct.includes('csv') || ct.getBytes()[0] !== 0x50) {  // Not a zip/xlsx
    // Treat as CSV
    const csv = blob.getDataAsString('UTF-8');
    const rows = Utilities.parseCsv(csv);
    const allHeaders = rows[0] || [];
    const allRows    = rows.slice(1);
    return { allHeaders, allRows };
  }

  // Excel: convert via Drive Advanced Service
  const tempName = 'kunal_mis_temp_' + Date.now();
  const xlsxBlob = blob.setName(tempName + '.xlsx')
                       .setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  let convertedId = null;
  try {
    const converted = Drive.Files.insert(
      { title: tempName, mimeType: 'application/vnd.google-apps.spreadsheet' },
      xlsxBlob,
      { convert: true }
    );
    convertedId = converted.id;

    const ss         = SpreadsheetApp.openById(convertedId);
    const sheet      = ss.getSheets()[0];
    const raw        = sheet.getDataRange().getValues();
    const allHeaders = raw[0] ? raw[0].map(String) : [];
    const allRows    = raw.slice(1);

    return { allHeaders, allRows };

  } finally {
    if (convertedId) {
      try { Drive.Files.remove(convertedId); } catch(e) {}
    }
  }
}


// ── COLUMN MAPPING ────────────────────────────────────────────────────────────
function mapColumns_(exportHeaders) {
  // Normalise for fuzzy match
  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const normExport = exportHeaders.map(norm);

  const colIdx      = [];
  const finalHeaders = [];

  COLS_WANTED.forEach(wanted => {
    const normWanted = norm(wanted);
    // Exact match first
    let idx = normExport.indexOf(normWanted);
    // Partial match fallback
    if (idx === -1) {
      idx = normExport.findIndex(h =>
        h.includes(normWanted) || normWanted.includes(h)
      );
    }
    colIdx.push(idx);          // -1 = not found → blank column
    finalHeaders.push(wanted); // always use the user's preferred label
  });

  const found  = colIdx.filter(i => i !== -1).length;
  const missed = COLS_WANTED.filter((_, i) => colIdx[i] === -1);
  Logger.log('Columns matched: ' + found + ' / ' + COLS_WANTED.length);
  if (missed.length) Logger.log('Not found (will be blank): ' + missed.join(', '));

  return { colIdx, finalHeaders };
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

  if (rows.length)
    sheet.getRange(3,1,rows.length,headers.length).setValues(rows);

  sheet.setFrozenRows(2);
  headers.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  Logger.log('✅ Sheet written: ' + rows.length + ' rows');
}


// ── PIVOT ─────────────────────────────────────────────────────────────────────
function writePivot_(headers, rows) {
  const ss    = SpreadsheetApp.openById(CFG.sheetId);
  let   sheet = ss.getSheetByName(CFG.pivotTab) || ss.insertSheet(CFG.pivotTab);
  sheet.clearContents();

  const now  = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const sIdx = headers.indexOf('Current User Status');

  function section(groupCol, label) {
    const gIdx = headers.indexOf(groupCol);
    if (gIdx === -1) return [];
    const grp = {};
    rows.forEach(r => {
      const key = (r[gIdx] || '(blank)').toString().substring(0, 40);
      const st  = r[sIdx] || '';
      if (!grp[key]) grp[key] = {};
      grp[key][st] = (grp[key][st] || 0) + 1;
    });
    const out = [['── By ' + label + ' ──'],
                 [label, ...STATUSES, 'Total']];
    Object.entries(grp)
      .map(([k, v]) => [k, v, STATUSES.reduce((s, x) => s + (v[x] || 0), 0)])
      .sort((a, b) => b[2] - a[2])
      .forEach(([k, v, tot]) => out.push([k, ...STATUSES.map(s => v[s] || 0), tot]));
    out.push(['TOTAL', ...STATUSES.map(s => rows.filter(r => r[sIdx] === s).length), rows.length]);
    out.push(['']);
    return out;
  }

  let out = [['Kunal Chauhan – MIS Pivot | ' + now], ['']];
  out = out.concat(section('User Email',               'Seller (Email)'));
  out = out.concat(section('Order Type (Air / Surface)','Courier Type'));
  out = out.concat(section('Placed Date (D-M-Y)',      'Placed Date'));
  out = out.concat(section('Courier Used',              'Courier'));
  out = out.concat(section('Pickup City',               'Pickup City'));

  const maxC = Math.max(...out.map(r => r.length || 1));
  sheet.getRange(1, 1, out.length, maxC).setValues(out);
  sheet.getRange(1, 1, 1, maxC).merge()
       .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');

  Logger.log('✅ Pivot written');
}


// ── EMAIL ─────────────────────────────────────────────────────────────────────
function sendEmail_(headers, rows) {
  const now   = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const sIdx  = headers.indexOf('Current User Status');
  const total = rows.length;

  // Status counts
  const counts = {};
  STATUSES.forEach(s => { counts[s] = rows.filter(r => r[sIdx] === s).length; });

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
    <p style="margin:0 0 14px;font-size:14px">
      <strong>Total orders:</strong> ${total}
    </p>
    <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:6px">
      <tr style="background:#f0f2f5">
        <th style="padding:10px 16px;text-align:left;font-size:13px">Status</th>
        <th style="padding:10px 16px;text-align:center;font-size:13px">Count</th>
      </tr>${statusRows}
    </table>
    <div style="text-align:center;margin:20px 0 8px">
      <a href="https://docs.google.com/spreadsheets/d/${CFG.sheetId}"
         style="display:inline-block;background:#34a853;color:#fff;
                padding:11px 28px;border-radius:6px;text-decoration:none;
                font-weight:700;font-size:14px">📊 Open Google Sheet</a>
    </div>
    <p style="margin:12px 0 0;font-size:12px;color:#999;text-align:center">
      Full ${headers.length}-column data attached as CSV (${total} rows)
    </p>
  </div>
</div>`;

  // Build CSV
  const esc = v => {
    const s = (v == null) ? '' : String(v);
    return (s.includes(',') || s.includes('"') || s.includes('\n'))
      ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = [headers.map(esc).join(','),
               ...rows.map(r => r.map(esc).join(','))].join('\n');

  const dateTag = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'ddMMMyyyy');
  GmailApp.sendEmail(
    CFG.emailTo,
    `ParcelX MIS | Kunal Chauhan | ${now}`,
    `${total} orders. CSV attached. View Google Sheet for full detail.`,
    {
      htmlBody    : html,
      cc          : CFG.emailCc,
      name        : 'ParcelX MIS',
      attachments : [Utilities.newBlob(csv, 'text/csv', 'Kunal_MIS_' + dateTag + '.csv')]
    }
  );
  Logger.log('✅ Email sent → ' + CFG.emailTo);
}


// ── TRIGGER ───────────────────────────────────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMIS')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runMIS').timeBased()
    .everyDays(1).atHour(9).inTimezone('Asia/Kolkata').create();

  Logger.log('✅ Daily 9 AM IST trigger set');
}


// ── DEBUG HELPERS ─────────────────────────────────────────────────────────────

/** Run this first to verify login + session is working */
function testConnection() {
  const cookies = getStoredCookies_();
  const headers = buildHeaders_(cookies);

  const resp = UrlFetchApp.fetch(CFG.panelBase + '/mis_report', {
    headers: headers,
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();

  Logger.log('Response code: ' + code);
  Logger.log('Page title check: ' + (body.match(/<title>(.*?)<\/title>/i)||['','(no title)'])[1]);
  Logger.log('Has "mis_report" in page: ' + body.includes('mis_report'));
  Logger.log('Session expired?: ' + (body.toLowerCase().includes('login') && body.toLowerCase().includes('password')));
  Logger.log('OTP page?: ' + body.toLowerCase().includes('enter otp'));
  Logger.log('First 500 chars: ' + body.substring(0,500));
}

/** Finds the actual export endpoint — run once after testConnection() passes */
function testFindExportUrl() {
  const cookies = getStoredCookies_();
  const headers = buildHeaders_(cookies);
  const { fromStr, toStr } = getDateRange_();

  // First trigger the search
  triggerSearch_(headers, fromStr, toStr);

  // Now probe all candidate export URLs
  const candidates = [
    '/mis_report/export_excel', '/mis_report/exportexcel', '/mis_report/export',
    '/mis_report/download', '/mis_report/excel', '/export/mis',
    '/mis_report?export=excel&from_date=' + fromStr + '&to_date=' + toStr + '&sales_poc=' + CFG.salesPoc + '&box8=1',
    '/mis_report/export_csv', '/mis_report/csv'
  ];

  candidates.forEach(path => {
    try {
      const resp = UrlFetchApp.fetch(CFG.panelBase + path, {
        headers: headers, muteHttpExceptions: true
      });
      const ct = resp.getHeaders()['Content-Type'] || '';
      Logger.log(resp.getResponseCode() + ' | ' + ct.substring(0,60) + '  →  ' + path);
    } catch(e) {
      Logger.log('ERROR | ' + path + ' | ' + e);
    }
  });
}

/** Once you find the export URL from testFindExportUrl(), add it here */
function setCustomExportUrl(url) {
  PropertiesService.getScriptProperties().setProperty('EXPORT_URL_OVERRIDE', url);
  Logger.log('Export URL saved: ' + url);
}
