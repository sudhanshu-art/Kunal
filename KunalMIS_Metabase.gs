/**
 * KunalMIS_Metabase.gs
 * Pulls Kunal Chauhan MIS data from Metabase API → Google Sheet → Email
 * No browser, no cookies, no OTP — just API credentials
 *
 * SETUP (one-time):
 *  1. Extensions → Apps Script → Services (+) → (nothing extra needed)
 *  2. Project Settings → Script Properties → Add:
 *       METABASE_USER  =  your mb.parcelx.in login email
 *       METABASE_PASS  =  your mb.parcelx.in password
 *  3. Run testConnection() to verify
 *  4. Run runMIS() to test full run
 *  5. Run setupTrigger() once for daily 9 AM IST
 */

// ── CONFIG ────────────────────────────────────────────────────────────────────
const CFG = {
  sheetId   : '1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE',
  dataTab   : 'MIS Data',
  pivotTab  : 'Pivot',
  emailTo   : 'kunal.chauhan@parcelx.in',
  emailCc   : 'sudhanshu@parcelx.in,sparsh@parcelx.in',
  alertEmail: 'sudhanshu@parcelx.in',
  mbUrl     : 'https://mb.parcelx.in',
  mbDb      : 2,
  daysBack  : 40
};

// ── SQL — all 43 columns for Kunal Chauhan (last 40 days) ────────────────────
const SQL = `
SELECT
  DATE_FORMAT(si.placed_date,     '%d-%m-%Y')           AS 'Placed Date (D-M-Y)',
  DATE_FORMAT(si.placed_date,     '%H:%i')               AS 'Placed Time',
  s.shipment_id                                           AS 'ParcelX Order ID',
  DATE_FORMAT(si.pickup_date,     '%d-%m-%Y')            AS 'Pickup Date (D-M-Y)',
  DATE_FORMAT(si.pickup_date,     '%H:%i')                AS 'Pickup Time',
  DATE_FORMAT(si.delivered_date,  '%d-%m-%Y')            AS 'Delivered Date',
  ss.status_title                                         AS 'Current User Status',
  s.waybill_number                                        AS 'Parcelx Tracking ID',
  COALESCE(si.pregen_waybill, s.ref_awb, '')             AS 'Pre Generated WayBill',
  ''                                                      AS 'RTO Waybill',
  COALESCE(CAST(s.client_order_id AS CHAR),
           s.invoice_number, '')                          AS 'Client Order ID / Invoice no',
  COALESCE(CAST(s.channel_order_id AS CHAR), '')          AS 'Channel Order Number',
  s.user_id                                               AS 'User Email',
  DATE_FORMAT(si.rts_date,        '%d-%m-%Y')            AS 'RTS Date',
  DATE_FORMAT(si.rto_date,        '%d-%m-%Y')            AS 'RTO Marked Date',
  DATE_FORMAT(si.first_ofd_date,  '%d-%m-%Y %H:%i')      AS 'First Attempt Date',
  DATE_FORMAT(si.last_ofd_date,   '%d-%m-%Y %H:%i')      AS 'Latest Attempt Date',
  'Kunal Chauhan'                                         AS 'Sales POC',
  COALESCE(
    JSON_UNQUOTE(JSON_EXTRACT(s.order_data,'$.product_name[0]')), ''
  )                                                       AS 'Item Name',
  COALESCE(
    JSON_UNQUOTE(JSON_EXTRACT(s.order_data,'$.product_quantity[0]')), ''
  )                                                       AS 'Item Qty',
  COALESCE(s.shipment_length, '')                         AS 'Item Length',
  COALESCE(s.shipment_height, '')                         AS 'Item Height',
  COALESCE(s.shipment_width,  '')                         AS 'Item Width',
  COALESCE(s.shipment_weight, '')                         AS 'Item Weight',
  ''                                                      AS 'CN Weight',
  ''                                                      AS 'DN Weight',
  ''                                                      AS 'Charged Weight',
  ''                                                      AS 'Updated Actual Weight',
  COALESCE(s.invoice_value, 0)                            AS 'Invoice Value',
  COALESCE(s.cod_amount,    0)                            AS 'COD',
  COALESCE(s.express_type,  '')                           AS 'Order Type (Air / Surface)',
  ''                                                      AS 'Courier Used',
  COALESCE(CAST(s.fulfilled_account AS CHAR), '')         AS 'Courier Account',
  ''                                                      AS 'Warehouse Name',
  ''                                                      AS 'Pickup Name',
  ''                                                      AS 'Pickup City',
  ''                                                      AS 'Pickup State',
  COALESCE(si.rto_reason,       '')                       AS 'RTO Reason',
  COALESCE(si.attempt_count,    0)                        AS 'Attempt Counts',
  DATE_FORMAT(si.ndr_first_date, '%d-%m-%Y %H:%i')       AS 'NDR First Date',
  COALESCE(si.ndr_first_remarks, '')                      AS 'NDR First Remarks',
  DATE_FORMAT(si.ndr_last_date,  '%d-%m-%Y %H:%i')       AS 'NDR Last Date',
  COALESCE(si.ndr_last_remarks,  '')                      AS 'NDR Last Remarks'

FROM shipments s
LEFT JOIN shipment_info   si  ON si.shipment_id = s.shipment_id
LEFT JOIN shipments_status ss ON ss.status_id   = s.user_status
                              AND ss.status_for  = 'User'
WHERE s.user_status IN (221, 222, 231, 232, 238)
  AND s.user_id IN (
      SELECT user_id FROM child_sales_poc_user_assign WHERE admin_id = 38
  )
  AND s.added_on >= DATE_SUB(NOW(), INTERVAL 40 DAY)
ORDER BY s.added_on DESC
LIMIT 5000
`;

const COLS = [
  'Placed Date (D-M-Y)','Placed Time','ParcelX Order ID',
  'Pickup Date (D-M-Y)','Pickup Time','Delivered Date',
  'Current User Status','Parcelx Tracking ID','Pre Generated WayBill',
  'RTO Waybill','Client Order ID / Invoice no','Channel Order Number',
  'User Email','RTS Date','RTO Marked Date',
  'First Attempt Date','Latest Attempt Date','Sales POC',
  'Item Name','Item Qty','Item Length','Item Height','Item Width','Item Weight',
  'CN Weight','DN Weight','Charged Weight','Updated Actual Weight',
  'Invoice Value','COD','Order Type (Air / Surface)',
  'Courier Used','Courier Account','Warehouse Name',
  'Pickup Name','Pickup City','Pickup State',
  'RTO Reason','Attempt Counts',
  'NDR First Date','NDR First Remarks','NDR Last Date','NDR Last Remarks'
];

const STATUSES  = ['Booked','Manifested','Not Picked','Out For Pickup','Pickup Pending'];
const STATUS_CLR = {
  'Booked':'#3498db','Manifested':'#9b59b6','Not Picked':'#e74c3c',
  'Out For Pickup':'#e67e22','Pickup Pending':'#f1c40f'
};


// ── MAIN ──────────────────────────────────────────────────────────────────────
function runMIS() {
  Logger.log('▶ Kunal MIS (Metabase) starting...');
  try {
    const rows = fetchFromMetabase_();
    Logger.log('Rows fetched: ' + rows.length);

    writeSheet_(rows);
    writePivot_(rows);
    sendEmail_(rows);
    Logger.log('✅ Done — ' + rows.length + ' rows');

  } catch (err) {
    Logger.log('❌ ' + err);
    handleError_(err);
    throw err;
  }
}


// ── METABASE API ──────────────────────────────────────────────────────────────
function getToken_() {
  const props = PropertiesService.getScriptProperties();
  const user  = props.getProperty('METABASE_USER');
  const pass  = props.getProperty('METABASE_PASS');
  if (!user || !pass)
    throw new Error('Set METABASE_USER and METABASE_PASS in Script Properties');

  const res = UrlFetchApp.fetch(CFG.mbUrl + '/api/session', {
    method          : 'post',
    contentType     : 'application/json',
    payload         : JSON.stringify({ username: user, password: pass }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200)
    throw new Error('Metabase login failed (' + res.getResponseCode() + '): ' + res.getContentText().substring(0, 200));

  return JSON.parse(res.getContentText()).id;
}

function fetchFromMetabase_() {
  const token = getToken_();
  Logger.log('Token: ' + token.substring(0, 8) + '...');

  const res = UrlFetchApp.fetch(CFG.mbUrl + '/api/dataset', {
    method          : 'post',
    contentType     : 'application/json',
    headers         : { 'X-Metabase-Session': token },
    payload         : JSON.stringify({
      database : CFG.mbDb,
      type     : 'native',
      native   : { query: SQL }
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200 && code !== 202)
    throw new Error('Metabase query failed (' + code + '): ' + res.getContentText().substring(0, 300));

  const json = JSON.parse(res.getContentText());
  if (json.status === 'failed') throw new Error('Query error: ' + json.error);

  // Map column names to rows
  const cols    = json.data.cols.map(c => c.display_name || c.name);
  const rawRows = json.data.rows;

  Logger.log('Metabase cols: ' + cols.length + ', rows: ' + rawRows.length);

  // Reorder to match COLS order
  const colIdx = COLS.map(wanted => {
    const norm  = s => s.toLowerCase().replace(/[^a-z0-9]/g, '');
    const nw    = norm(wanted);
    const nc    = cols.map(norm);
    let   idx   = nc.indexOf(nw);
    if (idx === -1) idx = nc.findIndex(h => h.includes(nw) || nw.includes(h));
    return idx;
  });

  const missed = COLS.filter((_, i) => colIdx[i] === -1);
  if (missed.length) Logger.log('Not matched (blank): ' + missed.join(', '));

  return rawRows.map(r => COLS.map((_, i) => colIdx[i] !== -1 ? (r[colIdx[i]] ?? '') : ''));
}


// ── WRITE SHEET ───────────────────────────────────────────────────────────────
function writeSheet_(rows) {
  const ss    = SpreadsheetApp.openById(CFG.sheetId);
  let   sheet = ss.getSheetByName(CFG.dataTab) || ss.insertSheet(CFG.dataTab, 0);
  sheet.clearContents();

  const now = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");

  // Title
  sheet.getRange(1,1).setValue('Kunal Chauhan MIS — Updated: ' + now);
  sheet.getRange(1,1,1,COLS.length).merge()
       .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);

  // Headers
  sheet.getRange(2,1,1,COLS.length).setValues([COLS])
       .setFontWeight('bold').setBackground('#e8ecf0');

  // Data
  if (rows.length)
    sheet.getRange(3,1,rows.length,COLS.length).setValues(rows);

  // Format
  sheet.setFrozenRows(2);
  COLS.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  if (rows.length)
    sheet.getRange(3,1,rows.length,COLS.length)
         .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  Logger.log('✅ Sheet written: ' + rows.length + ' rows');
}


// ── PIVOT ─────────────────────────────────────────────────────────────────────
function writePivot_(rows) {
  const ss    = SpreadsheetApp.openById(CFG.sheetId);
  let   sheet = ss.getSheetByName(CFG.pivotTab) || ss.insertSheet(CFG.pivotTab);
  sheet.clearContents();

  const now  = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const sIdx = COLS.indexOf('Current User Status');

  function section(col, label) {
    const gIdx = COLS.indexOf(col);
    if (gIdx === -1) return [];
    const grp = {};
    rows.forEach(r => {
      const k = String(r[gIdx] || '(blank)').substring(0, 40);
      const s = String(r[sIdx] || '');
      if (!grp[k]) grp[k] = {};
      grp[k][s] = (grp[k][s] || 0) + 1;
    });
    const out = [
      ['── By ' + label + ' ──'],
      [label, ...STATUSES, 'Total']
    ];
    Object.entries(grp)
      .map(([k, v]) => [k, v, STATUSES.reduce((s, x) => s + (v[x] || 0), 0)])
      .sort((a, b) => b[2] - a[2])
      .forEach(([k, v, tot]) => out.push([k, ...STATUSES.map(s => v[s] || 0), tot]));
    out.push(['TOTAL', ...STATUSES.map(s => rows.filter(r => r[sIdx] === s).length), rows.length]);
    out.push(['']);
    return out;
  }

  let out = [['Kunal Chauhan — MIS Pivot  |  ' + now], ['']];
  out = out.concat(section('User Email',               'Seller'));
  out = out.concat(section('Order Type (Air / Surface)','Courier Type'));
  out = out.concat(section('Placed Date (D-M-Y)',      'Placed Date'));
  out = out.concat(section('Courier Account',           'Courier Account'));
  out = out.concat(section('Current User Status',       'Status'));

  const maxC = Math.max(...out.map(r => r.length || 1));
  sheet.getRange(1, 1, out.length, maxC).setValues(out);
  sheet.getRange(1, 1, 1, maxC).merge()
       .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');

  Logger.log('✅ Pivot written');
}


// ── EMAIL ─────────────────────────────────────────────────────────────────────
function sendEmail_(rows) {
  const now   = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const sIdx  = COLS.indexOf('Current User Status');
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
    <p style="margin:0 0 14px;font-size:14px"><strong>Total orders:</strong> ${rows.length}</p>
    <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:6px">
      <tr style="background:#f0f2f5">
        <th style="padding:10px 16px;text-align:left;font-size:13px">Status</th>
        <th style="padding:10px 16px;text-align:center;font-size:13px">Count</th>
      </tr>${statusRows}
    </table>
    <div style="text-align:center;margin:20px 0 8px">
      <a href="https://docs.google.com/spreadsheets/d/${CFG.sheetId}"
         style="display:inline-block;background:#34a853;color:#fff;padding:11px 28px;
                border-radius:6px;text-decoration:none;font-weight:700;font-size:14px">
        📊 Open Google Sheet
      </a>
    </div>
    <p style="margin:12px 0 0;font-size:12px;color:#999;text-align:center">
      CSV attached — ${rows.length} rows · ${COLS.length} columns
    </p>
  </div>
</div>`;

  const esc = v => {
    const s = String(v == null ? '' : v);
    return (s.includes(',') || s.includes('"') || s.includes('\n'))
      ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  const csv = [COLS.map(esc).join(','), ...rows.map(r => r.map(esc).join(','))].join('\n');
  const tag = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'ddMMMyyyy');

  GmailApp.sendEmail(CFG.emailTo,
    'ParcelX MIS | Kunal Chauhan | ' + now,
    rows.length + ' orders. CSV attached.',
    {
      htmlBody    : html,
      cc          : CFG.emailCc,
      name        : 'ParcelX MIS',
      attachments : [Utilities.newBlob(csv, 'text/csv', 'Kunal_MIS_' + tag + '.csv')]
    }
  );
  Logger.log('✅ Email sent → ' + CFG.emailTo);
}


// ── ERROR HANDLER ─────────────────────────────────────────────────────────────
function handleError_(err) {
  const msg  = err.toString();
  const html = `
<div style="font-family:Arial,sans-serif;max-width:580px">
  <div style="background:#e74c3c;padding:20px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff">❌ Kunal MIS — Error</h2>
  </div>
  <div style="background:#fff;padding:20px;border:1px solid #ddd;border-top:none">
    <pre style="background:#f5f5f5;padding:12px;border-radius:4px;font-size:12px">${msg}</pre>
    <p>Check Apps Script logs for full details.</p>
  </div>
</div>`;
  try {
    GmailApp.sendEmail(CFG.alertEmail, '❌ Kunal MIS Error', msg.substring(0, 300),
      { htmlBody: html, name: 'ParcelX MIS Alert' });
  } catch(e) {}
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


// ── DEBUG ─────────────────────────────────────────────────────────────────────
/** Run this first to verify Metabase login works */
function testConnection() {
  const token = getToken_();
  Logger.log('✅ Metabase login OK — token: ' + token.substring(0, 8) + '...');
}

/** Run this to test data fetch without writing to sheet */
function testFetch() {
  const rows = fetchFromMetabase_();
  Logger.log('Total rows: ' + rows.length);
  Logger.log('Sample row 1: ' + rows[0]);
  Logger.log('Sample row 2: ' + rows[1]);
  const sIdx = COLS.indexOf('Current User Status');
  const counts = {};
  STATUSES.forEach(s => counts[s] = rows.filter(r => r[sIdx] === s).length);
  Logger.log('Status counts: ' + JSON.stringify(counts));
}
