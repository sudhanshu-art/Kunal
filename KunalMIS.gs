// ============================================================
// KUNAL MIS — Google Apps Script  (v3)
// Metabase API → Google Sheet → Email with CSV attachment
// ============================================================

const CONFIG = {
  sheetId  : "1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE",
  dataTab  : "MIS Data",
  pivotTab : "Pivot",
  emailTo  : "kunal.chauhan@parcelx.in",
  emailCc  : "sudhanshu@parcelx.in,sparsh@parcelx.in",
  mbUrl    : "https://mb.parcelx.in",
  mbDb     : 2
};

// ── SQL ──────────────────────────────────────────────────────
// admin_id=38 → Kunal Chauhan's seller group
const SQL = `
SELECT
  DATE_FORMAT(si.placed_date,      '%d-%m-%Y')       AS 'Placed Date (D-M-Y)',
  DATE_FORMAT(si.placed_date,      '%H:%i')           AS 'Placed Time',
  s.shipment_id                                        AS 'ParcelX Order ID',
  DATE_FORMAT(si.pickup_date,      '%d-%m-%Y')        AS 'Pickup Date (D-M-Y)',
  DATE_FORMAT(si.pickup_date,      '%H:%i')            AS 'Pickup Time',
  DATE_FORMAT(si.delivered_date,   '%d-%m-%Y %H:%i')  AS 'Delivered Date',
  ss.status_title                                      AS 'Current User Status',
  s.waybill_number                                     AS 'Parcelx Tracking ID',
  COALESCE(si.pregen_waybill, s.ref_awb, '')          AS 'Pre Generated WayBill',
  ''                                                   AS 'RTO Waybill',
  COALESCE(CAST(s.client_order_id AS CHAR),
           s.invoice_number, '')                       AS 'Client Order ID / Invoice no',
  COALESCE(CAST(s.channel_order_id AS CHAR), '')       AS 'Channel Order Number',
  s.user_id                                            AS 'User Email',
  DATE_FORMAT(si.rts_date,         '%d-%m-%Y')        AS 'RTS Date',
  DATE_FORMAT(si.rto_date,         '%d-%m-%Y')        AS 'RTO Marked Date',
  DATE_FORMAT(si.first_ofd_date,   '%d-%m-%Y %H:%i')  AS 'First Attempt Date',
  DATE_FORMAT(si.last_ofd_date,    '%d-%m-%Y %H:%i')  AS 'Latest Attempt Date',
  'Kunal Chauhan'                                      AS 'Sales POC',
  COALESCE(
    JSON_UNQUOTE(JSON_EXTRACT(s.order_data, '$.product_name[0]')), ''
  )                                                    AS 'Item Name',
  COALESCE(
    JSON_UNQUOTE(JSON_EXTRACT(s.order_data, '$.product_quantity[0]')), ''
  )                                                    AS 'Item Qty',
  COALESCE(s.shipment_length, '')                      AS 'Item Length',
  COALESCE(s.shipment_height, '')                      AS 'Item Height',
  COALESCE(s.shipment_width,  '')                      AS 'Item Width',
  COALESCE(s.shipment_weight, '')                      AS 'Item Weight',
  ''                                                   AS 'CN Weight',
  ''                                                   AS 'DN Weight',
  ''                                                   AS 'Charged Weight',
  ''                                                   AS 'Updated Actual Weight',
  COALESCE(s.invoice_value, 0)                         AS 'Invoice Value',
  COALESCE(s.cod_amount,    0)                         AS 'COD',
  COALESCE(s.express_type,  '')                        AS 'Order Type (Air / Surface)',
  ''                                                   AS 'Courier Used',
  COALESCE(s.fulfilled_account, '')                    AS 'Courier Account',
  ''                                                   AS 'Warehouse Name',
  ''                                                   AS 'Pickup Name',
  ''                                                   AS 'Pickup City',
  ''                                                   AS 'Pickup State',
  COALESCE(si.rto_reason,       '')                    AS 'RTO Reason',
  COALESCE(si.attempt_count,    0)                     AS 'Attempt Counts',
  DATE_FORMAT(si.ndr_first_date, '%d-%m-%Y %H:%i')    AS 'NDR First Date',
  COALESCE(si.ndr_first_remarks, '')                   AS 'NDR First Remarks',
  DATE_FORMAT(si.ndr_last_date,  '%d-%m-%Y %H:%i')    AS 'NDR Last Date',
  COALESCE(si.ndr_last_remarks,  '')                   AS 'NDR Last Remarks'

FROM shipments s
LEFT JOIN shipment_info si     ON si.shipment_id = s.shipment_id
LEFT JOIN shipments_status ss  ON ss.status_id   = s.user_status
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

const STATUS_LIST   = ['Booked','Manifested','Not Picked','Out For Pickup','Pickup Pending'];
const STATUS_COLORS = {
  'Booked':'#3498db','Manifested':'#9b59b6','Not Picked':'#e74c3c',
  'Out For Pickup':'#e67e22','Pickup Pending':'#f1c40f'
};

// ── MAIN ─────────────────────────────────────────────────────
function runMIS() {
  Logger.log('▶ Kunal MIS starting...');
  const data = fetchFromMetabase_();
  Logger.log('  Rows: ' + data.rows.length);
  writeSheet_(data);
  writePivot_(data);
  sendEmail_(data);
  Logger.log('✅ Done');
}

// ── METABASE ─────────────────────────────────────────────────
function getToken_() {
  const p    = PropertiesService.getScriptProperties();
  const user = p.getProperty('METABASE_USER');
  const pass = p.getProperty('METABASE_PASS');
  if (!user || !pass) throw new Error('Set METABASE_USER + METABASE_PASS in Script Properties');
  const res = UrlFetchApp.fetch(CONFIG.mbUrl + '/api/session', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ username: user, password: pass }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200)
    throw new Error('Metabase login failed: ' + res.getContentText());
  return JSON.parse(res.getContentText()).id;
}

function fetchFromMetabase_() {
  const token = getToken_();
  const res   = UrlFetchApp.fetch(CONFIG.mbUrl + '/api/dataset', {
    method: 'post', contentType: 'application/json',
    headers: { 'X-Metabase-Session': token },
    payload: JSON.stringify({ database: CONFIG.mbDb, type: 'native', native: { query: SQL } }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200 && code !== 202)
    throw new Error('Query failed (' + code + '): ' + res.getContentText().substring(0, 400));
  const json = JSON.parse(res.getContentText());
  if (json.status === 'failed') throw new Error('Query error: ' + json.error);
  return json.data;
}

// ── GOOGLE SHEET ─────────────────────────────────────────────
function writeSheet_(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  let   sheet = ss.getSheetByName(CONFIG.dataTab) || ss.insertSheet(CONFIG.dataTab, 0);
  sheet.clearContents();

  const now = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");

  // Title
  sheet.getRange(1,1).setValue('Kunal Chauhan MIS — Updated: ' + now);
  sheet.getRange(1,1,1,COLS.length).merge()
       .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');

  // Headers
  sheet.getRange(2,1,1,COLS.length).setValues([COLS])
       .setFontWeight('bold').setBackground('#e8ecf0');

  // Data
  if (data.rows.length)
    sheet.getRange(3,1,data.rows.length,COLS.length).setValues(data.rows);

  sheet.setFrozenRows(2);
  COLS.forEach((_,i) => sheet.autoResizeColumn(i+1));
  Logger.log('  Sheet written: ' + data.rows.length + ' rows');
}

// ── PIVOT ─────────────────────────────────────────────────────
function writePivot_(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  let   sheet = ss.getSheetByName(CONFIG.pivotTab) || ss.insertSheet(CONFIG.pivotTab);
  sheet.clearContents();

  const now   = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const cols  = data.cols.map(c => c.display_name || c.name);
  const rows  = data.rows;
  const sIdx  = cols.indexOf('Current User Status');

  function section(idx, label) {
    if (idx < 0) return [];
    const grp = {};
    rows.forEach(r => {
      const key = (r[idx] || '(blank)').toString().substring(0,40);
      const st  = r[sIdx] || '';
      if (!grp[key]) grp[key] = {};
      grp[key][st] = (grp[key][st]||0) + 1;
    });
    const out = [['── By ' + label + ' ──'],
                 [label, ...STATUS_LIST, 'Total']];
    Object.entries(grp)
      .map(([k,v]) => [k, v, STATUS_LIST.reduce((s,x)=>s+(v[x]||0),0)])
      .sort((a,b)=>b[2]-a[2])
      .forEach(([k,v,tot]) => out.push([k,...STATUS_LIST.map(s=>v[s]||0),tot]));
    out.push(['TOTAL',...STATUS_LIST.map(s=>rows.filter(r=>r[sIdx]===s).length),rows.length]);
    out.push(['']);
    return out;
  }

  let out = [['Kunal Chauhan – Pivot | ' + now],['']];
  out = out.concat(section(cols.indexOf('Sales POC'),            'Sales POC'));
  out = out.concat(section(cols.indexOf('Order Type (Air / Surface)'), 'Courier Type'));
  out = out.concat(section(cols.indexOf('Placed Date (D-M-Y)'), 'Placed Date'));

  sheet.getRange(1,1,out.length, Math.max(...out.map(r=>r.length||1))).setValues(out);
  sheet.getRange(1,1,1,6).merge().setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
}

// ── EMAIL ─────────────────────────────────────────────────────
function sendEmail_(data) {
  const now   = Utilities.formatDate(new Date(), 'Asia/Kolkata', "dd MMM yyyy, hh:mm a 'IST'");
  const cols  = data.cols.map(c => c.display_name || c.name);
  const rows  = data.rows;
  const sIdx  = cols.indexOf('Current User Status');
  const total = rows.length;

  // Status counts
  const counts = {};
  STATUS_LIST.forEach(s => { counts[s] = rows.filter(r => r[sIdx]===s).length; });

  const statusRows = STATUS_LIST.map(s => `
    <tr>
      <td style="padding:9px 16px;border-bottom:1px solid #f0f0f0;font-size:14px">${s}</td>
      <td style="padding:9px 16px;border-bottom:1px solid #f0f0f0;text-align:center">
        <span style="background:${STATUS_COLORS[s]};color:#fff;padding:3px 14px;
              border-radius:12px;font-weight:700;font-size:13px">${counts[s]}</span>
      </td>
    </tr>`).join('');

  const html = `
<div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;color:#333">
  <div style="background:#1a1a2e;padding:24px 28px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff;font-size:20px">📦 ParcelX MIS Report</h2>
    <p style="margin:6px 0 0;color:#aaa;font-size:13px">
      Kunal Chauhan &nbsp;·&nbsp; ${now} &nbsp;·&nbsp; Last 30 Days
    </p>
  </div>
  <div style="background:#fff;padding:22px 28px;border:1px solid #ddd;border-top:none">
    <p style="margin:0 0 14px;font-size:14px">
      <strong>Total orders in report:</strong> ${total}
    </p>
    <table style="width:100%;border-collapse:collapse;border-radius:6px;overflow:hidden">
      <tr style="background:#f0f2f5">
        <th style="padding:10px 16px;text-align:left;font-size:13px;font-weight:600">Status</th>
        <th style="padding:10px 16px;text-align:center;font-size:13px;font-weight:600">Count</th>
      </tr>${statusRows}
    </table>
    <div style="text-align:center;margin:20px 0 8px">
      <a href="https://docs.google.com/spreadsheets/d/${CONFIG.sheetId}"
         style="display:inline-block;background:#34a853;color:#fff;
                padding:11px 28px;border-radius:6px;text-decoration:none;
                font-weight:700;font-size:14px">
        📊 Open Full Sheet
      </a>
    </div>
    <p style="margin:12px 0 0;font-size:12px;color:#999;text-align:center">
      Full data (${COLS.length} columns) attached as CSV
    </p>
  </div>
</div>`;

  // ── Build CSV attachment ──────────────────────────────────
  function csvEscape(v) {
    const s = (v === null || v === undefined) ? '' : String(v);
    return (s.includes(',') || s.includes('"') || s.includes('\n'))
      ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  const csvLines = [COLS.map(csvEscape).join(',')];
  rows.forEach(r => csvLines.push(r.map(csvEscape).join(',')));
  const csvContent = csvLines.join('\n');

  // ── Send ─────────────────────────────────────────────────
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'ddMMMyyyy');
  GmailApp.sendEmail(
    CONFIG.emailTo,
    `ParcelX MIS | Kunal Chauhan | ${now}`,
    `${total} orders. See HTML for summary, CSV attached for full detail.`,
    {
      htmlBody    : html,
      cc          : CONFIG.emailCc,
      name        : 'ParcelX MIS',
      attachments : [Utilities.newBlob(csvContent, 'text/csv',
                      `Kunal_MIS_${dateStr}.csv`)]
    }
  );
  Logger.log('  Email sent → ' + CONFIG.emailTo);
}

// ── TRIGGER ───────────────────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMIS')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 9 AM IST daily
  ScriptApp.newTrigger('runMIS').timeBased()
    .everyDays(1).atHour(9).inTimezone('Asia/Kolkata').create();

  Logger.log('✅ Daily 9 AM IST trigger created');
}
