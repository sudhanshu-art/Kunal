// ============================================================
// KUNAL MIS — Google Apps Script
// Pulls live shipments from Metabase → Google Sheet → Email
// ============================================================

// ── SETTINGS (change these) ──────────────────────────────────
const CONFIG = {
  sheetId   : "1gUwZoO0v3Rjkx4XshHHuBUwPyMfLY7KaE7UfzBEPLyE",
  dataTab   : "MIS Data",
  pivotTab  : "Pivot",
  emailTo   : "kunal.chauhan@parcelx.in",
  emailCc   : "sudhanshu@parcelx.in,sparsh@parcelx.in",
  mbUrl     : "https://mb.parcelx.in",
  mbDb      : 2
};

// Metabase credentials — set these via:
// Extensions → Apps Script → Project Settings → Script Properties
// Keys: METABASE_USER  and  METABASE_PASS

// Status IDs → labels (from shipments_status table)
const STATUS_LABELS = {
  221: "Booked",
  222: "Manifested",
  231: "Not Picked",
  232: "Pickup Pending",
  238: "Out For Pickup"
};

// ── SQL QUERY ────────────────────────────────────────────────
// Fetches Kunal Chauhan's open orders (last 30 days)
// admin_id=38 maps to Kunal's seller group in child_sales_poc_user_assign
const SQL = `
SELECT
  s.shipment_id,
  s.waybill_number,
  ss.status_title        AS status,
  s.consignee_name       AS consignee,
  s.consignee_city       AS city,
  s.consignee_state      AS state,
  s.consignee_pincode    AS pincode,
  s.shipment_type        AS courier_type,
  s.payment_mode,
  COALESCE(s.cod_amount, 0) AS cod_amount,
  DATE(s.added_on)       AS placed_date,
  s.user_id              AS seller_id
FROM shipments s
JOIN shipments_status ss ON s.user_status = ss.status_id AND ss.status_for = 'User'
WHERE s.user_status IN (221, 222, 231, 232, 238)
  AND s.user_id IN (
      SELECT user_id FROM child_sales_poc_user_assign WHERE admin_id = 38
  )
  AND s.added_on >= DATE_SUB(NOW(), INTERVAL 30 DAY)
ORDER BY s.added_on DESC
LIMIT 5000
`;

// ── MAIN ─────────────────────────────────────────────────────
function runMIS() {
  try {
    Logger.log("▶ Starting Kunal MIS...");

    const data = fetchFromMetabase_();
    Logger.log(`  Fetched: ${data.rows.length} rows`);

    writeRawData_(data);
    writePivot_(data);
    sendEmail_(data);

    Logger.log("✅ Done!");
  } catch (err) {
    Logger.log("❌ " + err);
    throw err;
  }
}

// ── METABASE API ─────────────────────────────────────────────
function getToken_() {
  const props = PropertiesService.getScriptProperties();
  const user  = props.getProperty("METABASE_USER");
  const pass  = props.getProperty("METABASE_PASS");
  if (!user || !pass)
    throw new Error("Set METABASE_USER + METABASE_PASS in Script Properties");

  const res = UrlFetchApp.fetch(`${CONFIG.mbUrl}/api/session`, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ username: user, password: pass }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200)
    throw new Error("Metabase login failed: " + res.getContentText());
  return JSON.parse(res.getContentText()).id;
}

function fetchFromMetabase_() {
  const token = getToken_();

  const res = UrlFetchApp.fetch(`${CONFIG.mbUrl}/api/dataset`, {
    method: "post",
    contentType: "application/json",
    headers: { "X-Metabase-Session": token },
    payload: JSON.stringify({
      database: CONFIG.mbDb,
      type: "native",
      native: { query: SQL }
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200 && code !== 202)
    throw new Error("Query failed (" + code + "): " + res.getContentText().substring(0, 300));

  const json = JSON.parse(res.getContentText());
  if (json.status === "failed") throw new Error("Query error: " + json.error);
  return json.data;
}

// ── WRITE RAW DATA ───────────────────────────────────────────
function writeRawData_(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  let   sheet = ss.getSheetByName(CONFIG.dataTab);
  if (!sheet) sheet = ss.insertSheet(CONFIG.dataTab, 0);
  else        sheet.clearContents();

  const now  = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd MMM yyyy, hh:mm a 'IST'");
  const cols = data.cols.map(c => c.display_name || c.name);
  const rows = data.rows;

  // Title row
  sheet.getRange(1, 1).setValue("Kunal Chauhan MIS — Last Updated: " + now);
  sheet.getRange(1, 1, 1, cols.length).merge().setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");

  // Headers
  const hRow = sheet.getRange(2, 1, 1, cols.length);
  hRow.setValues([cols]).setFontWeight("bold").setBackground("#f0f2f5");

  // Data
  if (rows.length) sheet.getRange(3, 1, rows.length, cols.length).setValues(rows);

  // Format + freeze
  sheet.setFrozenRows(2);
  cols.forEach((_, i) => sheet.autoResizeColumn(i + 1));
  sheet.getRange(3, 1, Math.max(rows.length, 1), cols.length)
       .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  Logger.log(`  ✅ Raw data: ${rows.length} rows`);
}

// ── PIVOT SHEET ───────────────────────────────────────────────
function writePivot_(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  let   sheet = ss.getSheetByName(CONFIG.pivotTab);
  if (!sheet) sheet = ss.insertSheet(CONFIG.pivotTab);
  else        sheet.clearContents();

  const now   = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd MMM yyyy, hh:mm a 'IST'");
  const cols  = data.cols.map(c => c.name);
  const rows  = data.rows;

  const statusIdx  = cols.indexOf("status");
  const sellerIdx  = cols.indexOf("seller_id");
  const courierIdx = cols.indexOf("courier_type");
  const dateIdx    = cols.indexOf("placed_date");

  const STATUSES = ["Booked","Manifested","Not Picked","Out For Pickup","Pickup Pending"];

  let output = [
    [`Kunal Chauhan — MIS Pivot  |  ${now}`],
    [""],
  ];

  function pivotSection(idx, label) {
    if (idx === -1) return [];
    const groups = {};
    rows.forEach(r => {
      const key = (r[idx] || "(blank)").toString().substring(0, 30);
      const st  = r[statusIdx];
      if (!groups[key]) groups[key] = {};
      groups[key][st] = (groups[key][st] || 0) + 1;
    });

    const section = [[`── By ${label} ──`],
                     [label, ...STATUSES, "Total"]];
    
    Object.entries(groups)
      .map(([k, v]) => [k, v, STATUSES.reduce((s, x) => s + (v[x] || 0), 0)])
      .sort((a, b) => b[2] - a[2])
      .forEach(([k, v, tot]) =>
        section.push([k, ...STATUSES.map(s => v[s] || 0), tot])
      );

    section.push(["TOTAL", ...STATUSES.map(s => rows.filter(r => r[statusIdx] === s).length), rows.length]);
    section.push([""]);
    return section;
  }

  output = [...output, ...pivotSection(sellerIdx,  "Seller ID")];
  output = [...output, ...pivotSection(courierIdx, "Courier / Type")];
  output = [...output, ...pivotSection(dateIdx,    "Placed Date")];

  const maxCols = Math.max(...output.map(r => r.length));
  sheet.getRange(1, 1, output.length, maxCols).setValues(output);

  // Style header
  sheet.getRange(1, 1, 1, maxCols).merge().setBackground("#1a1a2e")
       .setFontColor("#ffffff").setFontWeight("bold");

  Logger.log("  ✅ Pivot created");
}

// ── EMAIL ────────────────────────────────────────────────────
function sendEmail_(data) {
  const cols     = data.cols.map(c => c.name);
  const rows     = data.rows;
  const statusIdx = cols.indexOf("status");
  const now      = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd MMM yyyy, hh:mm a 'IST'");

  const STATUSES = ["Booked","Manifested","Not Picked","Out For Pickup","Pickup Pending"];
  const COLORS   = {
    "Booked":"#3498db","Manifested":"#9b59b6","Not Picked":"#e74c3c",
    "Out For Pickup":"#e67e22","Pickup Pending":"#f1c40f"
  };

  const statusRows = STATUSES.map(s => {
    const n = rows.filter(r => r[statusIdx] === s).length;
    return `<tr>
      <td style="padding:9px 14px;border-bottom:1px solid #f0f0f0">${s}</td>
      <td style="padding:9px 14px;text-align:center">
        <span style="background:${COLORS[s]};color:#fff;padding:2px 14px;
              border-radius:12px;font-weight:700">${n}</span>
      </td></tr>`;
  }).join("");

  const html = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:#1a1a2e;padding:22px 28px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;color:#fff;font-size:20px">📦 ParcelX MIS Report</h2>
    <p style="margin:6px 0 0;color:#aaa;font-size:13px">Kunal Chauhan &nbsp;·&nbsp; ${now} &nbsp;·&nbsp; Last 30 Days</p>
  </div>
  <div style="background:#fff;padding:22px 28px;border:1px solid #ddd;border-top:none">
    <p style="margin:0 0 12px"><strong>Total orders:</strong> ${rows.length}</p>
    <table style="width:100%;border-collapse:collapse;background:#fafafa;border-radius:6px;overflow:hidden">
      <tr style="background:#f0f2f5">
        <th style="padding:9px 14px;text-align:left;font-size:13px">Status</th>
        <th style="padding:9px 14px;text-align:center;font-size:13px">Count</th>
      </tr>${statusRows}
    </table>
    <div style="text-align:center;margin-top:20px">
      <a href="https://docs.google.com/spreadsheets/d/${CONFIG.sheetId}"
         style="background:#34a853;color:#fff;padding:11px 28px;
                border-radius:6px;text-decoration:none;font-weight:700;font-size:14px">
        📊 Open Google Sheet
      </a>
    </div>
  </div>
  <div style="background:#f0f2f5;padding:12px 28px;text-align:center;border-radius:0 0 8px 8px">
    <p style="margin:0;font-size:11px;color:#999">Automated report · ParcelX MIS</p>
  </div>
</div>`;

  GmailApp.sendEmail(
    CONFIG.emailTo,
    `ParcelX MIS | Kunal Chauhan | ${now}`,
    `${rows.length} total orders. View HTML version for full report.`,
    { htmlBody: html, cc: CONFIG.emailCc, name: "ParcelX MIS" }
  );
  Logger.log(`  ✅ Email → ${CONFIG.emailTo}`);
}

// ── TRIGGER SETUP ─────────────────────────────────────────────
// Run this once to set the daily 9 AM trigger
function setupTrigger() {
  // Remove old triggers
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "runMIS")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 9 AM IST daily
  ScriptApp.newTrigger("runMIS")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone("Asia/Kolkata")
    .create();

  Logger.log("✅ Trigger set: runMIS runs daily at 9 AM IST");
}
