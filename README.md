# Kunal MIS — Google Apps Script (Recommended)

Pulls **live shipment data** directly from Metabase API → writes to Google Sheet → emails Kunal.

**No ParcelX panel login. No OTP. No browser. No server.**

---

## How it works
```
Google Apps Script
  → Metabase API (mb.parcelx.in)
    → Live shipments table (last 30 days, Kunal's sellers)
      → Google Sheet (raw data + pivot)
        → Gmail (kunal.chauhan@parcelx.in + CC)
```

## Setup (5 minutes)

### 1. Open Apps Script
Open your Google Sheet → **Extensions → Apps Script**

### 2. Paste the code
Delete everything → paste the full contents of `KunalMIS.gs`

### 3. Set credentials
⚙️ Project Settings → Script Properties → Add:

| Property | Value |
|---|---|
| `METABASE_USER` | your Metabase login email |
| `METABASE_PASS` | your Metabase password |

### 4. Test
Select `runMIS` → click ▶ Run → accept permissions

### 5. Schedule
Select `setupTrigger` → click ▶ Run

This sets **daily 9 AM IST** trigger. To add more times (1pm, 4pm, 8pm), run `setupTrigger` — it creates all 4.

---

## What it produces

| Output | Where |
|---|---|
| Raw data | Sheet tab "MIS Data" |
| Pivot (seller/courier/date) | Sheet tab "Pivot" |
| Email | kunal.chauhan@parcelx.in (CC: sudhanshu, sparsh) |

## Statuses tracked
`Booked · Manifested · Not Picked · Pickup Pending · Out For Pickup`
