# Kunal MIS Report — Automation

Auto-downloads the ParcelX MIS report for **Kunal Chauhan** (last 30 days),
pushes it to Google Sheets, and emails the team — 4x daily.

| Time | IST | UTC cron |
|------|-----|----------|
| Morning   | 09:00 | 03:30 |
| Afternoon | 13:00 | 07:30 |
| Evening   | 16:00 | 10:30 |
| Night     | 20:00 | 14:30 |

## GitHub Secrets to add (Settings → Secrets → Actions)

| Secret | Value |
|--------|-------|
| `PARCELX_BASIC_USER` | `pax` |
| `PARCELX_BASIC_PASS` | `cloud@4w5` |
| `PARCELX_FORM_USER`  | `Sudhanshu` |
| `PARCELX_FORM_PASS`  | `@shudh@@$%` |
| `SMTP_HOST`          | `smtp.gmail.com` |
| `SMTP_PORT`          | `587` |
| `SMTP_USER`          | `sudhanshu@parcelx.in` |
| `SMTP_PASS`          | Gmail App Password |
| `GOOGLE_SA_JSON`     | Full contents of service account JSON |

## Flow
1. Selenium logs into panel.parcelx.in (Basic Auth + form login)
2. MIS Report → Non-Mandatory Fields → Sales POC → Kunal Chauhan
3. Date range: today − 30 days → today
4. Search → Export Excel → download .xlsx
5. Upload to Google Sheet (gid=2029444177)
6. Email to kunal.chauhan@parcelx.in (CC: sudhanshu, sparsh)
