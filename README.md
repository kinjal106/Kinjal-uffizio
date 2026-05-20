# Uffizio KRA Dashboard — v3

STEP 1: Open google-apps-script.gs
        Replace "YOUR_SHEET_ID_HERE" on line 9 with your actual Sheet ID:
        1XZ32kMV32ASmgxXTF26sIPIoqRtB8ADOQ4j1__9sabM

STEP 2: Paste into Apps Script → Save → New deployment → Web app → Anyone → Deploy

STEP 3: Update SCRIPT_URL in index.html with your new URL

STEP 4: Upload index.html + data.js to GitHub

FIXES IN THIS VERSION:
✅ Sync no longer times out (removed slow per-row rating lookup)
✅ JSONP callback error fixed (no more "kra_cb is not defined")
✅ All date formats handled (DD-MMM-YY, YYYY-MM-DD, serial numbers)
✅ New tasks from sheet appear after sync
✅ Ratings save to KRA_Ratings sheet tab
