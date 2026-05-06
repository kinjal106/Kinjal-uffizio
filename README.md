# Uffizio Design Team — KRA Dashboard

A live KRA (Key Result Areas) rating dashboard for the Uffizio design team.
Rate tasks per employee across 7 KPI dimensions and get automatic KRA scores.

## Features
- ✅ All 99 tasks pre-loaded from your Excel sheet
- ⭐ Rate every task on 7 KPIs (weighted scoring)
- 📊 Auto KRA score calculation per member (out of 100)
- 🔄 Live sync with Google Sheets
- 📥 Export KRA scores as CSV at month end
- 💾 Ratings saved in browser + Google Sheet

## KPI Dimensions & Weights
| # | KPI | Weight |
|---|-----|--------|
| 1 | Self Cross Verifying / Rework | 15% |
| 2 | Time Management | 15% |
| 3 | Creativity / New Trends / Adaptability | 20% |
| 4 | Design Quality & Satisfaction | 15% |
| 5 | Dedication of Task & Behavior | 10% |
| 6 | Initiative / Proactiveness / Branding / Tools | 10% |
| 7 | AI Utilization & Result Efficiency | 15% |

## Setup — Connect to Google Sheets

### Step 1 — Upload Excel to Google Drive
- Go to drive.google.com
- Upload your Excel file
- Right-click → Open with → Google Sheets

### Step 2 — Open Apps Script
In Google Sheets → Extensions → Apps Script

### Step 3 — Paste the script
- Delete everything in the editor
- Copy everything from `google-apps-script.gs` in this repo
- Paste it → Save (Ctrl+S)

### Step 4 — Deploy as Web App
- Click **Deploy** → New deployment
- Type: **Web app**
- Execute as: **Me**
- Who has access: **Anyone**
- Click Deploy → Copy the URL

### Step 5 — Connect the dashboard
- Open `index.html` in Chrome
- Go to **Google Sheets Setup** tab (left sidebar)
- Paste the URL in Step 4 input box
- Click **Sync Sheet** in the top bar

## Daily Usage
| Action | How |
|--------|-----|
| Pull new tasks from sheet | Click **Sync Sheet** |
| Rate a task | Rate Tasks tab → click task → rate 7 KPIs |
| View KRA scores | KRA Summary tab → select member |
| Month-end export | Click **Export KRA** → downloads CSV |
| Add a new task | Click **+ Add Task** |

## Hosting on GitHub Pages
1. Push this repo to GitHub
2. Go to repo Settings → Pages
3. Source: **main branch / root**
4. Your dashboard is live at `https://yourusername.github.io/uffizio-kra-dashboard`

## Team Members
Harshil · Hinesh · Mansi · Vishal · Mayur · Kinjal
