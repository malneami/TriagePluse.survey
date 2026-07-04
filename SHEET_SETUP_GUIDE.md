# TriagePluse Survey — Google Sheet Setup Guide

## Overview

Survey responses are automatically sent to a Google Sheet via a Google Apps Script Web App. The existing webhook URL is already configured in `index.html`. This guide explains how to set up a **new** Google Sheet or update the existing one.

---

## Sheet Structure

When set up, the Google Sheet will contain **three tabs**:

| Sheet Tab | Contents |
|---|---|
| **Summary** | Survey metadata and quick overview |
| **Part A — Triage Nurses** | Responses from triage nurses (13 questions) |
| **Part B — Physicians & Managers** | Responses from physicians, charge nurses, and managers (9 questions) |

---

## Part A — Column Headers (Triage Nurses)

| Column | Field | Description |
|---|---|---|
| A | Timestamp | ISO 8601 submission time |
| B | Part | Always "A — Triage Nurse" |
| C | Q1 — Role | `nurse` |
| D | Q2 — Years of ED Experience | `lt1` / `1-3` / `4-7` / `8+` |
| E | Q3 — Time to CTAS Assignment | `under2` / `2-5` / `5-10` / `10+` |
| F | Q4 — Hardest Triage Challenge | `volume` / `language` / `ctas` / `noref` / `docs` / `other` |
| G | Q5 — Times Unsure (Last Month) | `never` / `1-2` / `3-5` / `5+` |
| H | Q6 — Action When Unsure | `colleague` / `judgment` / `lookup` / `safer` / `movedon` |
| I | Q7 — Used Aid for CTAS Decision | `yes-reg` / `yes-occ` / `no` |
| J | Q8 — What Aid Used | Free text (optional) |
| K | Q9 — Difficult Patient Story ⭐ | Free text (required, ≥10 chars) |
| L | Q10 — Resources at Triage Station | Comma-separated: `paper`, `screen`, `colleague`, `phone`, `nothing` |
| M | Q11 — Resources Actually Used | Free text (optional) |
| N | Q12 — Would Want Decision Support (1–5) | Likert scale 1–5 |
| O | Q13 — Additional Comments | Free text (optional) |

---

## Part B — Column Headers (Physicians & Managers)

| Column | Field | Description |
|---|---|---|
| A | Timestamp | ISO 8601 submission time |
| B | Part | Always "B — Physician/Manager" |
| C | Q1 — Role | `receiver` |
| D | Q2 — Years of ED Experience | `lt1` / `1-3` / `4-7` / `8+` |
| E | QB3 — Frequency of Re-assessment | `always` / `often` / `sometimes` / `rarely` |
| F | QB4 — CTAS Mismatch Frequency | `yes-reg` / `yes-many` / `yes-once` / `no` |
| G | QB5 — CTAS Mismatch Story ⭐ | Free text (required, ≥10 chars) |
| H | QB6 — Action When Level Seems Wrong | `silent` / `discuss` / `escalate` / `accept` |
| I | QB7 — Most Common Downstream Effect | `delay` / `crowd` / `reassess` / `conflict` / `wrong-workup` / `minimal` |
| J | QB8 — Confidence with Better Support (1–5) | Likert scale 1–5 |
| K | QB9 — Additional Comments | Free text (optional) |

---

## Setup Instructions (New Sheet)

### Step 1 — Create a Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a new spreadsheet
2. Name it: **TriagePluse Survey Responses — HMG/QID/1397**

### Step 2 — Open Apps Script

1. In the spreadsheet, click **Extensions → Apps Script**
2. Delete all existing code in the editor
3. Copy the entire contents of `google-apps-script.js` and paste it

### Step 3 — Run Setup

1. In the Apps Script editor, select the function `setupSheets` from the dropdown
2. Click **Run** (▶)
3. Grant permissions when prompted

### Step 4 — Deploy as Web App

1. Click **Deploy → New deployment**
2. Click the gear icon ⚙ next to "Type" and select **Web app**
3. Set:
   - **Description:** TriagePluse Survey Webhook
   - **Execute as:** Me (your account)
   - **Who has access:** Anyone
4. Click **Deploy**
5. Copy the **Web app URL** (looks like `https://script.google.com/macros/s/...../exec`)

### Step 5 — Update the Survey Page

1. Open `index.html`
2. Find the line: `var WEBHOOK_URL = '...'`
3. Replace the URL with your new Web app URL
4. Commit and push to GitHub

---

## Existing Webhook

The survey already has a webhook configured:
```
https://script.google.com/macros/s/AKfycbw2OBpDRPftQFqh7DmT1jCuXIprZSIiF_OvTIobO3tWy8hD4-n-O-UdC68ev-cQfVamCQ/exec
```

If this webhook is active and connected to a Google Sheet, responses are already being collected. To verify, check the Google Sheet associated with this Apps Script deployment.

---

## Testing the Webhook

After deployment, test by visiting the Web App URL in a browser. You should see:
```json
{"status":"ok","message":"TriagePluse webhook is live"}
```

---

## Sharing the Sheet

To share the response sheet with your team:
1. Click **Share** in the top-right corner of Google Sheets
2. Add team members' email addresses
3. Set permission level:
   - **Viewer** — for read-only access
   - **Editor** — for full access

For a public read-only link: **Share → Change to anyone with the link → Viewer**

---

*HMG Takhassusi Hospital — Emergency Department, Riyadh*  
*Survey ID: HMG/QID/1397*
