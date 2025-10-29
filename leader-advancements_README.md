# Leader Advancements Self-Service Portal

Spreadsheet-bound Google Apps Script for leaders to request/log awards & knots, report training, and manage service-pin tracking. Includes dropdown sync from data tabs, a MID allow-list gate, receipts & DL notifications, optional training reminders, and quick action links.

> **Repo path:** `leader-portal/src/leader-advancements.gs`  
> **Status:** Sanitized (no real emails/URLs/IDs/tokens)  
> **License:** MIT (see repo root)

---

## What this does

- **Form intake** (linked Google Form) → writes to **Submissions** sheet.
- **Leader directory** upsert on every submission (MID keyed).
- **Allow-list gate** for MIDs (prevents non-leaders from using the portal).
- **Receipts** to submitter + **DL notifications** (distribution list).
- **Dropdown sync** from data tabs (Awards, Training, Service Pin years).
- **“Youth knot” guard** (prevents one-time youth awards duplicates).
- **Optional reminders** (training expirations; service-pin digests & nudges).
- **Quick action links** (Approve/Deny/Need Info/Mark Issued) via Web App.

---

## Prerequisites

- A **Google Sheet** with these tabs:
  - `Data_Settings` (Key, Value)
  - `Data_Awards` (`Code`, `Name`)
  - `TrainingCatalog` (`Code`, `Friendly Title`, optional `Popular`)
  - `ServicePinYears` (`Years`) — A2:A… numeric values
  - `Leaders` and `Submissions` will be created/extended automatically
  - (Optional) `MID_AllowList` is managed by sync (see below)
- A **Google Form** **linked** to the Sheet. The script expects exact question titles (list below).
- Script editor: **Runtime V8**, **Exception Logging: Stackdriver**.
- Scopes (in `appsscript.json`):
  ```json
  {
    "timeZone": "America/New_York",
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "oauthScopes": [
      "https://www.googleapis.com/auth/script.external_request",
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/documents",
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/gmail.send"
    ]
  }
  ```

---

## Configure `Data_Settings`

Create a tab `Data_Settings` with header `Key | Value`. Add:

| Key                    | Example Value                                   | Notes |
|---|---|---|
| `ENV`                  | `PROD`                                          | Used in email footers. |
| `DL_LEADER_ADV`        | `leaders@example.org`                           | DL for advancement notifications. |
| `DL_WEBMASTER`         | `webmaster@example.org`                         | Gets MID allow-list alerts. |
| `RATE_LIMIT_MAX_HITS`  | `3`                                             | Per MID, per window. |
| `RATE_LIMIT_WINDOW_MIN`| `5`                                             | Minutes. |
| `PIN_LOOKAHEAD_DAYS`   | `45`                                            | Digest window. |
| `PIN_NUDGE_DAYS`       | `21`                                            | Pre-anniversary nudge. |
| `REMINDER_DEFAULT_DAYS`| `30`                                            | (Reserved) |
| `SEND_HISTORY_ALWAYS`  | `N`                                             | `Y` to auto-email history after each submission. |
| `MID_MASTER_URL`       | `https://docs.google.com/spreadsheets/d/REDACTED_ID/edit` | Master roster used to build allow-list. |

---

## Expected Form questions (exact titles)

Identity:
- `Email Address` *(built-in)*
- `Member ID (MID)`
- `First Name`
- `Last Name`
- `Unit(s)`
- `Primary Position`
- `Registered Since`

Requests/Logging:
- `Requested Item (Request)` *(dropdown)*
- `Details/Justification (Request)`
- `Proof URL (optional) (Request)`
- `Requested Item (Log)` *(dropdown)*
- `Date Issued (Self-Reported)`
- `Issued By (Self-Reported)`
- `Where Issued (Self-Reported)`
- `Proof URL (optional) (Log)`

Service Pin:
- `Years of registered service` *(dropdown from ServicePinYears)*
- `I attest the tenure is accurate`

Training:
- `Course (search/choose)` *(dropdown from TrainingCatalog)*
- `Completion Date`
- `Proof URL`
- `Remind me before expiry`

Review/Extras:
- `I’ve reviewed my answers and they’re correct.`
- `Email me my history now`
- `Lookup My History`

> The script auto-detects which submission type applies based on which fields are filled.

---

## Sheet columns (created/used)

**Submissions** (header is auto-ensured):
```
Timestamp, Environment, MID, Email, Name, Unit(s),
Submission Type, Requested Item, Details/Justification, Proof URL,
Status, Reviewer, Review Date, Issued?, Issued By, Issued Date,
Cost, Inventory Source, Notes,
Recognition Type, Date Issued (Self-Reported), Issued By (Self-Reported),
Where Issued (Self-Reported), Validation Status, De-dup Check
```

**Leaders** (auto-ensured/upserted by MID):
```
MID, First Name, Last Name, Email, Unit(s), Primary Position,
Registered Since (YYYY-MM-DD), Years of Service, Recommended Service Pin,
Last Pin Issued (Years), Last Submission, Admin Notes
```

`MID_AllowList` (managed by sync):
```
MID, First, Last
```

---

## Install & first-run checklist

1) **Paste code** into Script Editor (`leader-advancements.gs`) in the linked Sheet.
2) **Save**, then **Run** any function to authorize.
3) **onOpen menu**: In the Sheet, a menu “**Scouter Awards**” appears:
   - **Sync dropdowns from Data tab** → populates Form dropdowns from:
     - `Data_Awards`, `TrainingCatalog`, `ServicePinYears`
   - **Sync MID allowlist (from master)** → pulls from `MID_MASTER_URL` into `MID_AllowList`.
   - **Install 2:00 AM MID sync** → sets a daily trigger for the allow-list.
   - **Issue & Email (selected row)** → marks a row issued and emails recipient.
   - **Email history by MID…** → prompts and sends history table.

4) **Link the Form** to the Sheet (File → Form → Create / Link).  
   Use exact titles above, set the request/log options as dropdown/list/checkbox as appropriate.

5) **Triggers** (recommended):
   - Time-based **nightly** (or hourly) for:
     - `nightlySyncDropdowns` (optional)
     - `nightlyTrainingReminders` (optional)
     - `nightlyPinNudges` (optional)
     - `monthlyPinDigest` (monthly time-driven trigger)
   - **On form submit**: If you’re not using the Form link (rare), add an installable trigger for `onFormSubmit`.

---

## Action Links (Approve/Deny/Need Info/Mark Issued)

Optional quick-actions via a **Web App**:

1) Set a secret property:
   - Script Editor → **Project Settings** → **Script Properties**
   - Add `ACTION_TOKEN = <random string>` (do not commit to Git)
2) Deploy:
   - **Deploy → New deployment → Web app**
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone with link** (or your org)
   - Copy the Web App URL.
3) In your email templates (DL notifications), include `actionLinks_(mid, item)`.  
   The script will generate links like:
   ```
   https://script.google.com/macros/s/DEPLOYMENT_ID/exec?a=Approve&mid=...&item=...&t=ACTION_TOKEN
   ```

> The sanitized code uses placeholders (e.g., `leaders@example.org`, `DEPLOYMENT_ID`). Do not commit real tokens/URLs.

---

## Data tabs: how to populate

- **Data_Awards**:
  ```
  Code | Name
  REL-YOUTH | Religious Emblem (Youth)
  KNOT-AOL-YOUTH | Arrow of Light (Youth)
  ... others ...
  ```
  > Youth awards in the “one-time knot guard” are keyed via their Code.

- **TrainingCatalog**:
  ```
  Code | Friendly Title | Popular
  YPT | Youth Protection Training | Y
  IOLS | Introduction to Outdoor Leader Skills | Y
  ... (leave Popular blank to exclude from dropdown) ...
  ```

- **ServicePinYears**:
  ```
  Years
  1
  2
  3
  ...
  ```

---

## Testing matrix

- **Allow-list gate**:
  - Submit with a valid format MID **not** in `MID_AllowList` → expect polite rejection to submitter + alert to `DL_WEBMASTER`.
  - Submit with MID in `MID_AllowList` → normal flow.

- **Submission types** (one at a time):
  - Request Award / Knot
  - Log Previously Issued Award / Knot
  - Request Service Pin
  - Report Training Completion
  - Lookup My History

- **Duplicate youth knots**: re-request a youth knot already on file → expect guard email + DL notice.

- **Rate-limit**: submit > `RATE_LIMIT_MAX_HITS` within window → expect throttle message.

- **Issue & Email**: select a row in **Submissions**, run menu action, confirm status/issued columns & emails.

---

## Troubleshooting

- **Menu not showing** → Reload Sheet; ensure `onOpen()` exists and saved.
- **“Form item not found”** → Titles must match exactly (see list above).
- **“No Google Form linked”** → Link the Form to the Sheet (File → Form).
- **Web App links return token error** → set `ACTION_TOKEN` in Script Properties and redeploy.
- **Emails blocked** → first run any MailApp-using function to authorize Gmail scope.

---

## Security & Privacy

- **No real emails/URLs/IDs/tokens** are committed in this repo.  
- Use `Data_Settings` + Script Properties for live values.  
- Keep `ACTION_TOKEN` secret; rotate if leaked.  
- Limit Web App access as narrowly as possible.

---

## Optional: Manage with `clasp`

```
cd leader-portal/src
# If this is a new script
clasp create --type sheets --title "Leader Advancements (sanitized)" --rootDir .
# Or clone an existing Script ID (do not commit .clasp.json)
# clasp clone <SCRIPT_ID> --rootDir .
clasp push
```

Add `.clasp.json` to `.gitignore` so your Script ID isn’t committed.

---

## Changelog

- **v2025-10-22.1**: MID allow-list sync + guard, dropdown sync, training reminders, pin digest/nudges, action links, receipts & DL notices, de-dup note for logged items, youth knot guard.
