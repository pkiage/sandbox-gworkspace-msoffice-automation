# Email Attachment Indexer (Attachments from domain)

Gmail → Google Sheet indexer with Drive save, quick date ranges, and smart destination prompts

This add-on creates an **Attachments** menu in *this Google Sheet* and opens a **sidebar** where you can list Gmail attachments (and optionally save selected types to Drive). It now supports:

- **Date basis toggle:**  
  - **Received (message)** — only attachments from messages in your date window.  
  - **Threaded (conversation)** — if a thread is active in your window, include **all** attachments in that thread (even older ones).
- **Quick ranges:** **MTD**, **QTD**, **YTD**, **Since last run** (auto-fills dates).
- **Year + Quarter picker:** type a year and tick **Q1/Q2/Q3/Q4**, then **Apply** to set a range.
- **Destination prompts** when the active sheet isn’t blank (Append vs Create new tab) with **“Remember my choice next time.”**
- **Menu → Clear remembered choices** to reset stored decisions and sidebar defaults.

---

## Files

- `Code.js` — backend Apps Script (menu, sidebar handler, Gmail/Drive logic, destination prompts, quick ranges). In App Scripts js is gs.
- `Sidebar.html` — UI (date basis explanation, quick-range buttons, year+quarter picker, checkboxes, validation).

---

## Setup

1. Open your target Google Sheet → **Extensions → Apps Script**.
2. Create files **`Code.gs`** and **`Sidebar.html`** and paste the provided code.
3. **Save** and **reload** the sheet.
4. First run will ask for permissions.

---

## How to use

1. **Attachments → Open Sidebar**.
2. Pick **Date basis**:
   - **Received (message):** precise to your date window (recommended for reports).
   - **Threaded (conversation):** broader; includes all attachments from threads active in the window (good for audits).
3. Pick a **date window**:
   - Use **Quick ranges**: **MTD**, **QTD**, **YTD**, or **Since last run** (auto-fills).
   - Or choose **From date → now** *or* **Date range (inclusive)** and set **Start/End**.
   - Or use **Year & Quarters**: enter a year, check Q1–Q4, click **Apply** (fills Start/End).
4. (Optional) **Sender filter**: fill **either** a **domain** (`@xyz.com`) **or** a **specific email** (`name@xyz.com`). Leave blank for all.
5. (Optional) **Filters**:
   - **Include inline images** (unchecked by default).
   - **Min size (KB)** to skip tiny signature icons (e.g., 20).
6. (Optional) **Save to Drive**:
   - Tick **Save to Drive** and paste a **Drive folder URL**.
   - Choose which **file types** to **SAVE** (listing to the sheet is always logged; saving is restricted to checked types).
7. Click **Run**. Results write to the chosen sheet/tab.

---

## Destination behavior (where it writes)

- **Blank active sheet** → writes here (adds headers, formats, hides key column).
- **Non-blank with tool headers** → prompt: **Append here** or **Create a new tab** (can “remember” your choice).
- **Non-blank without headers** → prompt: **Create a new tab (recommended)** or **Append anyway** (can “remember”).
- **Very large/protected sheets** → safety fallback to **Create a new tab**.
- **Menu → Clear remembered choices** resets both the destination decisions and the sidebar defaults.

---

## Quick ranges (how they’re calculated)

- **MTD**: first day of the current month → now (From-date mode).  
- **QTD**: first day of the current calendar quarter (Jan/Apr/Jul/Oct) → now.  
- **YTD**: Jan 1 of the current year → now.  
- **Since last run**: uses the stored **last run date** (YYYY-MM-DD) → now.  
- **Year & Quarters**: choose year + any of **Q1 (Jan–Mar), Q2 (Apr–Jun), Q3 (Jul–Sep), Q4 (Oct–Dec)**; **Apply** fills an **inclusive** range spanning the selected quarters.

> **End date is inclusive**: the code uses `before:(End+1 day)` under the hood to capture the End day fully.

---

## Output columns

1. (Hidden) **Logged Key** — de-dupe token (`messageId#attachmentIndex#size`)  
2. **File Name**  
3. **File Type (MIME)**  
4. **Size (KB)**  
5. **Email Subject**  
6. **From**  
7. **To**  
8. **Date** (formatted in the spreadsheet timezone)  
9. **Message Link**  
10. **Thread Link**  
11. **Saved File URL** (if saving to Drive)

---

## Filters & saving

- **Listing vs Saving (hybrid)**: everything that passes **inline/min-size** is listed to the sheet; **saving** is further restricted to the **checked file types**.
- **File type groups** include PDF, Word, PPT, Excel, CSV/TSV, TXT, JSON, Images, Archives, Code, Other (catch-all).
- If no save types are checked, the run still logs a complete index; nothing is saved to Drive.

---

## De-duplication

- Prevents duplicates across runs on tool-managed tabs by checking the hidden **Logged Key** in Column A.

---

## “Remember my choice next time”

- Stores your **destination** decision **per user, per file** (no expiration).  
- Reset anytime via **Attachments → Clear remembered choices**.

---

## Permissions

- **Gmail** (read-only): list messages & attachments.
- **Drive** (write): only if you enable “Save to Drive”.
- **Sheets**: write rows, add/hide columns, freeze header row.

---

## Limits & tips

- Apps Script has execution time limits; for large mailboxes, use tighter ranges or filters.
- “Threaded” basis can multiply rows; “Received” is faster/smaller and better for month/quarter reports.
- Use **Min size (KB)** and keep **inline images** off unless needed.
- Keep a **dedicated tab** for clean headers + reliable de-dup.

---

## Privacy

All work happens inside your Google account; no external services are called.
