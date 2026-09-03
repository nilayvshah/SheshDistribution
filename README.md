# SZVS Jain Samaj — Shesh Distribution

Member entry system for the Shesh distribution. A single-file web app
(no build step, no external JavaScript libraries) backed by a Google
Apps Script web app and a Google Sheet.

## What it does

Scan a member QR code — or type the member ID, or search by name — and
the app shows the member's name, address, mobile and **OTP**. The
operator reads the OTP aloud and asks the person collecting to confirm
it before releasing the item. An optional comment can be added, then
the entry is saved.

A member can only be entered once. If they have already collected, the
entry page blocks the submit and points to the Duplicate Slip page,
which records a replacement against the same member.

## Files

| Path | What it is |
|---|---|
| `index.html` | The whole front end — markup, styles and script in one file |
| `google appscript/code.gs` | Apps Script backend, bound to the spreadsheet |

## Google Sheet — `SheshDistribution`

Three tabs, named exactly as below. The names are the three constants
at the top of `code.gs`; rename a tab and you must edit those lines.

### `master_data`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| memberid | New code | Name | Address | pin code | mobile | Manav Rahat | **otp** |

`otp` is new. `Manav Rahat` is carried over but no longer used by the app.

### `user_master`

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| userid | password | name | range_start | range_end | active | Is_Admin | **print_required** | **dup_access** |

- `active` — `NO` disables the login
- `Is_Admin` — `YES` unlocks the All Entries and Summary reports and the edit modal
- `print_required` — `NO` skips the print dialog after saving. Blank or `YES` shows it.
- `dup_access` — `YES` shows the Duplicate Slip page. Blank or anything else hides it. This is an explicit grant and admin status does **not** override it, so set it per user including for admins.
- `range_start` / `range_end` — the entry ID block for this desk

### `transaction_data`

| A | B | C | D | E | F | G | H | I | J | K |
|---|---|---|---|---|---|---|---|---|---|---|
| Entry id | Member_id | Member Name | OTP | Entry Method | Userid | Created on ( Date & time ) | remark | comment | changedby | updatedon |

- `remark` is set by the system: `Duplicate Slip` or blank
- `comment` is the operator's own free text
- Start this tab with the header row only

## Setup

1. Create the spreadsheet `SheshDistribution` with the three tabs above.
2. From **inside the spreadsheet**, open Extensions → Apps Script. The
   script must be container-bound; a standalone project will fail,
   because the code calls `SpreadsheetApp.getActiveSpreadsheet()`.
3. Paste in `google appscript/code.gs` and save.
4. Deploy → New deployment → **Web app**, execute as **Me**, access
   **Anyone**. Copy the `/exec` URL.
5. Paste that URL into `SCRIPT_URL` near the top of the script block in
   `index.html`.
6. Commit and serve `index.html` over HTTPS (GitHub Pages works).

When you change `code.gs` later, use **Manage deployments → edit → New
version**. Creating a *new* deployment issues a different URL and you
would have to update `index.html` again.

## Notes on how it works

**Entry IDs are allocated by the server**, inside a `LockService` block,
at the moment the row is written. The browser never predicts an ID. In
the previous version the browser held its own counter, so the same
login open on two devices produced colliding IDs.

**Duplicates are checked twice** — once when the code is scanned, so the
operator sees it immediately, and again on the server before the row is
appended. The second check is what actually enforces the rule.

**Scanner vs typing** is detected from keystroke timing: gaps under 80ms
mean a hardware scanner, which sets the entry method to QR. The operator
can override with the toggle.

**Printing** uses Web Bluetooth to an ESC/POS thermal printer whose name
starts with `HQ300`, on 58mm paper. Chrome only; Safari and Firefox do
not implement Web Bluetooth.

**Camera scanning** uses the browser's native `BarcodeDetector`. It works
on Android Chrome and is unavailable on iOS, where the hardware scanner
or manual entry should be used instead.

## Requirements

- Chrome on Android for scanning and printing
- The page must be served over HTTPS — camera and Bluetooth are blocked
  on plain HTTP
