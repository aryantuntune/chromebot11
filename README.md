# HP Gas OTP Bot

An automation bot for the **My HPGas** portal
(`https://myhpgas.in/myHPGas/PortalLogin.aspx`) that retrieves the **delivery
OTP** for a list of consumer accounts.

You give it two Excel files:

1. A **master** credentials database (consumer no., name, email, password, …).
2. A **targets** list — just the consumer IDs you actually want to process.

For each target consumer ID the bot:

1. Looks the ID up in the master to get that account's **email** + **password**.
2. Opens the HP Gas login page in a fresh Chrome window and fills the email.
3. **Pauses so you solve the CAPTCHA** (the only manual step).
4. Auto-fills the password, logs in, and opens **"View Cylinder Booking
   history"**.
5. Reads the latest booking's delivery cell — when a cylinder is **out for
   delivery** it shows `OutForDelivery (OTP: 7149)` — and extracts the **OTP**.
6. Writes that OTP to **two** output files and **saves after every account**.

The browser runs **visibly** on purpose — the CAPTCHA can only be solved by a
human.

---

## What you need first

1. **Python 3.9+** — already set up in this project's `.venv`.
2. **Google Chrome** (default). Edge, Brave, and Opera also work — see
   [Choosing a browser](#choosing-a-browser).

---

## Input files (put both in the `input/` folder)

The bot auto-detects which file is which:

| File | What it is | Shape |
|------|------------|-------|
| **Master** | The big credentials database | **≥ 5 columns**, no header. `A`=consumer no, `C`=email/login id, `E`=password |
| **Targets** | The consumer IDs to process this run | **single column** (`A`) of consumer IDs |

If several of either kind are present, the **most recently modified** one wins.
Every target consumer ID is matched against column **A** of the master to pull
that account's email (col **C**) and password (col **E**).

> Input spreadsheets are **git-ignored** — they hold real credentials and are
> never committed.

---

## How to run

From the project root:

```powershell
.\.venv\Scripts\python.exe run_bot.py
```

or just double-click **`run.bat`** (Windows).

A Chrome window opens per account. Solve the CAPTCHA shown; the bot does the
rest. **Stop any time** (close the window or press Ctrl+C) — progress is saved.

### Resume & retry

Re-running **resumes**: accounts already captured are skipped, and only ones
that **failed to load** (the portal's booking page is occasionally slow) are
retried. So you can stop and restart freely until every OTP is collected.

---

## Where the results go

Both are written to `output/` and updated after **every** account:

1. **`OTP_results.xlsx`** — a clean two-column sheet: **Consumer ID | OTP**.
2. **`<master-name>.result.xlsx`** — a full copy of the master with the **OTP
   added in a brand-new last column**, on each matched account's row.

### What lands in the OTP column

| Booking state | Value written |
|---------------|---------------|
| Out for delivery | the **OTP** digits, e.g. `7149` |
| Already delivered | `Delivered` |
| Booked, not yet out for delivery | `In process` |
| Page didn't load (retryable) | `NO_BOOKING_TABLE` |

---

## Choosing a browser

The portal's booking page can be flaky regardless of browser. The bot retries
and reloads the page up to 3× per account (no extra CAPTCHA needed), which
fixes most blank loads. If you still want to try a different browser:

```powershell
$env:BOT_BROWSER = "edge"     # chrome (default) | edge | brave | opera
.\.venv\Scripts\python.exe run_bot.py
```

---

## Run controls (environment variables)

| Variable | Effect |
|----------|--------|
| `BOT_BROWSER` | `chrome` (default), `edge`, `brave`, or `opera` |
| `BOT_MAX_ROWS` | Process only the first N targets (handy for a test run) |
| `BOT_DEBUG` | Save a screenshot + page text/HTML per account to `output/debug/` |
| `BOT_SETTLE_MS` | Pause (ms) after login before opening the booking page (default `3000`). The portal needs a moment to set up the session before the heavy booking grid will render — **raise this (e.g. `5000`) if the booking page often fails to load.** |
| `BOT_BETWEEN_MS` | Pause (ms) between accounts (default `3000`) so logins don't hammer the portal back-to-back. |

> **Tip — slow down to load more reliably.** The portal throttles rapid
> automated traffic and is slow to render the booking page right after a fast
> login. Pacing the bot like a human (the `BOT_SETTLE_MS` / `BOT_BETWEEN_MS`
> pauses) dramatically cuts booking-page load failures.

Example test run of the first 3 accounts with debug dumps:

```powershell
$env:BOT_MAX_ROWS = "3"; $env:BOT_DEBUG = "1"
.\.venv\Scripts\python.exe run_bot.py
```

---

## Troubleshooting

**"No master / targets file found."** Make sure both spreadsheets are in
`input/`: one with ≥ 5 columns (master) and one single-column list (targets).

**Booking page won't load for some accounts.** That's the portal being slow.
They're saved as `NO_BOOKING_TABLE`; just re-run — resume retries only those.

**CAPTCHA not solved in time.** You get 5 minutes per account; solve a bit
quicker or the account is recorded as failed and retried next run.

---

## Project layout

```
.
├── run_bot.py        ← the bot you run (entry point)
├── run.bat / .sh     ← start the bot
├── input/            ← put master + targets .xlsx here (git-ignored)
├── output/           ← OTP_results.xlsx + master.result.xlsx land here
└── src/              ← config + excel helpers (template scaffolding)
```

---

## Privacy & scope

This bot signs in **only** to the HP Gas portal, using **only** the credentials
in your master file. It performs no other logins and sends nothing anywhere
except the HP Gas site itself.
