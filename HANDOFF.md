# HP Gas OTP Bot — Setup & Operation Handoff

Everything needed to run this bot on a **new device**. All code is already in
this repo; this file covers the environment, the input files (which are *not*
in the repo), and the operational settings that make it reliable.

---

## 1. What the bot does

For each consumer ID in a **targets** list, it looks up that account's login in a
**master** credentials file, logs into the My HPGas portal (you solve the
CAPTCHA), opens **View Cylinder Booking history**, and reads the latest
booking's **delivery OTP**. It writes the OTP to two output files, saving after
every account.

---

## 2. Clone

```bash
git clone https://github.com/aryantuntune/chromebot11.git
cd chromebot11
```

The customized code is already here (entry point: `run_bot.py`).

---

## 3. Python environment (Windows)

1. **Python 3.9+** must be usable. On Windows the Microsoft Store can shadow the
   real Python — if typing `python` opens the Store or says "not found", either
   use the `py` launcher (below) or add the real install
   (`...\Programs\Python\Python3XX` and its `\Scripts`) to the **front** of your
   user PATH.
2. Create the virtual environment and install dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m playwright install chrome
```

Chrome, Edge, Brave, and Opera all work — see browser selection below.

---

## 4. Input files (add these to `input\` — they are NOT in the repo)

Put **two** Excel files in `input\`. The bot auto-detects which is which.

| File | Role | Shape |
|------|------|-------|
| **Master** | The big credentials database | **≥ 5 columns**, no header. `A` = consumer number, `C` = email / login id, `E` = login password |
| **Targets** | The consumer IDs to process this run | **single column** (`A`) of consumer IDs |

- Master = the file with ≥ 5 columns. Targets = the most-recently-modified
  single-column file. If you have several lists, the newest one wins.
- Each target consumer ID is matched against column **A** of the master to pull
  its email (col **C**) and password (col **E**).

> Input/output spreadsheets are git-ignored — they hold real credentials and are
> never committed.

---

## 5. Run

```powershell
.\.venv\Scripts\python.exe run_bot.py
```

A browser window opens per account. **Solve the CAPTCHA** in it; the bot fills
the password, logs in, opens the booking history, captures the OTP, and moves
on. Stop any time (close the window or Ctrl+C) — progress is saved.

### Outputs (in `output\`, saved after every account)

- **`OTP_results.xlsx`** — two columns: **Consumer ID | OTP**.
- **`<master-name>.result.xlsx`** — a full copy of the master with the **OTP in a
  new last column** on each matched row.

### What the OTP column contains

| Booking state | Value |
|---------------|-------|
| Out for delivery | the **OTP** digits, e.g. `7149` (read from the `OutForDelivery (OTP: 7149)` cell on the booking page) |
| Already delivered | `Delivered` |
| Booked, not yet out for delivery | `In process` |
| Page didn't load (retryable) | `NO_BOOKING_TABLE` |

---

## 6. IMPORTANT — pacing settings (the key to reliability)

The portal **throttles rapid automated logins** and needs a moment **after login
to set up the session** before the heavy booking page will render. Pacing the
bot like a human cut load-failures from **~33% to ~4%** in testing. **Always run
with these:**

| Env var | Recommended | Meaning |
|---------|-------------|---------|
| `BOT_SETTLE_MS` | `5000` | pause after login before opening the booking page |
| `BOT_BETWEEN_MS` | `4000` | pause between accounts |
| `BOT_BROWSER` | `edge` | `chrome` (default), `edge`, `brave`, or `opera` — all work |

```powershell
$env:BOT_BROWSER="edge"; $env:BOT_SETTLE_MS="5000"; $env:BOT_BETWEEN_MS="4000"
.\.venv\Scripts\python.exe run_bot.py
```

Other controls: `BOT_MAX_ROWS=N` (process only the first N — good for a test
run), `BOT_DEBUG=1` (save screenshots + page dumps to `output\debug\` for
troubleshooting).

---

## 7. Resume & retry (built in)

- **Re-running resumes:** already-captured accounts are skipped; only failures
  (`NO_BOOKING_TABLE` / `ERROR`) and not-yet-done accounts are processed. So you
  can stop and restart freely.
- **To retry ONLY the failures:** make a single-column `.xlsx` of the failed
  consumer IDs, drop it in `input\` (it becomes the newest targets file), and
  run. Raise the delays for a retry pass (e.g. `BOT_SETTLE_MS=8000`,
  `BOT_BETWEEN_MS=6000`) since failures usually coincide with the portal being
  slow.

---

## 8. When the portal degrades

After many rapid logins (~70+) the portal can start failing most requests
(cumulative throttling). If failures spike, **pause 20–30 minutes and resume**,
or raise the delays. This is the **server**, not the bot — switching browsers
won't help (all load the site equally fast).

---

## 9. Fixes already baked into the code

- **No page reloads on slow loads.** Reloading this ASP.NET portal resubmitted
  the console form and re-popped the "OK" modal, stranding the run on the
  customer console. The bot now re-clicks the left-nav link instead — a clean
  forward navigation.
- **OTP location.** The OTP is not in the "Status" column (which may say "In
  process"); it's in the delivery cell further right, formatted
  `OutForDelivery (OTP: NNNN)`. The bot reads that cell and extracts the digits.
