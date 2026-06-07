# HP Gas Sign-in Bot

A small, friendly automation bot for the **My HPGas** portal
(`https://myhpgas.in/myHPGas/PortalLogin.aspx`).

You give it an Excel file with a list of accounts (consumer number, email,
password). For each row it:

1. Opens the HP Gas login page in a **fresh Incognito Chrome window**.
2. Types the account's **email** into the Mobile/E-Mail box.
3. **Pauses so you solve the CAPTCHA** (the only manual step).
4. Once you solve it, it automatically types the **password**, clicks **Login**,
   dismisses the confirmation pop-up, and opens **"View Cylinder Booking
   history"**.
5. Writes a result into the **`code`** column and **saves after every row**.

It runs the browser **visibly** on purpose — the CAPTCHA can only be solved by a
human, so you need to see the page.

Works on **Windows, macOS, and Linux**.

---

## What you need first

1. **Python 3.9 or newer** — download from <https://python.org/downloads>.
   - On **Windows**, tick **"Add Python to PATH"** in the installer.
2. **Google Chrome** installed normally — <https://google.com/chrome>.
3. Your **input Excel file** (`.xlsx`) — see [Input file format](#input-file-format).

---

## Quick start

### Windows (easiest)

1. **Download / clone** this folder.
2. Double-click **`setup.bat`** — wait for it to finish (it sets everything up).
3. Put your `.xlsx` file into the **`input`** folder.
4. Double-click **`run.bat`** — the bot starts.

### macOS / Linux

```bash
# from inside the project folder, once:
chmod +x setup.sh run.sh   # only needed the first time
./setup.sh

# put your .xlsx into the input/ folder, then start the bot:
./run.sh
```

That's it. `setup` is run **once**; after that you only ever use `run`.

---

## How a run works (and how to stop)

When you start the bot:

- It reads the newest `.xlsx` in the **`input`** folder.
- For each row, an **Incognito Chrome window** opens with the email already
  filled in, and the terminal prints:

  ```
  ACTION REQUIRED — signing in as: someone@example.com
    In the visible Chrome window:
      1. Type the CAPTCHA code shown on the page.
      2. Press <Enter> in the CAPTCHA box.
    The bot continues automatically — nothing to do in this terminal.
  ```

- **You** type the CAPTCHA in the Chrome window and press Enter there.
  The password field appears, and the bot takes over automatically:
  password → Login → OK → booking history → save → next row.

**To stop early:**

- **Windows:** close the black window, or press **Ctrl + C** in it.
- **macOS / Linux:** press **Ctrl + C** in the terminal.

Stopping is safe — every completed row is already saved.

---

## Input file format

Put one `.xlsx` file in the **`input`** folder. It must have a header row and
these columns (the same layout as the sample file used to build the bot):

| Column | Header     | Meaning                                  | Used by the bot         |
|--------|------------|------------------------------------------|-------------------------|
| **A**  | `conno`    | Consumer number                          | (reference only)        |
| **B**  | `id`       | Email / Mobile used to log in            | typed into the login    |
| **C**  | `password` | Account password                         | typed after the CAPTCHA |
| **D**  | `code`     | Left empty — the bot writes results here | **output**              |

- Data starts on **row 2** (row 1 is the header).
- Rows with an empty email are skipped automatically.
- If there are several `.xlsx` files, the **most recently modified** one is used.

> The bot **never modifies your input file.** It writes a copy to the `output`
> folder instead.

---

## Where the results go

After each row, the bot saves a copy of the workbook to:

```
output/<your-file-name>.result.xlsx
```

(e.g. `output/shlal-otp-0605.result.xlsx`). The `code` column is filled in for
every row that was processed. The same file is overwritten as it progresses, so
you always have the latest results.

### What gets written to the `code` column?

Right now the bot records the first table row it finds on the booking-history
page (falling back to `LOGGED_IN_OK` if it can't find one), and `ERROR: ...` if
a row failed. **If you want a specific value captured** (a booking number, a
status, an OTP, etc.), open `run_bot.py`, find the `_capture_result` function,
and point it at the exact element — or just ask and it can be adjusted.

---

## Settings you might change

All in `run_bot.py`, near the top:

| Setting                     | Default | What it does                                   |
|-----------------------------|---------|------------------------------------------------|
| `HEADLESS`                  | `False` | Must stay `False` so you can solve the CAPTCHA |
| `CAPTCHA_SOLVE_TIMEOUT_MS`  | 300000  | How long (ms) you get per row to solve it      |
| `BROWSER_CHANNEL`           | chrome  | Use real Chrome; set to `None` for bundled one |

The window opens in **Incognito** mode (no saved cookies/history), and each row
also uses a brand-new isolated session.

---

## Uploading this to a Git repo

This project is ready to push. **Important:** the `.gitignore` is set up so your
spreadsheets are **never uploaded** — `input/*.xlsx`, `output/*.xlsx`, and any
`.xlsx` in the project root are ignored, because they contain real emails and
passwords. Double-check before pushing that no credential file is staged.

```bash
git init
git add .
git commit -m "HP Gas sign-in bot"
git branch -M main
git remote add origin <YOUR_REPO_URL>
git push -u origin main
```

Anyone who clones it then just runs `setup` once and `run` to use it — their own
`.xlsx` goes in `input/` and is never committed.

---

## Troubleshooting

**"Python is not recognized" (Windows).**
Python isn't on PATH. Reinstall from python.org and tick **"Add Python to
PATH"**, then re-run `setup.bat`.

**Chrome not found / won't launch.**
Install Google Chrome normally. If it's still not found, run this once inside the
project (after setup):
`.venv\Scripts\python -m playwright install chrome` (Windows) or
`.venv/bin/python -m playwright install chrome` (macOS/Linux).

**`ensurepip is not available` when creating the environment (some Linux).**
Install the venv package once: `sudo apt install python3-venv` (Debian/Ubuntu),
then re-run `./setup.sh`.

**"CAPTCHA not solved in time."**
You have 5 minutes per row by default. Solve it a bit quicker, or increase
`CAPTCHA_SOLVE_TIMEOUT_MS` in `run_bot.py`.

**The bot stops on one row but keeps going.**
That's by design — a failed row is recorded as `ERROR: ...` in the `code` column
and the bot moves to the next one, so one bad account never halts the whole run.

---

## Project layout

```
.
├── run_bot.py        ← the bot you run (entry point)
├── setup.bat / .sh   ← one-time setup
├── run.bat / .sh     ← start the bot
├── requirements.txt  ← Python dependencies
├── input/            ← put your .xlsx here (ignored by git)
├── output/           ← results land here (ignored by git)
├── src/
│   ├── config.py     ← config dataclasses + validation
│   └── excel_io.py   ← reads rows / writes & saves results
└── tests/            ← unit tests for the helpers above
```

> `src/main.py`, `src/browser.py`, and `src/scraper.py` are the original generic
> template the bot grew out of. The working entry point for the HP Gas flow is
> **`run_bot.py`**.

---

## Privacy & scope

This bot signs in **only** to the HP Gas portal, using **only** the credentials
in your Excel file. It performs no other logins (no Google, email, or cloud
accounts) and sends nothing anywhere except the HP Gas site itself.
