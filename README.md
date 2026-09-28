# Account Guardian

**A daily health check for your Google Ads accounts.**

Every day it looks at your ad accounts and finds money that is being wasted
quietly — things a normal report never shows you:

- A negative keyword that is blocking one of your own good keywords
- An ad group whose only ad got disapproved, so it silently stopped running
- Ads paying for clicks to a page that is broken, out of stock, or shows "not found"
- A winning campaign starved of budget while a losing one has money to spare

Each problem is **priced in dollars**, given an **exact fix**, and put on one
ranked to-do list in an Excel file. It also **saves your account's change
history every day** (who changed what, old value → new value), because Google
only keeps 30 days of it — Guardian keeps it forever.

**It is read-only.** It never changes anything in your ad account. It only
reads, reports, and recommends.

The whole tool is one file: `guardian.py`.

👉 Want to see what the report looks like without setting anything up? Open
[Account_Guardian_SAMPLE.xlsx](Account_Guardian_SAMPLE.xlsx), or make a fresh
one yourself:

```powershell
pip install google-ads openpyxl requests
python guardian.py --demo
```

---

## Setup — the 4 credentials you need

To read your real accounts, Google requires 4 values. Here is exactly where
each one comes from. You only do this once.

> **Before you start you need:**
> - A Google Ads **Manager account (MCC)** — the account that sits above your
>   ad accounts. If you don't have one, create it free at
>   https://ads.google.com/home/tools/manager-accounts/
> - A Google account that can sign in to that manager account.

### 1. `GOOGLE_ADS_DEVELOPER_TOKEN`

This is Google's permission slip to use the Ads API at all.

1. Sign in to your **Manager account (MCC)** at https://ads.google.com
2. Click **Admin** (gear icon) → **API Center**
   (older UI: **Tools & Settings → Setup → API Center**)
3. Fill in the short form and accept the terms.
4. Your developer token appears on that page — a string like `AbCdEfGh1234567890`.
   Copy it.

New tokens start with **"Test account only"** access. Apply for **Basic
access** on the same page (a short form, usually approved in 1–3 days) —
you need Basic access to read real (non-test) accounts.

### 2. `GOOGLE_ADS_CLIENT_ID` and `GOOGLE_ADS_CLIENT_SECRET`

These identify *this tool* to Google, like an app registration. They come
from Google Cloud (free, no billing needed for this).

1. Go to https://console.cloud.google.com and sign in with the same Google
   account.
2. Top bar → project dropdown → **New Project** → name it anything
   (e.g. "Account Guardian") → **Create**.
3. Menu → **APIs & Services → Library** → search **"Google Ads API"** →
   click it → **Enable**.
4. Menu → **APIs & Services → OAuth consent screen**:
   - User type: **External** → Create
   - App name: anything, add your email where asked → Save through the steps
   - Under **Audience / Test users**, click **Add users** and add your own
     Google email. (This lets you use the app while it's in "Testing" mode —
     you never need to publish it.)
5. Menu → **APIs & Services → Credentials** → **+ Create Credentials** →
   **OAuth client ID**:
   - Application type: **Desktop app**
   - Name: anything → **Create**
6. A box shows your **Client ID** (ends in `.apps.googleusercontent.com`)
   and **Client secret**. Copy both.

### 3. `GOOGLE_ADS_REFRESH_TOKEN`

This is the long-lived key that says "this Google account gave the tool
permission to read its ads data." You generate it once by logging in.

The easiest way is Google's own helper script:

```powershell
pip install google-ads
# download Google's official helper script:
curl -o generate_user_credentials.py https://raw.githubusercontent.com/googleads/google-ads-python/main/examples/authentication/generate_user_credentials.py
python generate_user_credentials.py --client_id "YOUR_CLIENT_ID" --client_secret "YOUR_CLIENT_SECRET"
```

1. A browser window opens — sign in with the Google account that can access
   your Manager account.
2. Google shows a scary "app isn't verified" warning because your app is in
   Testing mode — click **Continue** (it's your own app).
3. Approve the permission.
4. The script prints your **refresh token** in the terminal. Copy it.

### 4. `GOOGLE_ADS_LOGIN_CUSTOMER_ID`

This is simply your **Manager account's ID number**.

1. Sign in to https://ads.google.com with the manager account.
2. Look at the **top-right corner** — you'll see a number like `123-456-7890`.
3. That's it. (Guardian accepts it with or without the dashes.)

### The account list: `ACCOUNTS`

Tell Guardian which ad accounts to audit, as `Name:customer_id`:

- Each ad account's customer ID is shown in the top-right corner when you
  open that account in https://ads.google.com (e.g. `111-111-1111`).
- The name is just your own label — call it anything.

```
ACCOUNTS=Store1:111-111-1111,Store2:222-222-2222
```

If one account sits under a **different** manager account, add that manager's
ID as a third field: `Store3:333-333-3333:999-999-9999`.

### Optional: `GEMINI_API_KEY`

Only if you want an AI-written executive summary paragraph in the report.
Get a free key at https://aistudio.google.com/apikey (click **Create API
key**). Skip it or run with `--no-llm` — the report is complete either way.

---

## Putting it together

```powershell
python guardian.py --init        # creates the .env file (if not already there)
notepad .env                     # paste in the values from the steps above
python guardian.py --list-accounts   # sanity check: shows your account list
python guardian.py --probe Store1    # test the API connection on one account
python guardian.py                   # full audit → reports\Account_Guardian_<date>.xlsx
```

Your `.env` should end up looking like:

```
GOOGLE_ADS_DEVELOPER_TOKEN=AbCdEfGh1234567890
GOOGLE_ADS_CLIENT_ID=1234567890-abc123.apps.googleusercontent.com
GOOGLE_ADS_CLIENT_SECRET=GOCSPX-xxxxxxxxxxxxxxxx
GOOGLE_ADS_REFRESH_TOKEN=1//0abc-xxxxxxxxxxxxxxxx
GOOGLE_ADS_LOGIN_CUSTOMER_ID=123-456-7890
ACCOUNTS=Store1:111-111-1111,Store2:222-222-2222
```

**Never share or commit `.env`** — it's already in `.gitignore`.

### Run it every day

The change-history archive only stays complete if it runs daily. Schedule it
once with Windows Task Scheduler:

```powershell
schtasks /Create /TN "AccountGuardian" /SC DAILY /ST 06:30 /TR "python C:\path\to\guardian.py"
```

---

## What's in the report

One Excel file per run in `reports\`:

| Sheet | What it tells you |
|---|---|
| Executive Summary | Each account's health score (0–100), trend, and the top 20 problems |
| Findings Worklist | Every problem: severity, dollar cost, the exact fix, and a "Done?" column |
| Change Log | Who changed what and when, old value → new value |
| Destination Health | Every landing page checked, with its verdict and the spend sent to it |
| Budget & Pacing | Every budget, next to Google's suggested budget |
| Data Coverage | Anything Guardian couldn't read, listed honestly as a gap — never fake zeros |
| How To Read | Every rule and threshold explained |

## Useful commands

| Command | What it does |
|---|---|
| `python guardian.py` | Audit every configured account |
| `python guardian.py --accounts Store1` | Audit only some accounts |
| `python guardian.py --skip-url-checks` | Faster run — skip live landing-page checks |
| `python guardian.py --window 14` | Use a 14-day metric window instead of 30 |
| `python guardian.py --no-llm` | Skip the AI summary paragraph |
| `python guardian.py --demo` | Sample report from made-up data (no API needed) |
| `python guardian.py --self-test` | Run the built-in tests (no API needed) |

## Files

```
guardian.py                   the entire tool
.env                          your credentials (never committed)
Account_Guardian_SAMPLE.xlsx  sample output from --demo
history/                      daily change-history + budget snapshots (gitignored)
reports/                      one Excel report per run (gitignored)
```

## Roadmap

| Phase | Scope |
|---|---|
| 1 (done) | The 4 audit modules, health score, Excel report, daily change archiving |
| 2 | Join each change to before/after performance ("did that change help or hurt?") |
| 3 | Microsoft Advertising coverage |

## If something goes wrong

- **"Missing Google Ads credentials in .env"** — one of the 4 values above is
  empty in `.env`.
- **`--probe` fails with an auth error** — usually the developer token is
  still "Test account only" (apply for Basic access, step 1), or the refresh
  token was made with a Google account that can't see the manager account.
- **A `PermissionError` saving the report** — the Excel file is open. Close
  it and run again.
- Anything the API refuses to return shows up on the **Data Coverage** sheet
  with the reason — the report never silently shows zeros.
