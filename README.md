# OSY / ISY Case Router

A small web app that takes your **Case Tracker** file, finds cases that are
`Status = APVD` with an `OSY` or `ISY` fund code, checks them against your
**Caseload** file for duplicates, and gives you a small file of just the new
rows — already arranged in the right columns — to paste into Caseload.

Everything runs **in your browser**. No file is ever uploaded to a server —
this app has no backend and no database. That matters here because your
data has client names and IDs in it.

---

## How the matching works

For each row in the Case Tracker's monthly tab(s):

| Condition | Column |
|---|---|
| `Status` = `APVD` | Column O |
| `Fund` contains `OSY` or `ISY` | Column M (also matches `A/OSY`, `A/ISY`, etc.) |

If both are true, the row is a match.

## How columns map over to the Caseload tab

Only these fields have a matching column in Caseload, so only these are
copied. Everything else on a new row (App ID, Exit Date, PWE Placed, etc.)
is left blank for you to fill in later:

| Case Tracker column | → | Caseload column |
|---|---|---|
| State ID | → | State ID |
| Last Name | → | Last Name |
| First Name | → | First Name |
| Office | → | Office |
| Case Manager | → | Case Manager from PWE, S & C list |
| Assign Date | → | Date Assigned |
| Sign Due | → | Signature Due |
| Sign Date | → | Sign Date |
| Status | → | Status |
| Notes | → | Notes |

If you ever rename a column in either file and the app can't find a match,
it'll show a warning banner telling you which field it couldn't map — open
`lib/caseRouter.js` and adjust the `targetMatch` line for that field.

## Why the app only reads the "Caseload" tab from HCGY_Caseload.xlsx

That file is ~15MB because two tabs ("Data" and "Data 2") have formatting
applied to over a million mostly-empty rows. The app deliberately reads
*only* the Caseload tab and never touches or rewrites the rest of the file
— that keeps things fast and means there's zero risk of it corrupting your
formatting on the big file. The output is a brand-new small file you paste
from, by hand, into your master file.

---

## Running it yourself first (optional, but recommended)

You don't have to do this — you can skip straight to deploying — but if
you want to try it on your own computer first:

1. Install [Node.js](https://nodejs.org) (the LTS version) if you don't
   have it.
2. Open a terminal in this folder and run:
   ```
   npm install
   npm run dev
   ```
3. Open http://localhost:3000 in your browser.

---

## Deploying: GitHub → Vercel

### Part 1 — Put the code on GitHub

1. Go to [github.com](https://github.com) and click the **+** in the top
   right → **New repository**. Name it something like `osy-isy-router`.
   Leave it empty (no README/gitignore) since you already have those.
   Click **Create repository**.
2. On the next page, GitHub shows you commands under
   **"…or push an existing repository from the command line"**. Open a
   terminal in this project folder and run, one line at a time:
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/osy-isy-router.git
   git push -u origin main
   ```
   (Replace the URL with the one GitHub showed you.) It'll ask you to sign
   in the first time — follow the prompts.
3. Refresh the GitHub page — you should see all your files there.

### Part 2 — Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) and sign in (you can sign in
   directly with your GitHub account — easiest option).
2. Click **Add New… → Project**.
3. Find `osy-isy-router` in the list of your GitHub repos and click
   **Import**.
4. Vercel auto-detects this as a Next.js app — you don't need to change
   any settings. Click **Deploy**.
5. Wait about a minute. When it finishes, click the preview thumbnail (or
   **Visit**) — that's your live URL, something like
   `osy-isy-router.vercel.app`. Bookmark it.

That's it — anyone with that link can use the tool. Every time you `git
push` a change to the `main` branch in the future, Vercel automatically
redeploys it, so updates are just: edit code → commit → push.

---

## Using the app

1. Upload your Case Tracker file (Step 1) and your Caseload file (Step 2).
2. The app defaults to the most recent monthly tab that actually has data
   in it (it skips empty future-month placeholder tabs). Click other
   month tabs to add them to the scan if you want to check more than one
   month at once.
3. Click **Find matching cases**.
4. Rows already found in your Caseload tab (matched by State ID) are
   flagged **"On file"** and start **unchecked** — review them and check
   the box if you genuinely want to add it again. Everything else starts
   checked.
5. Click **Download new rows**. Open that file, select the data rows
   (skip the header row), copy them, then in `HCGY_Caseload.xlsx` go to
   the Caseload tab, click the first empty row, and use
   **Paste Special → Values** (this avoids carrying over any formatting
   from the small helper file).

## Project structure

```
app/
  layout.js      -- fonts + page metadata
  page.js         -- the UI (upload cards, tab picker, results table)
  globals.css     -- color palette / design tokens
lib/
  caseRouter.js   -- all the file-reading, filtering, and mapping logic
```

If the spreadsheet layouts change in the future, `lib/caseRouter.js` is
almost always the only file you'll need to edit.
# osy-isy-router
