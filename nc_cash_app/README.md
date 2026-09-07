# NC Cash Tracker

A small hosted app for tracking cash for NC KKBM LLC (New Caney Walk-On's).

- Upload a `.qbo` file — it verifies the account number, parses transactions,
  and inserts only new ones (deduped by the bank's FITID) into a SQLite database.
- Add manual **Projected** transactions (date, payee/payor, description,
  category, amount) right on the page.
- Download an always-current Excel export (Setup + Ledger tabs) any time.
- Transactions dated on/before the seeded "As-Of Date" are automatically
  skipped on upload — they're already baked into the starting balance, so
  importing them would double-count.

## Local run

```bash
pip install -r requirements.txt
python app.py
```

Visit `http://localhost:5000`.

## Deploying (same pattern as qbo_dedup on Render)

1. Push this folder to a **new** GitHub repo (e.g. `montgomerygunn-hub/nc-cash-tracker`).
2. In Render: **New → Web Service** → connect the repo.
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app`
3. **Important — persistent storage.** Render's default web service disk is
   *ephemeral*: it gets wiped on every redeploy or restart. Since the ledger
   lives in a SQLite file (`cash_tracker.db`) on disk, you need to attach a
   **Render Persistent Disk** (Render dashboard → your service → Disks → Add
   Disk) mounted at, e.g., `/var/data`, and then set an environment variable
   so the app writes the database there instead of next to the code:
   - Add env var `DB_DIR=/var/data` on the Render service (the app already
     reads this env var and will create/use `cash_tracker.db` there instead
     of next to the code)

   Without this step, every time you push a code update or Render restarts
   the service, **your transaction history would be wiped** and you'd be
   back to just the seeded starting balance. This is the one thing to not
   skip.

4. Once deployed, bookmark the Render URL (something like
   `nc-cash-tracker.onrender.com`) — that's your daily upload page.
5. Optional: point a subdomain of kkbm.net at it, same as planned for qbo_dedup.

## Files

- `app.py` — Flask app: routes, OFX/QBO parser, Excel export builder
- `templates/index.html` — the single page (upload, ledger table, add-projected form, settings)
- `requirements.txt` — Python dependencies
- `cash_tracker.db` — created automatically on first run (do not commit this to git)

## Notes

- The account is currently seeded as NC KKBM LLC, starting balance $50.00 as
  of 2026-09-07, account ending 4142. Change these under "Setup / Account
  Settings" on the page if anything changes.
- If a `.qbo` file's account number doesn't match the last-4 on file, the
  upload is rejected outright — nothing gets imported.
