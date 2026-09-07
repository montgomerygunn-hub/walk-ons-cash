"""
Walk-On's Cash Tracker (multi-account)
Upload a .qbo file per account, it auto-detects the account by its bank
account number, parses + dedupes transactions into SQLite, and lets you
download an always-current Excel export per account.
"""
import io
import os
import re
import sqlite3
from datetime import datetime

from flask import (
    Flask, request, render_template, send_file, redirect, url_for, flash, session
)

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.environ.get("DB_DIR", APP_DIR)
DB_PATH = os.path.join(DB_DIR, "cash_tracker.db")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

CATEGORIES = ["Sales", "COGS", "Payroll", "Vendor/AP", "Construction/CapEx",
              "Transfer", "Financing", "Fees", "Other"]

SCHEMA_VERSION = "2"

# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _table_exists(conn, name):
    return conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def _column_exists(conn, table, column):
    cols = [r["name"] for r in conn.execute(f"PRAGMA table_info({table})")]
    return column in cols


def init_db():
    conn = get_db()

    conn.executescript("""
    CREATE TABLE IF NOT EXISTS meta (
        key TEXT PRIMARY KEY,
        value TEXT
    );
    """)
    version_row = conn.execute("SELECT value FROM meta WHERE key='schema_version'").fetchone()
    current_version = version_row["value"] if version_row else None

    if current_version == SCHEMA_VERSION:
        conn.close()
        return

    if not _table_exists(conn, "accounts"):
        conn.executescript("""
        CREATE TABLE accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            starting_balance REAL NOT NULL DEFAULT 0,
            as_of_date TEXT NOT NULL,
            last4 TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)

    legacy_settings = None
    if _table_exists(conn, "settings"):
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
        legacy_settings = {r["key"]: r["value"] for r in rows}

    if legacy_settings and conn.execute("SELECT COUNT(*) c FROM accounts").fetchone()["c"] == 0:
        cur = conn.execute(
            "INSERT INTO accounts (name, starting_balance, as_of_date, last4) "
            "VALUES (?, ?, ?, ?)",
            (
                legacy_settings.get("account_name", "Account 1"),
                float(legacy_settings.get("starting_balance", "0")),
                legacy_settings.get("as_of_date", datetime.now().date().isoformat()),
                legacy_settings.get("account_last4", ""),
            ),
        )
        legacy_account_id = cur.lastrowid
    else:
        legacy_account_id = None

    if not _table_exists(conn, "transactions"):
        conn.executescript("""
        CREATE TABLE transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            account_id INTEGER NOT NULL REFERENCES accounts(id),
            txn_date TEXT NOT NULL,
            txn_type TEXT NOT NULL,
            payee TEXT,
            description TEXT,
            category TEXT,
            amount REAL NOT NULL,
            bank_id TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(account_id, bank_id)
        );
        """)
    elif not _column_exists(conn, "transactions", "account_id"):
        old_rows = conn.execute(
            "SELECT txn_date, txn_type, payee, description, category, amount, "
            "bank_id, created_at FROM transactions"
        ).fetchall()
        conn.execute("ALTER TABLE transactions RENAME TO transactions_old")
        conn.executescript("""
        CREATE TABLE transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            account_id INTEGER NOT NULL REFERENCES accounts(id),
            txn_date TEXT NOT NULL,
            txn_type TEXT NOT NULL,
            payee TEXT,
            description TEXT,
            category TEXT,
            amount REAL NOT NULL,
            bank_id TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(account_id, bank_id)
        );
        """)
        target_account_id = legacy_account_id or conn.execute(
            "SELECT id FROM accounts ORDER BY id LIMIT 1"
        ).fetchone()["id"]
        for r in old_rows:
            conn.execute(
                "INSERT INTO transactions (account_id, txn_date, txn_type, payee, "
                "description, category, amount, bank_id, created_at) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (target_account_id, r["txn_date"], r["txn_type"], r["payee"],
                 r["description"], r["category"], r["amount"], r["bank_id"],
                 r["created_at"]),
            )
        conn.execute("DROP TABLE transactions_old")

    if _table_exists(conn, "settings"):
        conn.execute("DROP TABLE settings")

    if conn.execute("SELECT COUNT(*) c FROM accounts").fetchone()["c"] == 0:
        conn.execute(
            "INSERT INTO accounts (name, starting_balance, as_of_date, last4) "
            "VALUES (?, ?, ?, ?)",
            ("NC KKBM LLC", 50.00, "2026-09-07", "4142"),
        )

    conn.execute(
        "INSERT INTO meta (key, value) VALUES ('schema_version', ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (SCHEMA_VERSION,),
    )
    conn.commit()
    conn.close()


def get_accounts():
    conn = get_db()
    rows = conn.execute("SELECT * FROM accounts ORDER BY name ASC").fetchall()
    conn.close()
    return rows


def get_account(account_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM accounts WHERE id=?", (account_id,)).fetchone()
    conn.close()
    return row


def get_selected_account_id():
    accounts = get_accounts()
    if not accounts:
        return None
    requested = request.args.get("account_id", type=int) or session.get("account_id")
    valid_ids = [a["id"] for a in accounts]
    if requested in valid_ids:
        session["account_id"] = requested
        return requested
    session["account_id"] = accounts[0]["id"]
    return accounts[0]["id"]


def account_balance(account_id):
    conn = get_db()
    acct = conn.execute("SELECT starting_balance FROM accounts WHERE id=?", (account_id,)).fetchone()
    if not acct:
        conn.close()
        return 0.0
    total = conn.execute(
        "SELECT COALESCE(SUM(amount), 0) s FROM transactions WHERE account_id=?",
        (account_id,),
    ).fetchone()["s"]
    conn.close()
    return acct["starting_balance"] + total


# ---------------------------------------------------------------------------
# QBO / OFX parsing
# ---------------------------------------------------------------------------

def parse_ofx_date(raw):
    raw = raw.strip()[:8]
    return datetime.strptime(raw, "%Y%m%d").date().isoformat()


def parse_qbo(file_bytes):
    text = file_bytes.decode("utf-8", errors="replace")

    def find_one(tag, blob):
        m = re.search(rf"<{tag}>([^\r\n<]*)", blob)
        return m.group(1).strip() if m else None

    acctid = find_one("ACCTID", text)
    bankid = find_one("BANKID", text)

    transactions = []
    for block in re.findall(r"<STMTTRN>(.*?)</STMTTRN>", text, re.S):
        dtposted = find_one("DTPOSTED", block)
        trnamt = find_one("TRNAMT", block)
        fitid = find_one("FITID", block)
        name = find_one("NAME", block) or find_one("N", block)
        memo = find_one("MEMO", block)
        if not (dtposted and trnamt and fitid):
            continue
        transactions.append({
            "date": parse_ofx_date(dtposted),
            "amount": float(trnamt),
            "fitid": fitid.strip(),
            "name": (name or "").strip(),
            "memo": (memo or "").strip(),
        })

    return {"acctid": acctid, "bankid": bankid, "transactions": transactions}


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def build_excel(account_id):
    FONT = "Arial"
    HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
    HEADER_FONT = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    TITLE_FONT = Font(name=FONT, size=14, bold=True, color="1F4E78")
    LABEL_FONT = Font(name=FONT, size=10, bold=True)
    INPUT_FONT = Font(name=FONT, size=11, bold=True, color="0000FF")
    NOTE_FONT = Font(name=FONT, size=9, italic=True, color="808080")
    THIN = Side(style="thin", color="B7B7B7")
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    ACTUAL_FILL = PatternFill("solid", fgColor="E2EFDA")
    PROJECTED_FILL = PatternFill("solid", fgColor="FFF2CC")

    acct = get_account(account_id)
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM transactions WHERE account_id=? ORDER BY txn_date ASC, id ASC",
        (account_id,),
    ).fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    setup = wb.active
    setup.title = "Setup"
    setup.sheet_view.showGridLines = False

    setup["B2"] = f"{acct['name']} — Cash Tracker Setup"
    setup["B2"].font = TITLE_FONT
    setup.merge_cells("B2:D2")

    setup["B4"] = "Account Name"; setup["B4"].font = LABEL_FONT
    setup["C4"] = acct["name"]; setup["C4"].font = INPUT_FONT

    setup["B5"] = "Starting Balance"; setup["B5"].font = LABEL_FONT
    setup["C5"] = acct["starting_balance"]; setup["C5"].font = INPUT_FONT
    setup["C5"].number_format = "$#,##0.00"

    setup["B6"] = "As-Of Date"; setup["B6"].font = LABEL_FONT
    setup["C6"] = acct["as_of_date"]; setup["C6"].font = INPUT_FONT

    setup["B7"] = "Account Number (last 4)"; setup["B7"].font = LABEL_FONT
    setup["C7"] = acct["last4"]; setup["C7"].font = INPUT_FONT

    setup["B8"] = "Current Cash Balance"; setup["B8"].font = LABEL_FONT
    running = acct["starting_balance"]
    for r in rows:
        running += r["amount"]
    setup["C8"] = running
    setup["C8"].number_format = "$#,##0.00"
    setup["C8"].font = Font(name=FONT, size=12, bold=True)

    setup["B10"] = "Generated by the Walk-On's Cash Tracker app — this file is a live export, re-download after any change."
    setup["B10"].font = NOTE_FONT

    for col, width in [("A", 3), ("B", 30), ("C", 20), ("D", 20)]:
        setup.column_dimensions[col].width = width

    ledger = wb.create_sheet("Ledger")
    ledger.sheet_view.showGridLines = False
    headers = ["Date", "Type", "Payee/Payor", "Description", "Category",
               "Amount", "Running Balance", "Bank ID (dedup key)"]
    for i, h in enumerate(headers, start=1):
        c = ledger.cell(row=1, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ledger.freeze_panes = "A2"
    widths = [12, 12, 22, 30, 16, 14, 16, 26]
    for i, w in enumerate(widths, start=1):
        ledger.column_dimensions[get_column_letter(i)].width = w

    running = acct["starting_balance"]
    r_idx = 2
    for r in rows:
        running += r["amount"]
        vals = [
            datetime.fromisoformat(r["txn_date"]).date(),
            r["txn_type"], r["payee"], r["description"], r["category"],
            r["amount"], running, r["bank_id"] or "",
        ]
        fill = ACTUAL_FILL if r["txn_type"] == "Actual" else PROJECTED_FILL
        for c_idx, v in enumerate(vals, start=1):
            cell = ledger.cell(row=r_idx, column=c_idx, value=v)
            cell.border = BORDER
            cell.fill = fill
            cell.font = Font(name=FONT, size=10)
            if c_idx == 1:
                cell.number_format = "mm/dd/yyyy"
            if c_idx in (6, 7):
                cell.number_format = "$#,##0.00;($#,##0.00)"
        r_idx += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, acct["name"]


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    accounts = get_accounts()
    if not accounts:
        return redirect(url_for("manage_accounts"))

    account_id = get_selected_account_id()
    acct = get_account(account_id)

    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM transactions WHERE account_id=? ORDER BY txn_date DESC, id DESC LIMIT 100",
        (account_id,),
    ).fetchall()
    conn.close()

    return render_template(
        "index.html",
        rows=rows,
        balance=account_balance(account_id),
        accounts=accounts,
        current_account=acct,
        categories=CATEGORIES,
    )


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("qbo_file")
    if not f or f.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    data = parse_qbo(f.read())
    acctid = data.get("acctid") or ""
    found_last4 = acctid[-4:] if len(acctid) >= 4 else acctid

    accounts = get_accounts()
    matches = [a for a in accounts if a["last4"] and a["last4"] == found_last4]

    if not matches:
        flash(
            f"This file is for account ending {found_last4 or '(unknown)'}, which "
            f"isn't set up yet. Add it under Manage Accounts first, then re-upload.",
            "error",
        )
        return redirect(url_for("index"))

    acct = matches[0]
    as_of_date = acct["as_of_date"]

    conn = get_db()
    inserted, dup_skipped, historical_skipped = 0, 0, 0
    for t in data["transactions"]:
        if t["date"] <= as_of_date:
            historical_skipped += 1
            continue
        try:
            conn.execute(
                "INSERT INTO transactions (account_id, txn_date, txn_type, payee, "
                "description, category, amount, bank_id) "
                "VALUES (?, ?, 'Actual', ?, ?, ?, ?, ?)",
                (acct["id"], t["date"], t["name"], t["memo"], "", t["amount"], t["fitid"]),
            )
            inserted += 1
        except sqlite3.IntegrityError:
            dup_skipped += 1
    conn.commit()
    conn.close()

    session["account_id"] = acct["id"]
    msg = (f"Matched account '{acct['name']}' (ending {found_last4}). "
           f"Imported {inserted} new transaction(s), skipped {dup_skipped} duplicate(s)")
    if historical_skipped:
        msg += f", skipped {historical_skipped} dated on/before the {as_of_date} starting balance"
    msg += "."
    flash(msg, "success" if inserted else "info")
    return redirect(url_for("index", account_id=acct["id"]))


@app.route("/add", methods=["POST"])
def add_transaction():
    account_id = request.form.get("account_id", type=int)
    txn_date = request.form.get("txn_date")
    payee = request.form.get("payee", "").strip()
    description = request.form.get("description", "").strip()
    category = request.form.get("category", "").strip()
    amount = request.form.get("amount", "").strip()

    if not account_id or not get_account(account_id):
        flash("Select a valid account first.", "error")
        return redirect(url_for("index"))

    try:
        amount_val = float(amount)
        datetime.strptime(txn_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        flash("Please provide a valid date and numeric amount.", "error")
        return redirect(url_for("index", account_id=account_id))

    conn = get_db()
    conn.execute(
        "INSERT INTO transactions (account_id, txn_date, txn_type, payee, description, "
        "category, amount, bank_id) VALUES (?, ?, 'Projected', ?, ?, ?, ?, NULL)",
        (account_id, txn_date, payee, description, category, amount_val),
    )
    conn.commit()
    conn.close()
    flash("Projected transaction added.", "success")
    return redirect(url_for("index", account_id=account_id))


@app.route("/delete/<int:txn_id>", methods=["POST"])
def delete_transaction(txn_id):
    conn = get_db()
    row = conn.execute(
        "SELECT txn_type, account_id FROM transactions WHERE id=?", (txn_id,)
    ).fetchone()
    account_id = row["account_id"] if row else None
    if row and row["txn_type"] == "Projected":
        conn.execute("DELETE FROM transactions WHERE id=?", (txn_id,))
        conn.commit()
        flash("Projected transaction removed.", "success")
    else:
        flash("Only Projected rows can be deleted here.", "error")
    conn.close()
    return redirect(url_for("index", account_id=account_id))


@app.route("/download/<int:account_id>")
def download(account_id):
    if not get_account(account_id):
        flash("Account not found.", "error")
        return redirect(url_for("index"))
    buf, account_name = build_excel(account_id)
    filename = f"{account_name.replace(' ', '_')}_Cash_Tracker.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/accounts", methods=["GET"])
def manage_accounts():
    return render_template("accounts.html", accounts=get_accounts())


@app.route("/accounts/new", methods=["POST"])
def create_account():
    name = request.form.get("name", "").strip()
    starting_balance = request.form.get("starting_balance", "0").strip()
    as_of_date = request.form.get("as_of_date", "").strip()
    last4 = request.form.get("last4", "").strip()

    try:
        starting_balance_val = float(starting_balance)
        datetime.strptime(as_of_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        flash("Please provide a valid starting balance and as-of date.", "error")
        return redirect(url_for("manage_accounts"))

    if not name:
        flash("Account name is required.", "error")
        return redirect(url_for("manage_accounts"))

    conn = get_db()
    cur = conn.execute(
        "INSERT INTO accounts (name, starting_balance, as_of_date, last4) VALUES (?, ?, ?, ?)",
        (name, starting_balance_val, as_of_date, last4),
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    session["account_id"] = new_id
    flash(f"Account '{name}' added.", "success")
    return redirect(url_for("index", account_id=new_id))


@app.route("/accounts/<int:account_id>/update", methods=["POST"])
def update_account(account_id):
    if not get_account(account_id):
        flash("Account not found.", "error")
        return redirect(url_for("manage_accounts"))

    name = request.form.get("name", "").strip()
    starting_balance = request.form.get("starting_balance", "").strip()
    as_of_date = request.form.get("as_of_date", "").strip()
    last4 = request.form.get("last4", "").strip()

    try:
        starting_balance_val = float(starting_balance)
        datetime.strptime(as_of_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        flash("Please provide a valid starting balance and as-of date.", "error")
        return redirect(url_for("manage_accounts"))

    conn = get_db()
    conn.execute(
        "UPDATE accounts SET name=?, starting_balance=?, as_of_date=?, last4=? WHERE id=?",
        (name, starting_balance_val, as_of_date, last4, account_id),
    )
    conn.commit()
    conn.close()
    flash(f"'{name}' updated.", "success")
    return redirect(url_for("manage_accounts"))


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
