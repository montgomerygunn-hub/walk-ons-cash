"""
NC KKBM LLC Cash Tracker
A small hosted app: upload a .qbo file, it parses + dedupes transactions into
a SQLite database, and you can download an always-current Excel export.
"""
import io
import os
import re
import sqlite3
from datetime import datetime, date

from flask import Flask, request, render_template, send_file, redirect, url_for, flash

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
# DB_DIR lets you point the database at a Render Persistent Disk mount
# (e.g. set env var DB_DIR=/var/data) so the ledger survives redeploys.
# Defaults to the app folder for local runs.
DB_DIR = os.environ.get("DB_DIR", APP_DIR)
DB_PATH = os.path.join(DB_DIR, "cash_tracker.db")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

CATEGORIES = ["Sales", "COGS", "Payroll", "Vendor/AP", "Construction/CapEx",
              "Transfer", "Financing", "Fees", "Other"]

# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );

    CREATE TABLE IF NOT EXISTS transactions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        txn_date TEXT NOT NULL,          -- YYYY-MM-DD
        txn_type TEXT NOT NULL,          -- 'Actual' or 'Projected'
        payee TEXT,
        description TEXT,
        category TEXT,
        amount REAL NOT NULL,
        bank_id TEXT UNIQUE,             -- FITID from .qbo, NULL for manual/projected entries
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    );
    """)
    # Seed default settings if empty
    defaults = {
        "account_name": "NC KKBM LLC",
        "starting_balance": "50.00",
        "as_of_date": "2026-09-07",
        "account_last4": "4142",
    }
    for k, v in defaults.items():
        conn.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (k, v)
        )
    conn.commit()
    conn.close()


def get_setting(key, default=None):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    conn = get_db()
    conn.execute(
        "INSERT INTO settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# QBO / OFX parsing
# ---------------------------------------------------------------------------

def parse_ofx_date(raw):
    """OFX dates look like 20260907171759 or 20260907 -- take just the date part."""
    raw = raw.strip()[:8]
    return datetime.strptime(raw, "%Y%m%d").date().isoformat()


def parse_qbo(file_bytes):
    """
    Very small, tolerant OFX/QBO parser. Returns:
        {
            "acctid": str or None,
            "bankid": str or None,
            "transactions": [ {date, amount, fitid, name, memo}, ... ]
        }
    OFX/QBO is SGML-like -- tags are not always closed, so we parse it
    with regexes rather than an XML parser.
    """
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

    ledgerbal = None
    m = re.search(r"<LEDGERBAL>.*?<BALAMT>([^\r\n<]*)", text, re.S)
    if m:
        try:
            ledgerbal = float(m.group(1).strip())
        except ValueError:
            ledgerbal = None

    return {"acctid": acctid, "bankid": bankid, "transactions": transactions, "ledgerbal": ledgerbal}


# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------

def build_excel():
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

    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM transactions ORDER BY txn_date ASC, id ASC"
    ).fetchall()
    account_name = get_setting("account_name")
    starting_balance = float(get_setting("starting_balance", "0"))
    as_of_date = get_setting("as_of_date")
    last4 = get_setting("account_last4")
    conn.close()

    wb = openpyxl.Workbook()
    setup = wb.active
    setup.title = "Setup"
    setup.sheet_view.showGridLines = False

    setup["B2"] = f"{account_name} — Cash Tracker Setup"
    setup["B2"].font = TITLE_FONT
    setup.merge_cells("B2:D2")

    setup["B4"] = "Account Name"; setup["B4"].font = LABEL_FONT
    setup["C4"] = account_name; setup["C4"].font = INPUT_FONT

    setup["B5"] = "Starting Balance"; setup["B5"].font = LABEL_FONT
    setup["C5"] = starting_balance; setup["C5"].font = INPUT_FONT
    setup["C5"].number_format = "$#,##0.00"

    setup["B6"] = "As-Of Date"; setup["B6"].font = LABEL_FONT
    setup["C6"] = as_of_date; setup["C6"].font = INPUT_FONT

    setup["B7"] = "Account Number (last 4)"; setup["B7"].font = LABEL_FONT
    setup["C7"] = last4; setup["C7"].font = INPUT_FONT

    setup["B8"] = "Current Cash Balance"; setup["B8"].font = LABEL_FONT
    running = starting_balance
    for r in rows:
        running += r["amount"]
    setup["C8"] = running
    setup["C8"].number_format = "$#,##0.00"
    setup["C8"].font = Font(name=FONT, size=12, bold=True)

    setup["B10"] = "Generated by NC Cash Tracker app — this file is a live export, re-download after any change."
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

    running = starting_balance
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
    return buf, account_name


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def index():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM transactions ORDER BY txn_date DESC, id DESC LIMIT 100"
    ).fetchall()
    starting_balance = float(get_setting("starting_balance", "0"))
    all_rows = conn.execute("SELECT amount FROM transactions").fetchall()
    balance = starting_balance + sum(r["amount"] for r in all_rows)
    conn.close()
    return render_template(
        "index.html",
        rows=rows,
        balance=balance,
        account_name=get_setting("account_name"),
        last4=get_setting("account_last4"),
        as_of_date=get_setting("as_of_date"),
        categories=CATEGORIES,
    )


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("qbo_file")
    if not f or f.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    data = parse_qbo(f.read())
    expected_last4 = get_setting("account_last4", "")
    acctid = data.get("acctid") or ""
    found_last4 = acctid[-4:] if len(acctid) >= 4 else acctid

    if expected_last4 and found_last4 and found_last4 != expected_last4:
        flash(
            f"Account mismatch! File is for account ending {found_last4}, "
            f"expected {expected_last4}. Nothing was imported.",
            "error",
        )
        return redirect(url_for("index"))

    as_of_date = get_setting("as_of_date", "1900-01-01")

    conn = get_db()
    inserted, dup_skipped, historical_skipped = 0, 0, 0
    for t in data["transactions"]:
        # Anything dated on or before the seeded as-of date is already baked
        # into the starting balance -- importing it would double-count.
        if t["date"] <= as_of_date:
            historical_skipped += 1
            continue
        try:
            conn.execute(
                "INSERT INTO transactions (txn_date, txn_type, payee, description, "
                "category, amount, bank_id) VALUES (?, 'Actual', ?, ?, ?, ?, ?)",
                (t["date"], t["name"], t["memo"], "", t["amount"], t["fitid"]),
            )
            inserted += 1
        except sqlite3.IntegrityError:
            # bank_id already exists -- duplicate, skip silently
            dup_skipped += 1
    conn.commit()
    conn.close()

    msg = f"Imported {inserted} new transaction(s), skipped {dup_skipped} duplicate(s)"
    if historical_skipped:
        msg += f", skipped {historical_skipped} dated on/before the {as_of_date} starting balance (already included in it)"
    msg += "."
    flash(msg, "success" if inserted else "info")
    return redirect(url_for("index"))


@app.route("/add", methods=["POST"])
def add_transaction():
    txn_date = request.form.get("txn_date")
    payee = request.form.get("payee", "").strip()
    description = request.form.get("description", "").strip()
    category = request.form.get("category", "").strip()
    amount = request.form.get("amount", "").strip()

    try:
        amount_val = float(amount)
        datetime.strptime(txn_date, "%Y-%m-%d")
    except (ValueError, TypeError):
        flash("Please provide a valid date and numeric amount.", "error")
        return redirect(url_for("index"))

    conn = get_db()
    conn.execute(
        "INSERT INTO transactions (txn_date, txn_type, payee, description, "
        "category, amount, bank_id) VALUES (?, 'Projected', ?, ?, ?, ?, NULL)",
        (txn_date, payee, description, category, amount_val),
    )
    conn.commit()
    conn.close()
    flash("Projected transaction added.", "success")
    return redirect(url_for("index"))


@app.route("/delete/<int:txn_id>", methods=["POST"])
def delete_transaction(txn_id):
    conn = get_db()
    row = conn.execute("SELECT txn_type FROM transactions WHERE id=?", (txn_id,)).fetchone()
    if row and row["txn_type"] == "Projected":
        conn.execute("DELETE FROM transactions WHERE id=?", (txn_id,))
        conn.commit()
        flash("Projected transaction removed.", "success")
    else:
        flash("Only Projected rows can be deleted here.", "error")
    conn.close()
    return redirect(url_for("index"))


@app.route("/download")
def download():
    buf, account_name = build_excel()
    filename = f"{account_name.replace(' ', '_')}_Cash_Tracker.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/settings", methods=["POST"])
def update_settings():
    for key in ("account_name", "starting_balance", "as_of_date", "account_last4"):
        val = request.form.get(key)
        if val is not None and val != "":
            set_setting(key, val)
    flash("Setup updated.", "success")
    return redirect(url_for("index"))


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
