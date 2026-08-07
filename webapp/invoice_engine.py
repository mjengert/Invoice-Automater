#!/usr/bin/env python3
"""
invoice_engine.py — Funky's Electrical

Shared business logic for the Invoice/Estimate tool: database access, fuzzy
customer search, invoice-number sequencing, Excel invoice generation,
PDF generation, printing (via Excel/AppleScript), and emailing (via
Mail.app/AppleScript).

This module is a direct extraction of the logic that already lives in
gui_automater.py (the original desktop app) — nothing about *how* an
invoice is built, numbered, or saved has changed. It's pulled out here so
the same, already-working logic can be driven by the new web dashboard
(app.py) instead of only the customtkinter desktop UI. gui_automater.py
itself is untouched and still works standalone.

Phase 2 adds estimates/quotes: a shared document-number sequence across
invoice/estimate/quote, prices that may be a firm number or free text
(a price range, e.g. "$600.00 - $700.00" — matching how these are typed
by hand today), and "convert to invoice" which keeps the same document
number rather than issuing a new one, exactly as done by hand today
(e.g. "Erin Barnes Estimate #3487" -> "Erin Barnes #3487" nine days
later, same number). Plain invoices with numeric-only prices behave
byte-for-byte the same as before.
"""

import calendar
import json
import os
import sqlite3
import subprocess
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter, range_boundaries
from rapidfuzz import process, fuzz
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)

# ─── Paths ──────────────────────────────────────────────────────────────────
# webapp/ lives one level below the Invoice-Automater root, alongside
# gui_automater.py, invoices.db, and the Invoice_template.xlsx file.

WEBAPP_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR   = os.path.dirname(WEBAPP_DIR)
DB_PATH    = os.path.join(BASE_DIR, "invoices.db")
INV_DIR    = os.path.join(BASE_DIR, "Invoice Excel Docs")
INV_PDF_DIR = os.path.join(BASE_DIR, "Invoice Excel PDFs")
TEMPLATE   = os.path.join(BASE_DIR, "Invoice_template.xlsx")
EMAIL_TO   = "earlybird0648@gmail.com"

COMPANY       = "Funky's Electrical"
COMPANY_ADDR  = "5924 Ashton Woods Cir., Milton, FL 32570"
COMPANY_PHONE = "(850) 207-1800"
SALESPERSON   = "Gary Funkhouser"
COMPANY_EMAIL = "funkyselectrical@gmail.com"
LICENSE_NUM   = "EC13005709"
COMPANY_WEBSITE = "www.funkyselectrical.com"

MONTHLY_DIR = os.path.join(BASE_DIR, "Monthly Logs")
PAYROLL_DIR = os.path.join(BASE_DIR, "Payroll")


# ─── Database ───────────────────────────────────────────────────────────────

def open_db():
    # Same base schema as gui_automater.py's open_db() — same file, same
    # customers/invoices tables, so the desktop app keeps working untouched.
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY, name TEXT UNIQUE, address TEXT, phone TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS invoices(
        id INTEGER PRIMARY KEY, number INTEGER UNIQUE, customer TEXT,
        job TEXT, total REAL, day TEXT)""")
    conn.commit()

    # Additive migration for Phase 2 (Estimates/Quotes). Existing rows all
    # default to doc_type='invoice', so old data — and the untouched desktop
    # app, which never references these columns — are unaffected.
    cols = [r[1] for r in conn.execute("PRAGMA table_info(invoices)").fetchall()]
    if "doc_type" not in cols:
        conn.execute("ALTER TABLE invoices ADD COLUMN doc_type TEXT DEFAULT 'invoice'")
        conn.execute("UPDATE invoices SET doc_type='invoice' WHERE doc_type IS NULL")
    if "amount_display" not in cols:
        conn.execute("ALTER TABLE invoices ADD COLUMN amount_display TEXT")
    if "line_items_json" not in cols:
        conn.execute("ALTER TABLE invoices ADD COLUMN line_items_json TEXT")

    # Additive migration for Phase 4 (Monthly Log). paid/paid_date are new,
    # nullable/defaulted columns on the existing invoices table — the
    # desktop app never references them, so it's unaffected.
    if "paid" not in cols:
        conn.execute("ALTER TABLE invoices ADD COLUMN paid INTEGER DEFAULT 0")
    if "paid_date" not in cols:
        conn.execute("ALTER TABLE invoices ADD COLUMN paid_date TEXT")
    conn.commit()

    # One-time backfill: every invoice that predates paid/unpaid tracking
    # actually being used should already read as paid (MJ's request); only
    # invoices added from here on should start unpaid. This is gated on a
    # persistent settings flag rather than on "paid was just added above" —
    # on the real shop Mac, paid/paid_date turned out to already exist from
    # an earlier deploy (added, but never backfilled, so every row sat at
    # paid=0/paid_date=NULL), so column-existence alone was not a reliable
    # "did this already run" signal. The flag makes this run exactly once
    # per database file no matter how far back its schema history goes.
    # settings is created here (in addition to further down) because this
    # check needs it before the rest of open_db() gets there.
    conn.execute("""CREATE TABLE IF NOT EXISTS settings(
        key TEXT PRIMARY KEY, value TEXT)""")
    already_backfilled = conn.execute(
        "SELECT 1 FROM settings WHERE key='paid_backfill_v1'").fetchone()
    if not already_backfilled:
        # Only rows paid/unpaid tracking never actually touched — a real
        # "mark unpaid" always sets paid_date back to NULL too, but so does
        # an invoice that's simply never been acted on, so restricting to
        # paid_date IS NULL here avoids re-flipping anything a person
        # deliberately marked unpaid after the feature was already live.
        conn.execute(
            "UPDATE invoices SET paid=1, paid_date=day "
            "WHERE (paid=0 OR paid IS NULL) AND paid_date IS NULL")
        conn.execute("""INSERT INTO settings(key, value) VALUES('paid_backfill_v1', '1')
            ON CONFLICT(key) DO UPDATE SET value=excluded.value""")
    conn.commit()

    # New table (Phase 4) — one row per Monthly Log line, either created
    # automatically when an invoice is marked paid (source='invoice', tied
    # back via invoice_number) or added by hand (source='manual'). A brand
    # new table can't collide with anything the desktop app uses.
    conn.execute("""CREATE TABLE IF NOT EXISTS monthly_entries(
        id INTEGER PRIMARY KEY,
        month TEXT NOT NULL,
        entry_date TEXT NOT NULL,
        customer TEXT NOT NULL,
        category TEXT NOT NULL,
        amount REAL NOT NULL,
        source TEXT NOT NULL DEFAULT 'manual',
        invoice_number INTEGER)""")
    conn.commit()

    # New tables (Phase 5 — Payroll Log). workers holds each worker's fixed
    # Workers' Comp classification (state/class code/description/net rate) —
    # stored once, auto-filled on the report, per the outline. payroll_entries
    # is the simple Date/Worker/Gross log. settings is a tiny key/value table
    # for report-level values like the Policy Number.
    conn.execute("""CREATE TABLE IF NOT EXISTS workers(
        id INTEGER PRIMARY KEY,
        name TEXT UNIQUE NOT NULL,
        state TEXT,
        class_code TEXT,
        code_description TEXT,
        net_rate REAL)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS payroll_entries(
        id INTEGER PRIMARY KEY,
        month TEXT NOT NULL,
        entry_date TEXT NOT NULL,
        worker TEXT NOT NULL,
        gross REAL NOT NULL,
        class_code TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS settings(
        key TEXT PRIMARY KEY,
        value TEXT)""")
    conn.commit()

    # Additive migration — payroll's Workers' Comp class code moved from
    # being fixed per-worker to being picked per payroll entry (there are
    # only ever the two codes below, and which one applies depends on what
    # kind of job the wages were for, not who did it).
    pe_cols = [r[1] for r in conn.execute("PRAGMA table_info(payroll_entries)").fetchall()]
    if "class_code" not in pe_cols:
        conn.execute("ALTER TABLE payroll_entries ADD COLUMN class_code TEXT")
        conn.commit()

    return conn


DOC_LABELS = {"invoice": "INVOICE", "estimate": "ESTIMATE", "quote": "QUOTE"}


def parse_amount(raw):
    """A line-item price may be a firm number, or (for estimates/quotes)
    free text such as a price range — that's exactly how these are typed
    into the real paperwork today (e.g. "$600.00 - $700.00"). Returns
    (numeric_value_or_None, display_text_or_None) — exactly one is set.
    """
    if raw is None or raw == "":
        return 0.0, None
    if isinstance(raw, (int, float)):
        return float(raw), None
    s = str(raw).strip()
    if not s:
        return 0.0, None
    cleaned = s.replace("$", "").replace(",", "").strip()
    try:
        return float(cleaned), None
    except ValueError:
        return None, s


def all_customers(conn):
    rows = conn.execute(
        "SELECT name, address, phone FROM customers ORDER BY name").fetchall()
    return [{"name": r[0], "address": r[1], "phone": r[2]} for r in rows]


def get_customer(conn, name):
    row = conn.execute(
        "SELECT name, address, phone FROM customers WHERE name = ?", (name,)).fetchone()
    return {"name": row[0], "address": row[1], "phone": row[2]} if row else None


def create_customer(conn, name, address, phone):
    """Add a customer on their own, with no invoice involved (Phase 3).
    Raises ValueError if a customer with this exact name already exists —
    callers should check find_similar_customers() first for a softer
    near-duplicate warning."""
    name = (name or "").strip()
    if not name:
        raise ValueError("Customer name is required.")
    if get_customer(conn, name):
        raise ValueError(f'A customer named "{name}" already exists.')
    conn.execute("INSERT INTO customers(name, address, phone) VALUES(?,?,?)",
                 (name, normalize_address(address or "") or None, phone or None))
    conn.commit()


def update_customer_contact(conn, name, address, phone):
    """Update address/phone only — the name is locked once a customer is
    created, since invoices.customer is a plain text field (not a foreign
    key) and a rename wouldn't cascade to past invoices."""
    cur = conn.execute("UPDATE customers SET address=?, phone=? WHERE name=?",
                        (normalize_address(address or "") or None, phone or None, name))
    conn.commit()
    if cur.rowcount == 0:
        raise ValueError(f'No customer named "{name}" found.')


def search_customers(conn, query=None, limit=300):
    """Plain substring search across name/address/phone, for the Customers
    list page (separate from fuzzy_search(), which is tuned for the
    invoice wizard's typo-tolerant autocomplete)."""
    if query:
        like = f"%{query}%"
        rows = conn.execute(
            """SELECT name, address, phone FROM customers
               WHERE name LIKE ? OR address LIKE ? OR phone LIKE ?
               ORDER BY name LIMIT ?""",
            (like, like, like, limit)).fetchall()
    else:
        rows = conn.execute(
            "SELECT name, address, phone FROM customers ORDER BY name LIMIT ?",
            (limit,)).fetchall()
    return [{"name": r[0], "address": r[1], "phone": r[2]} for r in rows]


def find_similar_customers(conn, name, threshold=80, limit=3):
    """Soft near-duplicate check used when adding a new customer — e.g.
    catches "Gary Sizemore" vs. "Gary Sizemoore" before two records for the
    same person pile up. Excludes an exact name match (that's a hard stop,
    handled by create_customer's ValueError, not a soft warning)."""
    name = (name or "").strip()
    if not name:
        return []
    customers = all_customers(conn)
    hits = fuzzy_search(customers, name, threshold=threshold, limit=limit + 1)
    return [c for c in hits if c["name"].lower() != name.lower()][:limit]


def customer_stats(conn, name):
    """Per-customer summary for the history page: total invoiced (real
    invoices only) and how many open estimates/quotes they have."""
    row = conn.execute(
        """SELECT COALESCE(SUM(total),0), COUNT(*) FROM invoices
           WHERE customer=? AND doc_type='invoice'""", (name,)).fetchone()
    total_invoiced, invoice_count = row[0], row[1]
    open_estimates = conn.execute(
        """SELECT COUNT(*) FROM invoices
           WHERE customer=? AND doc_type IN ('estimate','quote')""", (name,)).fetchone()[0]
    return {
        "total_invoiced": total_invoiced,
        "invoice_count": invoice_count,
        "open_estimates": open_estimates,
    }


def next_invoice_number(conn):
    try:
        wb = openpyxl.load_workbook(TEMPLATE, data_only=True)
        val = wb.active["B1"].value or ""
        wb.close()
        for j, ch in enumerate(val):
            if ch.isdigit():
                return int(val[j:]) + 1
    except Exception:
        pass
    row = conn.execute("SELECT MAX(number) FROM invoices").fetchone()
    return (row[0] or 3668) + 1


def fuzzy_search(customers, query, threshold=55, limit=8):
    if not query.strip():
        return customers[:12]
    names = [c["name"] for c in customers]
    hits = process.extract(query, names, scorer=fuzz.WRatio, limit=limit)
    result = []
    for name, score, _ in hits:
        if score >= threshold:
            c = next((x for x in customers if x["name"] == name), None)
            if c:
                result.append(c)
    return result


def normalize_address(raw):
    """Ensure address is two lines: 'Street\\nCity, ST Zip'.

    Handles addresses already containing a newline, or a plain comma-separated
    string in any of these formats:
        4431 Winners Gait Cir, Pace, FL 32571   → Street\\nCity, FL Zip
        4431 Winners Gait Cir, Pace, FL, 32571  → Street\\nCity, FL Zip
        4431 Winners Gait Cir, Pace FL 32571    → Street\\nCity FL Zip
    """
    if not raw:
        return ""
    raw = raw.strip()
    if "\n" in raw:
        return raw  # already formatted correctly

    parts = [p.strip() for p in raw.split(",")]
    if len(parts) >= 4:
        # State and zip are always the last two comma segments; anything
        # between the street and those last two (e.g. an apartment/suite
        # typed as its own segment: "123 Main St, Apt 4, Pace, FL, 32571")
        # folds into the city line instead of being silently dropped along
        # with the real zip that used to get cut off after parts[3].
        street = parts[0]
        zip_c = parts[-1]
        state = parts[-2].upper()
        city = ", ".join(parts[1:-2])
        return f"{street}\n{city}, {state} {zip_c}"
    elif len(parts) >= 2:
        idx = raw.index(",")
        return f"{raw[:idx].strip()}\n{raw[idx + 1:].strip()}"
    return raw


def safe_filename_component(name):
    """A customer name is used as-is in the DB and printed on the document
    itself, but a small number of characters can't appear in a filename —
    notably '/', which the filesystem reads as a path separator and would
    otherwise crash the save (confirmed against two real customers already
    on file: "Chad Miller c/o Santi Tefel" and "Lucy Hemming/Mother's
    Generator"). This only affects the on-disk filename; the customer's
    name is never altered anywhere else (DB, invoice content, UI)."""
    return (name or "").replace("/", "-").replace("\\", "-")


def _write_document_xlsx(customer, job_name, line_items, doc_num, doc_type):
    """Build the Excel document (invoice/estimate/quote) from the template
    and save it as the customer-facing file. Does NOT touch the template's
    running number — callers decide whether this document consumes a new
    number (save_document) or reuses one already consumed (convert_to_invoice).

    line_items: list of (detail, raw_price) — raw_price may be a plain
    number or free text (e.g. a price range). Returns
    (path, numeric_total, amount_display). amount_display is None unless
    at least one line item is free text, in which case it's the literal
    string written into the SUBTOTAL/TOTAL cells (those are live Excel
    formulas normally — formulas can't sum free text, so when a range is
    involved we overwrite them with a literal display value, the same way
    this is handled by hand on today's real paperwork).
    """
    label = DOC_LABELS.get(doc_type, "INVOICE")
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb.active

    ws["B1"] = f"{label} #{doc_num}"
    ws["B7"] = customer["name"]
    ws["B8"] = normalize_address(customer.get("address") or "")
    ws["B9"] = customer.get("phone") or ""
    ws["C7"] = job_name

    parsed = [(detail, *parse_amount(raw)) for detail, raw in line_items]

    if parsed:
        ws["B13"] = parsed[0][0]
        ws["C13"] = parsed[0][1] if parsed[0][1] is not None else parsed[0][2]

    n_extra = len(parsed) - 1
    if n_extra > 0:
        subtotal_row = 14
        for row in ws.iter_rows(min_row=13, max_row=40):
            for cell in row:
                if cell.value == "SUBTOTAL":
                    subtotal_row = cell.row

        ws.insert_rows(subtotal_row, n_extra)
        for i, (detail, num, disp) in enumerate(parsed[1:]):
            ws.cell(row=subtotal_row + i, column=2).value = detail
            ws.cell(row=subtotal_row + i, column=3).value = num if num is not None else disp

        for tbl in ws.tables.values():
            mc, mr, xc, xr = range_boundaries(tbl.ref)
            if xr >= subtotal_row - 1:
                tbl.ref = (f"{get_column_letter(mc)}{mr}:"
                           f"{get_column_letter(xc)}{xr + n_extra}")

    numeric_total = sum(num for _, num, disp in parsed if num is not None)
    has_range = any(disp is not None for _, num, disp in parsed)
    amount_display = None

    if has_range:
        parts = [f"${num:,.2f}" if num is not None else disp for _, num, disp in parsed]
        amount_display = parts[0] if len(parts) == 1 else " + ".join(parts)

        # Re-scan for SUBTOTAL/TOTAL by label (works whether or not rows
        # were inserted above) and overwrite with literal text.
        for row in ws.iter_rows(min_row=13, max_row=60):
            for cell in row:
                if cell.value in ("SUBTOTAL", "TOTAL"):
                    ws.cell(row=cell.row, column=3).value = amount_display

    os.makedirs(INV_DIR, exist_ok=True)
    out = os.path.join(INV_DIR, f"{safe_filename_component(customer['name'])} #{doc_num}.xlsx")
    wb.save(out)
    return out, numeric_total, amount_display


def save_document(conn, customer, job_name, line_items, doc_num, doc_type="invoice"):
    """Create a brand-new document (invoice, estimate, or quote), advancing
    the shared running number. For doc_type='invoice' with all-numeric
    prices this is byte-for-byte identical to the original save_invoice().
    """
    # advance the running number on the template — invoices, estimates, and
    # quotes all share one sequence (matches how numbers are used by hand:
    # an estimate and the invoice it becomes carry the SAME number, but a
    # fresh document always gets the next unused one).
    wb_tmpl = openpyxl.load_workbook(TEMPLATE)
    ws_tmpl = wb_tmpl.active
    raw = ws_tmpl["B1"].value or f"INVOICE #{doc_num - 1}"
    prefix = ""
    for j, ch in enumerate(raw):
        if ch.isdigit():
            prefix = raw[:j]
            break
    ws_tmpl["B1"] = f"{prefix}{doc_num}"
    wb_tmpl.save(TEMPLATE)
    wb_tmpl.close()

    path, total, amount_display = _write_document_xlsx(customer, job_name, line_items, doc_num, doc_type)

    conn.execute("""INSERT INTO customers(name, address, phone) VALUES(?,?,?)
        ON CONFLICT(name) DO UPDATE SET
        address=excluded.address, phone=excluded.phone""",
        (customer["name"], customer.get("address"), customer.get("phone")))
    conn.execute("""INSERT OR IGNORE INTO invoices(number,customer,job,total,day,doc_type,amount_display,line_items_json)
        VALUES(?,?,?,?,?,?,?,?)""",
        (doc_num, customer["name"], job_name, total, date.today().isoformat(),
         doc_type, amount_display, json.dumps(line_items)))
    conn.commit()
    return path


def save_invoice(conn, customer, job_name, line_items, inv_num):
    """Back-compat wrapper — a plain invoice is just save_document(..., doc_type='invoice')."""
    return save_document(conn, customer, job_name, line_items, inv_num, doc_type="invoice")


def convert_to_invoice(conn, number, customer, job_name, line_items):
    """Turn an existing estimate/quote into an invoice, keeping the SAME
    document number — the running sequence already consumed it when the
    estimate was created, exactly like "Erin Barnes Estimate #3487" became
    "Erin Barnes #3487" nine days later, same number. Does NOT advance the
    template's counter. Updates the existing DB row in place (one row per
    document number) rather than inserting a second one.
    """
    customer = {
        "name": customer["name"],
        "address": normalize_address(customer.get("address") or ""),
        "phone": customer.get("phone") or "",
    }
    path, total, amount_display = _write_document_xlsx(customer, job_name, line_items, number, "invoice")

    conn.execute("""INSERT INTO customers(name, address, phone) VALUES(?,?,?)
        ON CONFLICT(name) DO UPDATE SET
        address=excluded.address, phone=excluded.phone""",
        (customer["name"], customer.get("address"), customer.get("phone")))
    conn.execute("""UPDATE invoices
        SET customer=?, job=?, total=?, day=?, doc_type='invoice', amount_display=?, line_items_json=?
        WHERE number=?""",
        (customer["name"], job_name, total, date.today().isoformat(),
         amount_display, json.dumps(line_items), number))
    conn.commit()
    return path


def update_document(conn, number, customer, job_name, line_items):
    """Edit an already-saved invoice/estimate/quote in place — fix a typo in
    the job description, correct a price, or reassign it to a different
    customer — without reopening Excel. Keeps the same document number and
    doc_type (doc_type is intentionally not editable here — "Convert to
    Invoice" is the dedicated, deliberate action for turning an estimate
    into an invoice; this is just correcting what's already on the
    document). The customer-facing .xlsx is regenerated in place; the
    original creation date (`day`) is left untouched since editing isn't a
    new business event the way creating or converting a document is.

    If the customer's name changes, the old customer-named file is removed
    so there's no orphaned duplicate sitting next to the new one (best
    effort — if the old file happens to be open/locked, the edit still
    succeeds, just without that cleanup).

    If this invoice was already marked paid, its linked Monthly Log entry
    (source='invoice', tied back by invoice_number) has its customer name
    and amount synced to match — otherwise the Monthly Log and the invoice
    it came from could silently drift apart, the same class of bug the
    Aug 6 hardening pass (section 17) was built to catch. This is a no-op
    UPDATE (0 rows) for an unpaid invoice or a non-invoice document.
    """
    existing = conn.execute(
        "SELECT customer, doc_type FROM invoices WHERE number=?", (number,)).fetchone()
    if not existing:
        raise ValueError(f"No document #{number} found.")
    old_customer_name, doc_type = existing

    customer = {
        "name": (customer.get("name") or "").strip(),
        "address": normalize_address(customer.get("address") or ""),
        "phone": customer.get("phone") or "",
    }
    if not customer["name"]:
        raise ValueError("Customer name is required.")

    old_path = os.path.join(INV_DIR, f"{safe_filename_component(old_customer_name)} #{number}.xlsx")
    path, total, amount_display = _write_document_xlsx(customer, job_name, line_items, number, doc_type)

    if customer["name"] != old_customer_name and os.path.abspath(old_path) != os.path.abspath(path):
        try:
            if os.path.exists(old_path):
                os.remove(old_path)
        except OSError:
            pass

    conn.execute("""INSERT INTO customers(name, address, phone) VALUES(?,?,?)
        ON CONFLICT(name) DO UPDATE SET
        address=excluded.address, phone=excluded.phone""",
        (customer["name"], customer.get("address"), customer.get("phone")))
    conn.execute("""UPDATE invoices
        SET customer=?, job=?, total=?, amount_display=?, line_items_json=?
        WHERE number=?""",
        (customer["name"], job_name, total, amount_display, json.dumps(line_items), number))
    conn.execute("""UPDATE monthly_entries SET customer=?, amount=?
        WHERE source='invoice' AND invoice_number=?""",
        (customer["name"], total, number))
    conn.commit()
    return path


def get_document(conn, number):
    row = conn.execute(
        """SELECT number, customer, job, total, day, doc_type, amount_display, line_items_json
           FROM invoices WHERE number = ?""", (number,)).fetchone()
    if not row:
        return None
    customer = get_customer(conn, row[1]) or {"name": row[1], "address": "", "phone": ""}
    try:
        line_items = json.loads(row[7]) if row[7] else []
    except (ValueError, TypeError):
        line_items = []
    return {
        "number": row[0], "customer": customer, "job": row[2], "total": row[3],
        "day": row[4], "doc_type": row[5] or "invoice", "amount_display": row[6],
        "line_items": line_items,
    }


def list_documents(conn, doc_types=("invoice",), query=None, limit=200, customer=None):
    """Searchable list of documents restricted to the given doc_type(s).
    customer, when given, is an exact-match filter (used by the Customers
    history page) rather than a substring — it's matched against the same
    denormalized text column invoices.customer already uses everywhere else.
    """
    placeholders = ",".join("?" for _ in doc_types)
    params = list(doc_types)
    sql = f"""SELECT number, customer, job, total, day, doc_type, amount_display, paid, paid_date
               FROM invoices WHERE doc_type IN ({placeholders})"""
    if customer:
        sql += " AND customer = ?"
        params.append(customer)
    if query:
        like = f"%{query}%"
        sql += " AND (customer LIKE ? OR job LIKE ? OR CAST(number AS TEXT) LIKE ?)"
        params += [like, like, like]
    sql += " ORDER BY number DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    return [
        {"number": r[0], "customer": r[1], "job": r[2], "total": r[3], "day": r[4],
         "doc_type": r[5] or "invoice", "amount_display": r[6], "paid": bool(r[7]), "paid_date": r[8]}
        for r in rows
    ]


def list_invoices(conn, query=None, limit=200):
    """Back-compat wrapper — searchable list of past invoices only."""
    return list_documents(conn, doc_types=("invoice",), query=query, limit=limit)


def home_stats(conn):
    """Home landing stats (Phase 6): paid income this month, outstanding
    (unpaid) total, customer/invoice/estimate counts, and recent activity
    across all document types. "Paid this month" is pulled from the
    Monthly Log itself (monthly_totals) rather than recomputed separately,
    so the number on Home always matches the Monthly Log page exactly —
    no risk of the two drifting apart.
    """
    month = date.today().isoformat()[:7]
    paid = monthly_totals(conn, month)

    unpaid_row = conn.execute(
        """SELECT COALESCE(SUM(total),0), COUNT(*) FROM invoices
           WHERE doc_type='invoice' AND (paid IS NULL OR paid=0)""").fetchone()
    unpaid_total, unpaid_count = unpaid_row[0], unpaid_row[1]

    total_customers = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    total_invoices = conn.execute(
        "SELECT COUNT(*) FROM invoices WHERE doc_type='invoice'").fetchone()[0]
    open_estimates = conn.execute(
        "SELECT COUNT(*) FROM invoices WHERE doc_type IN ('estimate','quote')").fetchone()[0]

    recent = conn.execute(
        """SELECT number, customer, job, total, day, doc_type, amount_display, paid
           FROM invoices ORDER BY id DESC LIMIT 8""").fetchall()

    return {
        "month_paid_total": paid["total"],
        "month_paid_count": paid["count"],
        "unpaid_total": unpaid_total,
        "unpaid_count": unpaid_count,
        "total_customers": total_customers,
        "total_invoices": total_invoices,
        "open_estimates": open_estimates,
        "recent": [
            {"number": r[0], "customer": r[1], "job": r[2], "total": r[3], "day": r[4],
             "doc_type": r[5] or "invoice", "amount_display": r[6], "paid": bool(r[7])}
            for r in recent
        ],
    }


# ─── Monthly Log (Phase 4) ──────────────────────────────────────────────────
# Reproduces the "Name | G/E | Gross" sheet MJ has been keeping by hand
# (see Weird/July 2025.xlsx) — filled automatically when an invoice is
# marked paid, or added manually for cash jobs / anything off-book.

def month_label(month):
    """'2025-07' -> 'July 2025' — matches the real sheet's title cell."""
    try:
        y, m = month.split("-")
        return f"{calendar.month_name[int(m)]} {y}"
    except Exception:
        return month


def mark_invoice_paid(conn, number, category, paid_date=None):
    """Mark a saved invoice paid and drop a matching row into the Monthly
    Log for whatever month it was paid in. category is 'G' (Generator) or
    'E' (Electrical) — that split isn't derivable from the invoice itself,
    so the caller (the "Mark Paid" dialog) always supplies it. Returns the
    month ('YYYY-MM') the entry landed in.
    """
    category = (category or "").strip().upper()
    if category not in ("G", "E"):
        raise ValueError("Category must be G (Generator) or E (Electrical).")
    paid_date = (paid_date or "").strip() or date.today().isoformat()

    row = conn.execute(
        "SELECT customer, total, doc_type FROM invoices WHERE number=?", (number,)).fetchone()
    if not row:
        raise ValueError(f"No invoice #{number} found.")
    customer_name, total, doc_type = row
    if doc_type != "invoice":
        raise ValueError("Only invoices can be marked paid — convert the estimate/quote first.")

    month = paid_date[:7]
    # The paid=0 check is folded into the UPDATE's WHERE clause (rather than
    # a separate SELECT-then-UPDATE) so a double-click or network retry that
    # fires two mark-paid requests for the same invoice can't both pass the
    # guard — the second one matches zero rows and gets a clean "already
    # paid" error instead of silently inserting a second monthly_entries row
    # and double-counting the income.
    cur = conn.execute(
        "UPDATE invoices SET paid=1, paid_date=? WHERE number=? AND (paid IS NULL OR paid=0)",
        (paid_date, number))
    if cur.rowcount == 0:
        raise ValueError(f"Invoice #{number} is already marked paid.")
    conn.execute("""INSERT INTO monthly_entries(month, entry_date, customer, category, amount, source, invoice_number)
        VALUES(?,?,?,?,?,?,?)""",
        (month, paid_date, customer_name, category, total or 0.0, "invoice", number))
    conn.commit()
    return month


def unmark_invoice_paid(conn, number):
    """The reverse of mark_invoice_paid — flips a paid invoice back to
    unpaid. Also removes its linked Monthly Log row if one exists (mirrors
    delete_monthly_entry's reverse sync, so the two never drift apart), but
    doesn't require one — an invoice paid via the one-time backfill (see
    open_db's paid_backfill_v1) never got a monthly_entries row in the
    first place, so this is the only way to unmark one of those.
    The paid=1 check is folded into the UPDATE's WHERE clause, same
    atomicity reasoning as mark_invoice_paid, so a double-click can't both
    "succeed" and double-delete a monthly_entries row.
    """
    cur = conn.execute(
        "UPDATE invoices SET paid=0, paid_date=NULL WHERE number=? AND paid=1", (number,))
    if cur.rowcount == 0:
        raise ValueError(f"Invoice #{number} is not currently marked paid.")
    conn.execute(
        "DELETE FROM monthly_entries WHERE source='invoice' AND invoice_number=?", (number,))
    conn.commit()


def add_manual_monthly_entry(conn, entry_date, customer, category, amount):
    """A Monthly Log row not tied to any invoice — cash jobs, anything
    off-book, or old entries being backfilled by hand."""
    customer = (customer or "").strip()
    category = (category or "").strip().upper()
    if not customer:
        raise ValueError("Name is required.")
    if category not in ("G", "E"):
        raise ValueError("Category must be G (Generator) or E (Electrical).")
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        raise ValueError("Gross amount must be a number.")
    entry_date = (entry_date or "").strip() or date.today().isoformat()
    month = entry_date[:7]
    conn.execute("""INSERT INTO monthly_entries(month, entry_date, customer, category, amount, source, invoice_number)
        VALUES(?,?,?,?,?,'manual',NULL)""",
        (month, entry_date, customer, category, amount))
    conn.commit()
    return month


def update_monthly_entry(conn, entry_id, entry_date, customer, category, amount):
    """Edit an existing Monthly Log row. Allowed for both manual entries and
    invoice-sourced ones — an invoice-sourced entry keeps its `source`/
    `invoice_number` link (so deleting it later still correctly un-marks the
    invoice), but its date/name/category/amount can be corrected here
    without touching the underlying invoice itself."""
    row = conn.execute("SELECT id FROM monthly_entries WHERE id=?", (entry_id,)).fetchone()
    if not row:
        raise ValueError("Entry not found.")
    customer = (customer or "").strip()
    category = (category or "").strip().upper()
    if not customer:
        raise ValueError("Name is required.")
    if category not in ("G", "E"):
        raise ValueError("Category must be G (Generator) or E (Electrical).")
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        raise ValueError("Gross amount must be a number.")
    entry_date = (entry_date or "").strip() or date.today().isoformat()
    month = entry_date[:7]
    conn.execute(
        "UPDATE monthly_entries SET month=?, entry_date=?, customer=?, category=?, amount=? WHERE id=?",
        (month, entry_date, customer, category, amount, entry_id))
    conn.commit()
    return month


def list_monthly_entries(conn, month):
    rows = conn.execute(
        """SELECT id, entry_date, customer, category, amount, source, invoice_number
           FROM monthly_entries WHERE month=? ORDER BY entry_date, id""", (month,)).fetchall()
    return [
        {"id": r[0], "entry_date": r[1], "customer": r[2], "category": r[3],
         "amount": r[4], "source": r[5], "invoice_number": r[6]}
        for r in rows
    ]


def monthly_totals(conn, month):
    row = conn.execute(
        """SELECT COALESCE(SUM(amount),0), COUNT(*),
                  COALESCE(SUM(CASE WHEN category='G' THEN amount ELSE 0 END),0),
                  COALESCE(SUM(CASE WHEN category='E' THEN amount ELSE 0 END),0)
           FROM monthly_entries WHERE month=?""", (month,)).fetchone()
    return {"total": row[0], "count": row[1], "g_total": row[2], "e_total": row[3]}


def delete_monthly_entry(conn, entry_id):
    """Remove a Monthly Log row. If it was auto-created by marking an
    invoice paid, also un-mark that invoice — otherwise the invoice would
    be stuck "paid" with no corresponding entry anywhere."""
    row = conn.execute(
        "SELECT source, invoice_number FROM monthly_entries WHERE id=?", (entry_id,)).fetchone()
    if not row:
        raise ValueError("Entry not found.")
    source, invoice_number = row
    conn.execute("DELETE FROM monthly_entries WHERE id=?", (entry_id,))
    if source == "invoice" and invoice_number:
        conn.execute("UPDATE invoices SET paid=0, paid_date=NULL WHERE number=?", (invoice_number,))
    conn.commit()


# Colors pulled directly from Invoice_template.xlsx's own theme (its heading
# accent and body-text accent, read out of the template's theme1.xml) so the
# Monthly Log export reads as the same document family as the real invoices
# instead of a differently-styled report — warm rust heading, dark-olive
# body text, clean white background, no solid color banners.
_INVOICE_ACCENT = "FFA85914"        # heading color ("Funky's Electrical", "BILL TO", "FOR")
_INVOICE_ACCENT_MUTED = "FF615A22"  # body/label text color
_INVOICE_LINE = "FFDDD4C0"          # thin warm-grey rule, echoes the invoice's understated dividers
_HEADER_FONT = "Georgia"
_BODY_FONT = "Arial"


def export_monthly_xlsx(conn, month):
    """Build a Monthly Log workbook for one month, styled to match the real
    Invoice_template.xlsx look (Georgia headings in the invoice's own rust
    accent color, clean white background, thin rule dividers instead of
    solid color banners) rather than standing apart with its own separate
    visual identity. One row per entry, plus a Total row and Generator/
    Electrical subtotal rows below it with a live SUM/SUMIF formula, styled
    like the invoice's own SUBTOTAL/SALES TAX/TOTAL block.
    """
    entries = list_monthly_entries(conn, month)
    totals = monthly_totals(conn, month)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly Log"
    ws.sheet_view.showGridLines = False

    # Fit-to-one-page-wide — belt-and-suspenders so "Download PDF"/"Email
    # PDF" never silently split a table across pages when printed/exported.
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 3

    thin_rule = Border(bottom=Side(style="thin", color=_INVOICE_LINE))
    accent_rule = Border(bottom=Side(style="thin", color=_INVOICE_ACCENT))
    total_rule = Border(top=Side(style="thin", color=_INVOICE_ACCENT))

    # Company header — same three lines the invoice itself opens with.
    ws.merge_cells("B2:D2")
    ws["B2"] = COMPANY
    ws["B2"].font = Font(name=_HEADER_FONT, size=26, color=_INVOICE_ACCENT)
    ws.row_dimensions[2].height = 38

    ws["B3"] = COMPANY_ADDR
    ws["B3"].font = Font(name=_BODY_FONT, size=11, color=_INVOICE_ACCENT_MUTED)
    ws["B4"] = COMPANY_PHONE
    ws["B4"].font = Font(name=_BODY_FONT, size=11, color=_INVOICE_ACCENT_MUTED)

    # "MONTHLY LOG" / month — same weight & placement as the invoice's own
    # "BILL TO" / "FOR" section headers.
    ws["B6"] = "MONTHLY LOG"
    ws["B6"].font = Font(name=_HEADER_FONT, size=15, color=_INVOICE_ACCENT)
    ws["C6"] = month_label(month)
    ws["C6"].font = Font(name=_BODY_FONT, size=12, bold=True, color=_INVOICE_ACCENT_MUTED)
    ws.row_dimensions[6].height = 22

    header_row = 8
    for col, label in (("B", "NAME"), ("C", "G/E"), ("D", "GROSS")):
        cell = ws[f"{col}{header_row}"]
        cell.value = label
        cell.font = Font(name=_BODY_FONT, size=12, bold=True, color=_INVOICE_ACCENT_MUTED)
        cell.border = accent_rule
        if col == "D":
            cell.alignment = Alignment(horizontal="right")
    ws.row_dimensions[header_row].height = 20

    row = header_row + 1
    for e in entries:
        ws[f"B{row}"] = e["customer"]
        ws[f"C{row}"] = e["category"]
        ws[f"D{row}"] = e["amount"]
        ws[f"D{row}"].number_format = '"$"#,##0.00'
        ws[f"D{row}"].alignment = Alignment(horizontal="right")
        for col in ("B", "C", "D"):
            cell = ws[f"{col}{row}"]
            cell.font = Font(name=_BODY_FONT, size=11, color="FF262117")
            cell.border = thin_rule
        ws.row_dimensions[row].height = 20
        row += 1

    first_data_row = header_row + 1
    last_data_row = row - 1
    if not entries:
        row += 1  # keep the Total row visually separated even with no data yet

    total_row = row
    ws[f"C{total_row}"] = "Total"
    ws[f"C{total_row}"].alignment = Alignment(horizontal="right")
    ws[f"D{total_row}"] = (f"=SUM(D{first_data_row}:D{last_data_row})" if entries else 0)
    for col in ("B", "C", "D"):
        cell = ws[f"{col}{total_row}"]
        cell.font = Font(name=_BODY_FONT, size=13, bold=True, color=_INVOICE_ACCENT_MUTED)
        cell.border = total_rule
    ws[f"D{total_row}"].number_format = '"$"#,##0.00'
    ws[f"D{total_row}"].alignment = Alignment(horizontal="right")
    ws.row_dimensions[total_row].height = 24

    for label, cat in (("Generator (G) Total", "G"), ("Electrical (E) Total", "E")):
        row = total_row + (1 if cat == "G" else 2)
        ws[f"C{row}"] = label
        ws[f"C{row}"].alignment = Alignment(horizontal="right")
        ws[f"D{row}"] = (f'=SUMIF(C{first_data_row}:C{last_data_row},"{cat}",D{first_data_row}:D{last_data_row})'
                          if entries else 0)
        for col in ("B", "C", "D"):
            cell = ws[f"{col}{row}"]
            cell.font = Font(name=_BODY_FONT, size=11, color=_INVOICE_ACCENT_MUTED)
        ws[f"D{row}"].number_format = '"$"#,##0.00'
        ws[f"D{row}"].alignment = Alignment(horizontal="right")
        ws.row_dimensions[row].height = 18

    os.makedirs(MONTHLY_DIR, exist_ok=True)
    out = os.path.join(MONTHLY_DIR, f"Monthly Log - {month_label(month)}.xlsx")
    wb.save(out)
    return out


# ─── Payroll Log (Phase 5) ───────────────────────────────────────────────
# Simple Date/Worker/Gross entry. Workers' Comp classification used to be a
# fixed profile per worker, but the shop only ever uses two class codes and
# which one applies depends on what kind of work the wages were for, not
# who did it — so (as of the Aug 6, 2026 payroll redesign) the code is
# chosen per entry instead, from this fixed, always-the-same-two-codes
# list. State is a single report-level setting ("payroll_state", default
# "FL") rather than a per-worker field, since every worker files under the
# same policy/state — matches how the real Payroll Report March 2026.xlsx
# shows the same state on every group row regardless of who's in it.
PAYROLL_CODES = {
    # Rates are the decimal form of "dollars per $100 of payroll" — e.g.
    # 5190's real rate is $2.969 per $100, so 2.969/100 = 0.02969 here.
    # Premium = Subject-to-WC wages x this rate (see export_payroll_xlsx).
    "5190": {"description": "Wiring", "rate": 0.02969},
    "3724": {"description": "Machinery", "rate": 0.0355},
}
PAYROLL_CODE_ORDER = ["5190", "3724"]
DEFAULT_PAYROLL_STATE = "FL"


def get_setting(conn, key, default=None):
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    return row[0] if row else default


def set_setting(conn, key, value):
    conn.execute("""INSERT INTO settings(key, value) VALUES(?,?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value""", (key, value))
    conn.commit()


def create_worker(conn, name):
    """Workers are now just a simple name roster (autocomplete convenience
    on the Payroll Log's "Worker" field) — Workers' Comp classification no
    longer lives here, see PAYROLL_CODES above."""
    name = (name or "").strip()
    if not name:
        raise ValueError("Worker name is required.")
    if conn.execute("SELECT 1 FROM workers WHERE name=?", (name,)).fetchone():
        raise ValueError(f'A worker named "{name}" already exists.')
    conn.execute("INSERT INTO workers(name) VALUES(?)", (name,))
    conn.commit()


def update_worker(conn, old_name, name):
    """Rename a worker — workers are a small, owner-managed list (unlike
    customers, which have hundreds of historical invoices pinned to their
    name), so a rename safely cascades into any payroll_entries already
    logged under the old name."""
    name = (name or "").strip()
    if not name:
        raise ValueError("Worker name is required.")
    if not conn.execute("SELECT 1 FROM workers WHERE name=?", (old_name,)).fetchone():
        raise ValueError(f'No worker named "{old_name}" found.')
    if name != old_name and conn.execute("SELECT 1 FROM workers WHERE name=?", (name,)).fetchone():
        raise ValueError(f'A worker named "{name}" already exists.')
    conn.execute("UPDATE workers SET name=? WHERE name=?", (name, old_name))
    if name != old_name:
        conn.execute("UPDATE payroll_entries SET worker=? WHERE worker=?", (name, old_name))
    conn.commit()


def get_worker(conn, name):
    row = conn.execute("SELECT name FROM workers WHERE name=?", (name,)).fetchone()
    return {"name": row[0]} if row else None


def list_workers(conn):
    rows = conn.execute("SELECT name FROM workers ORDER BY name").fetchall()
    return [{"name": r[0]} for r in rows]


def add_payroll_entry(conn, entry_date, worker, class_code, gross):
    worker = (worker or "").strip()
    if not worker:
        raise ValueError("Worker is required.")
    class_code = (class_code or "").strip()
    if class_code not in PAYROLL_CODES:
        raise ValueError("Choose a class code (Wiring or Machinery).")
    try:
        gross = float(gross)
    except (TypeError, ValueError):
        raise ValueError("Gross must be a number.")
    entry_date = (entry_date or "").strip() or date.today().isoformat()
    month = entry_date[:7]
    conn.execute(
        "INSERT INTO payroll_entries(month, entry_date, worker, gross, class_code) VALUES(?,?,?,?,?)",
        (month, entry_date, worker, gross, class_code))
    conn.commit()
    return month


def update_payroll_entry(conn, entry_id, entry_date, worker, class_code, gross):
    """Edit an existing Payroll Log row."""
    row = conn.execute("SELECT id FROM payroll_entries WHERE id=?", (entry_id,)).fetchone()
    if not row:
        raise ValueError("Entry not found.")
    worker = (worker or "").strip()
    if not worker:
        raise ValueError("Worker is required.")
    class_code = (class_code or "").strip()
    if class_code not in PAYROLL_CODES:
        raise ValueError("Choose a class code (Wiring or Machinery).")
    try:
        gross = float(gross)
    except (TypeError, ValueError):
        raise ValueError("Gross must be a number.")
    entry_date = (entry_date or "").strip() or date.today().isoformat()
    month = entry_date[:7]
    conn.execute(
        "UPDATE payroll_entries SET month=?, entry_date=?, worker=?, gross=?, class_code=? WHERE id=?",
        (month, entry_date, worker, gross, class_code, entry_id))
    conn.commit()
    return month


def list_payroll_entries(conn, month):
    rows = conn.execute(
        """SELECT id, entry_date, worker, gross, class_code FROM payroll_entries
           WHERE month=? ORDER BY entry_date, id""", (month,)).fetchall()
    return [{"id": r[0], "entry_date": r[1], "worker": r[2], "gross": r[3], "class_code": r[4]} for r in rows]


def delete_payroll_entry(conn, entry_id):
    cur = conn.execute("DELETE FROM payroll_entries WHERE id=?", (entry_id,))
    conn.commit()
    if cur.rowcount == 0:
        raise ValueError("Entry not found.")


def payroll_totals(conn, month):
    row = conn.execute(
        "SELECT COALESCE(SUM(gross),0), COUNT(*) FROM payroll_entries WHERE month=?", (month,)).fetchone()
    return {"total": row[0], "count": row[1]}


def export_payroll_xlsx(conn, month):
    """Build the Workers' Comp Payroll Report for one month, styled to match
    the invoice/Monthly Log look (Georgia headings in the invoice's own
    rust accent color, thin rule dividers) instead of the plain black/bold
    look it had before. A Date/Employee/Gross log on the left, grouped by
    the two fixed Workers' Comp class codes (see PAYROLL_CODES) — both
    codes always appear, even with $0 in wages that month, matching the
    real Payroll Report March 2026.xlsx (every group row carries the same
    state, and both codes are always laid out, not just whichever one saw
    activity). Reg. Wages and Subject to WC are set equal to Gross — the
    real file does the same wherever it actually computed a premium (no
    wage caps or exclusions in play here, matching "per-worker yearly tax
    totals not needed" from the outline).
    """
    entries = list_payroll_entries(conn, month)
    policy_number = get_setting(conn, "policy_number", "")
    state = get_setting(conn, "payroll_state", DEFAULT_PAYROLL_STATE) or DEFAULT_PAYROLL_STATE

    y, m = month.split("-")
    first_day = date(int(y), int(m), 1)
    last_day = date(int(y), int(m), calendar.monthrange(int(y), int(m))[1])
    period = f"{first_day.strftime('%B %-d, %Y') if os.name != 'nt' else first_day.strftime('%B %d, %Y')} - {last_day.strftime('%B %-d, %Y') if os.name != 'nt' else last_day.strftime('%B %d, %Y')}"

    # Both class codes always get a group — fixed order, fixed rate — even
    # with zero entries that month (matches PAYROLL_CODES being "already
    # defined in every doc" rather than only appearing once used).
    groups = {code: [] for code in PAYROLL_CODE_ORDER}
    for e in entries:
        code = e.get("class_code")
        if code in groups:
            groups[code].append(e)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Payroll Report"
    ws.sheet_view.showGridLines = False

    # Landscape + fit-to-one-page-WIDE (still, so the 9 columns of one row
    # never get split across pages — a row cut in half is far less
    # readable than a shrunk-down row). Height is left unconstrained
    # (fitToHeight=0) so, with the extra row spacing added below, the
    # report is free to spill onto a second page top-to-bottom instead of
    # everything being squeezed to fit on one.
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 9
    ws.column_dimensions["E"].width = 13
    ws.column_dimensions["F"].width = 20
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 17
    ws.column_dimensions["I"].width = 13

    thin_rule = Border(bottom=Side(style="thin", color=_INVOICE_LINE))
    accent_rule = Border(bottom=Side(style="thin", color=_INVOICE_ACCENT))
    total_rule = Border(top=Side(style="thin", color=_INVOICE_ACCENT))
    currency = '"$"#,##0.00'

    # Company header — same as the invoice and Monthly Log. Merged out to
    # column C (not just B) so "Funky's Electrical" and the period line
    # have room to breathe instead of clipping against column D's content.
    ws.merge_cells("A1:C1")
    ws["A1"] = COMPANY
    ws["A1"].font = Font(name=_HEADER_FONT, size=22, color=_INVOICE_ACCENT)
    ws.row_dimensions[1].height = 32
    ws["A2"] = COMPANY_ADDR
    ws["A2"].font = Font(name=_BODY_FONT, size=10.5, color=_INVOICE_ACCENT_MUTED)

    ws.merge_cells("A3:C3")
    ws["A3"] = f"Payroll: {period}"
    ws["A3"].font = Font(name=_HEADER_FONT, size=12.5, color=_INVOICE_ACCENT)
    ws["D1"] = "PAYROLL REPORT"
    ws["D1"].font = Font(name=_HEADER_FONT, size=15, color=_INVOICE_ACCENT)
    ws["D2"] = f"Policy Number #{policy_number}" if policy_number else "Policy Number #(not set)"
    ws["D2"].font = Font(name=_BODY_FONT, size=11, bold=True, color=_INVOICE_ACCENT_MUTED)
    ws["D3"] = f"Reporting Period: {period}"
    ws["D3"].font = Font(name=_BODY_FONT, size=11, bold=True, color=_INVOICE_ACCENT_MUTED)

    headers = [("A", "Date"), ("B", "Employee"), ("C", "Gross"), ("D", "State"),
               ("E", "Class Code"), ("F", "Code Description"), ("G", "Reg. Wages"),
               ("H", "Subject to WC"), ("I", "Net Rate")]
    header_row = 6
    ws.print_title_rows = f"{header_row}:{header_row}"
    for col, label in headers:
        cell = ws[f"{col}{header_row}"]
        cell.value = label
        cell.font = Font(name=_BODY_FONT, size=12, bold=True, color=_INVOICE_ACCENT_MUTED)
        cell.border = accent_rule
    ws.row_dimensions[header_row].height = 22

    row = header_row + 1
    grand_total_gross = 0.0
    grand_total_premium = 0.0

    for code in PAYROLL_CODE_ORDER:
        rows = groups[code]
        info = PAYROLL_CODES[code]
        group_start = row
        for e in rows:
            ws[f"A{row}"] = date.fromisoformat(e["entry_date"])
            ws[f"A{row}"].number_format = "m/d/yyyy"
            ws[f"B{row}"] = e["worker"]
            ws[f"C{row}"] = e["gross"]
            ws[f"C{row}"].number_format = currency
            for col in ("A", "B", "C"):
                cell = ws[f"{col}{row}"]
                cell.font = Font(name=_BODY_FONT, size=11, color="FF262117")
                cell.border = thin_rule
            ws.row_dimensions[row].height = 19
            row += 1
        group_end = row - 1
        group_gross_ref = f"SUM(C{group_start}:C{group_end})" if rows else None

        # Total row first, directly under the entries it closes off (A/C
        # only — nothing in D onward, so the right-aligned currency total
        # never sits flush against left-aligned text in the very next
        # column, which is what made an earlier version read as one mashed
        # value like "$2,970.75FL").
        group_total_row = row
        ws[f"A{group_total_row}"] = f"{info['description']} Total:"
        ws[f"A{group_total_row}"].font = Font(name=_BODY_FONT, size=12, bold=True, color=_INVOICE_ACCENT_MUTED)
        ws[f"C{group_total_row}"] = f"={group_gross_ref}" if group_gross_ref else 0
        ws[f"C{group_total_row}"].number_format = currency
        ws[f"C{group_total_row}"].font = Font(name=_BODY_FONT, size=12, bold=True, color=_INVOICE_ACCENT_MUTED)
        ws[f"C{group_total_row}"].border = total_rule
        ws[f"A{group_total_row}"].border = total_rule
        ws.row_dimensions[group_total_row].height = 21
        row += 1

        # Classification row — state / class code / description / reg
        # wages / subject to WC / net rate, all group-level, on its own
        # row so it has room to breathe instead of sharing a row with the
        # total above.
        class_row = row
        ws[f"D{class_row}"] = state
        ws[f"E{class_row}"] = code
        ws[f"F{class_row}"] = info["description"]
        ws[f"G{class_row}"] = f"={group_gross_ref}" if group_gross_ref else 0
        ws[f"G{class_row}"].number_format = currency
        ws[f"H{class_row}"] = f"={group_gross_ref}" if group_gross_ref else 0
        ws[f"H{class_row}"].number_format = currency
        rate = info["rate"]
        ws[f"I{class_row}"] = rate
        ws[f"I{class_row}"].number_format = "0.0000"
        for col in ("D", "E", "F", "G", "H", "I"):
            ws[f"{col}{class_row}"].font = Font(name=_BODY_FONT, size=11, color=_INVOICE_ACCENT_MUTED)
        ws.row_dimensions[class_row].height = 19
        row += 1

        premium_row = row
        ws[f"D{premium_row}"] = "Premium:"
        ws[f"D{premium_row}"].font = Font(name=_BODY_FONT, size=11, bold=True, color=_INVOICE_ACCENT_MUTED)
        ws[f"I{premium_row}"] = f"=H{class_row}*I{class_row}"
        ws[f"I{premium_row}"].number_format = currency
        ws[f"I{premium_row}"].font = Font(name=_BODY_FONT, size=11, bold=True, color=_INVOICE_ACCENT_MUTED)
        ws.row_dimensions[premium_row].height = 19

        group_gross = sum(e["gross"] for e in rows)
        grand_total_gross += group_gross
        grand_total_premium += group_gross * rate

        row = premium_row + 3  # two blank spacer rows between groups

    grand_row = row
    ws[f"A{grand_row}"] = "Total:"
    ws[f"A{grand_row}"].font = Font(name=_BODY_FONT, size=13, bold=True, color=_INVOICE_ACCENT)
    ws[f"C{grand_row}"] = grand_total_gross
    ws[f"C{grand_row}"].number_format = currency
    ws[f"C{grand_row}"].font = Font(name=_BODY_FONT, size=13, bold=True, color=_INVOICE_ACCENT)
    ws[f"H{grand_row}"] = "Total Premium:"
    ws[f"H{grand_row}"].font = Font(name=_BODY_FONT, size=13, bold=True, color=_INVOICE_ACCENT)
    ws[f"I{grand_row}"] = grand_total_premium
    ws[f"I{grand_row}"].number_format = currency
    ws[f"I{grand_row}"].font = Font(name=_BODY_FONT, size=13, bold=True, color=_INVOICE_ACCENT)

    if not entries:
        note_row = grand_row + 2
        ws[f"A{note_row}"] = "No payroll entries for this month yet — both class codes still show at $0."
        ws[f"A{note_row}"].font = Font(name=_BODY_FONT, size=10, italic=True, color=_INVOICE_ACCENT_MUTED)
        note_row += 1
    else:
        note_row = grand_row + 2

    ws[f"A{note_row}"] = ("Reg. Wages and Subject to WC are set equal to Gross (no wage caps or "
                           "exclusions applied). Verify against your policy's actual wage-basis rules.")
    ws[f"A{note_row}"].font = Font(name=_BODY_FONT, size=9, italic=True, color="FF6B7280")

    os.makedirs(PAYROLL_DIR, exist_ok=True)
    out = os.path.join(PAYROLL_DIR, f"Payroll Report {month_label(month)}.xlsx")
    wb.save(out)
    return out


# ─── Reports (Phase 6) ───────────────────────────────────────────────────
# Month-by-month income, G/E breakdown, and top customers — all built on
# top of data the earlier phases already collect (Monthly Log entries,
# invoice totals), no new tables needed.

def monthly_income_series(conn, months=12):
    """The last N months of paid income (Total/Generator/Electrical), most
    recent last, including months with no entries at all — makes for a
    clean trend without gaps."""
    today = date.today()
    seq = []
    y, m = today.year, today.month
    for i in range(months - 1, -1, -1):
        mm = m - i
        yy = y
        while mm <= 0:
            mm += 12
            yy -= 1
        seq.append(f"{yy:04d}-{mm:02d}")
    result = []
    for mo in seq:
        t = monthly_totals(conn, mo)
        result.append({"month": mo, "month_label": month_label(mo),
                        "total": t["total"], "g_total": t["g_total"], "e_total": t["e_total"]})
    return result


def top_customers(conn, limit=10):
    """Customers ranked by total invoiced (real invoices only)."""
    rows = conn.execute(
        """SELECT customer, COALESCE(SUM(total),0) AS t, COUNT(*) AS c
           FROM invoices WHERE doc_type='invoice'
           GROUP BY customer ORDER BY t DESC LIMIT ?""", (limit,)).fetchall()
    return [{"customer": r[0], "total": r[1], "count": r[2]} for r in rows]


def ge_alltime_totals(conn):
    """All-time Generator vs. Electrical split, from every Monthly Log
    entry ever recorded (paid income only, matching the Monthly Log's own
    definition of income)."""
    row = conn.execute(
        """SELECT COALESCE(SUM(CASE WHEN category='G' THEN amount ELSE 0 END),0),
                  COALESCE(SUM(CASE WHEN category='E' THEN amount ELSE 0 END),0),
                  COALESCE(SUM(amount),0)
           FROM monthly_entries""").fetchone()
    return {"g_total": row[0], "e_total": row[1], "total": row[2]}


# ─── Browser preview PDF (reportlab, cross-platform, no Excel needed) ───────
# Used only for the in-browser "Preview" button — the actual saved document
# is always the Excel file built in save_invoice(). Print/email still use
# Excel's own renderer via xlwings below, exactly as the desktop app does.

_BLUE = colors.HexColor("#1a5276")
_LGRAY = colors.HexColor("#f2f3f4")


def _ps(name, **kw):
    return ParagraphStyle(name, **kw)


def generate_preview_pdf(customer, job_name, line_items, inv_num, doc_type="invoice", out_path=None):
    """Render a quick-look PDF for the browser preview pane. line_items may
    contain free-text (range) prices for estimates/quotes."""
    label = DOC_LABELS.get(doc_type, "INVOICE")
    out = out_path or os.path.join(BASE_DIR, f".preview-{inv_num}.pdf")
    doc = SimpleDocTemplate(out, pagesize=letter,
                             leftMargin=0.75 * inch, rightMargin=0.75 * inch,
                             topMargin=0.75 * inch, bottomMargin=0.75 * inch)
    W = letter[0] - 1.5 * inch

    today = date.today().strftime("%B %d, %Y")
    parsed = [(detail, *parse_amount(raw)) for detail, raw in line_items]
    numeric_total = sum(num for _, num, disp in parsed if num is not None)
    has_range = any(disp is not None for _, num, disp in parsed)
    if has_range:
        parts = [f"${num:,.2f}" if num is not None else disp for _, num, disp in parsed]
        total_display = parts[0] if len(parts) == 1 else " + ".join(parts)
    else:
        total_display = f"${numeric_total:,.2f}"

    hdr = Table([
        [Paragraph(COMPANY, _ps("co", fontSize=14, fontName="Helvetica-Bold", textColor=_BLUE)),
         Paragraph(f"{label} #{inv_num}", _ps("in", fontSize=20, fontName="Helvetica-Bold", textColor=_BLUE, alignment=TA_RIGHT))],
        [Paragraph(COMPANY_ADDR, _ps("ca", fontSize=8, textColor=colors.grey)),
         Paragraph(today, _ps("dt", fontSize=8, textColor=colors.grey, alignment=TA_RIGHT))],
        [Paragraph(COMPANY_PHONE, _ps("cp", fontSize=8, textColor=colors.grey)), ""],
    ], colWidths=[W * 0.55, W * 0.45])
    hdr.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 10),  # extra clearance under the large invoice number
    ]))

    addr_html = (customer.get("address") or "").replace("\n", "<br/>")
    bill = Table([
        [Paragraph("BILL TO", _ps("bh", fontSize=9, fontName="Helvetica-Bold", textColor=colors.white)),
         Paragraph("FOR", _ps("fh", fontSize=9, fontName="Helvetica-Bold", textColor=colors.white))],
        [Paragraph(customer["name"], _ps("cn", fontSize=10, fontName="Helvetica-Bold")),
         Paragraph(job_name, _ps("jn", fontSize=10, fontName="Helvetica-Bold"))],
        [Paragraph(addr_html, _ps("ad", fontSize=9)), ""],
        [Paragraph(customer.get("phone") or "", _ps("ph", fontSize=9)), ""],
    ], colWidths=[W * 0.55, W * 0.45])
    bill.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), _BLUE),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 4), ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
        ("TOPPADDING", (0, 1), (-1, -1), 3), ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))

    rows = [[Paragraph("Details", _ps("dh", fontSize=9, fontName="Helvetica-Bold", textColor=colors.white)),
             Paragraph("Amount", _ps("ah", fontSize=9, fontName="Helvetica-Bold", textColor=colors.white, alignment=TA_RIGHT))]]
    for i, (detail, num, disp) in enumerate(parsed):
        amt_text = f"${num:,.2f}" if num is not None else disp
        rows.append([
            Paragraph(detail, _ps(f"d{i}", fontSize=9)),
            Paragraph(amt_text, _ps(f"p{i}", fontSize=9, alignment=TA_RIGHT)),
        ])
    rows += [
        ["", ""],
        [Paragraph("SUBTOTAL", _ps("sub", fontSize=9, fontName="Helvetica-Bold")),
         Paragraph(total_display, _ps("sv2", fontSize=9, fontName="Helvetica-Bold", alignment=TA_RIGHT))],
        [Paragraph("SALES TAX", _ps("tx", fontSize=9)),
         Paragraph("N/A", _ps("txv", fontSize=9, alignment=TA_RIGHT))],
        [Paragraph("TOTAL", _ps("tot", fontSize=11, fontName="Helvetica-Bold", textColor=_BLUE)),
         Paragraph(total_display, _ps("tv", fontSize=11, fontName="Helvetica-Bold", textColor=_BLUE, alignment=TA_RIGHT))],
    ]
    row_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), _BLUE),
        ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LINEBELOW", (0, -1), (-1, -1), 1.5, _BLUE),
        ("LINEABOVE", (0, -1), (-1, -1), 0.5, colors.lightgrey),
    ]
    for i in range(len(parsed)):
        if i % 2 == 0:
            row_styles.append(("BACKGROUND", (0, i + 1), (-1, i + 1), _LGRAY))
    items = Table(rows, colWidths=[W * 0.72, W * 0.28])
    items.setStyle(TableStyle(row_styles))

    fs = _ps("ft", fontSize=7, textColor=colors.grey)
    doc.build([
        hdr,
        HRFlowable(width="100%", thickness=2, color=_BLUE, spaceAfter=8),
        bill, Spacer(1, 8), items, Spacer(1, 16),
        HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey, spaceAfter=4),
        Paragraph(f"Make all checks payable to {COMPANY}", fs),
        Paragraph(f"{SALESPERSON},  {COMPANY_PHONE},  {COMPANY_EMAIL}  |  License # {LICENSE_NUM}", fs),
    ])
    return out


# ─── PDF via xlwings (Excel renders it) — macOS + Excel only ────────────────

def generate_pdf_from_xlsx(xlsx_path, out_dir=None):
    """Export xlsx to PDF using Excel's own renderer via xlwings. By default
    the PDF is written right next to the source .xlsx — the right behavior
    for Monthly Log/Payroll Report exports, which already live in their own
    dedicated folders (MONTHLY_DIR/PAYROLL_DIR). Pass out_dir to write the
    PDF somewhere else instead — used for invoice/estimate/quote PDFs,
    which get their own folder (INV_PDF_DIR) separate from the .xlsx files
    (section 27)."""
    import xlwings as xw
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        pdf_name = os.path.basename(xlsx_path).replace(".xlsx", ".pdf")
        pdf_path = os.path.join(out_dir, pdf_name)
    else:
        pdf_path = xlsx_path.replace(".xlsx", ".pdf")
    app = xw.App(visible=False)
    try:
        wb = app.books.open(os.path.abspath(xlsx_path))
        wb.to_pdf(pdf_path)
        wb.close()
    finally:
        app.quit()
    return pdf_path


# ─── Email via Mail.app — macOS only ─────────────────────────────────────────

def _applescript_quote(s):
    """Escape a string for safe interpolation into an AppleScript
    double-quoted string literal. Without this, a customer name or file
    path containing a '"' (or a stray backslash) breaks out of the literal
    — at best a cryptic AppleScript error, at worst arbitrary AppleScript
    execution on the shop Mac spliced in via the customer-name field."""
    return (s or "").replace("\\", "\\\\").replace('"', '\\"')


def _sender_line(from_address):
    """AppleScript line that sets an outgoing message's sender account, or
    "" to leave it unset (Mail.app then uses its own default account —
    the original, pre-section-24 behavior). Mail.app can only send AS AN
    ACCOUNT IT ALREADY HAS CONFIGURED (Mail > Settings > Accounts on the
    shop Mac) — this tells Mail which of those existing accounts to use,
    it can't invent a new outgoing account or send as an arbitrary address
    Mail doesn't already have credentials for."""
    if not from_address:
        return ""
    return f'set sender of msg to "{_applescript_quote(from_address)}"'


def email_invoice(pdf_path, inv_num, customer_name, from_address=None):
    """Attach PDF to a new Mail.app message and open it ready to send."""
    abs_pdf = _applescript_quote(os.path.abspath(pdf_path))
    subject = _applescript_quote(f"{customer_name} #{inv_num}")
    sender_line = _sender_line(from_address)
    steps = [
        "-e", 'tell application "Mail"',
        "-e", (f'set msg to make new outgoing message with properties '
               f'{{subject:"{subject}", content:"", visible:true}}'),
        "-e", "tell msg",
    ]
    if sender_line:
        steps += ["-e", sender_line]
    steps += [
        "-e", f'make new to recipient with properties {{address:"{EMAIL_TO}"}}',
        "-e", f'make new attachment with properties {{file name:POSIX file "{abs_pdf}"}}',
        "-e", "end tell",
        "-e", "activate",
        "-e", "end tell",
    ]
    result = subprocess.run(["osascript"] + steps, capture_output=True, text=True, timeout=30)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())


def email_invoice_auto(pdf_path, inv_num, customer_name, to_address, from_address=None):
    """Bulk-send variant of email_invoice() — one email per invoice, sent
    immediately instead of left open in Mail.app for manual review. This
    skips the human-in-the-loop step the single-invoice "Email PDF" action
    normally gives, so callers should only reach this after an explicit,
    visible confirmation (invoice count + destination address) — never as
    a silent side effect. from_address (the "send_from_email" setting,
    section 24) picks which of Mail.app's own configured accounts sends
    it; left as None, Mail.app's default account is used, same as before
    this option existed."""
    abs_pdf = _applescript_quote(os.path.abspath(pdf_path))
    subject = _applescript_quote(f"{customer_name} #{inv_num}")
    to = _applescript_quote(to_address)
    sender_line = _sender_line(from_address)
    steps = [
        "-e", 'tell application "Mail"',
        "-e", (f'set msg to make new outgoing message with properties '
               f'{{subject:"{subject}", content:"", visible:false}}'),
        "-e", "tell msg",
    ]
    if sender_line:
        steps += ["-e", sender_line]
    steps += [
        "-e", f'make new to recipient with properties {{address:"{to}"}}',
        "-e", f'make new attachment with properties {{file name:POSIX file "{abs_pdf}"}}',
        "-e", "delay 1",
        "-e", "end tell",
        "-e", "send msg",
        "-e", "end tell",
    ]
    result = subprocess.run(["osascript"] + steps, capture_output=True, text=True, timeout=30)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())


def email_report_auto(pdf_path, subject, to_address, from_address=None):
    """Email a Monthly Log or Payroll Report PDF as its own message, sent
    immediately — the report equivalent of email_invoice_auto(), just with
    a caller-supplied subject instead of an invoice number/customer name.
    Used by the "Email PDF" action on the Monthly Log and Payroll Log
    pages, always after an explicit on-screen confirmation (destination +
    which month) — never as a silent side effect. from_address works the
    same way as in email_invoice_auto() — see that docstring."""
    abs_pdf = _applescript_quote(os.path.abspath(pdf_path))
    subject_q = _applescript_quote(subject)
    to = _applescript_quote(to_address)
    sender_line = _sender_line(from_address)
    steps = [
        "-e", 'tell application "Mail"',
        "-e", (f'set msg to make new outgoing message with properties '
               f'{{subject:"{subject_q}", content:"", visible:false}}'),
        "-e", "tell msg",
    ]
    if sender_line:
        steps += ["-e", sender_line]
    steps += [
        "-e", f'make new to recipient with properties {{address:"{to}"}}',
        "-e", f'make new attachment with properties {{file name:POSIX file "{abs_pdf}"}}',
        "-e", "delay 1",
        "-e", "end tell",
        "-e", "send msg",
        "-e", "end tell",
    ]
    result = subprocess.run(["osascript"] + steps, capture_output=True, text=True, timeout=30)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())


# ─── Print via Excel — macOS + Excel only ────────────────────────────────────

def print_xlsx(xlsx_path):
    """Open xlsx in Microsoft Excel and send to the default printer."""
    abs_path = _applescript_quote(os.path.abspath(xlsx_path))
    result = subprocess.run(
        [
            "osascript",
            "-e", 'tell application "Microsoft Excel"',
            "-e", "activate",
            "-e", f'open POSIX file "{abs_path}"',
            "-e", "delay 4",
            "-e", "print out active sheet",
            "-e", "delay 2",
            "-e", "close active workbook saving no",
            "-e", "end tell",
        ],
        capture_output=True, text=True, timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip())
