#!/usr/bin/env python3
"""
Funky's Electrical — Bookkeeping Dashboard
Phase 1: web dashboard shell + the existing Invoice tool, modern UI.
Phase 2: Estimates / Quotes — type toggle, price ranges, convert-to-invoice.

Run with:  python3 app.py
Then open: http://localhost:5001  (or http://<this-mac's-name>.local:5001
from any other machine on the shop network)

Port 5001 (not the default Flask 5000) because macOS's built-in AirPlay
Receiver claims port 5000 by default on modern macOS — starting on 5001
avoids that conflict entirely instead of asking everyone to turn AirPlay
Receiver off.
"""

import glob
import os
import platform
import threading
from datetime import date

from flask import Flask, render_template, request, jsonify, send_file, abort, redirect, url_for

import invoice_engine as engine

app = Flask(__name__)

# Excel/Mail.app automation (xlwings + AppleScript) can only run one at a
# time on this machine — serialize those actions so two people clicking
# Print at once don't collide.
_automation_lock = threading.Lock()

# The next document number is read from Invoice_template.xlsx's B1 cell,
# then the save writes the bumped number back to that same cell before
# building the actual invoice/estimate file. Two people saving at the same
# moment could otherwise both read the same "next" number — one save's DB
# row would be silently dropped (INSERT OR IGNORE on the UNIQUE number
# column) while its Excel file still got written to disk, and the two
# concurrent template read/writes could corrupt the template file itself.
# Serializing the whole fetch-number + save step closes that race for this
# single-process dev server (the only way this app is actually run).
_save_lock = threading.Lock()

IS_MAC = platform.system() == "Darwin"

PLACEHOLDER_INFO = {}


# ─── Helpers ──────────────────────────────────────────────────────────────

def invoice_file_path(customer_name, number):
    """Deterministic path for a document saved by the current dashboard —
    save_document()/update_document() write to
    INV_DIR/'{customer} #{number}.xlsx' (section 27). Older files saved
    before that folder existed may still sit directly in BASE_DIR or one of
    the other legacy folders — see locate_invoice_file() below, which
    checks all of them.
    """
    return os.path.join(engine.INV_DIR, f"{engine.safe_filename_component(customer_name)} #{number}.xlsx")


def locate_invoice_file(customer_name, number):
    """Best-effort lookup for invoices that may live in INV_DIR ('Invoice
    Excel Docs', where the dashboard saves new ones as of section 27),
    BASE_DIR (the root folder — where the dashboard saved them before
    section 27, and where the older desktop app still saves its own), or
    'Invoice Excel PDFs' (a legacy location seen on the real Mac). Prefers
    an exact filename match, falls back to a '#<number>' glob.
    """
    exact = invoice_file_path(customer_name, number)
    if os.path.isfile(exact):
        return exact

    search_dirs = [
        engine.BASE_DIR,
        engine.INV_DIR,
        engine.INV_PDF_DIR,
    ]
    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        hits = glob.glob(os.path.join(d, f"*#{number}.xlsx"))
        if hits:
            return hits[0]
        hits = glob.glob(os.path.join(d, f"*#{number}.pdf"))
        if hits:
            return hits[0]
    return None


def parse_line_items(raw):
    """raw: list of {"detail": str, "price": number|str} from the client.
    Price is passed through as typed — engine.parse_amount() decides
    whether it's a firm number or free text (e.g. a price range), so this
    stays agnostic to document type.
    """
    items = []
    for row in raw or []:
        detail = (row.get("detail") or "").strip()
        price_raw = row.get("price")
        if isinstance(price_raw, str):
            price_raw = price_raw.strip()
        if detail or price_raw not in (None, ""):
            items.append((detail, price_raw))
    return items


def validate_prices_for_type(doc_type, line_items):
    """Invoices need a firm number — price ranges are an estimate/quote
    thing. Returns an error string, or None if everything checks out."""
    if doc_type != "invoice":
        return None
    for _, raw in line_items:
        num, disp = engine.parse_amount(raw)
        if disp is not None:
            return (f"Invoices need a firm price — “{disp}” isn't a number. "
                     "Price ranges are for Estimates/Quotes.")
    return None


def numeric_total(line_items):
    return sum((engine.parse_amount(raw)[0] or 0.0) for _, raw in line_items)


def default_entry_date(month):
    """Default date for the "Add Entry" date field on the Monthly/Payroll
    Log. Always defaulting to the real today's date was misleading once the
    page's own month picker was changed to a different month — the add-entry
    date field kept showing today regardless, so an entry added while
    looking at, say, June could silently land in August's log instead (each
    entry's month is derived from its own date, not from whichever month
    happens to be on screen). If the month being viewed is the real current
    month, today is still the right default (correct day-of-month too);
    otherwise default to the 1st of the viewed month so the date field
    always starts inside the month the user is actually looking at."""
    if month == date.today().isoformat()[:7]:
        return date.today().isoformat()
    return f"{month}-01"


# ─── Page routes (server-rendered shell) ───────────────────────────────────

@app.route("/")
def index():
    return home()


@app.route("/home")
def home():
    conn = engine.open_db()
    try:
        stats = engine.home_stats(conn)
    finally:
        conn.close()
    return render_template("home.html", active="home", stats=stats)


@app.route("/invoices")
def invoices_list():
    q = request.args.get("q", "").strip()
    conn = engine.open_db()
    try:
        rows = engine.list_documents(conn, doc_types=("invoice",), query=q or None)
        bulk_email_to = engine.get_setting(conn, "bulk_email_to", "") or ""
    finally:
        conn.close()
    return render_template("invoices_list.html", active="invoices", rows=rows, q=q,
                            bulk_email_to=bulk_email_to)


@app.route("/estimates")
def estimates_list():
    q = request.args.get("q", "").strip()
    conn = engine.open_db()
    try:
        rows = engine.list_documents(conn, doc_types=("estimate", "quote"), query=q or None)
    finally:
        conn.close()
    return render_template("estimates_list.html", active="estimates", rows=rows, q=q)


@app.route("/customers")
def customers_list():
    q = request.args.get("q", "").strip()
    conn = engine.open_db()
    try:
        rows = engine.search_customers(conn, query=q or None)
    finally:
        conn.close()
    return render_template("customers_list.html", active="customers", rows=rows, q=q)


@app.route("/customers/new")
def customer_new_form():
    conn = engine.open_db()
    try:
        places_key = engine.get_setting(conn, "google_places_api_key", "") or ""
    finally:
        conn.close()
    return render_template("customer_form.html", active="customers", mode="new", customer=None,
                            google_places_api_key=places_key)


@app.route("/customers/<name>")
def customer_detail(name):
    conn = engine.open_db()
    try:
        customer = engine.get_customer(conn, name)
        if not customer:
            abort(404)
        stats = engine.customer_stats(conn, name)
        history = engine.list_documents(
            conn, doc_types=("invoice", "estimate", "quote"), customer=name, limit=500)
    finally:
        conn.close()
    return render_template("customer_detail.html", active="customers",
                            customer=customer, stats=stats, history=history)


@app.route("/customers/<name>/edit")
def customer_edit_form(name):
    conn = engine.open_db()
    try:
        customer = engine.get_customer(conn, name)
        places_key = engine.get_setting(conn, "google_places_api_key", "") or ""
    finally:
        conn.close()
    if not customer:
        abort(404)
    return render_template("customer_form.html", active="customers", mode="edit", customer=customer,
                            google_places_api_key=places_key)


@app.route("/monthly")
def monthly_log():
    month = request.args.get("month", "").strip() or date.today().isoformat()[:7]
    conn = engine.open_db()
    try:
        entries = engine.list_monthly_entries(conn, month)
        totals = engine.monthly_totals(conn, month)
        customers = engine.all_customers(conn)
        monthly_email_to = engine.get_setting(conn, "monthly_email_to", "") or ""
    finally:
        conn.close()
    return render_template("monthly.html", active="monthly", month=month,
                            month_label=engine.month_label(month), entries=entries,
                            totals=totals, customers=customers, default_entry_date=default_entry_date(month),
                            monthly_email_to=monthly_email_to)


@app.route("/monthly/export")
def monthly_export():
    month = request.args.get("month", "").strip() or date.today().isoformat()[:7]
    conn = engine.open_db()
    try:
        path = engine.export_monthly_xlsx(conn, month)
    finally:
        conn.close()
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/monthly/export/pdf")
def monthly_export_pdf():
    month = request.args.get("month", "").strip() or date.today().isoformat()[:7]
    if not IS_MAC:
        return ("PDF export needs Excel, so this only works when the dashboard "
                "is running on the shop Mac."), 400
    conn = engine.open_db()
    try:
        xlsx_path = engine.export_monthly_xlsx(conn, month)
    finally:
        conn.close()
    try:
        with _automation_lock:
            pdf_path = engine.generate_pdf_from_xlsx(xlsx_path)
    except Exception as e:
        return (str(e)), 500
    return send_file(pdf_path, as_attachment=True, download_name=os.path.basename(pdf_path))


@app.route("/reports")
def reports():
    conn = engine.open_db()
    try:
        series = engine.monthly_income_series(conn, months=12)
        top = engine.top_customers(conn, limit=10)
        ge = engine.ge_alltime_totals(conn)
    finally:
        conn.close()
    return render_template("reports.html", active="reports", series=series, top_customers=top, ge=ge)


@app.route("/settings")
def settings_page():
    """One place for every setting that used to be scattered across the
    page it's used on (bulk email on Invoices, report emails on Monthly/
    Payroll, Policy Number/State on Payroll, the Places API key on
    Customers) — all of it is just rows in the same settings key/value
    table, so this route is nothing more than reading all of them at once.
    """
    conn = engine.open_db()
    try:
        values = {
            "send_from_email": engine.get_setting(conn, "send_from_email", "") or "",
            "bulk_email_to": engine.get_setting(conn, "bulk_email_to", "") or "",
            "monthly_email_to": engine.get_setting(conn, "monthly_email_to", "") or "",
            "payroll_email_to": engine.get_setting(conn, "payroll_email_to", "") or "",
            "policy_number": engine.get_setting(conn, "policy_number", "") or "",
            "payroll_state": engine.get_setting(conn, "payroll_state", engine.DEFAULT_PAYROLL_STATE) or engine.DEFAULT_PAYROLL_STATE,
            "google_places_api_key": engine.get_setting(conn, "google_places_api_key", "") or "",
        }
    finally:
        conn.close()
    return render_template("settings.html", active="settings", **values)


@app.route("/payroll")
def payroll_log():
    month = request.args.get("month", "").strip() or date.today().isoformat()[:7]
    conn = engine.open_db()
    try:
        entries = engine.list_payroll_entries(conn, month)
        totals = engine.payroll_totals(conn, month)
        workers = engine.list_workers(conn)
        policy_number = engine.get_setting(conn, "policy_number", "")
        payroll_state = engine.get_setting(conn, "payroll_state", engine.DEFAULT_PAYROLL_STATE)
        payroll_email_to = engine.get_setting(conn, "payroll_email_to", "") or ""
    finally:
        conn.close()
    return render_template("payroll.html", active="payroll", month=month,
                            month_label=engine.month_label(month), entries=entries,
                            totals=totals, workers=workers, policy_number=policy_number,
                            payroll_state=payroll_state, payroll_email_to=payroll_email_to,
                            payroll_codes=engine.PAYROLL_CODES, payroll_code_order=engine.PAYROLL_CODE_ORDER,
                            default_entry_date=default_entry_date(month))


@app.route("/payroll/workers")
def payroll_workers():
    conn = engine.open_db()
    try:
        workers = engine.list_workers(conn)
    finally:
        conn.close()
    return render_template("payroll_workers.html", active="payroll", workers=workers)


@app.route("/payroll/workers/new")
def worker_new_form():
    return render_template("worker_form.html", active="payroll", mode="new", worker=None)


@app.route("/payroll/workers/<name>/edit")
def worker_edit_form(name):
    conn = engine.open_db()
    try:
        worker = engine.get_worker(conn, name)
    finally:
        conn.close()
    if not worker:
        abort(404)
    return render_template("worker_form.html", active="payroll", mode="edit", worker=worker)


@app.route("/payroll/export")
def payroll_export():
    month = request.args.get("month", "").strip() or date.today().isoformat()[:7]
    conn = engine.open_db()
    try:
        path = engine.export_payroll_xlsx(conn, month)
    finally:
        conn.close()
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/payroll/export/pdf")
def payroll_export_pdf():
    month = request.args.get("month", "").strip() or date.today().isoformat()[:7]
    if not IS_MAC:
        return ("PDF export needs Excel, so this only works when the dashboard "
                "is running on the shop Mac."), 400
    conn = engine.open_db()
    try:
        xlsx_path = engine.export_payroll_xlsx(conn, month)
    finally:
        conn.close()
    try:
        with _automation_lock:
            pdf_path = engine.generate_pdf_from_xlsx(xlsx_path)
    except Exception as e:
        return (str(e)), 500
    return send_file(pdf_path, as_attachment=True, download_name=os.path.basename(pdf_path))


@app.route("/documents/new")
def document_new():
    doc_type = request.args.get("type", "invoice")
    if doc_type not in ("invoice", "estimate", "quote"):
        doc_type = "invoice"
    convert_number = request.args.get("convert", type=int)
    customer_name = request.args.get("customer", "").strip() or None
    active = "estimates" if (convert_number or doc_type != "invoice") else "invoices"
    conn = engine.open_db()
    try:
        places_key = engine.get_setting(conn, "google_places_api_key", "") or ""
    finally:
        conn.close()
    return render_template("document_new.html", active=active,
                            doc_type=doc_type, convert_number=convert_number,
                            customer_name=customer_name, google_places_api_key=places_key)


@app.route("/documents/<int:number>/edit")
def document_edit(number):
    conn = engine.open_db()
    try:
        doc = engine.get_document(conn, number)
    finally:
        conn.close()
    if not doc:
        abort(404)
    active = "invoices" if doc["doc_type"] == "invoice" else "estimates"
    return render_template("document_new.html", active=active,
                            doc_type=doc["doc_type"], convert_number=None,
                            customer_name=None, edit_number=number)


@app.route("/invoices/new")
def invoices_new():
    # Old Phase 1 URL — kept working, just points at the generalized wizard.
    return redirect(url_for("document_new", type="invoice"))


@app.route("/<section>")
def placeholder(section):
    info = PLACEHOLDER_INFO.get(section)
    if not info:
        abort(404)
    return render_template("placeholder.html", active=section, info=info)


# ─── JSON API ───────────────────────────────────────────────────────────────

@app.route("/api/customers/search")
def api_customer_search():
    q = request.args.get("q", "")
    conn = engine.open_db()
    try:
        customers = engine.all_customers(conn)
    finally:
        conn.close()
    return jsonify(engine.fuzzy_search(customers, q))


@app.route("/api/customers/<name>", methods=["GET"])
def api_get_customer(name):
    conn = engine.open_db()
    try:
        customer = engine.get_customer(conn, name)
    finally:
        conn.close()
    if not customer:
        return jsonify({"ok": False, "error": f'No customer named "{name}" found.'}), 404
    return jsonify({"ok": True, "customer": customer})


@app.route("/api/customers", methods=["POST"])
def api_create_customer():
    """Add a customer on their own (Phase 3) — no invoice involved. Soft
    near-duplicate confirmation: if the name closely matches an existing
    customer (but isn't an exact match), the client gets a 409 with
    needs_confirm + a list of similar names, and must resend with
    confirm: true to save anyway. An exact-name duplicate is always a
    hard error — there's no "confirm" past that."""
    data = request.get_json(force=True) or {}
    name = (data.get("name") or "").strip()
    address = (data.get("address") or "").strip()
    phone = (data.get("phone") or "").strip()
    confirm = bool(data.get("confirm"))

    if not name:
        return jsonify({"ok": False, "error": "Customer name is required."}), 400

    conn = engine.open_db()
    try:
        if engine.get_customer(conn, name):
            return jsonify({"ok": False, "error": f'A customer named "{name}" already exists.'}), 409

        if not confirm:
            similar = engine.find_similar_customers(conn, name)
            if similar:
                return jsonify({
                    "ok": False,
                    "needs_confirm": True,
                    "similar": similar,
                    "error": "This looks similar to an existing customer — save anyway?",
                }), 409

        engine.create_customer(conn, name, address, phone)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 409
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()

    return jsonify({"ok": True, "customer": {"name": name, "address": engine.normalize_address(address), "phone": phone}})


@app.route("/api/customers/<name>", methods=["POST"])
def api_update_customer(name):
    """Update address/phone only — name editing is out of scope (see
    update_customer_contact's docstring)."""
    data = request.get_json(force=True) or {}
    address = (data.get("address") or "").strip()
    phone = (data.get("phone") or "").strip()
    conn = engine.open_db()
    try:
        engine.update_customer_contact(conn, name, address, phone)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "customer": {"name": name, "address": engine.normalize_address(address), "phone": phone}})


@app.route("/api/invoices/<int:number>/mark-paid", methods=["POST"])
def api_mark_invoice_paid(number):
    data = request.get_json(force=True) or {}
    category = (data.get("category") or "").strip().upper()
    paid_date = (data.get("date") or "").strip() or None

    conn = engine.open_db()
    try:
        month = engine.mark_invoice_paid(conn, number, category, paid_date=paid_date)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "month": month, "month_label": engine.month_label(month)})


@app.route("/api/invoices/<int:number>/unmark-paid", methods=["POST"])
def api_unmark_invoice_paid(number):
    conn = engine.open_db()
    try:
        engine.unmark_invoice_paid(conn, number)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.route("/api/monthly/entries", methods=["POST"])
def api_add_monthly_entry():
    data = request.get_json(force=True) or {}
    conn = engine.open_db()
    try:
        month = engine.add_manual_monthly_entry(
            conn, data.get("entry_date"), data.get("customer"),
            data.get("category"), data.get("amount"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "month": month})


@app.route("/api/monthly/entries/<int:entry_id>/update", methods=["POST"])
def api_update_monthly_entry(entry_id):
    data = request.get_json(force=True) or {}
    conn = engine.open_db()
    try:
        month = engine.update_monthly_entry(
            conn, entry_id, data.get("entry_date"), data.get("customer"),
            data.get("category"), data.get("amount"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "month": month})


@app.route("/api/monthly/entries/<int:entry_id>/delete", methods=["POST"])
def api_delete_monthly_entry(entry_id):
    conn = engine.open_db()
    try:
        engine.delete_monthly_entry(conn, entry_id)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.route("/api/payroll/workers", methods=["POST"])
def api_create_worker():
    data = request.get_json(force=True) or {}
    conn = engine.open_db()
    try:
        engine.create_worker(conn, data.get("name"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 409
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.route("/api/payroll/workers/<name>", methods=["POST"])
def api_update_worker(name):
    data = request.get_json(force=True) or {}
    conn = engine.open_db()
    try:
        engine.update_worker(conn, name, data.get("name"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 409
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "name": (data.get("name") or "").strip()})


@app.route("/api/payroll/entries", methods=["POST"])
def api_add_payroll_entry():
    data = request.get_json(force=True) or {}
    conn = engine.open_db()
    try:
        month = engine.add_payroll_entry(
            conn, data.get("entry_date"), data.get("worker"), data.get("class_code"), data.get("gross"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "month": month})


@app.route("/api/payroll/entries/<int:entry_id>/update", methods=["POST"])
def api_update_payroll_entry(entry_id):
    data = request.get_json(force=True) or {}
    conn = engine.open_db()
    try:
        month = engine.update_payroll_entry(
            conn, entry_id, data.get("entry_date"), data.get("worker"), data.get("class_code"), data.get("gross"))
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True, "month": month})


@app.route("/api/payroll/entries/<int:entry_id>/delete", methods=["POST"])
def api_delete_payroll_entry(entry_id):
    conn = engine.open_db()
    try:
        engine.delete_payroll_entry(conn, entry_id)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 404
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.route("/api/settings", methods=["POST"])
def api_set_setting():
    data = request.get_json(force=True) or {}
    key = (data.get("key") or "").strip()
    value = data.get("value")
    if not key:
        return jsonify({"ok": False, "error": "Missing setting key."}), 400
    conn = engine.open_db()
    try:
        engine.set_setting(conn, key, value)
    finally:
        conn.close()
    return jsonify({"ok": True})


@app.route("/api/invoices/next-number")
def api_next_number():
    conn = engine.open_db()
    try:
        n = engine.next_invoice_number(conn)
    finally:
        conn.close()
    return jsonify({"number": n})


@app.route("/api/invoices/normalize-address", methods=["POST"])
def api_normalize_address():
    data = request.get_json(force=True) or {}
    return jsonify({"address": engine.normalize_address(data.get("address", ""))})


@app.route("/api/documents/<int:number>")
def api_get_document(number):
    conn = engine.open_db()
    try:
        doc = engine.get_document(conn, number)
    finally:
        conn.close()
    if not doc:
        return jsonify({"ok": False, "error": f"No document #{number} found."}), 404
    return jsonify({"ok": True, "document": doc})


@app.route("/api/documents", methods=["POST"])
def api_save_document():
    data = request.get_json(force=True) or {}
    customer = data.get("customer") or {}
    name = (customer.get("name") or "").strip()
    job_name = (data.get("job_name") or "").strip()
    doc_type = data.get("doc_type") or "invoice"
    if doc_type not in ("invoice", "estimate", "quote"):
        doc_type = "invoice"
    line_items = parse_line_items(data.get("line_items"))

    if not name:
        return jsonify({"ok": False, "error": "Customer name is required."}), 400
    if not job_name:
        return jsonify({"ok": False, "error": "Job name is required."}), 400
    if not line_items:
        return jsonify({"ok": False, "error": "At least one line item is required."}), 400

    price_error = validate_prices_for_type(doc_type, line_items)
    if price_error:
        return jsonify({"ok": False, "error": price_error}), 400

    customer = {
        "name": name,
        "address": engine.normalize_address(customer.get("address") or ""),
        "phone": customer.get("phone") or "",
    }

    conn = engine.open_db()
    try:
        with _save_lock:
            doc_num = engine.next_invoice_number(conn)
            path = engine.save_document(conn, customer, job_name, line_items, doc_num, doc_type=doc_type)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()

    return jsonify({
        "ok": True,
        "number": doc_num,
        "doc_type": doc_type,
        "customer": customer["name"],
        "job_name": job_name,
        "total": numeric_total(line_items),
        "filename": os.path.basename(path),
    })


@app.route("/api/documents/convert", methods=["POST"])
def api_convert_document():
    """Turn an existing estimate/quote into an invoice, keeping the same
    document number (the shared sequence already consumed it)."""
    data = request.get_json(force=True) or {}
    number = data.get("number")
    customer = data.get("customer") or {}
    name = (customer.get("name") or "").strip()
    job_name = (data.get("job_name") or "").strip()
    line_items = parse_line_items(data.get("line_items"))

    if not number:
        return jsonify({"ok": False, "error": "Missing document number."}), 400
    if not name:
        return jsonify({"ok": False, "error": "Customer name is required."}), 400
    if not job_name:
        return jsonify({"ok": False, "error": "Job name is required."}), 400
    if not line_items:
        return jsonify({"ok": False, "error": "At least one line item is required."}), 400

    price_error = validate_prices_for_type("invoice", line_items)
    if price_error:
        return jsonify({"ok": False, "error": price_error}), 400

    customer = {
        "name": name,
        "address": engine.normalize_address(customer.get("address") or ""),
        "phone": customer.get("phone") or "",
    }

    conn = engine.open_db()
    try:
        existing = engine.get_document(conn, number)
        if not existing:
            return jsonify({"ok": False, "error": f"No document #{number} found."}), 404
        path = engine.convert_to_invoice(conn, number, customer, job_name, line_items)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()

    return jsonify({
        "ok": True,
        "number": number,
        "doc_type": "invoice",
        "customer": customer["name"],
        "job_name": job_name,
        "total": numeric_total(line_items),
        "filename": os.path.basename(path),
    })


@app.route("/api/documents/<int:number>/update", methods=["POST"])
def api_update_document(number):
    """Edit an already-saved invoice/estimate/quote — customer, job name,
    and line items/pricing are all editable; the document number and
    doc_type stay fixed (converting an estimate to an invoice is its own
    dedicated action, not part of this)."""
    data = request.get_json(force=True) or {}
    customer = data.get("customer") or {}
    name = (customer.get("name") or "").strip()
    job_name = (data.get("job_name") or "").strip()
    line_items = parse_line_items(data.get("line_items"))

    if not name:
        return jsonify({"ok": False, "error": "Customer name is required."}), 400
    if not job_name:
        return jsonify({"ok": False, "error": "Job name is required."}), 400
    if not line_items:
        return jsonify({"ok": False, "error": "At least one line item is required."}), 400

    customer = {
        "name": name,
        "address": engine.normalize_address(customer.get("address") or ""),
        "phone": customer.get("phone") or "",
    }

    conn = engine.open_db()
    try:
        existing = engine.get_document(conn, number)
        if not existing:
            return jsonify({"ok": False, "error": f"No document #{number} found."}), 404
        price_error = validate_prices_for_type(existing["doc_type"], line_items)
        if price_error:
            return jsonify({"ok": False, "error": price_error}), 400
        path = engine.update_document(conn, number, customer, job_name, line_items)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    finally:
        conn.close()

    return jsonify({
        "ok": True,
        "number": number,
        "doc_type": existing["doc_type"],
        "customer": customer["name"],
        "job_name": job_name,
        "total": numeric_total(line_items),
        "filename": os.path.basename(path),
    })


@app.route("/api/documents/preview", methods=["POST"])
def api_preview_pdf():
    """Render a quick-look PDF from an in-progress (not-yet-saved) draft."""
    data = request.get_json(force=True) or {}
    customer = data.get("customer") or {}
    name = (customer.get("name") or "Customer").strip() or "Customer"
    job_name = (data.get("job_name") or "").strip()
    doc_type = data.get("doc_type") or "invoice"
    line_items = parse_line_items(data.get("line_items"))
    doc_num = data.get("number") or "____"

    customer = {
        "name": name,
        "address": engine.normalize_address(customer.get("address") or ""),
        "phone": customer.get("phone") or "",
    }
    if not line_items:
        line_items = [("", 0.0)]

    tmp_path = os.path.join(engine.BASE_DIR, f".preview-draft-{os.getpid()}.pdf")
    try:
        engine.generate_preview_pdf(customer, job_name, line_items, doc_num, doc_type=doc_type, out_path=tmp_path)
        return send_file(tmp_path, mimetype="application/pdf", as_attachment=False,
                          download_name=f"Preview #{doc_num}.pdf")
    finally:
        # send_file streams before the finally block's file removal would run in
        # some setups; guard with exists() and best-effort cleanup afterward.
        try:
            if os.path.exists(tmp_path):
                threading.Timer(5.0, lambda: os.path.exists(tmp_path) and os.remove(tmp_path)).start()
        except Exception:
            pass


@app.route("/api/invoices/action", methods=["POST"])
def api_invoice_action():
    """Print / email / open-in-Excel for an invoice that's already been saved."""
    data = request.get_json(force=True) or {}
    customer_name = (data.get("customer") or "").strip()
    number = data.get("number")
    action = data.get("action")

    if not customer_name or not number or action not in ("print", "email", "open"):
        return jsonify({"ok": False, "error": "Missing customer, number, or action."}), 400

    if not IS_MAC:
        return jsonify({
            "ok": False,
            "error": "Print/email/open need Excel and Mail.app, so this only works "
                     "when the dashboard is running on the shop Mac.",
        }), 400

    path = locate_invoice_file(customer_name, number)
    if not path:
        return jsonify({"ok": False, "error": f"Couldn't find the saved file for #{number}."}), 404

    try:
        with _automation_lock:
            if action == "print":
                engine.print_xlsx(path)
                return jsonify({"ok": True, "message": "Sent to the default printer."})
            elif action == "email":
                pdf = engine.generate_pdf_from_xlsx(path, out_dir=engine.INV_PDF_DIR)
                conn = engine.open_db()
                try:
                    from_address = (engine.get_setting(conn, "send_from_email") or "").strip() or None
                finally:
                    conn.close()
                engine.email_invoice(pdf, number, customer_name, from_address=from_address)
                return jsonify({"ok": True, "message": f"Email opened for {engine.EMAIL_TO}."})
            elif action == "open":
                import subprocess
                subprocess.run(["open", path])
                return jsonify({"ok": True, "message": "Opened in Excel."})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/invoices/bulk-email", methods=["POST"])
def api_bulk_email_invoices():
    """Send the PDF for each selected invoice as its own individual email
    to the saved "bulk email address" setting — no manual per-email "click
    send" step. The address is read from the server's own saved setting
    (not trusted from the request body) so a request can't redirect where
    real customer paperwork gets sent just by passing a different "to".
    The UI is expected to have already shown an explicit confirmation
    (count + destination) before calling this — this endpoint itself does
    not ask again."""
    data = request.get_json(force=True) or {}
    items = data.get("items") or []
    if not items:
        return jsonify({"ok": False, "error": "No invoices selected."}), 400

    if not IS_MAC:
        return jsonify({
            "ok": False,
            "error": "Emailing needs Excel and Mail.app, so this only works "
                     "when the dashboard is running on the shop Mac.",
        }), 400

    conn = engine.open_db()
    try:
        to_address = (engine.get_setting(conn, "bulk_email_to") or "").strip()
        from_address = (engine.get_setting(conn, "send_from_email") or "").strip() or None
    finally:
        conn.close()
    if not to_address:
        return jsonify({"ok": False, "error": "Set a bulk email address first, then try again."}), 400

    sent, failed = [], []
    with _automation_lock:
        for item in items:
            customer_name = (item.get("customer") or "").strip()
            number = item.get("number")
            if not customer_name or not number:
                failed.append({"number": number, "customer": customer_name, "error": "Missing customer or number."})
                continue
            try:
                path = locate_invoice_file(customer_name, number)
                if not path:
                    raise RuntimeError(f"Couldn't find the saved file for #{number}.")
                pdf = engine.generate_pdf_from_xlsx(path, out_dir=engine.INV_PDF_DIR)
                engine.email_invoice_auto(pdf, number, customer_name, to_address, from_address=from_address)
                sent.append({"number": number, "customer": customer_name})
            except Exception as e:
                failed.append({"number": number, "customer": customer_name, "error": str(e)})

    return jsonify({"ok": True, "to": to_address, "sent": sent, "failed": failed})


@app.route("/api/monthly/email-pdf", methods=["POST"])
def api_email_monthly_pdf():
    """Email the current month's Monthly Log as a PDF to the saved
    'monthly_email_to' address — sent immediately, after the UI has
    already shown an explicit confirmation (month + destination)."""
    data = request.get_json(force=True) or {}
    month = (data.get("month") or "").strip() or date.today().isoformat()[:7]

    if not IS_MAC:
        return jsonify({
            "ok": False,
            "error": "Emailing needs Excel and Mail.app, so this only works "
                     "when the dashboard is running on the shop Mac.",
        }), 400

    conn = engine.open_db()
    try:
        to_address = (engine.get_setting(conn, "monthly_email_to") or "").strip()
        from_address = (engine.get_setting(conn, "send_from_email") or "").strip() or None
        xlsx_path = engine.export_monthly_xlsx(conn, month)
    finally:
        conn.close()
    if not to_address:
        return jsonify({"ok": False, "error": "Set an email address first, then try again."}), 400

    month_label = engine.month_label(month)
    try:
        with _automation_lock:
            pdf_path = engine.generate_pdf_from_xlsx(xlsx_path)
            engine.email_report_auto(pdf_path, f"Monthly Log — {month_label}", to_address, from_address=from_address)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    return jsonify({"ok": True, "to": to_address, "month_label": month_label})


@app.route("/api/payroll/email-pdf", methods=["POST"])
def api_email_payroll_pdf():
    """Email the current month's Payroll Report as a PDF to the saved
    'payroll_email_to' address — sent immediately, after the UI has
    already shown an explicit confirmation (month + destination)."""
    data = request.get_json(force=True) or {}
    month = (data.get("month") or "").strip() or date.today().isoformat()[:7]

    if not IS_MAC:
        return jsonify({
            "ok": False,
            "error": "Emailing needs Excel and Mail.app, so this only works "
                     "when the dashboard is running on the shop Mac.",
        }), 400

    conn = engine.open_db()
    try:
        to_address = (engine.get_setting(conn, "payroll_email_to") or "").strip()
        from_address = (engine.get_setting(conn, "send_from_email") or "").strip() or None
        xlsx_path = engine.export_payroll_xlsx(conn, month)
    finally:
        conn.close()
    if not to_address:
        return jsonify({"ok": False, "error": "Set an email address first, then try again."}), 400

    month_label = engine.month_label(month)
    try:
        with _automation_lock:
            pdf_path = engine.generate_pdf_from_xlsx(xlsx_path)
            engine.email_report_auto(pdf_path, f"Payroll Report — {month_label}", to_address, from_address=from_address)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500
    return jsonify({"ok": True, "to": to_address, "month_label": month_label})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False, threaded=True)
