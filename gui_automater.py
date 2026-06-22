#!/usr/bin/env python3
"""Invoice Maker — Funky's Electrical"""

import os, sqlite3, subprocess, threading
from datetime import date
import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
import customtkinter as ctk
from rapidfuzz import process, fuzz

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DB_PATH   = os.path.join(BASE_DIR, "invoices.db")
INV_DIR   = os.path.join(BASE_DIR, "Invoice Excel Docs")
TEMPLATE  = os.path.join(BASE_DIR, "Invoice_template.xlsx")

# ─── Database ──────────────────────────────────────────────────────────────────

def open_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS customers(
        id INTEGER PRIMARY KEY, name TEXT UNIQUE, address TEXT, phone TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS invoices(
        id INTEGER PRIMARY KEY, number INTEGER UNIQUE, customer TEXT,
        job TEXT, total REAL, day TEXT)""")
    conn.commit()
    return conn


def needs_migration(conn):
    return conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0] == 0


def migrate(conn, on_progress=None):
    if not os.path.isdir(INV_DIR):
        return
    files = [f for f in os.listdir(INV_DIR) if f.endswith(".xlsx")]
    total = len(files)
    for i, fname in enumerate(files):
        if on_progress:
            on_progress(i / total, fname[:50])
        try:
            wb = openpyxl.load_workbook(
                os.path.join(INV_DIR, fname), read_only=True, data_only=True)
            ws = wb.active
            raw_num = ws["B1"].value or ""
            name    = ws["B7"].value
            addr    = ws["B8"].value
            phone   = ws["B9"].value
            job     = ws["C7"].value
            price   = ws["C13"].value
            wb.close()

            if not name:
                continue

            inv_num = None
            for j, ch in enumerate(raw_num):
                if ch.isdigit():
                    try:
                        inv_num = int(raw_num[j:])
                    except ValueError:
                        pass
                    break

            conn.execute("""INSERT INTO customers(name, address, phone) VALUES(?,?,?)
                ON CONFLICT(name) DO UPDATE SET
                address=COALESCE(excluded.address, address),
                phone=COALESCE(excluded.phone, phone)""",
                (name,
                 str(addr).strip() if addr else None,
                 str(phone).strip() if phone else None))

            if inv_num:
                conn.execute("""INSERT OR IGNORE INTO invoices(number,customer,job,total,day)
                    VALUES(?,?,?,?,?)""",
                    (inv_num, name, str(job) if job else None,
                     float(price) if isinstance(price, (int, float)) else None,
                     date.today().isoformat()))
        except Exception:
            pass

    conn.commit()
    if on_progress:
        on_progress(1.0, "Done!")


def all_customers(conn):
    rows = conn.execute(
        "SELECT name, address, phone FROM customers ORDER BY name").fetchall()
    return [{"name": r[0], "address": r[1], "phone": r[2]} for r in rows]


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
    hits  = process.extract(query, names, scorer=fuzz.WRatio, limit=limit)
    result = []
    for name, score, _ in hits:
        if score >= threshold:
            c = next((x for x in customers if x["name"] == name), None)
            if c:
                result.append(c)
    return result


def save_invoice(conn, customer, job_name, line_items, inv_num):
    # ── Step 1: advance the invoice number on the template only ──────────────
    wb_tmpl = openpyxl.load_workbook(TEMPLATE)
    ws_tmpl = wb_tmpl.active
    raw = ws_tmpl["B1"].value or f"INVOICE #{inv_num - 1}"
    prefix = ""
    for j, ch in enumerate(raw):
        if ch.isdigit():
            prefix = raw[:j]
            break
    ws_tmpl["B1"] = f"{prefix}{inv_num}"
    wb_tmpl.save(TEMPLATE)   # template only ever gets the new number
    wb_tmpl.close()

    # ── Step 2: build the actual invoice from the freshly-saved template ─────
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb.active

    ws["B7"] = customer["name"]
    ws["B8"] = normalize_address(customer.get("address") or "")
    ws["B9"] = customer.get("phone") or ""
    ws["C7"] = job_name

    # First line item goes in the existing detail row
    if line_items:
        ws["B13"] = line_items[0][0]
        ws["C13"] = line_items[0][1]

    # Insert extra rows before SUBTOTAL
    n_extra = len(line_items) - 1
    if n_extra > 0:
        subtotal_row = 14
        for row in ws.iter_rows(min_row=13, max_row=40):
            for cell in row:
                if cell.value == "SUBTOTAL":
                    subtotal_row = cell.row

        ws.insert_rows(subtotal_row, n_extra)
        for i, (detail, price) in enumerate(line_items[1:]):
            ws.cell(row=subtotal_row + i, column=2).value = detail
            ws.cell(row=subtotal_row + i, column=3).value = price

        # Expand table bounds so SUBTOTAL formula covers new rows
        for tbl in ws.tables.values():
            mc, mr, xc, xr = range_boundaries(tbl.ref)
            if xr >= subtotal_row - 1:
                tbl.ref = (f"{get_column_letter(mc)}{mr}:"
                           f"{get_column_letter(xc)}{xr + n_extra}")

    out = os.path.join(BASE_DIR, f"{customer['name']} #{inv_num}.xlsx")
    wb.save(out)   # new invoice file only — template is untouched

    total = sum(p for _, p in line_items)
    conn.execute("""INSERT INTO customers(name, address, phone) VALUES(?,?,?)
        ON CONFLICT(name) DO UPDATE SET
        address=excluded.address, phone=excluded.phone""",
        (customer["name"], customer.get("address"), customer.get("phone")))
    conn.execute("""INSERT OR IGNORE INTO invoices(number,customer,job,total,day)
        VALUES(?,?,?,?,?)""",
        (inv_num, customer["name"], job_name, total, date.today().isoformat()))
    conn.commit()
    return out


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
        # "Street, City, ST, Zip" → "Street\nCity, ST Zip"
        street = parts[0]
        city   = parts[1]
        state  = parts[2].upper()
        zip_c  = parts[3]
        return f"{street}\n{city}, {state} {zip_c}"
    elif len(parts) >= 2:
        # "Street, City ST Zip"  or  "Street, City, ST Zip"
        # Split at the very first comma; everything after becomes line 2
        idx = raw.index(",")
        return f"{raw[:idx].strip()}\n{raw[idx + 1:].strip()}"
    return raw


# ─── Print via Excel ───────────────────────────────────────────────────────────

def print_xlsx(xlsx_path):
    """Open xlsx in Microsoft Excel and send to the default printer.

    Uses Excel's 'print out' AppleScript command — prints silently with
    correct formatting, no random symbols.
    """
    abs_path = os.path.abspath(xlsx_path)
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


# ─── UI constants ──────────────────────────────────────────────────────────────

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

H1    = ("Helvetica", 24, "bold")
H2    = ("Helvetica", 13)
LBL   = ("Helvetica", 12)
SM    = ("Helvetica", 11)
GREEN = "#2ecc71"
RED   = "#e74c3c"
GRAY  = "#3a3a3a"

# ─── App shell ─────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Invoice Maker — Funky's Electrical")
        self.geometry("760x580")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        self.conn          = open_db()
        self.customers     = []
        self.customer      = None
        self.invoice_path  = None   # saved .xlsx path

        self._frames = {}
        for Cls in (LoadingFrame, SearchFrame, NewCustomerFrame,
                    JobFrame, ConfirmFrame, DoneFrame):
            f = Cls(self)
            f.place(relx=0, rely=0, relwidth=1, relheight=1)
            self._frames[Cls.KEY] = f

        self.show("loading")
        threading.Thread(target=self._startup, daemon=True).start()

    def show(self, key):
        frame = self._frames[key]
        frame.tkraise()
        if hasattr(frame, "on_show"):
            self.after(0, frame.on_show)

    def _startup(self):
        # Open a thread-local connection; SQLite objects can't cross thread boundaries
        conn = open_db()
        if needs_migration(conn):
            self._frames["loading"].set_msg("Importing invoice history from Excel files…")
            migrate(conn, on_progress=self._frames["loading"].set_progress)
        self.customers = all_customers(conn)
        conn.close()
        self.after(0, lambda: self.show("search"))


# ─── Loading ───────────────────────────────────────────────────────────────────

class LoadingFrame(ctk.CTkFrame):
    KEY = "loading"

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        ctk.CTkLabel(self, text="Invoice Maker", font=("Helvetica", 30, "bold")).pack(pady=(90, 4))
        ctk.CTkLabel(self, text="Funky's Electrical", font=H2, text_color="gray").pack()
        self._bar = ctk.CTkProgressBar(self, width=400)
        self._bar.pack(pady=(50, 8))
        self._bar.set(0)
        self._lbl = ctk.CTkLabel(self, text="Starting…", font=SM, text_color="gray")
        self._lbl.pack()

    def set_progress(self, v, msg=""):
        self.after(0, lambda: (self._bar.set(v), self._lbl.configure(text=msg)))

    def set_msg(self, msg):
        self.after(0, lambda: self._lbl.configure(text=msg))


# ─── Customer search ───────────────────────────────────────────────────────────

class SearchFrame(ctk.CTkFrame):
    KEY = "search"

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.app = parent

        ctk.CTkLabel(self, text="New Invoice", font=H1).pack(pady=(36, 4))
        ctk.CTkLabel(self, text="Search by customer name", font=H2, text_color="gray").pack(pady=(0, 16))

        self._entry = ctk.CTkEntry(self, width=440, height=40,
                                   placeholder_text="Start typing a name…", font=LBL)
        self._entry.pack()
        self._entry.bind("<KeyRelease>", lambda _: self._refresh())

        self._scroll = ctk.CTkScrollableFrame(self, width=580, height=230, label_text="")
        self._scroll.pack(pady=8, padx=90)

        ctk.CTkButton(self, text="+ New Customer", width=180, fg_color=GRAY,
                      hover_color="#555", command=self._go_new).pack(pady=4)

    def on_show(self):
        self._entry.delete(0, "end")
        self._refresh()
        self._entry.focus()

    def _refresh(self):
        q    = self._entry.get()
        hits = fuzzy_search(self.app.customers, q)
        for w in self._scroll.winfo_children():
            w.destroy()
        if not hits:
            ctk.CTkLabel(self._scroll, text="No matches.", text_color="gray", font=SM).pack(pady=10)
            return
        for cust in hits:
            row = ctk.CTkFrame(self._scroll, fg_color="#2b2b2b", corner_radius=8)
            row.pack(fill="x", pady=3, padx=2)
            info = ctk.CTkFrame(row, fg_color="transparent")
            info.pack(side="left", fill="x", expand=True, padx=12, pady=8)
            ctk.CTkLabel(info, text=cust["name"],
                         font=("Helvetica", 12, "bold")).pack(anchor="w")
            ctk.CTkLabel(info, text=cust.get("address") or "",
                         font=SM, text_color="gray").pack(anchor="w")
            ctk.CTkLabel(info, text=cust.get("phone") or "",
                         font=SM, text_color="gray").pack(anchor="w")
            ctk.CTkButton(row, text="Select →", width=88,
                          command=lambda c=cust: self._select(c)).pack(side="right", padx=12)

    def _select(self, cust):
        self.app.customer = cust
        self.app.show("job")

    def _go_new(self):
        self.app._frames["new_customer"].setup(self._entry.get())
        self.app.show("new_customer")


# ─── New customer form ─────────────────────────────────────────────────────────

class NewCustomerFrame(ctk.CTkFrame):
    KEY = "new_customer"

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.app = parent

        ctk.CTkLabel(self, text="New Customer", font=H1).pack(pady=(40, 4))
        ctk.CTkLabel(self, text="No existing record — enter contact details",
                     font=H2, text_color="gray").pack(pady=(0, 24))

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack()

        fields = [
            ("Full Name",  "First Last"),
            ("Address",    "123 Main St, City, ST 00000"),
            ("Phone",      "(850) 555-1234"),
        ]
        self._entries = {}
        for i, (label, ph) in enumerate(fields):
            ctk.CTkLabel(form, text=label, font=LBL, width=90, anchor="w").grid(
                row=i, column=0, padx=12, pady=8, sticky="w")
            e = ctk.CTkEntry(form, width=340, placeholder_text=ph)
            e.grid(row=i, column=1, padx=12, pady=8)
            self._entries[label] = e

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=24)
        ctk.CTkButton(btns, text="← Back", width=100, fg_color="transparent",
                      border_width=1,
                      command=lambda: self.app.show("search")).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Continue →", width=140,
                      command=self._submit).pack(side="left", padx=10)

    def setup(self, prefill=""):
        self._entries["Full Name"].delete(0, "end")
        self._entries["Full Name"].insert(0, prefill)
        self._entries["Address"].delete(0, "end")
        self._entries["Phone"].delete(0, "end")

    def _submit(self):
        name    = self._entries["Full Name"].get().strip()
        addr_raw = self._entries["Address"].get().strip()
        phone   = self._entries["Phone"].get().strip()
        self.app.customer = {
            "name": name,
            "address": normalize_address(addr_raw),
            "phone": phone,
        }
        self.app.show("job")


# ─── Job details ───────────────────────────────────────────────────────────────

class JobFrame(ctk.CTkFrame):
    KEY = "job"

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.app   = parent
        self._rows = []   # (detail_var, price_var, frame)

        self._subheader = ctk.CTkLabel(self, text="", font=H2, text_color="gray")

        ctk.CTkLabel(self, text="Job Details", font=H1).pack(pady=(30, 4))
        self._subheader.pack(pady=(0, 16))

        # Job name row
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=80)
        ctk.CTkLabel(top, text="Job Name", font=LBL, width=88, anchor="w").pack(side="left")
        self._job_name = ctk.CTkEntry(top, width=390,
                                      placeholder_text="e.g. Generator Maintenance")
        self._job_name.pack(side="left", padx=8)

        # Line items header
        ih = ctk.CTkFrame(self, fg_color="transparent")
        ih.pack(fill="x", padx=80, pady=(14, 2))
        ctk.CTkLabel(ih, text="Line Items", font=("Helvetica", 12, "bold")).pack(side="left")
        ctk.CTkButton(ih, text="+ Add Row", width=88, height=26, font=SM,
                      command=self._add_row).pack(side="right")

        self._rows_frame = ctk.CTkScrollableFrame(self, height=168)
        self._rows_frame.pack(fill="x", padx=80, pady=4)

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=14)
        ctk.CTkButton(btns, text="← Back", width=100, fg_color="transparent",
                      border_width=1,
                      command=lambda: self.app.show("search")).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Review →", width=140,
                      command=self._submit).pack(side="left", padx=10)

    def on_show(self):
        if self.app.customer:
            self._subheader.configure(text=f"Customer: {self.app.customer['name']}")
        self._job_name.configure(border_color=["gray75", "gray30"])
        self._job_name.delete(0, "end")
        for _, _, f in self._rows:
            f.destroy()
        self._rows.clear()
        self._add_row()

    def _add_row(self):
        idx = len(self._rows)
        f   = ctk.CTkFrame(self._rows_frame, fg_color="#2b2b2b", corner_radius=6)
        f.pack(fill="x", pady=3, padx=2)

        ctk.CTkLabel(f, text=f"{idx + 1}.", font=SM, width=22).pack(side="left", padx=(8, 0))
        dv = ctk.StringVar()
        pv = ctk.StringVar()
        ctk.CTkEntry(f, textvariable=dv, width=290,
                     placeholder_text="Description").pack(side="left", padx=6, pady=6)
        ctk.CTkLabel(f, text="$", font=LBL).pack(side="left")
        ctk.CTkEntry(f, textvariable=pv, width=90,
                     placeholder_text="0.00").pack(side="left", padx=(2, 6), pady=6)

        def rm(frame=f):
            frame.destroy()
            self._rows = [(d, p, fr) for d, p, fr in self._rows if fr is not frame]

        ctk.CTkButton(f, text="✕", width=28, height=28, fg_color="#555",
                      hover_color=RED, command=rm).pack(side="left", padx=4)
        self._rows.append((dv, pv, f))

    def _submit(self):
        job_name = self._job_name.get().strip()
        if not job_name:
            self._job_name.configure(border_color=RED)
            return

        items = []
        for dv, pv, _ in self._rows:
            d = dv.get().strip()
            p_str = pv.get().strip()
            if d or p_str:
                try:
                    p = float(p_str) if p_str else 0.0
                except ValueError:
                    p = 0.0
                items.append((d, p))

        if not items:
            return

        inv_num = next_invoice_number(self.app.conn)
        self.app._frames["confirm"].setup(self.app.customer, job_name, items, inv_num)
        self.app.show("confirm")


# ─── Confirm ───────────────────────────────────────────────────────────────────

class ConfirmFrame(ctk.CTkFrame):
    KEY = "confirm"

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.app    = parent
        self._data  = None
        self._dyn   = []   # dynamically built widgets

    def setup(self, customer, job_name, items, inv_num):
        self._data = (customer, job_name, items, inv_num)
        for w in self._dyn:
            w.destroy()
        self._dyn.clear()

        def add(w):
            self._dyn.append(w)
            return w

        add(ctk.CTkLabel(self, text="Confirm Invoice", font=H1)).pack(pady=(28, 4))
        add(ctk.CTkLabel(self, text=f"Invoice #{inv_num}",
                         font=H2, text_color="#5dade2")).pack(pady=(0, 14))

        card = add(ctk.CTkFrame(self, fg_color="#2b2b2b", corner_radius=12))
        card.pack(fill="x", padx=80, pady=4)

        # Customer section
        ctk.CTkLabel(card, text="CUSTOMER",
                     font=("Helvetica", 10), text_color="gray").pack(anchor="w", padx=16, pady=(12, 2))
        ctk.CTkLabel(card, text=customer["name"],
                     font=("Helvetica", 13, "bold")).pack(anchor="w", padx=16)
        ctk.CTkLabel(card, text=customer.get("address") or "",
                     font=SM, text_color="gray").pack(anchor="w", padx=16)
        ctk.CTkLabel(card, text=customer.get("phone") or "",
                     font=SM, text_color="gray").pack(anchor="w", padx=16, pady=(0, 8))

        ctk.CTkFrame(card, height=1, fg_color="#444").pack(fill="x", padx=16)

        # Job section
        ctk.CTkLabel(card, text="JOB",
                     font=("Helvetica", 10), text_color="gray").pack(anchor="w", padx=16, pady=(8, 2))
        ctk.CTkLabel(card, text=job_name,
                     font=("Helvetica", 13, "bold")).pack(anchor="w", padx=16, pady=(0, 6))

        for detail, price in items:
            r = ctk.CTkFrame(card, fg_color="transparent")
            r.pack(fill="x", padx=16, pady=1)
            ctk.CTkLabel(r, text=f"  • {detail}", font=SM).pack(side="left")
            ctk.CTkLabel(r, text=f"${price:,.2f}", font=SM, text_color=GREEN).pack(side="right")

        total = sum(p for _, p in items)
        ctk.CTkFrame(card, height=1, fg_color="#444").pack(fill="x", padx=16, pady=6)
        tr = ctk.CTkFrame(card, fg_color="transparent")
        tr.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkLabel(tr, text="TOTAL", font=("Helvetica", 12, "bold")).pack(side="left")
        ctk.CTkLabel(tr, text=f"${total:,.2f}",
                     font=("Helvetica", 12, "bold"), text_color=GREEN).pack(side="right")

        btns = add(ctk.CTkFrame(self, fg_color="transparent"))
        btns.pack(pady=18)
        ctk.CTkButton(btns, text="← Edit", width=100, fg_color="transparent",
                      border_width=1,
                      command=lambda: self.app.show("job")).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Save Invoice ✓", width=160,
                      fg_color="#27ae60", hover_color="#229954",
                      command=self._save).pack(side="left", padx=10)

    def _save(self):
        customer, job_name, items, inv_num = self._data
        path = save_invoice(self.app.conn, customer, job_name, items, inv_num)
        self.app.invoice_path = path
        self.app.customers = all_customers(self.app.conn)
        self.app.show("done")


# ─── Done ──────────────────────────────────────────────────────────────────────

class DoneFrame(ctk.CTkFrame):
    KEY = "done"

    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self.app = parent

        ctk.CTkLabel(self, text="✓  Invoice Saved!",
                     font=("Helvetica", 26, "bold"), text_color=GREEN).pack(pady=(80, 8))
        ctk.CTkLabel(self, text="The invoice has been saved successfully.",
                     font=H2, text_color="gray").pack()

        self._status = ctk.CTkLabel(self, text="", font=SM, text_color="gray")
        self._status.pack(pady=8)

        btns = ctk.CTkFrame(self, fg_color="transparent")
        btns.pack(pady=20)

        ctk.CTkButton(btns, text="🖨  Print", width=160,
                      command=self._print).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="📂  Open Excel", width=160, fg_color=GRAY,
                      hover_color="#555", command=self._open).pack(side="left", padx=10)

        ctk.CTkButton(self, text="New Invoice", width=160,
                      command=self._new).pack(pady=8)

    def on_show(self):
        self._status.configure(text="", text_color="gray")

    def _print(self):
        if not self.app.invoice_path:
            return
        self._status.configure(text="Sending to printer via Excel…", text_color="gray")
        self.update()
        try:
            print_xlsx(self.app.invoice_path)
            self._status.configure(text="Sent to default printer.", text_color=GREEN)
        except Exception as e:
            self._status.configure(text=f"Print error: {e}", text_color=RED)

    def _open(self):
        if self.app.invoice_path:
            subprocess.run(["open", self.app.invoice_path])

    def _new(self):
        self.app.invoice_path = None
        self.app.show("search")


# ─── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
