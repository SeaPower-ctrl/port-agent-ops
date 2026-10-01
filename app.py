#!/usr/bin/env python3
"""
Port Agent Ops - DO Tracker (v2, with logins)
------------------------------------------------
Same shared Delivery Order board as before, now with:
  - Login required to view or use the board.
  - First time the app ever runs, it asks you to create the first
    Admin account (that's you).
  - Admins can add more staff accounts (Settings > Manage Users).
  - Mobile-friendly layout - works fine on a phone browser.

Run:
    pip install -r requirements.txt
    python app.py   (or: py app.py on Windows)

Then open http://localhost:5000
"""

import os
import re
import csv
import io
from datetime import datetime
from functools import wraps
from flask import Flask, request, jsonify, g, render_template_string, session, redirect, url_for, send_file
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
import xlrd
import psycopg2
import psycopg2.extras

DATABASE_URL = os.environ.get("DATABASE_URL", "")

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET_KEY", "change-this-secret-key-later")

# Shared Sea Power logo (used in-page and as the browser-tab favicon).
LOGO_B64 = "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"


class DBWrapper:
    """Thin wrapper so the rest of the app can keep using SQLite-style
    '?' placeholders and db.execute(...).fetchone()/fetchall(), while
    actually talking to Postgres underneath."""

    def __init__(self, conn):
        self.conn = conn

    def execute(self, query, params=()):
        cur = self.conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(query.replace("?", "%s"), params)
        return cur

    def commit(self):
        self.conn.commit()

    def close(self):
        self.conn.close()


def get_db():
    if "db" not in g:
        conn = psycopg2.connect(DATABASE_URL)
        g.db = DBWrapper(conn)
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    conn = psycopg2.connect(DATABASE_URL)
    cur = conn.cursor()
    cur.execute(
        """CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT DEFAULT 'staff',
            created_at TEXT DEFAULT ''
        )"""
    )
    cur.execute(
        """CREATE TABLE IF NOT EXISTS records (
            bl_number TEXT PRIMARY KEY,
            consignee TEXT DEFAULT '',
            port TEXT DEFAULT '',
            vessel TEXT DEFAULT '',
            invoice_issued INTEGER DEFAULT 0,
            invoice_by TEXT DEFAULT '',
            invoice_at TEXT DEFAULT '',
            approval_received INTEGER DEFAULT 0,
            approval_by TEXT DEFAULT '',
            approval_at TEXT DEFAULT '',
            do_issued INTEGER DEFAULT 0,
            do_by TEXT DEFAULT '',
            do_at TEXT DEFAULT '',
            remarks TEXT DEFAULT '',
            created_at TEXT DEFAULT ''
        )"""
    )
    # Existing databases (already deployed) won't have these columns yet -
    # add them if missing, so this upgrade doesn't require wiping the data.
    cur.execute("ALTER TABLE records ADD COLUMN IF NOT EXISTS port TEXT DEFAULT ''")
    cur.execute("ALTER TABLE records ADD COLUMN IF NOT EXISTS vessel TEXT DEFAULT ''")
    # Each record is owned by whichever staff account created it - DO Tracker
    # is per-staff (admin sees everything, staff only see their own).
    cur.execute("ALTER TABLE records ADD COLUMN IF NOT EXISTS created_by TEXT DEFAULT ''")
    cur.execute(
        """CREATE TABLE IF NOT EXISTS vessels (
            name TEXT PRIMARY KEY,
            mmsi TEXT DEFAULT '',
            updated_by TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )"""
    )
    # "operator" is set once, when the vessel is first added, and never
    # overwritten afterward - it's whoever entered the vessel originally.
    cur.execute("ALTER TABLE vessels ADD COLUMN IF NOT EXISTS operator TEXT DEFAULT ''")
    # Direct Delivery Classifier - its own table, completely separate from
    # DO Tracker's records. A BL gets a row here the moment it's classified,
    # whether or not it's ever been on the DO Tracker board.
    cur.execute(
        """CREATE TABLE IF NOT EXISTS direct_delivery (
            bl_number TEXT PRIMARY KEY,
            is_direct INTEGER NOT NULL,
            reason TEXT DEFAULT '',
            classified_by TEXT DEFAULT '',
            classified_at TEXT DEFAULT ''
        )"""
    )
    # needs_review: the classifier flags a BL instead of silently trusting
    # a shaky read - a value close enough to the 30MT/12m line that a small
    # parsing slip could flip the verdict, a header where more than one
    # column plausibly looked like "the" weight column, or a weight so
    # large it got dropped by the sanity cap rather than risk misreading
    # it. None of these mean the answer is wrong - just that it's worth a
    # human glance rather than blind trust.
    cur.execute("ALTER TABLE direct_delivery ADD COLUMN IF NOT EXISTS needs_review INTEGER DEFAULT 0")
    cur.execute("ALTER TABLE direct_delivery ADD COLUMN IF NOT EXISTS review_note TEXT DEFAULT ''")
    conn.commit()
    cur.close()
    conn.close()


def any_users_exist():
    db = get_db()
    return db.execute("SELECT 1 FROM users LIMIT 1").fetchone() is not None


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not any_users_exist():
            return redirect(url_for("setup"))
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


def admin_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if session.get("role") != "admin":
            return "Admins only.", 403
        return f(*args, **kwargs)
    return wrapper


# ---------- Auth routes ----------

@app.route("/setup", methods=["GET", "POST"])
def setup():
    if any_users_exist():
        return redirect(url_for("login"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if not username or not password:
            error = "Please fill in both fields."
        else:
            db = get_db()
            db.execute(
                "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, 'admin', ?)",
                (username, generate_password_hash(password), datetime.utcnow().strftime("%Y-%m-%d %H:%M")),
            )
            db.commit()
            return redirect(url_for("login"))
    return render_template_string(SETUP_HTML, error=error)


@app.route("/login", methods=["GET", "POST"])
def login():
    if not any_users_exist():
        return redirect(url_for("setup"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("index"))
        error = "Wrong username or password."
    return render_template_string(LOGIN_HTML, error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------- Main app ----------

@app.route("/")
@login_required
def index():
    return render_template_string(HUB_HTML, username=session.get("username"), role=session.get("role"))


@app.route("/do-tracker")
@login_required
def do_tracker_page():
    return render_template_string(PAGE_HTML, username=session.get("username"), role=session.get("role"))


@app.route("/vessel-tracker")
@login_required
def vessel_tracker_page():
    return render_template_string(VESSEL_TRACKER_HTML, username=session.get("username"), role=session.get("role"))


@app.route("/kpi")
@login_required
def kpi_page():
    return render_template_string(KPI_HTML, username=session.get("username"), role=session.get("role"))


@app.route("/direct-delivery")
@login_required
def direct_delivery_page():
    return render_template_string(DIRECT_DELIVERY_HTML, username=session.get("username"), role=session.get("role"))


@app.route("/users")
@login_required
@admin_required
def users_page():
    db = get_db()
    users = db.execute("SELECT id, username, role, created_at FROM users ORDER BY created_at").fetchall()
    return render_template_string(USERS_HTML, users=users, username=session.get("username"))


@app.route("/api/users", methods=["POST"])
@login_required
@admin_required
def add_user():
    data = request.get_json(force=True)
    username = data.get("username", "").strip()
    password = data.get("password", "")
    role = data.get("role", "staff")
    if role not in ("admin", "staff"):
        role = "staff"
    if not username or not password:
        return jsonify({"error": "Missing fields"}), 400
    db = get_db()
    try:
        db.execute(
            "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)",
            (username, generate_password_hash(password), role, datetime.utcnow().strftime("%Y-%m-%d %H:%M")),
        )
        db.commit()
    except psycopg2.IntegrityError:
        db.conn.rollback()
        return jsonify({"error": "Username already exists"}), 400
    return jsonify({"ok": True})


@app.route("/api/users/<int:user_id>", methods=["DELETE"])
@login_required
@admin_required
def delete_user(user_id):
    if user_id == session.get("user_id"):
        return jsonify({"error": "Can't delete your own account while logged in"}), 400
    db = get_db()
    db.execute("DELETE FROM users WHERE id = ?", (user_id,))
    db.commit()
    return jsonify({"ok": True})


# ---------- DO Tracker API ----------

@app.route("/api/records", methods=["GET"])
@login_required
def list_records():
    db = get_db()
    if session.get("role") == "admin":
        rows = db.execute("SELECT * FROM records ORDER BY created_at DESC").fetchall()
    else:
        rows = db.execute(
            "SELECT * FROM records WHERE created_by = ? ORDER BY created_at DESC",
            (session.get("username"),),
        ).fetchall()
    return jsonify([dict(r) for r in rows])


def _parse_ts(s):
    """Parses the app's 'YYYY-MM-DD HH:MM' timestamp strings. Returns None
    for blank/unparseable values instead of raising, since plenty of older
    or in-progress records have empty *_at fields."""
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d %H:%M")
    except ValueError:
        return None


@app.route("/api/kpi", methods=["GET"])
@login_required
def kpi_data():
    """Port Agent KPI - built entirely from the DO Tracker board, no
    separate data entry. Three things a port agent's manager would actually
    want to see:
      - turnaround: how long BLs take to move through each stage, on average
      - backlog: what's stuck right now, oldest first
      - workload: how many BLs each agent is carrying (admin only - staff
        only ever see their own records anyway, so a "workload" breakdown
        for them would just be a breakdown of one)
    """
    db = get_db()
    is_admin = session.get("role") == "admin"
    if is_admin:
        rows = db.execute("SELECT * FROM records ORDER BY created_at ASC").fetchall()
    else:
        rows = db.execute(
            "SELECT * FROM records WHERE created_by = ? ORDER BY created_at ASC",
            (session.get("username"),),
        ).fetchall()
    records = [dict(r) for r in rows]

    def avg_hours(deltas):
        if not deltas:
            return None
        return round(sum(deltas) / len(deltas) / 3600.0, 1)

    invoice_hrs, approval_hrs, do_hrs = [], [], []
    now = datetime.utcnow()
    pending_invoice, pending_approval, pending_do = [], [], []
    per_agent = {}

    for r in records:
        created = _parse_ts(r.get("created_at"))
        agent = r.get("created_by") or "(unknown)"
        bucket = per_agent.setdefault(agent, {"total": 0, "complete": 0, "pending": 0, "turnarounds": []})
        bucket["total"] += 1

        complete = bool(r.get("invoice_issued") and r.get("approval_received") and r.get("do_issued"))
        if complete:
            bucket["complete"] += 1
        else:
            bucket["pending"] += 1

        if created:
            inv_at = _parse_ts(r.get("invoice_at"))
            appr_at = _parse_ts(r.get("approval_at"))
            do_at = _parse_ts(r.get("do_at"))
            if inv_at:
                invoice_hrs.append((inv_at - created).total_seconds())
            if appr_at:
                approval_hrs.append((appr_at - created).total_seconds())
            if do_at:
                do_hrs.append((do_at - created).total_seconds())
                bucket["turnarounds"].append((do_at - created).total_seconds())

            days_open = round((now - created).total_seconds() / 86400.0, 1)
            entry = {
                "bl_number": r.get("bl_number"), "port": r.get("port"), "vessel": r.get("vessel"),
                "created_by": agent, "days_open": days_open,
            }
            if not r.get("invoice_issued"):
                pending_invoice.append(entry)
            if not r.get("approval_received"):
                pending_approval.append(entry)
            if not r.get("do_issued"):
                pending_do.append(entry)

    for lst in (pending_invoice, pending_approval, pending_do):
        lst.sort(key=lambda e: -e["days_open"])

    workload = None
    if is_admin:
        workload = [
            {
                "agent": agent, "total": b["total"], "complete": b["complete"], "pending": b["pending"],
                "avg_turnaround_hours": avg_hours(b["turnarounds"]),
            }
            for agent, b in sorted(per_agent.items(), key=lambda kv: -kv[1]["total"])
        ]

    return jsonify({
        "total_bls": len(records),
        "turnaround": {
            "avg_hours_to_invoice": avg_hours(invoice_hrs),
            "avg_hours_to_approval": avg_hours(approval_hrs),
            "avg_hours_to_do": avg_hours(do_hrs),
        },
        "backlog": {
            "pending_invoice": pending_invoice[:15],
            "pending_approval": pending_approval[:15],
            "pending_do": pending_do[:15],
            "counts": {
                "pending_invoice": len(pending_invoice),
                "pending_approval": len(pending_approval),
                "pending_do": len(pending_do),
            },
        },
        "workload": workload,
    })


def _owns_record(bl_number):
    """Admins can touch any record. Staff can only touch records they
    created themselves."""
    if session.get("role") == "admin":
        return True
    db = get_db()
    row = db.execute("SELECT created_by FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
    return row is not None and row["created_by"] == session.get("username")


@app.route("/api/manifest", methods=["POST"])
@login_required
def submit_manifest():
    data = request.get_json(force=True)
    lines = data.get("lines", "")
    db = get_db()
    added = 0
    for raw in lines.splitlines():
        raw = raw.strip()
        if not raw:
            continue
        # Only the BL number matters now - if the user still pastes
        # "BL, something" (old habit), just take the BL part.
        bl_number = raw.split(",", 1)[0].strip().upper()
        consignee = ""
        if not bl_number:
            continue
        existing = db.execute("SELECT 1 FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
        if existing:
            continue
        db.execute(
            "INSERT INTO records (bl_number, consignee, created_at, created_by) VALUES (?, ?, ?, ?)",
            (bl_number, consignee, datetime.utcnow().strftime("%Y-%m-%d %H:%M"), session.get("username")),
        )
        added += 1
    db.commit()
    return jsonify({"added": added})


# Header names we'll recognize for the BL Number column in an uploaded
# manifest. Matching is case-insensitive and ignores spaces/punctuation.
# (Consignee is intentionally no longer tracked - only the BL number matters.)
BL_HEADER_WORDS = ["blnumber", "bl", "billoflading", "billofladingno", "bl no", "blno"]


def _normalize_header(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def _extract_bl_numbers_from_rows(rows, allow_no_header_fallback=True):
    """rows: a list of rows, each row a sequence of cell values (any type,
    already-stringifiable). Looks for a header naming the BL Number column
    in the first 5 rows (same recognized header words used for Excel).

    If no such header is found:
      - allow_no_header_fallback=False skips this table entirely (used for
        sheets/tables beyond the first one in a multi-sheet/multi-table
        file - real manifests are commonly one main BL-list sheet plus
        per-BL "attachment" sheets, e.g. a vehicle's chassis/engine-number
        list, which have an ITEM/serial column but no BL Number column at
        all; blindly reading their column A as BL numbers turns "1, 2, 3,
        4..." into fake BL records).
      - allow_no_header_fallback=True (the default - used for a lone
        sheet/table, or plain pasted text) falls back to column A as the
        BL number. As a safety net even then, if the values that column A
        produces are themselves just "1", "2", "3", "4", ... (a serial/
        item counter, not real BL data), this is skipped too.

    Returns a flat list of upper-cased BL number strings (not deduped)."""
    if not rows:
        return []

    bl_col = None
    header_row_index = None
    for i, row in enumerate(rows[:5]):
        for col_index, cell in enumerate(row):
            norm = _normalize_header(cell)
            # Exact match only - a loose "startswith" here used to also
            # match unrelated things like a "B/L ATTACHMENT" sheet title
            # (normalizes to "blattachment", which starts with "bl"),
            # falsely treating it as a real header and reading junk data
            # out of the column below it.
            if norm and any(norm == w.replace(" ", "") for w in BL_HEADER_WORDS):
                bl_col = col_index
                header_row_index = i
        if bl_col is not None:
            break

    if bl_col is None:
        if not allow_no_header_fallback:
            return []
        bl_col = 0
        data_rows = rows
    else:
        data_rows = rows[header_row_index + 1:]

    out = []
    for row in data_rows:
        if bl_col >= len(row):
            continue
        raw_bl = row[bl_col]
        if raw_bl is None or str(raw_bl).strip() == "":
            continue
        raw_text = str(raw_bl).strip()
        # A cell occasionally lists more than one BL number stacked on
        # separate lines inside it (e.g. a shared-contact/remarks row that
        # covers two BLs handled by the same person) - split those apart
        # rather than keeping the newline embedded in one garbled
        # "BL1\nBL2" record, which would land on the board as its own
        # fake, unmatched entry alongside the two real ones.
        for line in raw_text.splitlines():
            candidate = line.strip().upper()
            if not candidate:
                continue
            # Skip a summary row ("TOTAL:", "GRAND TOTAL", "SUB TOTAL",
            # "VOYAGE TOTAL", "PAGE TOTAL"...) sitting in the same column as
            # the real BL numbers - manifests commonly have one or more of
            # these (a running subtotal per page, plus a grand/voyage total
            # at the end), and none of them are real BLs. A substring check
            # catches every "___ TOTAL" variant rather than only the exact
            # phrases seen so far.
            norm_candidate = _normalize_header(candidate)
            candidate_check = re.sub(r"\s+", "", candidate)
            if "total" in norm_candidate or any(w in candidate_check for w in ("合计", "总计", "汇总", "小计")):
                continue
            out.append(candidate)

    if header_row_index is None:
        sample = out[:5]
        if sample == [str(n) for n in range(1, len(sample) + 1)]:
            return []

    return out


_DTRKR_NON_MANIFEST_SHEET_RE = re.compile(
    r"CONTACT|REMARK|NOTES?\b|ADDRESS", re.IGNORECASE
)


def _dtrkr_is_non_manifest_sheet_name(name):
    """A sheet whose own name marks it as per-BL reference info (a contact
    list, remarks, notes...) rather than the shipment's actual BL data.
    These sheets commonly reuse a "BL NO." column just to key their rows to
    a BL, which would otherwise look exactly like a real manifest table and
    get re-extracted as if it were one - at best adding nothing but repeat
    work (the same BL numbers already found on the main sheet), at worst
    adding garbage (one shared-contact row can cover two BLs via a
    multi-line cell, or append extra text like "BL085(085 123-125)")."""
    return bool(_DTRKR_NON_MANIFEST_SHEET_RE.search(str(name or "")))


def _extract_bl_numbers_from_lines(text):
    """Fallback for formats with no real table (a .docx with no tables, or
    a .pdf page with no detectable table/borders - common for manifests
    exported or printed without visible grid lines). One BL per non-empty
    line, taking whatever is before the first comma if present, then run
    through the same header-recognition as tabular rows so a stray header
    line like "BL Number" at the top doesn't get inserted as a record."""
    rows = []
    for raw in text.splitlines():
        raw = raw.strip()
        if not raw:
            continue
        rows.append([raw.split(",", 1)[0].strip()])
    return _extract_bl_numbers_from_rows(rows)


_ZIP_MAGIC = b"PK\x03\x04"
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _load_excel_sheets_by_content(raw_bytes, filename):
    """Reads an Excel file's sheets as [(sheet_name, rows), ...], trusting
    the file's actual bytes over its extension. A very common real-world
    mismatch: a modern .xlsx (which is really a ZIP archive) gets
    saved/forwarded/attached with an old ".xls" name - email clients and
    "Save As" dialogs do this constantly - and a strict extension check
    then hands it to the wrong library (xlrd, which only reads the old
    binary format) and fails outright with an unhelpful error, even though
    the file is perfectly readable. This checks the real file signature
    first and only falls back to the extension if the content doesn't
    look like either known format."""
    if raw_bytes[:4] == _ZIP_MAGIC:
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        return [(s.title, list(s.iter_rows(values_only=True))) for s in wb.worksheets]
    if raw_bytes[:8] == _OLE_MAGIC:
        book = xlrd.open_workbook(file_contents=raw_bytes)
        return [(s.name, [s.row_values(r) for r in range(s.nrows)]) for s in book.sheets()]
    lower = filename.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        return [(s.title, list(s.iter_rows(values_only=True))) for s in wb.worksheets]
    if lower.endswith(".xls"):
        book = xlrd.open_workbook(file_contents=raw_bytes)
        return [(s.name, [s.row_values(r) for r in range(s.nrows)]) for s in book.sheets()]
    raise ValueError(f"{filename} isn't a recognizable Excel file")


@app.route("/api/manifest/upload", methods=["POST"])
@login_required
def upload_manifest_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file received"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400

    filename = file.filename.lower()
    # The whole manifest gets tagged with the Port and Vessel it was
    # uploaded for, so the board can be organized Port > Vessel. Always
    # stored upper-case for consistency, even if the field's own
    # uppercasing-as-you-type got bypassed somehow (e.g. a pasted value).
    port = request.form.get("port", "").strip().upper()
    vessel = request.form.get("vessel", "").strip().upper()

    bl_numbers = []

    try:
        if filename.endswith((".xlsx", ".xlsm", ".xls")):
            # Go through every sheet (not just the first/active one) so BLs
            # aren't missed if the file has multiple tabs or was last saved
            # on a different sheet. Only the FIRST sheet gets the "no
            # header -> assume column A" fallback: a real-world manifest is
            # commonly one main BL-list sheet plus per-BL "attachment"
            # sheets (e.g. a vehicle's chassis/engine-number list) that
            # have their own ITEM/serial column but no BL data at all -
            # falling back on those would read "1, 2, 3, 4..." as BL
            # numbers. The sheets are read by sniffing the file's real
            # content rather than trusting its extension - a very common
            # mismatch is a modern .xlsx saved/forwarded with an old .xls
            # name, which would otherwise fail outright.
            sheets = _load_excel_sheets_by_content(file.read(), filename)
            for i, (sheet_name, rows) in enumerate(sheets):
                # A per-BL contact/remarks tab (beyond the main sheet) is
                # reference info, not shipment data - skip it entirely so it
                # can't re-add (or garble) BLs already found on the main
                # sheet. See _dtrkr_is_non_manifest_sheet_name.
                if i > 0 and _dtrkr_is_non_manifest_sheet_name(sheet_name):
                    continue
                bl_numbers.extend(_extract_bl_numbers_from_rows(rows, allow_no_header_fallback=(i == 0)))

        elif filename.endswith(".csv"):
            text = file.read().decode("utf-8-sig", errors="ignore")
            rows = list(csv.reader(io.StringIO(text)))
            bl_numbers.extend(_extract_bl_numbers_from_rows(rows))

        elif filename.endswith(".docx"):
            import docx
            document = docx.Document(file)
            if document.tables:
                for i, table in enumerate(document.tables):
                    rows = [[cell.text for cell in row.cells] for row in table.rows]
                    bl_numbers.extend(_extract_bl_numbers_from_rows(rows, allow_no_header_fallback=(i == 0)))
            else:
                full_text = "\n".join(p.text for p in document.paragraphs)
                bl_numbers.extend(_extract_bl_numbers_from_lines(full_text))

        elif filename.endswith(".pdf"):
            import pdfplumber
            all_tables = []
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    for table in (page.extract_tables() or []):
                        if table:
                            all_tables.append(table)
                if all_tables:
                    for i, table in enumerate(all_tables):
                        bl_numbers.extend(_extract_bl_numbers_from_rows(table, allow_no_header_fallback=(i == 0)))
                else:
                    for page in pdf.pages:
                        bl_numbers.extend(_extract_bl_numbers_from_lines(page.extract_text() or ""))

        else:
            return jsonify({"error": "Unsupported file type. Please upload .xlsx, .xls, .csv, .docx or .pdf."}), 400
    except Exception:
        return jsonify({"error": "Couldn't read that file. Make sure it isn't corrupted or password-protected."}), 400

    db = get_db()
    added = 0
    skipped = 0
    for bl_number in bl_numbers:
        if not bl_number:
            continue
        existing = db.execute("SELECT 1 FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
        if existing:
            skipped += 1
            continue
        db.execute(
            "INSERT INTO records (bl_number, port, vessel, created_at, created_by) VALUES (?, ?, ?, ?, ?)",
            (bl_number, port, vessel, datetime.utcnow().strftime("%Y-%m-%d %H:%M"), session.get("username")),
        )
        added += 1

    db.commit()
    return jsonify({"added": added, "skipped": skipped})


# ---------- Direct Delivery Classifier ----------
# Rule: a BL is Direct Delivery if any item in its packing list is over 30MT
# (heavy lift) or over 12m long (oversize) - UNLESS that item has wheels it
# can drive off on, or is a coil, in which case it doesn't need a low-bed
# trailer and so is NOT direct delivery even though it's heavy/oversize.
# A BL can span several sub-sheets/files in the packing list (e.g.
# ...103____1.xlsx through ...103____4.xlsx are all one BL "103"), so
# classification is done BL-wise after grouping ACROSS every file/sheet in
# one upload, not per sub-sheet. Completely separate feature from DO
# Tracker - it classifies whatever BLs are in the packing list, whether or
# not they're on the board.
#
# Real packing lists from different mills/suppliers vary wildly (English or
# Chinese headers, a per-piece table or a per-BL summary table, weight given
# per-piece or as a lot/BL total, a combined "size" column instead of a
# separate length column, a two-row header where the units (MT/PCS/etc) are
# on the row under the labels, sometimes several tables stacked in one
# sheet, letterhead/cover blocks repeated mid-sheet on multi-page exports,
# and occasionally no table at all - just labelled cells). The parsing
# below is written to cope with all of that rather than one fixed template.

DD_WEIGHT_MT_THRESHOLD = 30.0
DD_LENGTH_M_THRESHOLD = 12.0

# A single packing-list line item this heavy/long is never real - it means
# a header got misread and some unrelated number (an invoice number, a
# date serial, a lot subtotal) was picked up as if it were the weight or
# length of one piece. Anything past this is dropped rather than trusted.
DD_SANITY_MAX_WEIGHT_MT = 500.0
DD_SANITY_MAX_LENGTH_M = 100.0

# Keyword lists are matched as SUBSTRINGS of the (normalized) header text,
# not exact matches - real headers are things like "GROSS WEIGHT/MT" or
# "毛重(MT)", never a bare "weight". English keywords are matched against an
# a-z0-9-only normalization; Chinese keywords are matched against the header
# with whitespace/punctuation stripped but characters kept as-is.
DD_WEIGHT_GROSS_WORDS = ["grossweight", "grosswt", "gw", "毛重"]
DD_WEIGHT_NET_WORDS = ["netweight", "networt", "nw", "净重"]
DD_WEIGHT_GENERIC_WORDS = ["weight", "wt", "重量"]
DD_LENGTH_WORDS = ["length", "长度"]
DD_SPEC_WORDS = ["size", "spec", "specification", "dimension", "dimensions", "规格", "尺寸"]
DD_QTY_WORDS = [
    "noofpc", "noofpcs", "noofcoils", "noofcoil", "numberofcoils", "numberofcoil",
    "numberofpcs", "numberofpc", "numberofbundles", "numberofbundle", "bundles", "bundle",
    "qnty", "qty", "quantity", "pcs", "pieces", "件数", "数量", "支数",
]
DD_DESC_HEADER_WORDS = ["description", "desc", "cargo", "commodity", "itemdescription", "goodsdescription", "cargodescription", "货名", "品名", "货物名称"]
# NOTE: deliberately no bare "bl" here - as a raw substring it false-matches
# ordinary words like "TABLE" or "DOUBLE". A short B/L-style token is
# matched separately by _DD_BL_HEADER_RE below.
DD_BL_HEADER_WORDS = ["blnumber", "billoflading", "billofladingno", "blno", "提单号"]
_DD_BL_HEADER_RE = re.compile(r"^(BILL|B[./]?\s?L[./]?)\s?(NO|NUMBER|#|$)", re.IGNORECASE)

# A "TOTAL weight" column (e.g. "total(KGS)", "Total Weight(KG)", "合计重量")
# is a LOT total (qty x per-unit weight), kept separate from the per-unit
# "weight(KGS)" column some templates also carry. When both are present,
# the per-unit column must NOT be divided by qty again - see the
# weight_total cross-check in _extract_classification_rows.
_DD_TOTAL_WEIGHT_RE = re.compile(
    r"TOTAL.{0,15}(WEIGHT|WT|KGS?|吨)|(WEIGHT|WT|KGS?).{0,15}TOTAL|总重|合计重量|总计重量|总重量",
    re.IGNORECASE,
)

# Some headers state their unit explicitly ("weight(KGS)", "毛重(MT)") - that
# beats guessing the unit from the raw number's magnitude, which misreads a
# real sub-1000 KG figure (e.g. 447 kg) as if it were already in MT.
_DD_UNIT_KG_RE = re.compile(r"\bKGS?\b|千克|公斤", re.IGNORECASE)
_DD_UNIT_MT_RE = re.compile(r"\bM\.?\s?T\.?S?\b|\bTONNES?\b|\bTONS?\b|吨", re.IGNORECASE)


def _dd_header_weight_unit(cell):
    """Returns "kg"/"mt" if this header cell states its weight unit
    explicitly, else None (fall back to the magnitude heuristic)."""
    if not isinstance(cell, str) or not cell.strip():
        return None
    if _DD_UNIT_KG_RE.search(cell):
        return "kg"
    if _DD_UNIT_MT_RE.search(cell):
        return "mt"
    return None

# Exception keywords (English + common Chinese) - if the item's description
# (or the sheet's overall goods/product-description line) contains any of
# these, it's excluded from Direct Delivery even if it trips the
# weight/length threshold. Wire rod and coiled steel are always shipped as
# coils.
DD_EXCEPTION_WORDS = [
    "wheel", "wheels", "self-propelled", "self propelled", "tyre", "tire", "tyres", "tires",
    "trailer mounted", "drive off", "roll on", "coil", "coils", "wire rod", "wire rods",
    "hrc", "crc", "hot rolled coil", "cold rolled coil", "steel coil",
    "轮", "车轮", "自走", "自行", "钢卷", "卷材", "卷", "线材",
]

# Short unit tokens on a "units row" directly under the real header (e.g.
# "(MT)" under a column literally labelled "QUANTITY" - some suppliers use
# "quantity" to mean the tonnage, not a piece count). A confirmed unit
# always wins over a guess from the label text above it.
_DD_WEIGHT_UNIT_RE = re.compile(r"(?<![A-Za-z])(MT|M\s?\.?\s?T|TONNES?|TONS?|KGS?)(?![A-Za-z])", re.IGNORECASE)
_DD_QTY_UNIT_RE = re.compile(r"(?<![A-Za-z])(PCS?|SETS?|NOS?|COILS?|BDLS?|BUNDLES?)(?![A-Za-z])", re.IGNORECASE)
DD_QTY_UNIT_CHINESE = ("支数", "件数", "卷数", "支", "件")

# A header like "Package(COIL)" or "Packing(PCS)" - the unit named in the
# parentheses makes it a qty column, but a bare "PACKAGE NO." (an ID/serial
# column, not a count) must NOT match, so this needs the parenthesised
# unit specifically rather than the word "package" alone.
_DD_QTY_PAREN_RE = re.compile(r"PACKAGE\S*\s*\(\s*(COILS?|PCS?|BDLS?|BUNDLES?|SETS?|NOS?)\s*\)", re.IGNORECASE)

# Rows that are pure letterhead/cover-page metadata (invoice no, contract
# no, shipping marks, page numbers...) rather than cargo data. These show
# up ABOVE a table's real header, but on multi-page exports the whole
# cover block repeats again mid-sheet before every new page's header - if
# not skipped, whatever happens to sit in the old table's weight/length
# column positions on those rows gets misread as a cargo line.
DD_NOISE_ROW_WORDS = [
    "invoiceno", "contractno", "poinvoiceno", "pocontractno", "lcno", "lc/no",
    "shippingmark", "invoicedate", "shipmentfrom", "shipmentto",
    "portofloading", "portofdischarg", "descriptionofgoods", "shippingterms",
    "towhomitmayconcern", "foraccountandrisk", "termsofprice", "countryoforigin",
    "deliveryinvoicing", "packinglistno", "packinglistdate", "mainconslygroup",
]

# A discharge port naming one of these countries/hubs means the sheet is
# for an entirely different shipment that happens to share the workbook
# (suppliers sometimes leave old tabs for other consignees in a reused
# file) - not a Saudi-bound BL at all, so its numbers shouldn't be mixed
# into this classification.
DD_NON_SAUDI_DISCHARGE_MARKERS = [
    "IRAQ", "UMM QASR", "KUWAIT", "QATAR", "DOHA", "BAHRAIN", "OMAN", "SOHAR",
    "UAE", "DUBAI", "ABU DHABI", "JEBEL ALI", "EGYPT", "INDIA", "PAKISTAN",
]
DD_SAUDI_DISCHARGE_MARKERS = [
    "SAUDI", "KSA", "JEDDAH", "JEDDA", "DAMMAM", "JUBAIL", "YANBU", "RIYADH",
    "RAS TANURA", "DUBA", "KING ABDUL",
]


def _dd_norm_ascii(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def _dd_norm_raw(text):
    """Keeps non-ASCII characters (Chinese headers) intact, strips the
    whitespace/punctuation that varies between files."""
    return re.sub(r"[\s.\-_:：（）()/\\]+", "", str(text or ""))


def _dd_header_matches(cell, keywords):
    """True if this header cell names one of the given columns. `keywords`
    can mix plain-English words (matched as a substring of the a-z0-9
    normalization) and Chinese terms (matched as a substring with
    whitespace/punctuation stripped but characters preserved)."""
    if not isinstance(cell, str) or not cell.strip():
        return False
    ascii_norm = _dd_norm_ascii(cell)
    raw_norm = _dd_norm_raw(cell)
    for kw in keywords:
        if kw.isascii():
            k = _dd_norm_ascii(kw)
            if k and k in ascii_norm:
                return True
        elif kw in raw_norm:
            return True
    return False


def _dd_is_exception(description):
    text = str(description or "")
    text_lower = text.lower()
    for w in DD_EXCEPTION_WORDS:
        if w.isascii():
            if w.lower() in text_lower:
                return True
        elif w in text:
            return True
    return False


def _dd_parse_number(raw):
    if raw is None:
        return None
    try:
        return float(str(raw).replace(",", "").strip())
    except (TypeError, ValueError):
        m = re.search(r"[\d.]+", str(raw).replace(",", ""))
        if not m:
            return None
        try:
            return float(m.group())
        except ValueError:
            return None


def _dd_parse_weight_mt(raw, unit_hint=None):
    """Best-effort: a bare number is assumed to already be in MT if it's
    under 1000, or in KG (divided down to MT) if 1000 or over - this holds
    for realistic single-item cargo weights either way it's written. A
    sub-1000 KG figure (e.g. a 447kg item) defeats that guess, so when the
    column's own header states its unit explicitly (unit_hint, from
    _dd_header_weight_unit), that always wins over the magnitude guess.
    Note: this may still be a LOT total awaiting division by a qty column -
    the sanity cap is applied by the caller only once the final per-item
    figure is known, not here."""
    v = _dd_parse_number(raw)
    if v is None or v <= 0:
        return None
    if unit_hint == "kg":
        return v / 1000.0
    if unit_hint == "mt":
        return v
    return v / 1000.0 if v >= 1000 else v


def _dd_parse_length_m(raw):
    """Takes a number out of a dedicated length cell and scales it to
    meters: explicit "mm"/"cm" in the text wins, otherwise falls back to a
    magnitude guess (packing lists almost always give length in mm)."""
    text = str(raw or "")
    v = _dd_parse_number(text)
    if v is None or v <= 0:
        return None
    lower = text.lower()
    if "mm" in lower:
        m = v / 1000.0
    elif "cm" in lower:
        m = v / 100.0
    elif v > 1000:
        m = v / 1000.0
    elif v > 100:
        m = v / 100.0
    else:
        m = v
    return m if m <= DD_SANITY_MAX_LENGTH_M else None


def _dd_parse_spec_length_mm(raw):
    """A combined "size"/"规格" cell like "20*2440*4880" (thickness x width x
    length, mm) or "0.35*1000" (thickness x width only - a coil, no length).
    Returns the length in meters only when THREE numbers are present; two
    numbers means there's no length dimension at all (typical for coils)."""
    text = str(raw or "")
    nums = re.findall(r"[\d.]+", text.replace(",", ""))
    if len(nums) < 3:
        return None
    try:
        length_mm = float(nums[2])
    except ValueError:
        return None
    if length_mm <= 0:
        return None
    m = length_mm / 1000.0
    return m if m <= DD_SANITY_MAX_LENGTH_M else None


def _dd_bl_root(text):
    """Strips a trailing sub-sheet letter suffix off a BL/sheet name, e.g.
    "004A" -> "004", "004-I" -> "004", so every sub-sheet of one BL groups
    together. Left as-is if it doesn't end in digits-then-letters."""
    s = str(text or "").strip().upper()
    m = re.match(r"^(.*\d)[\s\-_]?[A-Z]{1,2}$", s)
    return m.group(1) if m else s


def _dd_filename_bl_hint(filename):
    """Many real packing lists don't have a BL column in the table at all -
    the BL/reference number is only ever written in the FILE NAME (e.g.
    "ZSM2603XGJD103____1.xlsx" -> the BL is "ZSM2603XGJD103", the trailing
    "____1" is just that file's part/page number). The real reference is
    always the clean token before the first run of junk - underscores,
    the file extension, or a "-1"/"-2" suffix appended after more
    underscores/words. Falls back to the bare filename stem if nothing
    looks like a clean token."""
    stem = re.sub(r"\.(xlsx|xlsm|xls|csv|pdf|docx)$", "", str(filename or ""), flags=re.IGNORECASE)
    stem = stem.strip()
    if not stem:
        return ""
    head = re.split(r"[_\s]", stem, 1)[0].strip()
    return (head or stem).upper()


_DD_GENERIC_SHEET_RE = re.compile(
    r"^(SHEET\s*\d*|PAGE\s*\d*|PACKING\s*LIST\s*\d*|COMMERCIAL\s*INVOICE|INVOICE\s*\d*|"
    r"DETAILED\s*PACKING\s*LIST\s*REFER(ENCE)?\s*\(?\d*\)?|SHIPPING\s*MARKS?|信息|箱单|装箱单|packing)$",
    re.IGNORECASE,
)


def _dd_is_generic_sheet_name(name):
    """True for a sheet/tab name that's just a template placeholder
    ("Sheet1", "PAGE 2", "INVOICE"...) rather than a real per-sheet BL or
    reference number. A workbook where every sheet is named like this is
    one BL split across pages/sections, so the file NAME should be used as
    the BL hint instead; a workbook with real per-sheet names (each one a
    distinct BL, e.g. "CH26207BJED028") should keep using those."""
    s = str(name or "").strip()
    return bool(_DD_GENERIC_SHEET_RE.match(s))


_DD_GOODS_LABEL_RE = re.compile(
    r"(?:PRODUCT\s*DESCRIPTION|DESCRIPTION\s*OF\s*GOODS|GOODS|COMMODITY|CARGO)\s*[:：]\s*([^\n\r]+)",
    re.IGNORECASE,
)
_DD_GOODS_BARE_LABEL_RE = re.compile(
    r"^(?:PRODUCT\s*DESCRIPTION|DESCRIPTION\s*OF\s*GOODS|NAME\s*OF\s*COMMODITY|GOODS|COMMODITY|CARGO)\s*:?\s*$",
    re.IGNORECASE,
)


def _dd_find_goods_line(rows, max_scan=30):
    """Packing lists sometimes name the cargo once in a label line above
    the table rather than in a per-row description column. Used as a
    sheet-wide fallback so the coil/wheeled exception can still be
    checked. Handles the label and the value being in the same cell
    ("GOODS:HOT ROLLED STEEL COIL", or buried inside a longer multi-line
    cell like "PRODUCT DESCRIPTION: HOT ROLLED STEEL COILS") AND the label
    sitting alone in one cell with the value in the next cell along in the
    same row (a common layout: col A = "DESCRIPTION OF GOODS", col C =
    the actual product name, with an empty col B between them)."""
    for row in rows[:max_scan]:
        for cell in row:
            if not isinstance(cell, str):
                continue
            m = _DD_GOODS_LABEL_RE.search(cell)
            if m:
                line = m.group(1).strip()
                if line:
                    return line
        # Label-alone-in-its-own-cell layout: take the next non-empty cell
        # in the same row as the value.
        cells = list(row)
        for idx, cell in enumerate(cells):
            if not isinstance(cell, str) or not _DD_GOODS_BARE_LABEL_RE.match(cell.strip()):
                continue
            for later in cells[idx + 1:]:
                if isinstance(later, str) and later.strip():
                    return later.strip()
    return ""


def _dd_is_totals_row(row):
    """A subtotal/grand-total row ("TOTAL", "SUBTOTAL", "合计", "总计",
    "汇总", "小计") - these carry a LOT total, not one item's figures, and
    mark the end of that table segment."""
    for cell in row:
        if not isinstance(cell, str) or not cell.strip():
            continue
        text = cell.strip()
        if re.match(r"^(GRAND\s+)?(SUB)?\s*TOTAL\s*:?\s*$", text, re.IGNORECASE):
            return True
        compact = re.sub(r"\s+", "", text)
        if compact in ("合计", "总计", "汇总", "小计"):
            return True
        return False  # only the row's first non-empty cell counts
    return False


def _dd_is_noise_row(row):
    """A pure letterhead/cover-page row (invoice no, contract no, shipping
    marks, page x/y...) rather than a cargo data row. These repeat mid-sheet
    on multi-page packing lists, between one page's table and the next."""
    cells = [c for c in row if isinstance(c, str) and c.strip()]
    if not cells:
        return False
    label_like = 0
    for c in cells:
        norm = _dd_norm_ascii(c)
        if any(w in norm for w in DD_NOISE_ROW_WORDS):
            return True
        if re.match(r"^[A-Za-z][A-Za-z /]{1,30}:\s*$", c.strip()):
            label_like += 1
    return label_like >= 2


def _dd_scan_unit_row(row):
    """Reads a "units" row sitting directly under a table's label row (e.g.
    "(MT)"/"(M)"/"(MM)" or "PCS"/"COILS" under COMMODITY/LENGTH/QUANTITY
    labels). Returns {column_index: "weight"|"qty"} for confirmed unit
    tokens only - this is used to override an ambiguous label-only guess
    such as a "QUANTITY" column that's actually the tonnage."""
    out = {}
    for col_index, cell in enumerate(row):
        if not isinstance(cell, str) or not cell.strip():
            continue
        text = re.sub(r"[.,]", "", cell.strip())
        if _DD_WEIGHT_UNIT_RE.search(text):
            out[col_index] = "weight"
        elif _DD_QTY_UNIT_RE.search(text) or any(w in cell for w in DD_QTY_UNIT_CHINESE):
            out[col_index] = "qty"
    return out


def _dd_sheet_is_non_saudi(rows, max_scan=25):
    """Some supplier workbooks leave tabs from a completely different
    shipment/consignee mixed into the same file (a reused template). If a
    sheet's own "PORT OF DISCHARG(E/ING)" line names a non-Saudi country or
    hub, its cargo has nothing to do with this BL and shouldn't be counted.
    Only acts when a discharge port is explicitly found and it clearly
    names a non-Saudi place; otherwise (no such line, or it's ambiguous)
    the sheet is processed as normal."""
    for row in rows[:max_scan]:
        row_has_label = False
        row_text_parts = []
        for cell in row:
            if not isinstance(cell, str):
                continue
            row_text_parts.append(cell)
            if "DISCHARG" in cell.upper():
                row_has_label = True
        if not row_has_label:
            continue
        # The port name is sometimes in the SAME cell as the label
        # ("PORT OF DISCHARGE: JEDDAH") and sometimes in a separate cell
        # further along the same row ("PORT OF DISCHARGING:", then later,
        # "UMM QASR,IRAQ") - checking the whole row's text covers both.
        upper = " ".join(row_text_parts).upper()
        if any(m in upper for m in DD_SAUDI_DISCHARGE_MARKERS):
            return False
        if any(m in upper for m in DD_NON_SAUDI_DISCHARGE_MARKERS):
            return True
    return False


def _dd_build_header_candidates(row):
    """Scans one row's string cells for every column that NAMES a tracked
    field, returning ({category: [column_index, ...]}, {col_index: unit})
    - every match is kept (not just the first) so a later resolution step
    can correctly hand a column to the right category even when two
    categories' keywords both landed on it (e.g. "QUANTITY" meaning
    tonnage, see _dd_scan_unit_row)."""
    cats = {"bl": [], "weight_gross": [], "weight_net": [], "weight_generic": [],
            "weight_total": [], "length": [], "spec": [], "qty": [], "desc": []}
    weight_units = {}
    for col_index, cell in enumerate(row):
        if not isinstance(cell, str) or not cell.strip():
            continue
        if _dd_header_matches(cell, DD_BL_HEADER_WORDS) or _DD_BL_HEADER_RE.match(cell.strip()):
            cats["bl"].append(col_index)
        if _DD_TOTAL_WEIGHT_RE.search(cell):
            cats["weight_total"].append(col_index)
        if _dd_header_matches(cell, DD_WEIGHT_GROSS_WORDS):
            cats["weight_gross"].append(col_index)
        if _dd_header_matches(cell, DD_WEIGHT_NET_WORDS):
            cats["weight_net"].append(col_index)
        if _dd_header_matches(cell, DD_WEIGHT_GENERIC_WORDS):
            cats["weight_generic"].append(col_index)
        if _dd_header_matches(cell, DD_LENGTH_WORDS):
            cats["length"].append(col_index)
        if _dd_header_matches(cell, DD_SPEC_WORDS):
            cats["spec"].append(col_index)
        if _dd_header_matches(cell, DD_QTY_WORDS) or _DD_QTY_PAREN_RE.search(cell):
            cats["qty"].append(col_index)
        if _dd_header_matches(cell, DD_DESC_HEADER_WORDS):
            cats["desc"].append(col_index)
        unit = _dd_header_weight_unit(cell)
        if unit:
            weight_units[col_index] = unit
    return cats, weight_units


def _dd_is_header_row(cats):
    has_weight = bool(cats["weight_gross"] or cats["weight_net"] or cats["weight_generic"])
    has_other = bool(cats["length"] or cats["spec"] or cats["qty"] or cats["bl"])
    return has_weight and has_other


def _dd_resolve_header(cats, unit_overrides, weight_units=None):
    """Turns the raw candidate lists (plus any confirmed units-row signal)
    into one final {category: column_index} mapping, each column used at
    most once. A units-row signal for a column overrides a conflicting
    label-only guess for that SAME column (e.g. "QUANTITY" mislabeled as
    qty gets corrected to weight), while other label matches (e.g. a
    genuinely separate "NUMBER OF BUNDLES" column) are unaffected.

    "weight_total" (a column explicitly labelled as a TOTAL weight, e.g.
    "total(KGS)") is resolved before the per-unit weight categories so it
    never gets claimed as the primary weight column when a genuinely
    separate per-unit column also exists. But if it's the ONLY weight-ish
    column on this header, it IS the primary weight column (some templates
    only have a lot-total weight, meant to be divided by qty as before), so
    it's folded back into weight_generic in that case."""
    cats = {k: list(v) for k, v in cats.items()}
    for col, kind in unit_overrides.items():
        if kind == "weight":
            cats["qty"] = [c for c in cats["qty"] if c != col]
        elif kind == "qty":
            cats["weight_generic"] = [c for c in cats["weight_generic"] if c != col]
            cats["weight_net"] = [c for c in cats["weight_net"] if c != col]
            cats["weight_gross"] = [c for c in cats["weight_gross"] if c != col]
            cats["weight_total"] = [c for c in cats["weight_total"] if c != col]

    found = {}
    used = set()
    for cat in ("bl", "weight_total", "weight_gross", "weight_net", "length", "spec", "qty", "weight_generic", "desc"):
        for col in cats[cat]:
            if col not in used:
                found[cat] = col
                used.add(col)
                break

    for col, kind in unit_overrides.items():
        if col in used:
            continue
        if kind == "weight" and not any(k in found for k in ("weight_gross", "weight_net", "weight_generic", "weight_total")):
            found["weight_generic"] = col
            used.add(col)
        elif kind == "qty" and "qty" not in found:
            found["qty"] = col
            used.add(col)

    if "weight_total" in found and not any(k in found for k in ("weight_gross", "weight_net", "weight_generic")):
        # No separate per-unit weight column exists - this total-labelled
        # column IS the one weight figure we have, so treat it as the
        # ordinary (qty-divisible) primary weight column instead.
        found["weight_generic"] = found.pop("weight_total")

    # More than one column plausibly read as "the" per-unit weight (distinct
    # from the separate weight_total cross-check column) means the header
    # was genuinely ambiguous - we still have to pick one by priority, but
    # it's worth flagging rather than asserting it with full confidence.
    weight_like_cols = set(cats["weight_gross"]) | set(cats["weight_net"]) | set(cats["weight_generic"])
    found["_ambiguous_weight"] = len(weight_like_cols) > 1

    found["_weight_units"] = dict(weight_units or {})
    return found


def _extract_classification_rows(rows, sheet_bl_hint=None):
    """Scans a grid of cell values for one or more tables (a sheet can have
    several header+data blocks stacked on top of each other, e.g. once per
    page of a multi-page export) and returns one {bl, weight_mt, length_m,
    description} dict per usable data row.

    A row is treated as a new header whenever its string cells name a
    weight/spec/qty/BL column - this re-reads the column layout each time
    it changes, rather than assuming one fixed table per sheet. Right after
    a header is found, the next row is checked for unit tokens (MT/PCS/...)
    that refine or correct the column mapping before any data is read."""
    if not rows:
        return []
    if _dd_sheet_is_non_saudi(rows):
        return []

    goods_line = _dd_find_goods_line(rows)

    out = []
    cols = None  # current column mapping, or None until a header is found
    i = 0
    n = len(rows)
    while i < n:
        row = rows[i]
        cats, weight_units = _dd_build_header_candidates(row)
        # Peek at the next row for unit tokens (MT/PCS/...) BEFORE deciding
        # whether this row even qualifies as a header - some templates put
        # the weight/qty units on that row and leave the label row itself
        # with no weight-sounding word at all (e.g. "QUANTITY"/"(MT)" split
        # across the two rows).
        unit_overrides = {}
        if i + 1 < n:
            next_cats, _next_units = _dd_build_header_candidates(rows[i + 1])
            if not _dd_is_header_row(next_cats):
                unit_overrides = _dd_scan_unit_row(rows[i + 1])
        has_weight = bool(cats["weight_gross"] or cats["weight_net"] or cats["weight_generic"] or cats["weight_total"]) or "weight" in unit_overrides.values()
        has_other = bool(cats["length"] or cats["spec"] or cats["qty"] or cats["bl"]) or "qty" in unit_overrides.values()
        if has_weight and has_other:
            cols = _dd_resolve_header(cats, unit_overrides, weight_units)
            i += 2 if unit_overrides else 1
            continue

        if cols is None:
            i += 1
            continue
        if _dd_is_totals_row(row):
            cols = None
            i += 1
            continue
        if _dd_is_noise_row(row):
            i += 1
            continue

        def cell_at(key):
            idx = cols.get(key)
            return row[idx] if idx is not None and idx < len(row) and row[idx] not in (None, "") else None

        def weight_unit_for(key):
            idx = cols.get(key)
            return cols.get("_weight_units", {}).get(idx) if idx is not None else None

        bl_cell = cell_at("bl")
        if bl_cell is not None:
            bl = str(bl_cell).strip().upper()
        elif "bl" in cols:
            # This table has its own per-row BL column, but this row's cell
            # is blank - almost always a subtotal/blank row, not real cargo.
            i += 1
            continue
        else:
            bl = sheet_bl_hint or ""
        if not bl:
            i += 1
            continue
        bl_check = re.sub(r"\s+", "", bl)
        if any(w in bl_check for w in ("合计", "总计", "汇总", "小计")) or "TOTAL" in bl_check.upper():
            i += 1
            continue

        if cell_at("weight_gross") is not None:
            weight_key = "weight_gross"
        elif cell_at("weight_net") is not None:
            weight_key = "weight_net"
        else:
            weight_key = "weight_generic"
        weight_cell = cell_at(weight_key)
        weight_mt = _dd_parse_weight_mt(weight_cell, unit_hint=weight_unit_for(weight_key))

        length_m = None
        length_cell = cell_at("length")
        if length_cell is not None:
            length_m = _dd_parse_length_m(length_cell)
        elif cell_at("spec") is not None:
            length_m = _dd_parse_spec_length_mm(cell_at("spec"))

        qty = _dd_parse_number(cell_at("qty"))

        # A separate column explicitly labelled as a TOTAL weight (e.g.
        # "total(KGS)") lets us tell a per-unit weight column apart from a
        # lot-total one: if total =~ weight x qty, "weight" is ALREADY
        # per-unit and must not be divided again.
        skip_division = False
        total_cell = cell_at("weight_total")
        if total_cell is not None and weight_mt is not None and qty and qty > 0:
            total_mt = _dd_parse_weight_mt(total_cell, unit_hint=weight_unit_for("weight_total"))
            if total_mt is not None:
                expected_total = weight_mt * qty
                if expected_total > 0 and abs(total_mt - expected_total) <= max(0.05 * expected_total, 0.01):
                    skip_division = True

        division_guessed = bool(weight_mt is not None and qty and qty > 0 and not skip_division and total_cell is None)
        if weight_mt is not None and qty and qty > 0 and not skip_division:
            weight_mt = weight_mt / qty

        sanity_dropped = False
        # Sanity check the FINAL per-item weight only, once any lot-total
        # has already been divided down by its piece/bundle/coil count. A
        # figure this far out isn't just "uncertain" - it's almost
        # certainly a misread, so it's dropped rather than kept with a
        # caveat, but the BL is still flagged so a human knows something
        # on it couldn't be trusted.
        if weight_mt is not None and weight_mt > DD_SANITY_MAX_WEIGHT_MT:
            weight_mt = None
            sanity_dropped = True

        description = str(cell_at("desc")).strip() if cell_at("desc") is not None else goods_line

        if weight_mt is None and length_m is None and not sanity_dropped:
            i += 1
            continue

        # The actual decision on whether any of this is worth flagging
        # happens in _dd_classify_groups, once it's known whether the item
        # is wheeled/coil-excepted (an excepted item's exact weight doesn't
        # change the verdict, so there's nothing to double check either
        # way) and whether it's the item that's actually driving the
        # result. These are just the raw signals.
        out.append({
            "bl": _dd_bl_root(bl), "weight_mt": weight_mt, "length_m": length_m,
            "description": description, "sanity_dropped": sanity_dropped,
            "ambiguous_weight": bool(cols.get("_ambiguous_weight")), "division_guessed": division_guessed,
        })
        i += 1

    if not out:
        out = _dd_freetext_fallback(rows, sheet_bl_hint, goods_line)
    return out


_DD_FREETEXT_WEIGHT_LABEL_RE = re.compile(
    r"(?:G\.?\s*W\.?|GROSS\s*WEIGHT|N\.?\s*W\.?|NET\s*WEIGHT|TOTAL\s*WEIGHT)\s*[:：]?\s*"
    r"([\d,]+\.?\d*)\s*(M\.?\s?T\.?S?|TONNES?|TONS?|KGS?)\b",
    re.IGNORECASE,
)


def _dd_freetext_fallback(rows, sheet_bl_hint, goods_line):
    """Some packing lists aren't a table at all - just labelled cells like
    "G.W.:161.261 MT", "Gross Weight: 4500 KGS" or "95 PIECES" scattered on
    the sheet (or, for an OCR'd scan/.doc, scattered across noisy
    recognized text). Deliberately narrow and label-anchored rather than
    "find any big number" - a loose guess here is exactly the kind of
    wrong-weight bug this classifier has already been burned by once, so a
    weight is only ever taken from text that explicitly names it."""
    if not sheet_bl_hint:
        return []
    blob_cells = [str(c) for row in rows for c in row if isinstance(c, str)]
    blob = " | ".join(blob_cells)

    weight_match = _DD_FREETEXT_WEIGHT_LABEL_RE.search(blob)
    if not weight_match:
        return []
    raw_value, unit = weight_match.group(1), weight_match.group(2)
    weight_mt = _dd_parse_weight_mt(raw_value, unit_hint=_dd_header_weight_unit(unit))
    if not weight_mt or weight_mt > DD_SANITY_MAX_WEIGHT_MT:
        return []

    pieces_match = re.search(r"(\d+)\s*PIECES", blob, re.IGNORECASE)
    qty = _dd_parse_number(pieces_match.group(1)) if pieces_match else None
    if qty and qty > 0:
        weight_mt = weight_mt / qty

    description = goods_line or blob[:200]
    return [{
        "bl": _dd_bl_root(sheet_bl_hint), "weight_mt": weight_mt, "length_m": None,
        "description": description, "sanity_dropped": False,
        "ambiguous_weight": False, "division_guessed": False, "freetext": True,
    }]


def _dd_classify_groups(items):
    """Groups classification rows by BL and applies the Direct Delivery
    rule, returning {bl_root: (is_direct, reason, needs_review, review_note)}.
    needs_review is never a claim that the verdict is wrong - only that
    some input to it (a borderline value, an ambiguous column, a dropped
    outlier, a free-text guess) wasn't clean enough to trust blind, and a
    quick look at the source file is worth it."""
    by_bl = {}
    for item in items:
        by_bl.setdefault(item["bl"], []).append(item)

    results = {}
    for bl, group in by_bl.items():
        best_trigger = None  # (weight_mt, length_m, description) of the strongest qualifying item
        only_exception_triggers = True
        any_trigger = False
        review_flags = []
        for item in group:
            w, l, d = item["weight_mt"], item["length_m"], item["description"]
            excepted = _dd_is_exception(d)

            # A wheeled/coil item's exact weight or length doesn't change
            # its verdict either way, so there's nothing worth a human
            # double-checking there - only a non-excepted item close
            # enough to the 30MT/12m line (or missing data near it) can
            # actually flip the outcome.
            if not excepted:
                weight_relevant = w is not None and w >= DD_WEIGHT_MT_THRESHOLD * 0.85
                length_relevant = l is not None and l >= DD_LENGTH_M_THRESHOLD * 0.85
                if w is not None and DD_WEIGHT_MT_THRESHOLD * 0.85 <= w <= DD_WEIGHT_MT_THRESHOLD * 1.15:
                    review_flags.append(f"weight ({w:.1f}MT) is close to the 30MT line")
                if l is not None and DD_LENGTH_M_THRESHOLD * 0.85 <= l <= DD_LENGTH_M_THRESHOLD * 1.15:
                    review_flags.append(f"length ({l:.1f}m) is close to the 12m line")
                if item.get("ambiguous_weight") and weight_relevant:
                    review_flags.append("more than one column looked like the weight column")
                if item.get("division_guessed") and weight_relevant:
                    review_flags.append("weight was estimated by dividing a lot total by quantity (unverified)")
                if item.get("sanity_dropped"):
                    review_flags.append("a weight over 500MT was found on this BL and ignored as likely misread")
                if item.get("freetext") and weight_relevant:
                    review_flags.append("read from free text near the 30MT line - please double check the source file")

            triggers = (w is not None and w > DD_WEIGHT_MT_THRESHOLD) or (l is not None and l > DD_LENGTH_M_THRESHOLD)
            if not triggers:
                continue
            any_trigger = True
            if excepted:
                continue
            only_exception_triggers = False
            if best_trigger is None or (w or 0) > (best_trigger[0] or 0) or (l or 0) > (best_trigger[1] or 0):
                best_trigger = (w, l, d)

        needs_review = bool(review_flags)
        # De-dupe while keeping order, and cap so the note stays readable.
        seen = set()
        unique_flags = [f for f in review_flags if not (f in seen or seen.add(f))]
        review_note = "; ".join(unique_flags[:3])

        if best_trigger is not None:
            w, l, d = best_trigger
            detail = f"{w:.1f}MT" if w and w > DD_WEIGHT_MT_THRESHOLD else f"{l:.1f}m"
            reason = f"Direct delivery - item at {detail}" + (f" ({d[:40]})" if d else "")
            results[bl] = (True, reason, needs_review, review_note)
        elif any_trigger and only_exception_triggers:
            results[bl] = (False, "Heavy/oversize item(s) are wheeled or coiled - exception applies", needs_review, review_note)
        else:
            results[bl] = (False, "All items under 30MT and 12m", needs_review, review_note)
    return results


def _dd_sheet_bl_hint(sheet_name, filename_hint):
    """The BL hint to fall back on when a table has no per-row BL column of
    its own. Real, distinct per-sheet names (one BL per tab) win; a generic
    template name ("Sheet1", "PAGE 2"...) means the whole file is one BL,
    so the file name is used instead."""
    if not _dd_is_generic_sheet_name(sheet_name):
        root = _dd_bl_root(sheet_name)
        if root:
            return root
    return filename_hint


def _dd_is_invoice_only_sheet_name(name):
    n = str(name or "").upper()
    return "INVOICE" in n and "PACKING" not in n


def _dd_skip_redundant_invoice_sheets(names):
    """A workbook that has BOTH a commercial-invoice tab and a packing-list
    tab is describing the same shipment twice - the invoice tab often has
    no reliable per-piece/per-bundle count for its lot-total weight (it's a
    pricing document, not a packaging one), which can turn a whole lot's
    weight into a phantom single-item weight. When a real packing-list tab
    is present, the invoice-only tabs are skipped in favor of it. Returns
    the set of sheet names (as given) to skip."""
    has_packing = any("PACKING" in str(n or "").upper() for n in names)
    if not has_packing:
        return set()
    return {n for n in names if _dd_is_invoice_only_sheet_name(n)}


def _dd_extract_pdf_bytes(pdf_bytes, filename_hint):
    """Runs the existing table/text PDF scan over raw PDF bytes - shared by
    real uploaded PDFs and by a .doc converted to PDF first (see below). A
    page with neither a real table nor a text layer is almost always a
    scanned/embedded PICTURE of the packing list (common for .doc files
    that paste in a spreadsheet as an embedded object) rather than an
    empty page, so as a last resort that page is rasterized and OCR'd too."""
    import pdfplumber
    items = []
    blank_pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            tables = page.extract_tables() or []
            if tables:
                for table in tables:
                    items.extend(_extract_classification_rows(table, sheet_bl_hint=filename_hint))
                continue
            text = page.extract_text() or ""
            if text.strip():
                rows = [[line] for line in text.split("\n")]
                items.extend(_extract_classification_rows(rows, sheet_bl_hint=filename_hint))
            else:
                blank_pages.append(page_index)

    if blank_pages:
        try:
            import pytesseract
            from pdf2image import convert_from_bytes
        except ImportError:
            return items
        try:
            images = convert_from_bytes(pdf_bytes, dpi=200)
        except Exception:
            return items
        for page_index in blank_pages:
            if page_index >= len(images):
                continue
            try:
                text = pytesseract.image_to_string(images[page_index])
            except Exception:
                continue
            if text.strip():
                # OCR text is too noisy for the column-index table logic -
                # a garbled line can trivially "look like" a header/weight
                # cell and hand back a plausible but wrong number (exactly
                # the failure mode already fixed once this round). Only the
                # label-anchored freetext fallback is trusted here.
                rows = [[line] for line in text.split("\n")]
                goods_line = _dd_find_goods_line(rows)
                items.extend(_dd_freetext_fallback(rows, filename_hint, goods_line))
    return items


def _dd_office_doc_to_pdf_bytes(raw_bytes, suffix):
    """Converts a legacy .doc (or other office file) to PDF using
    LibreOffice headless, so its table layout is preserved and can be read
    with the same pdfplumber path as a native PDF - a hand-rolled binary
    .doc parser is too easy to get subtly wrong (which, for this
    classifier, means a silently wrong weight - worse than just not
    reading the file). Requires the `soffice` binary on the server; the
    caller turns a missing binary or a conversion failure into a clear
    per-file "couldn't read" message instead of crashing the whole
    request."""
    import subprocess
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, f"input{suffix}")
        with open(src, "wb") as f:
            f.write(raw_bytes)
        subprocess.run(
            ["soffice", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmp, src],
            check=True, timeout=90, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        pdf_path = os.path.join(tmp, "input.pdf")
        with open(pdf_path, "rb") as f:
            return f.read()


def _dd_ocr_image_to_rows(file_storage):
    """OCRs a scanned packing list (PNG/JPG/etc.) into text lines. Requires
    pytesseract/Pillow AND the `tesseract-ocr` system binary on the server
    - missing either raises, which the caller turns into a clear per-file
    message."""
    import pytesseract
    from PIL import Image
    img = Image.open(file_storage)
    text = pytesseract.image_to_string(img)
    return [[line] for line in text.split("\n")]


def _dd_extract_from_upload(file_storage):
    """Reads every sheet/page of one uploaded file and returns its
    classification rows. Raises on a file that can't be read at all, but
    an unsupported extension or an empty result just yields no rows so one
    bad file in a multi-file upload doesn't sink the rest."""
    filename = (file_storage.filename or "").strip()
    lower = filename.lower()
    filename_hint = _dd_filename_bl_hint(filename)
    items = []

    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        # Sniff the real file content rather than trusting the extension -
        # a modern .xlsx saved/forwarded with an old .xls name (or vice
        # versa) is common in practice and would otherwise fail outright.
        sheets = _load_excel_sheets_by_content(file_storage.read(), filename)
        skip_names = _dd_skip_redundant_invoice_sheets([name for name, _ in sheets])
        for sheet_name, rows in sheets:
            if sheet_name in skip_names:
                continue
            items.extend(_extract_classification_rows(rows, sheet_bl_hint=_dd_sheet_bl_hint(sheet_name, filename_hint)))
    elif lower.endswith(".csv"):
        text = file_storage.read().decode("utf-8-sig", errors="ignore")
        rows = list(csv.reader(io.StringIO(text)))
        items.extend(_extract_classification_rows(rows, sheet_bl_hint=filename_hint))
    elif lower.endswith(".pdf"):
        items.extend(_dd_extract_pdf_bytes(file_storage.read(), filename_hint))
    elif lower.endswith((".doc", ".docx")):
        try:
            pdf_bytes = _dd_office_doc_to_pdf_bytes(file_storage.read(), os.path.splitext(lower)[1])
        except FileNotFoundError:
            raise ValueError(f"Couldn't read {filename} - reading .doc/.docx files needs LibreOffice installed on the server.")
        except Exception:
            raise ValueError(f"Couldn't convert {filename} to a readable format.")
        items.extend(_dd_extract_pdf_bytes(pdf_bytes, filename_hint))
    elif lower.endswith((".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")):
        try:
            rows = _dd_ocr_image_to_rows(file_storage)
        except ImportError:
            raise ValueError(f"Couldn't read {filename} - scanning images needs OCR support (pytesseract/tesseract) installed on the server.")
        except Exception:
            raise ValueError(f"Couldn't OCR {filename} - the image may be too low-resolution or unclear to read.")
        # OCR text is too noisy to trust with the column-index table logic
        # (see _dd_extract_pdf_bytes) - only the label-anchored freetext
        # fallback is used for it.
        goods_line = _dd_find_goods_line(rows)
        items.extend(_dd_freetext_fallback(rows, filename_hint, goods_line))
    else:
        raise ValueError(f"Unsupported file type: {filename}")
    return items


@app.route("/api/manifest/classify", methods=["POST"])
@login_required
def classify_manifest():
    """Upload one or more packing lists and every BL across all of them
    gets classified and saved - completely standalone, no dependency on DO
    Tracker's board at all. A BL that's split across several files (e.g.
    the same reference number's pages 1-4 uploaded as separate files) is
    still classified as ONE BL, since every file's items are pooled before
    grouping."""
    files = request.files.getlist("file")
    if not files or all(not f.filename for f in files):
        return jsonify({"error": "No file received"}), 400

    items = []
    failed = []
    for file in files:
        if not file.filename:
            continue
        filename = file.filename
        try:
            file_items = _dd_extract_from_upload(file)
        except ValueError as e:
            failed.append(str(e) if str(e) else filename)
            continue
        except Exception:
            failed.append(filename)
            continue
        items.extend(file_items)

    if not items:
        msg = "Couldn't find a weight or length column in " + ("that file" if len(files) == 1 else "any of those files") + " - classification needs at least one of those."
        if failed:
            msg = f"Couldn't read {', '.join(failed)}. " + msg
        return jsonify({"error": msg}), 400

    classified = _dd_classify_groups(items)

    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    results = []
    for bl, (is_direct, reason, needs_review, review_note) in classified.items():
        db.execute(
            """INSERT INTO direct_delivery (bl_number, is_direct, reason, classified_by, classified_at, needs_review, review_note)
               VALUES (?, ?, ?, ?, ?, ?, ?)
               ON CONFLICT (bl_number) DO UPDATE SET is_direct = EXCLUDED.is_direct, reason = EXCLUDED.reason,
                   classified_by = EXCLUDED.classified_by, classified_at = EXCLUDED.classified_at,
                   needs_review = EXCLUDED.needs_review, review_note = EXCLUDED.review_note""",
            (bl, 1 if is_direct else 0, reason, session.get("username"), now, 1 if needs_review else 0, review_note),
        )
        results.append({"bl": bl, "direct": is_direct, "reason": reason, "needs_review": needs_review, "review_note": review_note})
    db.commit()
    return jsonify({"classified": results, "failed": failed})


@app.route("/api/direct-delivery", methods=["GET"])
@login_required
def list_direct_delivery():
    """Every BL that's been classified as Direct Delivery, for the results
    table - so a refresh doesn't lose what was just uploaded. Not-direct
    BLs aren't shown here at all. Admins see everything; staff only see
    what they classified."""
    db = get_db()
    if session.get("role") == "admin":
        rows = db.execute(
            "SELECT bl_number, is_direct, reason, classified_by, classified_at, needs_review, review_note "
            "FROM direct_delivery WHERE is_direct = 1 ORDER BY classified_at DESC, bl_number"
        ).fetchall()
    else:
        rows = db.execute(
            "SELECT bl_number, is_direct, reason, classified_by, classified_at, needs_review, review_note "
            "FROM direct_delivery WHERE is_direct = 1 AND classified_by = ? ORDER BY classified_at DESC, bl_number",
            (session.get("username"),),
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/direct-delivery/review", methods=["GET"])
@login_required
def list_direct_delivery_review():
    """Every BL flagged needs_review, regardless of is_direct - a
    borderline weight/length, an ambiguous header, a dropped sanity-cap
    outlier, or a free-text-only read. Shown separately from the main
    Direct Delivery list so an uncertain "not direct" doesn't just
    disappear (the main list only ever shows is_direct=1 rows)."""
    db = get_db()
    if session.get("role") == "admin":
        rows = db.execute(
            "SELECT bl_number, is_direct, reason, classified_by, classified_at, review_note "
            "FROM direct_delivery WHERE needs_review = 1 ORDER BY classified_at DESC, bl_number"
        ).fetchall()
    else:
        rows = db.execute(
            "SELECT bl_number, is_direct, reason, classified_by, classified_at, review_note "
            "FROM direct_delivery WHERE needs_review = 1 AND classified_by = ? ORDER BY classified_at DESC, bl_number",
            (session.get("username"),),
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/direct-delivery/<path:bl>", methods=["DELETE"])
@login_required
def delete_direct_delivery(bl):
    """Removes one BL from the results list - for clearing out stale rows
    left over from an earlier test/upload (e.g. a bogus BL name or a weight
    that was misread before a classifier bug was fixed). Admins can remove
    any row; staff can only remove rows they classified themselves."""
    db = get_db()
    if session.get("role") == "admin":
        db.execute("DELETE FROM direct_delivery WHERE bl_number = ?", (bl,))
    else:
        db.execute(
            "DELETE FROM direct_delivery WHERE bl_number = ? AND classified_by = ?",
            (bl, session.get("username")),
        )
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<path:bl_number>/toggle", methods=["POST"])
@login_required
def toggle_status(bl_number):
    if not _owns_record(bl_number.upper()):
        return "Not your record.", 403
    data = request.get_json(force=True)
    field = data.get("field")
    value = 1 if data.get("value") else 0
    user = session.get("username", "Unknown")

    allowed = {
        "invoice_issued": ("invoice_by", "invoice_at"),
        "approval_received": ("approval_by", "approval_at"),
        "do_issued": ("do_by", "do_at"),
    }
    if field not in allowed:
        return jsonify({"error": "invalid field"}), 400

    by_field, at_field = allowed[field]
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M") if value else ""
    by_val = user if value else ""

    db = get_db()
    db.execute(
        f"UPDATE records SET {field} = ?, {by_field} = ?, {at_field} = ? WHERE bl_number = ?",
        (value, by_val, now, bl_number.upper()),
    )
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<path:bl_number>/remarks", methods=["POST"])
@login_required
def update_remarks(bl_number):
    if not _owns_record(bl_number.upper()):
        return "Not your record.", 403
    data = request.get_json(force=True)
    remarks = data.get("remarks", "")
    db = get_db()
    db.execute("UPDATE records SET remarks = ? WHERE bl_number = ?", (remarks, bl_number.upper()))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<path:bl_number>", methods=["DELETE"])
@login_required
def delete_record(bl_number):
    if not _owns_record(bl_number.upper()):
        return "Not your record.", 403
    db = get_db()
    db.execute("DELETE FROM records WHERE bl_number = ?", (bl_number.upper(),))
    db.commit()
    return jsonify({"ok": True})


def _restore_one_record(db, data):
    """Re-inserts a record with all its original fields (rather than a
    bare fresh row). Shared by the single 'Undo after delete' restore and
    the bulk-undo-after-clear-all restore. Returns True if it inserted a
    new row, False if that BL number already exists (a no-op, not an
    error - the whole point of Undo is to be safe to click more than
    once)."""
    bl_number = str(data.get("bl_number", "")).strip().upper()
    if not bl_number:
        return False

    existing = db.execute("SELECT 1 FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
    if existing:
        return False

    db.execute(
        """INSERT INTO records
           (bl_number, consignee, port, vessel, invoice_issued, invoice_by, invoice_at,
            approval_received, approval_by, approval_at, do_issued, do_by, do_at,
            remarks, created_at, created_by)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            bl_number,
            data.get("consignee", ""),
            data.get("port", ""),
            data.get("vessel", ""),
            1 if data.get("invoice_issued") else 0,
            data.get("invoice_by", ""),
            data.get("invoice_at", ""),
            1 if data.get("approval_received") else 0,
            data.get("approval_by", ""),
            data.get("approval_at", ""),
            1 if data.get("do_issued") else 0,
            data.get("do_by", ""),
            data.get("do_at", ""),
            data.get("remarks", ""),
            data.get("created_at", ""),
            data.get("created_by") or session.get("username"),
        ),
    )
    return True


@app.route("/api/records/restore", methods=["POST"])
@login_required
def restore_record():
    """Used by the 'Undo' notice after a delete - re-inserts a record with
    all its original fields, rather than a bare fresh row."""
    data = request.get_json(force=True)
    db = get_db()
    if not str(data.get("bl_number", "")).strip():
        return jsonify({"error": "missing bl_number"}), 400
    inserted = _restore_one_record(db, data)
    db.commit()
    return jsonify({"ok": True, "note": None if inserted else "already exists"})


@app.route("/api/records/bulk-delete", methods=["POST"])
@login_required
def bulk_delete_records():
    """Removes many BLs in one call - the "Remove all" buttons on a
    vessel/port group, or the whole board, use this instead of firing one
    DELETE per row (the difference matters once a manifest has 100+ BLs).
    Staff can only delete their own records even if other BLs were passed
    in (ownership is still checked per-row); returns the full data of
    whatever it actually deleted so the client can offer an Undo that
    restores exactly those rows."""
    data = request.get_json(force=True)
    bl_numbers = [str(b).strip().upper() for b in data.get("bl_numbers", []) if str(b).strip()]
    if not bl_numbers:
        return jsonify({"error": "No BL numbers given"}), 400

    db = get_db()
    deleted = []
    for bl_number in bl_numbers:
        if not _owns_record(bl_number):
            continue
        row = db.execute("SELECT * FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
        if row is None:
            continue
        deleted.append(dict(row))
        db.execute("DELETE FROM records WHERE bl_number = ?", (bl_number,))
    db.commit()
    return jsonify({"deleted": deleted})


@app.route("/api/records/bulk-restore", methods=["POST"])
@login_required
def bulk_restore_records():
    """Undo counterpart to bulk-delete - re-inserts every record passed in
    (skipping any that already exist, same as the single restore)."""
    data = request.get_json(force=True)
    items = data.get("records", [])
    db = get_db()
    restored = 0
    for item in items:
        if _restore_one_record(db, item):
            restored += 1
    db.commit()
    return jsonify({"restored": restored})


@app.route("/api/groups/rename", methods=["POST"])
@login_required
def rename_group():
    """Renaming a Port or Vessel group header updates every BL record
    filed under it - lets the whole board be reorganized without editing
    each BL one by one."""
    data = request.get_json(force=True)
    group_type = data.get("type")
    old_port = data.get("old_port", "")
    new_value = data.get("new_value", "").strip()
    db = get_db()
    is_admin = session.get("role") == "admin"
    if group_type == "port":
        if is_admin:
            db.execute("UPDATE records SET port = ? WHERE port = ?", (new_value, old_port))
        else:
            db.execute(
                "UPDATE records SET port = ? WHERE port = ? AND created_by = ?",
                (new_value, old_port, session.get("username")),
            )
    elif group_type == "vessel":
        old_vessel = data.get("old_vessel", "")
        if is_admin:
            db.execute(
                "UPDATE records SET vessel = ? WHERE port = ? AND vessel = ?",
                (new_value, old_port, old_vessel),
            )
        else:
            db.execute(
                "UPDATE records SET vessel = ? WHERE port = ? AND vessel = ? AND created_by = ?",
                (new_value, old_port, old_vessel, session.get("username")),
            )
    else:
        return jsonify({"error": "invalid type"}), 400
    db.commit()
    return jsonify({"ok": True})


# ---------- Vessel Tracker (standalone list, live positions via MarineTraffic's free embed) ----------
# This list is fully independent of DO Tracker's records - vessels are
# added/removed here directly. Whoever adds a vessel becomes its "operator"
# permanently (shown everywhere), even after DO Tracker becomes per-staff.

@app.route("/api/vessels", methods=["GET"])
@login_required
def list_vessels():
    db = get_db()
    rows = db.execute("SELECT name, mmsi, operator FROM vessels ORDER BY name").fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/vessels", methods=["POST"])
@login_required
def add_vessel():
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"error": "Missing vessel name."}), 400
    db = get_db()
    existing = db.execute("SELECT 1 FROM vessels WHERE name = ?", (name,)).fetchone()
    if existing:
        return jsonify({"error": "That vessel is already on the list."}), 400
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    db.execute(
        "INSERT INTO vessels (name, mmsi, operator, updated_by, updated_at) VALUES (?, '', ?, ?, ?)",
        (name, session.get("username"), session.get("username"), now),
    )
    db.commit()
    return jsonify({"ok": True, "name": name, "mmsi": "", "operator": session.get("username")})


@app.route("/api/vessels/<path:name>", methods=["DELETE"])
@login_required
def delete_vessel(name):
    db = get_db()
    db.execute("DELETE FROM vessels WHERE name = ?", (name,))
    db.commit()
    return jsonify({"ok": True})


# Label variants (normalized: lowercase, letters/digits only) recognized in
# a "Ship's Particulars" sheet, wherever the label and its value happen to
# sit - these sheets aren't a fixed template, so we scan every cell.
VESSEL_NAME_LABELS = {"shipsname", "vesselname", "shipname", "nameofvessel", "nameofship"}
VESSEL_MMSI_LABELS = {"mmsi", "mmsino"}


def _normalize_label(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


def _extract_vessel_particulars(rows):
    """Given a grid of cell values (list of rows), find the vessel's name
    and MMSI by scanning for a recognizable label cell and taking the next
    non-empty cell after it in the same row as the value."""
    name = None
    mmsi = None
    for row in rows:
        for i, cell in enumerate(row):
            key = _normalize_label(cell)
            if not key:
                continue
            if name is None and key in VESSEL_NAME_LABELS:
                for v in row[i + 1:]:
                    if v is not None and str(v).strip() != "":
                        name = str(v).strip()
                        break
            if mmsi is None and key in VESSEL_MMSI_LABELS:
                for v in row[i + 1:]:
                    if v is not None and str(v).strip() != "":
                        try:
                            mmsi = str(int(float(v)))
                        except (TypeError, ValueError):
                            mmsi = re.sub(r"[^0-9]", "", str(v))
                        break
    return name, mmsi


@app.route("/api/vessels/upload", methods=["POST"])
@login_required
def upload_vessel_particulars():
    """Drag a Ship's Particulars file (.xls or .xlsx) straight in and this
    pulls out the vessel name + MMSI and adds/updates it on the tracker -
    no manual typing needed."""
    if "file" not in request.files:
        return jsonify({"error": "No file received"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    fname = file.filename.lower()

    rows = []
    try:
        if fname.endswith(".xls"):
            wb = xlrd.open_workbook(file_contents=file.read())
            for sheet in wb.sheets():
                for r in range(sheet.nrows):
                    rows.append([sheet.cell_value(r, c) for c in range(sheet.ncols)])
        elif fname.endswith((".xlsx", ".xlsm")):
            wb = openpyxl.load_workbook(file, data_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    rows.append(list(row))
        else:
            return jsonify({"error": "Please upload an .xls or .xlsx file."}), 400
    except Exception:
        return jsonify({"error": "Couldn't read that file - make sure it's a valid Excel file."}), 400

    name, mmsi = _extract_vessel_particulars(rows)
    if not name:
        return jsonify({"error": "Couldn't find a vessel name in that file. Try adding it manually."}), 400
    if mmsi and (not mmsi.isdigit() or len(mmsi) != 9):
        mmsi = None  # found something but it doesn't look like a real MMSI - don't fail the whole import over it

    db = get_db()
    existing = db.execute("SELECT operator FROM vessels WHERE name = ?", (name,)).fetchone()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    if existing:
        if mmsi:
            db.execute(
                "UPDATE vessels SET mmsi = ?, updated_by = ?, updated_at = ? WHERE name = ?",
                (mmsi, session.get("username"), now, name),
            )
    else:
        db.execute(
            "INSERT INTO vessels (name, mmsi, operator, updated_by, updated_at) VALUES (?, ?, ?, ?, ?)",
            (name, mmsi or "", session.get("username"), session.get("username"), now),
        )
    db.commit()
    return jsonify({"ok": True, "name": name, "mmsi": mmsi or "", "had_mmsi": bool(mmsi)})


@app.route("/api/vessels/mmsi", methods=["GET"])
@login_required
def get_vessel_mmsi():
    db = get_db()
    rows = db.execute("SELECT name, mmsi FROM vessels").fetchall()
    return jsonify({r["name"]: r["mmsi"] for r in rows if r["mmsi"]})


@app.route("/api/vessels/mmsi", methods=["POST"])
@login_required
def set_vessel_mmsi():
    data = request.get_json(force=True)
    name = (data.get("vessel") or "").strip()
    mmsi = (data.get("mmsi") or "").strip()
    if not name:
        return jsonify({"error": "Missing vessel name."}), 400
    if mmsi and (not mmsi.isdigit() or len(mmsi) != 9):
        return jsonify({"error": "MMSI must be exactly 9 digits."}), 400
    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    if mmsi:
        # operator is only set on the initial INSERT (whoever adds the vessel
        # first) and deliberately left out of the ON CONFLICT update, so a
        # later MMSI edit never reassigns ownership.
        db.execute(
            """INSERT INTO vessels (name, mmsi, operator, updated_by, updated_at) VALUES (?, ?, ?, ?, ?)
               ON CONFLICT (name) DO UPDATE SET mmsi = EXCLUDED.mmsi, updated_by = EXCLUDED.updated_by, updated_at = EXCLUDED.updated_at""",
            (name, mmsi, session.get("username"), session.get("username"), now),
        )
    else:
        db.execute(
            "UPDATE vessels SET mmsi = '', updated_by = ?, updated_at = ? WHERE name = ?",
            (session.get("username"), now, name),
        )
    db.commit()
    return jsonify({"ok": True, "vessel": name, "mmsi": mmsi})


# ---------- Templates ----------

AUTH_STYLE = """
<style>
  :root {
    --bg: #f2f4f7;
    --card: #ffffff;
    --text: #1c2b3a;
    --muted: #7a8794;
    --border: #e6e9ed;
    --navy: #123a56;
    --navy-deep: #0b2740;
    --navy-light: #1f5c85;
    --gold: #c9a227;
    --gold-light: #e0bd53;
    --danger: #d1483f;
    --danger-bg: #fbeceb;
    --shadow-md: 0 20px 60px rgba(11,39,64,0.16);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23;
    --card: #1a232f;
    --text: #e9eef3;
    --muted: #93a1b1;
    --border: #29323f;
    --navy: #3f86ba;
    --navy-deep: #274a67;
    --navy-light: #5aa2d1;
    --gold: #e3bb4c;
    --gold-light: #f0cf72;
    --danger: #e2685f;
    --danger-bg: #3a2220;
    --shadow-md: 0 20px 60px rgba(0,0,0,0.55);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  html, body { height: 100%; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text);
    margin: 0; display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 20px;
    position: relative; overflow: hidden;
    transition: background-color .3s ease, color .3s ease;
  }

  /* Ambient drifting gradient blobs */
  .blob {
    position: fixed; border-radius: 50%; filter: blur(60px); z-index: 0; pointer-events: none;
    opacity: .55; transition: opacity .3s ease;
  }
  .blob1 { width: 420px; height: 420px; top: -140px; left: -120px; background: radial-gradient(circle, var(--navy-light), transparent 70%); animation: drift1 16s ease-in-out infinite; }
  .blob2 { width: 380px; height: 380px; bottom: -160px; right: -100px; background: radial-gradient(circle, var(--gold), transparent 70%); opacity: .35; animation: drift2 20s ease-in-out infinite; }
  :root[data-theme="dark"] .blob { opacity: .28; }
  @keyframes drift1 { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(40px,30px) scale(1.08); } }
  @keyframes drift2 { 0%,100% { transform: translate(0,0) scale(1); } 50% { transform: translate(-30px,-25px) scale(1.1); } }

  /* Theme toggle, top right */
  .theme-switch { position: fixed; top: 20px; right: 20px; z-index: 5; display: inline-flex; width: 54px; height: 29px; cursor: pointer; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track {
    position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center;
    justify-content: space-between; padding: 0 7px;
    background: linear-gradient(135deg,#8fcaf0,#f4d58d); transition: background .3s ease;
  }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob {
    position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%;
    background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1);
  }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .box {
    position: relative; z-index: 1;
    background: color-mix(in srgb, var(--card) 92%, transparent);
    backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
    border: 1px solid var(--border); border-radius: 20px; padding: 34px 30px;
    width: 100%; max-width: 380px; box-shadow: var(--shadow-md);
    animation: card-in .55s cubic-bezier(.16,1,.3,1) both;
  }
  @keyframes card-in {
    from { opacity: 0; transform: translateY(22px) scale(.97); }
    to { opacity: 1; transform: translateY(0) scale(1); }
  }

  .brand-mark {
    display: flex; flex-direction: column; align-items: center; text-align: center; margin-bottom: 22px;
  }
  .brand-mark img {
    height: 56px; width: auto; margin-bottom: 12px;
    animation: mark-in 2.1s .1s cubic-bezier(.22,.7,.2,1) both;
  }
  /* Logo pops in, then swings and settles like a compass needle finding its heading */
  @keyframes mark-in {
    0%   { opacity: 0; transform: scale(.6) rotate(-16deg); }
    30%  { opacity: 1; transform: scale(1) rotate(15deg); }
    48%  { transform: scale(1) rotate(-10deg); }
    64%  { transform: scale(1) rotate(6deg); }
    80%  { transform: scale(1) rotate(-3deg); }
    92%  { transform: scale(1) rotate(1deg); }
    100% { opacity: 1; transform: scale(1) rotate(0deg); }
  }
  .brand-mark .co { font-size: 11px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: var(--gold); }
  .brand-mark .tag { font-size: 11.5px; color: var(--muted); margin-top: 2px; }

  h1 { font-size: 19px; margin: 0 0 4px; text-align: center; letter-spacing: -0.01em; }
  .sub { color: var(--muted); font-size: 13px; margin-bottom: 22px; text-align: center; }

  label { font-size: 12px; font-weight: 600; display: block; margin-bottom: 6px; margin-top: 16px; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; }
  input, select {
    width: 100%; padding: 11px 13px; border: 1px solid var(--border); border-radius: 10px;
    font-size: 14.5px; font-family: inherit; background: var(--bg); color: var(--text);
    transition: border-color .2s ease, background .2s ease, box-shadow .25s ease, transform .15s ease;
  }
  input:focus, select:focus {
    outline: none; border-color: var(--gold); background: var(--card);
    transform: translateY(-1px);
    box-shadow: 0 0 0 4px color-mix(in srgb, var(--gold) 22%, transparent),
                0 0 16px color-mix(in srgb, var(--gold) 35%, transparent);
  }
  button {
    position: relative;
    width: 100%; background: var(--navy); color: #fff; border: none; border-radius: 999px;
    padding: 13px; font-size: 14px; font-weight: 700; margin-top: 24px; cursor: pointer;
    transition: background .15s ease, transform .08s ease;
  }
  button:hover { background: var(--navy-light); }
  button:active { transform: scale(.98); }

  /* Loading state - shown while a form submit is in flight */
  button.loading { color: transparent; pointer-events: none; }
  button.loading::after {
    content: ""; position: absolute; left: 50%; top: 50%; width: 18px; height: 18px;
    margin: -9px 0 0 -9px; border: 2.5px solid rgba(255,255,255,.35); border-top-color: #fff;
    border-radius: 50%; animation: btn-spin .65s linear infinite;
  }
  @keyframes btn-spin { to { transform: rotate(360deg); } }

  /* Staggered entrance for the form fields, one after another */
  form > label, form > input, form > .pw-wrap, form > button, form > .error {
    animation: field-in .5s ease both;
  }
  form > label:nth-of-type(1) { animation-delay: .18s; }
  form > input:nth-of-type(1), form > .pw-wrap:nth-of-type(1) { animation-delay: .24s; }
  form > label:nth-of-type(2) { animation-delay: .30s; }
  form > input:nth-of-type(2), form > .pw-wrap:nth-of-type(2) { animation-delay: .36s; }
  form > button { animation-delay: .44s; }
  @keyframes field-in {
    from { opacity: 0; transform: translateY(10px); }
    to { opacity: 1; transform: translateY(0); }
  }
  .error {
    background: var(--danger-bg); color: var(--danger); padding: 10px 12px; border-radius: 10px;
    font-size: 13px; margin-top: 16px; text-align: center; font-weight: 600;
    animation: shake .35s ease;
  }
  @keyframes shake {
    10%,90% { transform: translateX(-1px); } 20%,80% { transform: translateX(2px); }
    30%,50%,70% { transform: translateX(-4px); } 40%,60% { transform: translateX(4px); }
  }

  /* Password show/hide toggle */
  .pw-wrap { position: relative; }
  .pw-wrap input { padding-right: 42px; }
  .pw-toggle {
    position: absolute; right: 5px; top: 50%; transform: translateY(-50%);
    width: 32px; height: 32px; margin: 0; padding: 0; background: none; border: none;
    display: flex; align-items: center; justify-content: center; cursor: pointer;
    color: var(--muted); border-radius: 8px; transition: color .15s ease, background .15s ease;
  }
  .pw-toggle:hover { color: var(--text); background: color-mix(in srgb, var(--border) 70%, transparent); }
  .pw-toggle:active { transform: translateY(-50%) scale(.92); }
  .pw-toggle svg { width: 18px; height: 18px; pointer-events: none; }
</style>
<script>
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light'; // default to light for first-time visitors; once they toggle, localStorage remembers it
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}

function togglePw(btn) {
  const wrap = btn.closest('.pw-wrap');
  const input = wrap.querySelector('input');
  const eye = btn.querySelector('.icon-eye');
  const eyeOff = btn.querySelector('.icon-eye-off');
  const showing = input.type === 'password';
  input.type = showing ? 'text' : 'password';
  eye.style.display = showing ? 'none' : '';
  eyeOff.style.display = showing ? '' : 'none';
  btn.setAttribute('aria-label', showing ? 'Hide password' : 'Show password');
}

/* Show a spinner on the submit button while the request is in flight,
   so it's clear something is happening after clicking Sign In / Create. */
window.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('form').forEach((form) => {
    form.addEventListener('submit', () => {
      if (!form.checkValidity()) return;
      const btn = form.querySelector('button[type="submit"]');
      if (btn) { btn.classList.add('loading'); btn.disabled = true; }
    });
  });
});
</script>
"""

THEME_TOGGLE_SNIPPET = """
  <label class="theme-switch" title="Toggle dark mode">
    <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
    <span class="theme-track">
      <span class="theme-icon sun">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
      </span>
      <span class="theme-icon moon">
        <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
      </span>
      <span class="theme-knob"></span>
    </span>
  </label>
"""

PW_TOGGLE_BTN = """<button type="button" class="pw-toggle" onclick="togglePw(this)" tabindex="-1" aria-label="Show password">
      <svg class="icon-eye" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7-11-7-11-7Z"/><circle cx="12" cy="12" r="3"/></svg>
      <svg class="icon-eye-off" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" style="display:none"><path d="M17.94 17.94A10.94 10.94 0 0 1 12 19c-7 0-11-7-11-7a21.6 21.6 0 0 1 5.06-6.17M9.9 4.24A10.94 10.94 0 0 1 12 4c7 0 11 7 11 7a21.6 21.6 0 0 1-2.16 3.19M14.12 14.12a3 3 0 1 1-4.24-4.24"/><path d="M1 1l22 22"/></svg>
    </button>"""


SETUP_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Set up</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">""" + AUTH_STYLE + """</head><body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
""" + THEME_TOGGLE_SNIPPET + """
<div class="box">
  <div class="brand-mark">
    <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
    <span class="co">Compass</span>
    <span class="tag">Sea Power Marine Services Co. Ltd</span>
  </div>
  <h1>Welcome to Compass</h1>
  <div class="sub">First time here - create the Admin account to get started.</div>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Choose a username</label>
    <input type="text" name="username" required autofocus>
    <label>Choose a password</label>
    <div class="pw-wrap">
      <input type="password" name="password" required>
      """ + PW_TOGGLE_BTN + """
    </div>
    <button type="submit">Create Admin Account</button>
  </form>
</div>
</body></html>
"""

LOGIN_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Sign In</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">""" + AUTH_STYLE + """</head><body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
""" + THEME_TOGGLE_SNIPPET + """
<div class="box">
  <div class="brand-mark">
    <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
    <span class="co">Compass</span>
    <span class="tag">Sea Power Marine Services Co. Ltd</span>
  </div>
  <h1>Sign in to Compass</h1>
  <div class="sub">Your shared workspace.</div>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Username</label>
    <input type="text" name="username" required autofocus>
    <label>Password</label>
    <div class="pw-wrap">
      <input type="password" name="password" required>
      """ + PW_TOGGLE_BTN + """
    </div>
    <button type="submit">Sign In</button>
  </form>
</div>
</body></html>
"""

USERS_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Manage Users</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --danger: #d1483f; --danger-bg: #fbeceb;
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --danger: #e2685f; --danger-bg: #3a2220;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 32px;
    transition: background-color .25s ease, color .25s ease;
  }
  .topbar {
    position: sticky; top: 0; z-index: 50; display: flex; justify-content: space-between; align-items: center;
    gap: 12px; flex-wrap: wrap; padding: 14px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px); -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand img { height: 32px; width: auto; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .app-name { font-size: 14.5px; font-weight: 700; color: var(--text); }
  .brand-text .app-tag { font-size: 11px; color: var(--muted); }
  .topbar-right { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
  .topbar-right a { color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px; padding: 6px 12px; border-radius: 20px; transition: background .15s ease; }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }

  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track { position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center; justify-content: space-between; padding: 0 7px; background: linear-gradient(135deg,#8fcaf0,#f4d58d); transition: background .3s ease; }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob { position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1); }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .card { background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 18px; margin-bottom: 16px; box-shadow: var(--shadow-sm); }
  .card-label { font-size: 12px; font-weight: 600; color: var(--navy); text-transform: uppercase; letter-spacing: .04em; margin-bottom: 10px; }
  :root[data-theme="dark"] .card-label { color: var(--navy-light); }
  table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  th, td { text-align: left; padding: 12px 10px; border-bottom: 1px solid var(--border); }
  th { color: var(--muted); font-weight: 600; font-size: 10.5px; text-transform: uppercase; letter-spacing: .05em; background: color-mix(in srgb, var(--border) 40%, transparent); }
  tbody tr:last-child td { border-bottom: none; }
  .role-pill { display: inline-block; font-size: 10.5px; font-weight: 700; padding: 2px 9px; border-radius: 999px; text-transform: uppercase; letter-spacing: .03em; }
  .role-pill.admin { background: color-mix(in srgb, var(--gold) 18%, transparent); color: var(--gold); }
  .role-pill.staff { background: color-mix(in srgb, var(--navy-light) 14%, transparent); color: var(--navy-light); }
  input[type=text], input[type=password] {
    border: 1px solid var(--border); border-radius: 10px; padding: 10px 12px; font-size: 14px; font-family: inherit;
    background: var(--bg); color: var(--text); transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
  }
  input:focus { outline: none; border-color: var(--navy-light); background: var(--card); box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 20%, transparent); }
  select { border: 1px solid var(--border); border-radius: 10px; padding: 10px 12px; font-size: 14px; font-family: inherit; background: var(--bg); color: var(--text); }
  button { background: var(--navy); color: #fff; border: none; border-radius: 999px; padding: 10px 18px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background .15s ease, transform .08s ease; }
  button:hover { background: var(--navy-light); }
  button:active { transform: scale(.97); }
  .del { background: none; color: var(--danger); font-size: 12px; font-weight: 600; padding: 5px 10px; border-radius: 999px; }
  .del:hover { background: var(--danger-bg); }
  .row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; pointer-events: none; max-width: min(320px, calc(100vw - 40px)); }
  #toastHost .toast { pointer-events: auto; }
  .toast { background: var(--navy-deep); color: #fff; padding: 11px 16px; border-radius: 12px; font-size: 13px; display: flex; align-items: center; gap: 14px; box-shadow: 0 10px 30px rgba(0,0,0,0.25); animation: toast-in .18s ease-out; max-width: 320px; }
  .toast.error { background: var(--danger); }
  .toast a { color: var(--gold-light); font-weight: 700; text-decoration: none; cursor: pointer; white-space: nowrap; }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(8px); } }
</style>
<script>
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light'; // default to light for first-time visitors; once they toggle, localStorage remembers it
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}
</script>
</head><body>
  <div class="topbar">
    <a href="/" class="brand" style="text-decoration:none;">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Manage Users</span>
      </div>
    </a>
    <div class="topbar-right">
      <label class="theme-switch" title="Toggle dark mode">
        <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
        <span class="theme-track">
          <span class="theme-icon sun">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
          </span>
          <span class="theme-icon moon">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
          </span>
          <span class="theme-knob"></span>
        </span>
      </label>
      <a href="/">&larr; Back to board</a>
      <span style="padding:6px 4px;">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="card">
    <div class="card-label">Add a user</div>
    <div class="row">
      <input type="text" id="newUsername" placeholder="Username">
      <input type="password" id="newPassword" placeholder="Password">
      <select id="newRole"><option value="staff">Staff</option><option value="admin">Admin</option></select>
      <button onclick="addUser()">Add User</button>
    </div>
  </div>
  <div class="card">
    <table>
      <thead><tr><th>Username</th><th>Role</th><th>Created</th><th></th></tr></thead>
      <tbody>
        {% for u in users %}
        <tr>
          <td>{{ u['username'] }}</td>
          <td><span class="role-pill {{ u['role'] }}">{{ u['role'] }}</span></td>
          <td class="local-time" data-utc="{{ u['created_at'] }}">{{ u['created_at'] }}</td>
          <td><button class="del" onclick='delUser({{ u["id"] }}, {{ u["username"]|tojson }})'>Remove</button></td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
  <div id="toastHost"></div>
<script>
// The server stores "Created" timestamps as naive UTC - convert each one to
// the viewer's own timezone before displaying it.
document.querySelectorAll('.local-time').forEach(el => {
  const raw = el.dataset.utc;
  if (!raw) return;
  const iso = raw.includes('T') ? raw : raw.replace(' ', 'T') + ':00Z';
  const d = new Date(iso);
  if (!isNaN(d.getTime())) {
    el.textContent = d.toLocaleString(undefined, {
      year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
    });
  }
});

function showToast(message, opts) {
  opts = opts || {};
  const host = document.getElementById('toastHost');
  const el = document.createElement('div');
  el.className = 'toast' + (opts.error ? ' error' : '');
  const text = document.createElement('span');
  text.textContent = message;
  el.appendChild(text);
  host.appendChild(el);
  const duration = opts.duration || 4000;
  const timer = setTimeout(dismiss, duration);
  function dismiss() {
    clearTimeout(timer);
    el.classList.add('fading');
    setTimeout(() => el.remove(), 220);
  }
}
async function addUser() {
  const username = document.getElementById('newUsername').value.trim();
  const password = document.getElementById('newPassword').value;
  const role = document.getElementById('newRole').value;
  if (!username || !password) { showToast('Fill in username and password', {error:true}); return; }
  const res = await fetch('/api/users', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({username, password, role})});
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not add that user.', {error:true}); return; }
  showToast('User ' + username + ' added.');
  setTimeout(() => location.reload(), 500);
}
async function delUser(id, username) {
  const res = await fetch('/api/users/' + id, {method:'DELETE'});
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.error) {
    showToast(data.error || 'Could not remove that user.', {error:true, duration: 5000});
    return;
  }
  showToast('Removed user ' + (username || '') + '.');
  setTimeout(() => location.reload(), 500);
}
</script>
</body></html>
"""

HUB_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05);
    --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25);
    --shadow-md: 0 10px 30px rgba(0,0,0,0.35);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 48px;
    transition: background-color .25s ease, color .25s ease;
  }
  .topbar {
    position: sticky; top: 0; z-index: 50; display: flex; justify-content: space-between; align-items: center;
    gap: 12px; flex-wrap: wrap; padding: 14px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px); -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand img { height: 32px; width: auto; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .app-name { font-size: 14.5px; font-weight: 700; color: var(--text); }
  .brand-text .app-tag { font-size: 11px; color: var(--muted); }
  .topbar-right { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
  .topbar-right a { color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px; padding: 6px 12px; border-radius: 20px; transition: background .15s ease; }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }

  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track { position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center; justify-content: space-between; padding: 0 7px; background: linear-gradient(135deg,#8fcaf0,#f4d58d); transition: background .3s ease; }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob { position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1); }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .hero { padding: 28px 4px 8px; }
  .hero .eyebrow { font-size: 12px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: var(--gold); margin-bottom: 6px; }
  :root[data-theme="dark"] .hero .eyebrow { color: var(--gold-light); }
  .hero h1 { font-size: 26px; margin: 0 0 6px; letter-spacing: -0.01em; }
  .hero p { color: var(--muted); margin: 0; font-size: 14.5px; }

  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 16px; margin-top: 24px; }
  .tile {
    display: flex; flex-direction: column; gap: 12px; background: var(--card); border: 1px solid var(--border);
    border-radius: 18px; padding: 20px; text-decoration: none; color: var(--text); box-shadow: var(--shadow-sm);
    transition: transform .18s cubic-bezier(.16,1,.3,1), box-shadow .18s ease, border-color .18s ease;
    animation: tile-in .5s cubic-bezier(.16,1,.3,1) both;
  }
  .tile:hover { transform: translateY(-3px); box-shadow: var(--shadow-md); border-color: var(--navy-light); }
  .tile:active { transform: translateY(-1px) scale(.99); }
  @keyframes tile-in { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
  .tile .tile-icon {
    width: 44px; height: 44px; border-radius: 13px; display: flex; align-items: center; justify-content: center;
    background: color-mix(in srgb, var(--navy) 12%, transparent); color: var(--navy);
  }
  :root[data-theme="dark"] .tile .tile-icon { color: var(--navy-light); background: color-mix(in srgb, var(--navy-light) 18%, transparent); }
  .tile .tile-icon svg { width: 22px; height: 22px; }
  .tile h3 { margin: 0; font-size: 16px; font-weight: 700; }
  .tile p { margin: 0; color: var(--muted); font-size: 13px; line-height: 1.45; }
  .tile .tile-go { margin-top: auto; font-size: 12.5px; font-weight: 700; color: var(--navy); display: flex; align-items: center; gap: 4px; }
  :root[data-theme="dark"] .tile .tile-go { color: var(--navy-light); }
  .tile .tile-go svg { width: 13px; height: 13px; transition: transform .18s ease; }
  .tile:hover .tile-go svg { transform: translateX(3px); }

  .tile.soon { opacity: .55; cursor: default; }
  .tile.soon:hover { transform: none; box-shadow: var(--shadow-sm); border-color: var(--border); }
  .tile.soon .tile-go { color: var(--muted); }
  .soon-pill { font-size: 9.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; color: var(--muted); background: color-mix(in srgb, var(--border) 60%, transparent); padding: 2px 8px; border-radius: 999px; align-self: flex-start; }
</style>
</head><body>
  <div class="topbar">
    <a href="/" class="brand" style="text-decoration:none;">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Sea Power Marine Services Co. Ltd</span>
      </div>
    </a>
    <div class="topbar-right">
      <label class="theme-switch" title="Toggle dark mode">
        <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
        <span class="theme-track">
          <span class="theme-icon sun">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
          </span>
          <span class="theme-icon moon">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
          </span>
          <span class="theme-knob"></span>
        </span>
      </label>
      {% if role == 'admin' %}<a href="/users">Manage Users</a>{% endif %}
      <span style="padding:6px 4px;">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="hero">
    <div class="eyebrow">Compass</div>
    <h1>What are you working on?</h1>
    <p>Pick a workspace below. More will show up here as they're added.</p>
  </div>

  <div class="grid">
    <a class="tile" href="/do-tracker">
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 2h6a1 1 0 0 1 1 1v1h1a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h1V3a1 1 0 0 1 1-1Z"/><path d="M9 12l2 2 4-4"/></svg>
      </div>
      <h3>DO Tracker</h3>
      <p>Track invoice, approval and DO status per Bill of Lading, organized by port and vessel.</p>
      <span class="tile-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></span>
    </a>

    <a class="tile" href="/vessel-tracker">
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 21c2 1 4 1 6 0s4-1 6 0 4 1 6 0M4 18l1-9 2-3h10l2 3 1 9M9 6V3h6v3"/></svg>
      </div>
      <h3>Vessel Tracker</h3>
      <p>Live positions for your vessels, straight from MarineTraffic - its own list, tagged by operator.</p>
      <span class="tile-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></span>
    </a>

    <a class="tile" href="/kpi">
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M18 17V9M13 17V5M8 17v-4"/></svg>
      </div>
      <h3>Port Agent KPI</h3>
      <p>Turnaround times, pending backlog and workload, built from the DO Tracker board.</p>
      <span class="tile-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></span>
    </a>

    <a class="tile" href="/direct-delivery">
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7l9-4 9 4-9 4-9-4z"/><path d="M3 7v10l9 4 9-4V7"/><path d="M12 11v10"/></svg>
      </div>
      <h3>Direct Delivery Classifier</h3>
      <p>Upload a cargo packing list and see which BLs need direct delivery, by weight and size.</p>
      <span class="tile-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></span>
    </a>
  </div>

<script>
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light'; // default to light for first-time visitors; once they toggle, localStorage remembers it
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}
</script>
</body></html>
"""

VESSEL_TRACKER_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Vessel Tracker</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --danger: #d1483f; --danger-bg: #fbeceb;
    --ok: #1c8a5a; --ok-bg: #e7f5ee;
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05);
    --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --danger: #e2685f; --danger-bg: #3a2220;
    --ok: #3ecb8e; --ok-bg: #163329;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25);
    --shadow-md: 0 10px 30px rgba(0,0,0,0.35);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 32px;
    transition: background-color .25s ease, color .25s ease;
  }
  .topbar {
    position: sticky; top: 0; z-index: 50; display: flex; justify-content: space-between; align-items: center;
    gap: 12px; flex-wrap: wrap; padding: 14px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px); -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand img { height: 32px; width: auto; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .app-name { font-size: 14.5px; font-weight: 700; color: var(--text); }
  .brand-text .app-tag { font-size: 11px; color: var(--muted); }
  .topbar-right { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
  .topbar-right a { color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px; padding: 6px 12px; border-radius: 20px; transition: background .15s ease; }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }

  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track { position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center; justify-content: space-between; padding: 0 7px; background: linear-gradient(135deg,#8fcaf0,#f4d58d); transition: background .3s ease; }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob { position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1); }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .page-head { padding: 4px 4px 18px; }
  .page-head .eyebrow { font-size: 12px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: var(--gold); margin-bottom: 6px; }
  :root[data-theme="dark"] .page-head .eyebrow { color: var(--gold-light); }
  .page-head h1 { font-size: 22px; margin: 0 0 6px; letter-spacing: -0.01em; }
  .page-head p { color: var(--muted); margin: 0; font-size: 13.5px; }

  .layout { display: grid; grid-template-columns: 320px 1fr; gap: 16px; align-items: start; }
  @media (max-width: 860px) { .layout { grid-template-columns: 1fr; } }

  .panel { background: var(--card); border: 1px solid var(--border); border-radius: 18px; box-shadow: var(--shadow-sm); }
  .sidebar { padding: 14px; max-height: calc(100vh - 150px); overflow-y: auto; scrollbar-gutter: stable; }
  .sidebar input[type=text] {
    width: 100%; padding: 10px 13px; border: 1px solid var(--border); border-radius: 10px;
    font-size: 13.5px; font-family: inherit; background: var(--bg); color: var(--text); margin-bottom: 12px;
  }
  .sidebar input:focus { outline: none; border-color: var(--navy-light); box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 20%, transparent); }
  .add-vessel-row { display: flex; gap: 8px; margin-bottom: 12px; }
  .add-vessel-row input[type=text] { margin-bottom: 0; }
  .add-vessel-row .btn { flex-shrink: 0; }

  .dropzone {
    display: flex; align-items: center; gap: 12px; cursor: pointer;
    border: 1.5px dashed var(--border); border-radius: 14px; padding: 14px;
    margin-bottom: 12px; transition: border-color .15s ease, background .15s ease;
  }
  .dropzone:hover, .dropzone.dragover {
    border-color: var(--navy-light); background: color-mix(in srgb, var(--navy-light) 6%, transparent);
  }
  .dropzone-icon {
    width: 34px; height: 34px; border-radius: 10px; background: var(--ok-bg); color: var(--ok);
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .dropzone-icon svg { width: 18px; height: 18px; }
  .dropzone-text { font-size: 12.5px; color: var(--text); }
  .dropzone-text b { font-weight: 700; }
  .dropzone-sub { font-size: 11px; color: var(--muted); margin-top: 2px; }
  .dropzone-filename { font-size: 11px; color: var(--navy-light); font-weight: 600; margin-top: 3px; }

  .port-head { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); padding: 10px 6px 6px; }
  .vessel-row {
    display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 10px 10px;
    border-radius: 12px; cursor: pointer; transition: background .12s ease; margin-bottom: 2px;
  }
  .vessel-row:hover { background: var(--bg); }
  .vessel-row:hover .row-del { opacity: 1; }
  .vessel-row.active { background: color-mix(in srgb, var(--navy) 12%, transparent); }
  :root[data-theme="dark"] .vessel-row.active { background: color-mix(in srgb, var(--navy-light) 20%, transparent); }
  .vname-wrap { overflow: hidden; min-width: 0; }
  .vessel-row .vname { display: block; font-size: 13.5px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .vessel-row .voperator { display: block; font-size: 11px; color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .row-del {
    flex-shrink: 0; width: 20px; height: 20px; border-radius: 50%; border: none; background: transparent;
    color: var(--muted); font-size: 15px; line-height: 1; cursor: pointer; opacity: 0; transition: opacity .12s ease, background .12s ease, color .12s ease;
    display: flex; align-items: center; justify-content: center;
  }
  .row-del:hover { background: var(--danger-bg); color: var(--danger); opacity: 1; }
  .pill { font-size: 9.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; padding: 3px 8px; border-radius: 999px; white-space: nowrap; flex-shrink: 0; }
  .pill.live { background: var(--ok-bg); color: var(--ok); }
  .pill.live::before { content: ''; display: inline-block; width: 6px; height: 6px; border-radius: 50%; background: var(--ok); margin-right: 5px; animation: pulse 1.8s ease-in-out infinite; }
  .pill.none { background: color-mix(in srgb, var(--border) 70%, transparent); color: var(--muted); }
  @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: .35; } }
  .empty-side { color: var(--muted); font-size: 13px; padding: 20px 8px; text-align: center; }

  .mainpanel { min-height: 560px; display: flex; flex-direction: column; overflow: hidden; }
  .map-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 16px 18px; border-bottom: 1px solid var(--border); flex-wrap: wrap; }
  .map-head h2 { margin: 0; font-size: 16.5px; }
  .map-head .sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .map-actions { display: flex; align-items: center; gap: 8px; }
  .zoom-ctl { display: flex; align-items: center; gap: 2px; background: var(--bg); border: 1px solid var(--border); border-radius: 999px; padding: 3px; margin-right: 4px; }
  .zbtn { width: 26px; height: 26px; border-radius: 50%; border: none; background: transparent; color: var(--navy); font-size: 16px; font-weight: 700; line-height: 1; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: background .12s ease; }
  :root[data-theme="dark"] .zbtn { color: var(--navy-light); }
  .zbtn:hover { background: var(--card); }
  .zbtn:disabled { opacity: .35; cursor: default; }
  .zbtn:disabled:hover { background: transparent; }
  .zlabel { font-size: 11.5px; font-weight: 700; color: var(--muted); width: 20px; text-align: center; }
  .btn { background: var(--navy); color: #fff; border: none; border-radius: 999px; padding: 8px 15px; font-size: 12.5px; font-weight: 600; cursor: pointer; transition: background .15s ease, transform .08s ease; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
  .btn:hover { background: var(--navy-light); }
  .btn:active { transform: scale(.97); }
  .btn.ghost { background: none; color: var(--navy); border: 1px solid var(--border); }
  :root[data-theme="dark"] .btn.ghost { color: var(--navy-light); }
  .btn.ghost:hover { background: var(--border); }
  .btn.ghost.danger { color: var(--danger); }
  .btn.ghost.danger:hover { background: var(--danger-bg); }
  .btn svg { width: 13px; height: 13px; }

  .map-body { flex: 1; position: relative; min-height: 480px; background: color-mix(in srgb, var(--border) 30%, transparent); }
  .map-body iframe { position: absolute; inset: 0; width: 100%; height: 100%; border: 0; }

  .empty-state, .setup-state { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 10px; text-align: center; padding: 30px; }
  .empty-state svg, .setup-state svg { width: 46px; height: 46px; color: var(--muted); opacity: .5; }
  .empty-state h3, .setup-state h3 { margin: 4px 0 0; font-size: 16px; }
  .empty-state p, .setup-state p { margin: 0; color: var(--muted); font-size: 13px; max-width: 340px; }

  .mmsi-form { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; justify-content: center; }
  .mmsi-form input { padding: 10px 13px; border: 1px solid var(--border); border-radius: 10px; font-size: 14px; font-family: inherit; background: var(--card); color: var(--text); width: 190px; text-align: center; letter-spacing: .04em; }
  .mmsi-form input:focus { outline: none; border-color: var(--navy-light); box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 20%, transparent); }
  .mmsi-hint { font-size: 11.5px; color: var(--muted); margin-top: 8px; max-width: 320px; }
  .mmsi-hint a { color: var(--navy); }
  :root[data-theme="dark"] .mmsi-hint a { color: var(--navy-light); }

  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; pointer-events: none; max-width: min(320px, calc(100vw - 40px)); }
  #toastHost .toast { pointer-events: auto; }
  .toast { background: var(--navy-deep); color: #fff; padding: 11px 16px; border-radius: 12px; font-size: 13px; display: flex; align-items: center; gap: 14px; box-shadow: 0 10px 30px rgba(0,0,0,0.25); animation: toast-in .18s ease-out; max-width: 320px; }
  .toast.error { background: var(--danger); }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(8px); } }
</style>
</head><body>
  <div class="topbar">
    <a href="/" class="brand" style="text-decoration:none;">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Vessel Tracker</span>
      </div>
    </a>
    <div class="topbar-right">
      <label class="theme-switch" title="Toggle dark mode">
        <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
        <span class="theme-track">
          <span class="theme-icon sun">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
          </span>
          <span class="theme-icon moon">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
          </span>
          <span class="theme-knob"></span>
        </span>
      </label>
      <a href="/do-tracker">DO Tracker</a>
      <span style="padding:6px 4px;">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="page-head">
    <div class="eyebrow">Live AIS</div>
    <h1>Vessel Tracker</h1>
    <p>Real-time positions for the vessels you're tracking, pulled straight from MarineTraffic. Its own list - separate from DO Tracker.</p>
  </div>

  <div class="layout">
    <div class="panel sidebar">
      <div class="add-vessel-row">
        <input type="text" id="newVesselName" placeholder="Add a vessel name..." maxlength="80" onkeydown="if(event.key==='Enter')addVessel()">
        <button class="btn" onclick="addVessel()">Add</button>
      </div>
      <label class="dropzone" id="particularsDropzone" for="particularsFile">
        <div class="dropzone-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
            <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/>
          </svg>
        </div>
        <div>
          <div class="dropzone-text"><b>Drag in a Ship's Particulars file</b></div>
          <div class="dropzone-sub">.xls or .xlsx - name &amp; MMSI are read automatically</div>
          <div class="dropzone-filename" id="particularsFilename"></div>
        </div>
        <input type="file" id="particularsFile" accept=".xls,.xlsx,.xlsm" style="display:none" onchange="uploadParticulars()">
      </label>
      <input type="text" id="searchBox" placeholder="Search vessel..." oninput="renderList()">
      <div id="vesselList"></div>
    </div>

    <div class="panel mainpanel">
      <div class="map-head" id="mapHead" style="display:none;">
        <div>
          <h2 id="mhName">-</h2>
          <div class="sub" id="mhSub">-</div>
        </div>
        <div class="map-actions">
          <div class="zoom-ctl">
            <button class="zbtn" id="zoomOut" onclick="adjustZoom(-1)" title="Zoom out">&minus;</button>
            <span class="zlabel" id="zoomLabel">12</span>
            <button class="zbtn" id="zoomIn" onclick="adjustZoom(1)" title="Zoom in">+</button>
          </div>
          <button class="btn ghost" id="editMmsiBtn" onclick="showMmsiForm()">Edit MMSI</button>
          <a class="btn ghost" id="openExternal" href="#" target="_blank" rel="noopener">
            Open in MarineTraffic
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M7 17L17 7M7 7h10v10"/></svg>
          </a>
          <button class="btn ghost danger" id="removeVesselBtn" onclick="removeSelectedVessel()" title="Remove vessel">Remove</button>
        </div>
      </div>
      <div class="map-body" id="mapBody">
        <div class="empty-state" id="noSelection">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M2 21c2 1 4 1 6 0s4-1 6 0 4 1 6 0M4 18l1-9 2-3h10l2 3 1 9M9 6V3h6v3"/></svg>
          <h3>Add a vessel to get started</h3>
          <p>Type a vessel name on the left and hit Add, then give it an MMSI to start tracking it live.</p>
        </div>
      </div>
    </div>
  </div>

<div id="toastHost"></div>

<script>
let vessels = {};
let selected = null;

const DEFAULT_CENTER = { lat: 22.5, lon: 41.5 };
const MIN_ZOOM = 3, MAX_ZOOM = 16;
let currentZoom = 12;
try {
  const savedZoom = parseInt(localStorage.getItem('vt_zoom'), 10);
  if (savedZoom && savedZoom >= MIN_ZOOM && savedZoom <= MAX_ZOOM) currentZoom = savedZoom;
} catch (e) {}

function embedUrl(mmsi) {
  return 'https://www.marinetraffic.com/en/ais/embed/zoom:' + currentZoom +
    '/centery:' + DEFAULT_CENTER.lat + '/centerx:' + DEFAULT_CENTER.lon +
    '/maptype:4/shownames:true/mmsi:' + encodeURIComponent(mmsi) +
    '/shipid:0/fleet:/fleet_id:/vtypes:/showmenu:false/remember:false';
}

function adjustZoom(delta) {
  currentZoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, currentZoom + delta));
  try { localStorage.setItem('vt_zoom', currentZoom); } catch (e) {}
  updateZoomControls();
  const mmsi = selected && mmsiMap[selected];
  if (mmsi) showMap(mmsi);
}

function updateZoomControls() {
  const label = document.getElementById('zoomLabel');
  if (label) label.textContent = currentZoom;
  const outBtn = document.getElementById('zoomOut');
  const inBtn = document.getElementById('zoomIn');
  if (outBtn) outBtn.disabled = currentZoom <= MIN_ZOOM;
  if (inBtn) inBtn.disabled = currentZoom >= MAX_ZOOM;
}
function externalUrl(vesselName, mmsi) {
  if (mmsi) return 'https://www.marinetraffic.com/en/ais/details/ships/mmsi:' + encodeURIComponent(mmsi);
  return 'https://www.marinetraffic.com/en/ais/index/search/all?keyword=' + encodeURIComponent(vesselName);
}

async function loadData() {
  const res = await fetch('/api/vessels');
  const list = await res.json();
  vessels = {};
  list.forEach(v => { vessels[v.name] = { mmsi: v.mmsi || '', operator: v.operator || '' }; });
  renderList();
  if (selected && !vessels[selected]) {
    selected = null;
    document.getElementById('mapHead').style.display = 'none';
    document.getElementById('mapBody').innerHTML = '<div class="empty-state" id="noSelection"><h3>Add a vessel to get started</h3><p>Type a vessel name on the left and hit Add, then give it an MMSI to start tracking it live.</p></div>';
  }
}

function naturalCompare(a, b) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

function renderList() {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const names = Object.keys(vessels).filter(v => !q || v.toLowerCase().includes(q)).sort(naturalCompare);
  const listEl = document.getElementById('vesselList');
  if (names.length === 0) {
    listEl.innerHTML = '<div class="empty-side">No vessels yet - add one above to start tracking it.</div>';
    return;
  }
  listEl.innerHTML = names.map(v => {
    const info = vessels[v];
    const isActive = selected === v;
    return '<div class="vessel-row' + (isActive ? ' active' : '') + '" data-vessel="' + escapeHtml(v) + '">' +
      '<div class="vname-wrap"><span class="vname">' + escapeHtml(v) + '</span>' +
      (info.operator ? '<span class="voperator">Operator: ' + escapeHtml(info.operator) + '</span>' : '') + '</div>' +
      (info.mmsi ? '<span class="pill live">Live</span>' : '<span class="pill none">No MMSI</span>') +
      '<button class="row-del" data-del="' + escapeHtml(v) + '" title="Remove vessel">&times;</button>' +
      '</div>';
  }).join('');
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

document.getElementById('vesselList').addEventListener('click', (e) => {
  const delBtn = e.target.closest('.row-del');
  if (delBtn) {
    e.stopPropagation();
    removeVessel(delBtn.dataset.del);
    return;
  }
  const row = e.target.closest('.vessel-row');
  if (!row) return;
  selectVessel(row.dataset.vessel);
});

function selectVessel(name) {
  selected = name;
  renderList();
  const info = vessels[name] || { mmsi: '', operator: '' };
  document.getElementById('mapHead').style.display = 'flex';
  document.getElementById('mhName').textContent = name;
  document.getElementById('mhSub').textContent = info.operator ? ('Operator: ' + info.operator) : 'No operator on file';
  document.getElementById('openExternal').href = externalUrl(name, info.mmsi);
  updateZoomControls();
  if (info.mmsi) {
    showMap(info.mmsi);
  } else {
    showSetup(name);
  }
}

async function addVessel() {
  const input = document.getElementById('newVesselName');
  const name = input.value.trim();
  if (!name) return;
  const res = await fetch('/api/vessels', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name})
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.error) {
    showToast(data.error || 'Could not add that vessel.', {error: true});
    return;
  }
  input.value = '';
  await loadData();
  selectVessel(name);
  showToast('Added ' + name + '.');
}

async function removeVessel(name) {
  if (!confirm('Remove ' + name + ' from the tracker? This cannot be undone.')) return;
  await fetch('/api/vessels/' + encodeURIComponent(name), {method: 'DELETE'});
  if (selected === name) selected = null;
  await loadData();
  showToast('Removed ' + name + '.');
}

function removeSelectedVessel() {
  if (!selected) return;
  removeVessel(selected);
}

/* ---------- Ship's Particulars upload (drag & drop) ---------- */
const particularsDropzone = document.getElementById('particularsDropzone');
['dragenter', 'dragover'].forEach(evt => {
  particularsDropzone.addEventListener(evt, e => { e.preventDefault(); particularsDropzone.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(evt => {
  particularsDropzone.addEventListener(evt, e => { e.preventDefault(); particularsDropzone.classList.remove('dragover'); });
});
particularsDropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (!file) return;
  document.getElementById('particularsFile').files = e.dataTransfer.files;
  uploadParticulars();
});

async function uploadParticulars() {
  const fileInput = document.getElementById('particularsFile');
  const file = fileInput.files[0];
  if (!file) return;
  document.getElementById('particularsFilename').textContent = file.name;

  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch('/api/vessels/upload', { method: 'POST', body: formData });
  const data = await res.json().catch(() => ({}));
  fileInput.value = '';
  if (!res.ok || data.error) {
    showToast(data.error || 'Could not read that file.', {error: true});
    return;
  }
  await loadData();
  selectVessel(data.name);
  showToast(data.had_mmsi ? ('Added ' + data.name + ' with MMSI ' + data.mmsi + '.') : ('Added ' + data.name + ' - no MMSI found, add it manually.'));
}

function showMap(mmsi) {
  document.getElementById('mapBody').innerHTML =
    '<iframe src="' + embedUrl(mmsi) + '" loading="lazy" title="Live vessel position"></iframe>';
}

function showSetup(name) {
  document.getElementById('mapBody').innerHTML =
    '<div class="setup-state">' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 18l1-9 2-3h10l2 3 1 9M9 6V3h6v3M2 21c2 1 4 1 6 0s4-1 6 0 4 1 6 0"/></svg>' +
      '<h3>No MMSI on file for ' + escapeHtml(name) + '</h3>' +
      '<p>Add the vessel\\'s 9-digit MMSI number to start tracking its live position - free, no account needed on your end.</p>' +
      '<div class="mmsi-form">' +
        '<input type="text" id="mmsiInput" placeholder="e.g. 403123456" maxlength="9" inputmode="numeric">' +
        '<button class="btn" onclick="saveMmsi()">Save &amp; Track</button>' +
      '</div>' +
      '<div class="mmsi-hint">Find a vessel\\'s MMSI by searching its name on <a href="https://www.marinetraffic.com/en/ais/index/search/all" target="_blank" rel="noopener">marinetraffic.com</a> - it\\'s listed on the vessel\\'s details page.</div>' +
    '</div>';
}

function showMmsiForm() {
  if (!selected) return;
  showSetup(selected);
  const input = document.getElementById('mmsiInput');
  if (input) input.value = (vessels[selected] && vessels[selected].mmsi) || '';
}

async function saveMmsi() {
  const name = selected;
  if (!name) return;
  const val = document.getElementById('mmsiInput').value.trim();
  if (!/^[0-9]{9}$/.test(val)) {
    showToast('MMSI must be exactly 9 digits.', {error:true});
    return;
  }
  const res = await fetch('/api/vessels/mmsi', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({vessel: name, mmsi: val})
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.error) {
    showToast(data.error || 'Could not save MMSI.', {error:true});
    return;
  }
  if (!vessels[name]) vessels[name] = { mmsi: '', operator: '' };
  vessels[name].mmsi = val;
  showToast('Now tracking ' + name + '.');
  renderList();
  document.getElementById('openExternal').href = externalUrl(name, val);
  showMap(val);
}

function showToast(msg, opts) {
  opts = opts || {};
  const host = document.getElementById('toastHost');
  const el = document.createElement('div');
  el.className = 'toast' + (opts.error ? ' error' : '');
  el.textContent = msg;
  host.appendChild(el);
  setTimeout(() => {
    el.classList.add('fading');
    setTimeout(() => el.remove(), 220);
  }, opts.duration || 3200);
}

(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light'; // default to light for first-time visitors; once they toggle, localStorage remembers it
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}

loadData();
</script>
</body></html>
"""

KPI_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Port Agent KPI</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --danger: #d1483f; --danger-bg: #fbeceb; --ok: #1c8a5a; --ok-bg: #e7f5ee;
    --warn: #b8860b; --warn-bg: #fbf3df;
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05); --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --danger: #e2685f; --danger-bg: #3a2220; --ok: #3ecb8e; --ok-bg: #163329;
    --warn: #e3bb4c; --warn-bg: #362c14;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25); --shadow-md: 0 10px 30px rgba(0,0,0,0.35);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 32px;
    transition: background-color .25s ease, color .25s ease;
  }
  .topbar {
    position: sticky; top: 0; z-index: 50; display: flex; justify-content: space-between; align-items: center;
    gap: 12px; flex-wrap: wrap; padding: 14px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px); -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; text-decoration: none; }
  .brand img { height: 32px; width: auto; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .app-name { font-size: 14.5px; font-weight: 700; color: var(--text); }
  .brand-text .app-tag { font-size: 11px; color: var(--muted); }
  .topbar-right { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
  .topbar-right a { color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px; padding: 6px 12px; border-radius: 20px; transition: background .15s ease; }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }

  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track { position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center; justify-content: space-between; padding: 0 7px; background: linear-gradient(135deg,#8fcaf0,#f4d58d); transition: background .3s ease; }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob { position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1); }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .page-head { padding: 4px 4px 18px; }
  .page-head .eyebrow { font-size: 12px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: var(--gold); margin-bottom: 6px; }
  :root[data-theme="dark"] .page-head .eyebrow { color: var(--gold-light); }
  .page-head h1 { font-size: 22px; margin: 0 0 6px; letter-spacing: -0.01em; }
  .page-head p { color: var(--muted); margin: 0; font-size: 13.5px; }

  .metrics { display: flex; flex-wrap: wrap; gap: 14px; margin-bottom: 20px; }
  .metric-card { flex: 1; min-width: 160px; background: var(--card); border: 1px solid var(--border); border-radius: 16px; box-shadow: var(--shadow-sm); padding: 16px 18px; }
  .metric-card .m-label { font-size: 11.5px; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; font-weight: 700; margin-bottom: 6px; }
  .metric-card .m-value { font-size: 26px; font-weight: 700; }
  .metric-card .m-sub { font-size: 12px; color: var(--muted); margin-top: 2px; }

  .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 20px; }
  @media (max-width: 900px) { .kpi-grid { grid-template-columns: 1fr; } }

  .panel { background: var(--card); border: 1px solid var(--border); border-radius: 18px; box-shadow: var(--shadow-sm); padding: 16px 18px; }
  .panel h2 { font-size: 15px; margin: 0 0 2px; }
  .panel .panel-sub { font-size: 12px; color: var(--muted); margin: 0 0 12px; }

  .backlog-row { display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 9px 0; border-bottom: 1px solid var(--border); font-size: 13px; }
  .backlog-row:last-child { border-bottom: none; }
  .backlog-row .b-bl { font-weight: 700; }
  .backlog-row .b-meta { font-size: 11.5px; color: var(--muted); }
  .days-pill { font-size: 10.5px; font-weight: 700; padding: 3px 9px; border-radius: 999px; white-space: nowrap; }
  .days-pill.ok { background: var(--ok-bg); color: var(--ok); }
  .days-pill.warn { background: var(--warn-bg); color: var(--warn); }
  .days-pill.danger { background: var(--danger-bg); color: var(--danger); }
  .empty-note { color: var(--muted); font-size: 13px; padding: 10px 2px; }

  .workload-panel { margin-top: 4px; }
  table.workload { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  table.workload th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); padding: 8px 10px; border-bottom: 1px solid var(--border); }
  table.workload td { padding: 10px; border-bottom: 1px solid var(--border); }
  table.workload tr:last-child td { border-bottom: none; }

  .loading-note { color: var(--muted); font-size: 13.5px; padding: 30px 4px; text-align: center; }
</style>
</head>
<body>
  <div class="topbar">
    <a href="/" class="brand">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Port Agent KPI</span>
      </div>
    </a>
    <div class="topbar-right">
      <label class="theme-switch" title="Toggle dark mode">
        <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
        <span class="theme-track">
          <span class="theme-icon sun">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
          </span>
          <span class="theme-icon moon">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
          </span>
          <span class="theme-knob"></span>
        </span>
      </label>
      <a href="/do-tracker">DO Tracker</a>
      {% if role == 'admin' %}<a href="/users">Manage Users</a>{% endif %}
      <span style="padding:6px 4px;">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="page-head">
    <div class="eyebrow">Compass</div>
    <h1>Port Agent KPI</h1>
    <p>{% if role == 'admin' %}Built from the DO Tracker board - every staff member's records.{% else %}Built from the DO Tracker board - your own records.{% endif %}</p>
  </div>

  <div id="content">
    <div class="loading-note">Loading...</div>
  </div>

<script>
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light';
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}

const IS_ADMIN = {{ (role == 'admin')|tojson }};

function daysPillClass(days) {
  if (days >= 7) return 'danger';
  if (days >= 3) return 'warn';
  return 'ok';
}

function backlogRowsHtml(list, emptyMsg) {
  if (!list.length) return `<div class="empty-note">${emptyMsg}</div>`;
  return list.map(e => `
    <div class="backlog-row">
      <div>
        <div class="b-bl">${e.bl_number}</div>
        <div class="b-meta">${[e.port, e.vessel].filter(Boolean).join(' &middot; ') || 'No port/vessel set'}${IS_ADMIN ? ' &middot; ' + (e.created_by || 'unknown') : ''}</div>
      </div>
      <span class="days-pill ${daysPillClass(e.days_open)}">${e.days_open}d</span>
    </div>`).join('');
}

function fmtHours(h) {
  if (h === null || h === undefined) return '&ndash;';
  if (h < 48) return h + 'h';
  return (h / 24).toFixed(1) + 'd';
}

async function loadData() {
  const res = await fetch('/api/kpi');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const data = await res.json();

  const workloadHtml = (data.workload && data.workload.length) ? `
    <div class="panel workload-panel">
      <h2>Workload per agent</h2>
      <p class="panel-sub">Who's carrying what, right now.</p>
      <table class="workload">
        <thead><tr><th>Agent</th><th>Total BLs</th><th>Complete</th><th>Pending</th><th>Avg turnaround</th></tr></thead>
        <tbody>
          ${data.workload.map(w => `
            <tr>
              <td><b>${w.agent}</b></td>
              <td>${w.total}</td>
              <td>${w.complete}</td>
              <td>${w.pending}</td>
              <td>${fmtHours(w.avg_turnaround_hours)}</td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>` : '';

  document.getElementById('content').innerHTML = `
    <div class="metrics">
      <div class="metric-card">
        <div class="m-label">Total BLs</div>
        <div class="m-value">${data.total_bls}</div>
      </div>
      <div class="metric-card">
        <div class="m-label">Avg time to invoice</div>
        <div class="m-value">${fmtHours(data.turnaround.avg_hours_to_invoice)}</div>
        <div class="m-sub">from BL added</div>
      </div>
      <div class="metric-card">
        <div class="m-label">Avg time to approval</div>
        <div class="m-value">${fmtHours(data.turnaround.avg_hours_to_approval)}</div>
        <div class="m-sub">from BL added</div>
      </div>
      <div class="metric-card">
        <div class="m-label">Avg time to DO issued</div>
        <div class="m-value">${fmtHours(data.turnaround.avg_hours_to_do)}</div>
        <div class="m-sub">full turnaround</div>
      </div>
    </div>

    <div class="kpi-grid">
      <div class="panel">
        <h2>Pending invoice</h2>
        <p class="panel-sub">${data.backlog.counts.pending_invoice} BL(s) &middot; oldest first</p>
        ${backlogRowsHtml(data.backlog.pending_invoice, 'Nothing pending - invoices are all caught up.')}
      </div>
      <div class="panel">
        <h2>Pending approval</h2>
        <p class="panel-sub">${data.backlog.counts.pending_approval} BL(s) &middot; oldest first</p>
        ${backlogRowsHtml(data.backlog.pending_approval, 'Nothing pending - approvals are all caught up.')}
      </div>
      <div class="panel">
        <h2>Pending DO</h2>
        <p class="panel-sub">${data.backlog.counts.pending_do} BL(s) &middot; oldest first</p>
        ${backlogRowsHtml(data.backlog.pending_do, 'Nothing pending - all DOs are issued.')}
      </div>
    </div>

    ${workloadHtml}
  `;
}

loadData();
</script>
</body></html>
"""

DIRECT_DELIVERY_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Direct Delivery Classifier</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --danger: #d1483f; --danger-bg: #fbeceb; --ok: #1c8a5a; --ok-bg: #e7f5ee;
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05); --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --danger: #e2685f; --danger-bg: #3a2220; --ok: #3ecb8e; --ok-bg: #163329;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25); --shadow-md: 0 10px 30px rgba(0,0,0,0.35);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 32px;
    transition: background-color .25s ease, color .25s ease;
  }
  .topbar {
    position: sticky; top: 0; z-index: 50; display: flex; justify-content: space-between; align-items: center;
    gap: 12px; flex-wrap: wrap; padding: 14px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px); -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; text-decoration: none; }
  .brand img { height: 32px; width: auto; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .app-name { font-size: 14.5px; font-weight: 700; color: var(--text); }
  .brand-text .app-tag { font-size: 11px; color: var(--muted); }
  .topbar-right { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
  .topbar-right a { color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px; padding: 6px 12px; border-radius: 20px; transition: background .15s ease; }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }

  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track { position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center; justify-content: space-between; padding: 0 7px; background: linear-gradient(135deg,#8fcaf0,#f4d58d); transition: background .3s ease; }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob { position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1); }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .page-head { padding: 4px 4px 18px; }
  .page-head .eyebrow { font-size: 12px; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; color: var(--gold); margin-bottom: 6px; }
  :root[data-theme="dark"] .page-head .eyebrow { color: var(--gold-light); }
  .page-head h1 { font-size: 22px; margin: 0 0 6px; letter-spacing: -0.01em; }
  .page-head p { color: var(--muted); margin: 0; font-size: 13.5px; max-width: 640px; }

  .panel { background: var(--card); border: 1px solid var(--border); border-radius: 18px; box-shadow: var(--shadow-sm); padding: 18px 20px; margin-bottom: 18px; }
  .panel h2 { font-size: 15px; margin: 0 0 2px; }
  .panel .panel-sub { font-size: 12px; color: var(--muted); margin: 0 0 14px; }

  .dropzone {
    display: flex; align-items: center; gap: 12px; cursor: pointer;
    border: 1.5px dashed var(--border); border-radius: 14px; padding: 16px;
    transition: border-color .15s ease, background .15s ease;
  }
  .dropzone:hover, .dropzone.dragover {
    border-color: var(--navy-light); background: color-mix(in srgb, var(--navy-light) 6%, transparent);
  }
  .dropzone-icon {
    width: 36px; height: 36px; border-radius: 10px; background: var(--ok-bg); color: var(--ok);
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .dropzone-icon svg { width: 19px; height: 19px; }
  .dropzone-text { font-size: 13px; color: var(--text); }
  .dropzone-text b { font-weight: 700; }
  .dropzone-sub { font-size: 11.5px; color: var(--muted); margin-top: 2px; }

  .status-line { font-size: 13px; color: var(--muted); margin-top: 12px; min-height: 18px; }
  .status-line.ok { color: var(--ok); }
  .status-line.error { color: var(--danger); }

  table.dd-table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  table.dd-table th { text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); padding: 8px 10px; border-bottom: 1px solid var(--border); }
  table.dd-table td { padding: 10px; border-bottom: 1px solid var(--border); }
  table.dd-table tr:last-child td { border-bottom: none; }

  .dd-badge { display:inline-block; padding:3px 9px; border-radius:20px; font-size:11px; font-weight:700; letter-spacing:.3px; text-transform:uppercase; }
  .dd-badge.dd-yes { background: rgba(212,160,23,0.16); color:#8a6d1f; border:1px solid rgba(212,160,23,0.4); }
  :root[data-theme="dark"] .dd-badge.dd-yes { color: var(--gold-light); }
  .dd-badge.dd-no { background: rgba(120,120,120,0.12); color: var(--muted); border:1px solid var(--border); }

  .empty-note { color: var(--muted); font-size: 13px; padding: 10px 2px; }
  .unmatched-note { font-size: 12.5px; color: var(--muted); margin-top: 10px; padding: 10px 12px; background: var(--bg); border-radius: 10px; }
</style>
</head>
<body>
  <div class="topbar">
    <a href="/" class="brand">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Direct Delivery Classifier</span>
      </div>
    </a>
    <div class="topbar-right">
      <label class="theme-switch" title="Toggle dark mode">
        <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
        <span class="theme-track">
          <span class="theme-icon sun">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
          </span>
          <span class="theme-icon moon">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
          </span>
          <span class="theme-knob"></span>
        </span>
      </label>
      <a href="/do-tracker">DO Tracker</a>
      {% if role == 'admin' %}<a href="/users">Manage Users</a>{% endif %}
      <span style="padding:6px 4px;">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="page-head">
    <div class="eyebrow">Compass</div>
    <h1>Direct Delivery Classifier</h1>
    <p>Upload a cargo packing list and every BL over 30MT or 12m gets flagged as Direct Delivery - unless it's wheeled or a coil, in which case it doesn't need a low-bed trailer. This is fully standalone - separate from DO Tracker.</p>
  </div>

  <div class="panel">
    <h2>Classify packing lists</h2>
    <p class="panel-sub">.xlsx, .xls, .csv, .pdf, .doc/.docx, or a scanned photo (.png/.jpg) - select or drop as many files as you have at once. Reads the weight/length/description columns automatically, and a BL split across several files (same reference number, different pages) is grouped back into one BL.</p>
    <label class="dropzone" id="dropzone" for="classifyFile">
      <div class="dropzone-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
          <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/>
        </svg>
      </div>
      <div>
        <div class="dropzone-text"><b>Click to upload</b> or drag &amp; drop your packing lists</div>
        <div class="dropzone-sub">.xlsx, .xls, .csv, .pdf, .doc/.docx, .png/.jpg - multiple files at once is fine</div>
      </div>
      <input type="file" id="classifyFile" accept=".xlsx,.xlsm,.xls,.csv,.pdf,.doc,.docx,.png,.jpg,.jpeg,.bmp,.tif,.tiff" multiple style="display:none" onchange="uploadClassify()">
    </label>
    <div class="status-line" id="statusLine"></div>
  </div>

  <div class="panel" id="reviewPanel" style="display:none;">
    <h2>⚠ Needs a quick check</h2>
    <p class="panel-sub" id="reviewSub">Loading...</p>
    <div id="reviewBody"></div>
  </div>

  <div class="panel">
    <h2>Direct Delivery BLs</h2>
    <p class="panel-sub" id="resultsSub">Loading...</p>
    <div id="resultsBody"></div>
  </div>

<script>
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light';
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}

const dropzone = document.getElementById('dropzone');
['dragenter', 'dragover'].forEach(evt => {
  dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(evt => {
  dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.remove('dragover'); });
});
dropzone.addEventListener('drop', e => {
  if (e.dataTransfer.files.length) { document.getElementById('classifyFile').files = e.dataTransfer.files; uploadClassify(); }
});

function ddBadgeHtml(direct) {
  return `<span class="dd-badge ${direct ? 'dd-yes' : 'dd-no'}">${direct ? 'Direct Delivery' : 'Not direct'}</span>`;
}

async function uploadClassify() {
  const input = document.getElementById('classifyFile');
  const files = input.files;
  if (!files || !files.length) return;
  const status = document.getElementById('statusLine');
  status.className = 'status-line';
  status.textContent = files.length === 1 ? 'Classifying...' : `Classifying ${files.length} files...`;
  const fd = new FormData();
  for (const f of files) fd.append('file', f);
  try {
    const res = await fetch('/api/manifest/classify', {method: 'POST', body: fd});
    const data = await res.json();
    if (!res.ok) {
      status.className = 'status-line error';
      status.textContent = data.error || 'Could not classify those files.';
      input.value = '';
      return;
    }
    const classified = data.classified || [];
    const failed = data.failed || [];
    const direct = classified.filter(m => m.direct).length;
    const flagged = classified.filter(m => m.needs_review).length;
    status.className = failed.length ? 'status-line error' : 'status-line ok';
    let msg = classified.length
      ? `Classified ${classified.length} BL(s) - ${direct} direct delivery.`
      : 'No BLs could be read from those files.';
    if (flagged) msg += ` ${flagged} flagged for a quick manual check.`;
    if (failed.length) msg += ` Couldn't read: ${failed.join(', ')}.`;
    status.textContent = msg;
    input.value = '';
    await loadResults();
    await loadReview();
  } catch (e) {
    status.className = 'status-line error';
    status.textContent = 'Could not classify those files.';
    input.value = '';
  }
}

async function loadResults() {
  const res = await fetch('/api/direct-delivery');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const rows = await res.json();
  const sub = document.getElementById('resultsSub');
  const body = document.getElementById('resultsBody');
  if (!rows.length) {
    sub.textContent = 'No direct delivery BLs yet.';
    body.innerHTML = '<div class="empty-note">Upload a packing list above - any BL that comes back Direct Delivery will show up here.</div>';
    return;
  }
  sub.textContent = `${rows.length} direct delivery BL(s).`;
  body.innerHTML = `
    <table class="dd-table">
      <thead><tr><th>BL Number</th><th>Reason</th><th>Classified</th><th></th></tr></thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td><b>${r.bl_number}</b></td>
            <td style="color:var(--muted);">${r.reason || ''}</td>
            <td style="color:var(--muted);">${r.classified_by || ''}${r.classified_at ? ' - ' + r.classified_at : ''}</td>
            <td><button class="btn ghost danger" style="padding:4px 10px;font-size:11.5px;" onclick="removeDirectDelivery('${r.bl_number.replace(/'/g, "\\\\'")}')">Remove</button></td>
          </tr>`).join('')}
      </tbody>
    </table>`;
}

async function loadReview() {
  const res = await fetch('/api/direct-delivery/review');
  if (res.status === 401 || res.redirected) return;
  const rows = await res.json();
  const panel = document.getElementById('reviewPanel');
  const sub = document.getElementById('reviewSub');
  const body = document.getElementById('reviewBody');
  if (!rows.length) { panel.style.display = 'none'; return; }
  panel.style.display = '';
  sub.textContent = `${rows.length} BL(s) where something about the read was uncertain - not necessarily wrong, just worth a glance at the source file.`;
  body.innerHTML = `
    <table class="dd-table">
      <thead><tr><th>BL Number</th><th>Verdict</th><th>Why flagged</th><th></th></tr></thead>
      <tbody>
        ${rows.map(r => `
          <tr>
            <td><b>${r.bl_number}</b></td>
            <td>${ddBadgeHtml(!!r.is_direct)}</td>
            <td style="color:var(--muted);">${r.review_note || ''}</td>
            <td><button class="btn ghost danger" style="padding:4px 10px;font-size:11.5px;" onclick="removeDirectDelivery('${r.bl_number.replace(/'/g, "\\\\'")}', true)">Remove</button></td>
          </tr>`).join('')}
      </tbody>
    </table>`;
}

async function removeDirectDelivery(bl, fromReview) {
  if (!confirm(`Remove ${bl} from this list? This won't affect the source files, only this table.`)) return;
  try {
    await fetch(`/api/direct-delivery/${encodeURIComponent(bl)}`, {method: 'DELETE'});
  } catch (e) {}
  await loadResults();
  await loadReview();
}

loadResults();
loadReview();
</script>
</body></html>
"""

PAGE_HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Compass - DO Tracker</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root {
    --bg: #f2f4f7;
    --bg-glow: radial-gradient(circle at 15% -10%, rgba(31,92,133,0.10), transparent 45%),
                radial-gradient(circle at 100% 0%, rgba(201,162,39,0.08), transparent 40%);
    --card: #ffffff;
    --text: #1c2b3a;
    --muted: #7a8794;
    --border: #e6e9ed;
    --navy: #123a56;
    --navy-deep: #0b2740;
    --navy-light: #1f5c85;
    --gold: #c9a227;
    --gold-light: #e0bd53;
    --success: #1f9d55;
    --success-bg: #eaf7ef;
    --danger: #d1483f;
    --danger-bg: #fbeceb;
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05);
    --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23;
    --bg-glow: radial-gradient(circle at 15% -10%, rgba(63,134,186,0.14), transparent 45%),
                radial-gradient(circle at 100% 0%, rgba(227,187,76,0.08), transparent 40%);
    --card: #1a232f;
    --text: #e9eef3;
    --muted: #93a1b1;
    --border: #29323f;
    --navy: #3f86ba;
    --navy-deep: #274a67;
    --navy-light: #5aa2d1;
    --gold: #e3bb4c;
    --gold-light: #f0cf72;
    --success: #3ecb7d;
    --success-bg: #163627;
    --danger: #e2685f;
    --danger-bg: #3a2220;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25);
    --shadow-md: 0 12px 32px rgba(0,0,0,0.45);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  html { -webkit-font-smoothing: antialiased; overflow-y: scroll; scrollbar-gutter: stable; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg-glow), var(--bg);
    color: var(--text); margin: 0; padding: 0 16px 40px;
    transition: background-color .25s ease, color .25s ease;
  }

  /* Header */
  .topbar {
    position: sticky; top: 0; z-index: 50;
    display: flex; justify-content: space-between; align-items: center;
    gap: 12px; flex-wrap: wrap;
    padding: 12px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px);
    -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand img { height: 36px; width: auto; display: block; }
  .brand-text { display: flex; flex-direction: column; line-height: 1.15; }
  .brand-text .app-name { font-size: 15px; font-weight: 700; color: var(--text); letter-spacing: -0.01em; }
  .brand-text .app-tag { font-size: 11px; color: var(--muted); font-weight: 500; }
  .topbar-right { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
  .topbar-right a {
    color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px;
    padding: 6px 12px; border-radius: 20px; transition: background .15s ease;
  }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }
  .who { padding: 6px 4px; }
  .who b { color: var(--text); }

  /* Sun/moon theme switch */
  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; flex-shrink: 0; }
  .theme-switch input { opacity: 0; width: 0; height: 0; position: absolute; }
  .theme-track {
    position: absolute; inset: 0; border-radius: 999px; display: flex; align-items: center;
    justify-content: space-between; padding: 0 7px;
    background: linear-gradient(135deg,#8fcaf0,#f4d58d);
    transition: background .3s ease;
  }
  :root[data-theme="dark"] .theme-track { background: linear-gradient(135deg,#1f2b42,#33456a); }
  .theme-icon { width: 13px; height: 13px; color: #fff; opacity: .9; z-index: 1; }
  .theme-icon svg { width: 100%; height: 100%; }
  .theme-knob {
    position: absolute; top: 3px; left: 3px; width: 23px; height: 23px; border-radius: 50%;
    background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,0.3); transition: transform .3s cubic-bezier(.4,0,.2,1);
  }
  input:checked + .theme-track .theme-knob { transform: translateX(25px); background: #0b2740; }

  .sub { color: var(--muted); font-size: 13px; margin: 2px 0 18px; }

  /* Cards */
  .card {
    background: var(--card); border: 1px solid var(--border); border-radius: 16px;
    padding: 18px; margin-bottom: 16px; box-shadow: var(--shadow-sm);
    transition: background-color .25s ease, border-color .25s ease;
  }
  .card-label { font-size: 12px; font-weight: 600; color: var(--navy); text-transform: uppercase; letter-spacing: .04em; margin-bottom: 10px; }
  :root[data-theme="dark"] .card-label { color: var(--navy-light); }
  .row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }

  input[type=text] {
    border: 1px solid var(--border); border-radius: 10px; padding: 10px 12px;
    font-size: 14px; font-family: inherit; width: 100%; background: var(--bg);
    color: var(--text);
    transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
  }
  input[type=text]:focus {
    outline: none; border-color: var(--navy-light); background: var(--card);
    box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 20%, transparent);
  }
  .tag-fields { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
  .tag-fields > div { flex: 1; min-width: 180px; }
  .tag-fields label { display: block; font-size: 11px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; margin-bottom: 5px; }

  button {
    background: var(--navy); color: #fff; border: none; border-radius: 999px;
    padding: 10px 18px; font-size: 13px; font-weight: 600; cursor: pointer;
    transition: background .15s ease, transform .08s ease;
  }
  button:hover { background: var(--navy-light); }
  button:active { transform: scale(0.97); }
  button:disabled { background: var(--border); color: var(--muted); cursor: not-allowed; }
  button:disabled:hover { background: var(--border); }

  /* Excel dropzone */
  .dropzone {
    display: flex; align-items: center; gap: 14px; cursor: pointer;
    border: 1.5px dashed var(--border); border-radius: 14px; padding: 20px;
    transition: border-color .15s ease, background .15s ease;
  }
  .dropzone:hover, .dropzone.dragover {
    border-color: var(--navy-light); background: color-mix(in srgb, var(--navy-light) 6%, transparent);
  }
  .dropzone-icon {
    width: 42px; height: 42px; border-radius: 12px; background: var(--success-bg); color: var(--success);
    display: flex; align-items: center; justify-content: center; flex-shrink: 0;
  }
  .dropzone-icon svg { width: 22px; height: 22px; }
  .dropzone-text { font-size: 13.5px; color: var(--text); }
  .dropzone-text b { font-weight: 700; }
  .dropzone-sub { font-size: 12px; color: var(--muted); margin-top: 2px; }
  .dropzone-filename { font-size: 12px; color: var(--navy-light); font-weight: 600; margin-top: 4px; }

  /* Summary stats */
  .summary { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
  .stat {
    background: var(--card); border: 1px solid var(--border); border-radius: 14px;
    padding: 14px 16px; font-size: 12px; color: var(--muted); flex: 1; min-width: 130px;
    box-shadow: var(--shadow-sm); display: flex; align-items: center; gap: 12px;
  }
  .stat-icon {
    width: 36px; height: 36px; border-radius: 10px; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center;
    background: color-mix(in srgb, var(--navy-light) 12%, transparent); color: var(--navy-light);
  }
  .stat.gold .stat-icon { background: color-mix(in srgb, var(--gold) 16%, transparent); color: var(--gold); }
  .stat.done .stat-icon { background: var(--success-bg); color: var(--success); }
  .stat-icon svg { width: 19px; height: 19px; stroke: currentColor; fill: none; stroke-width: 1.8; stroke-linecap: round; stroke-linejoin: round; }
  .stat b { display: block; font-size: 21px; font-weight: 700; color: var(--text); letter-spacing: -0.01em; line-height: 1.2; }

  /* Port / Vessel group structure */
  .port-group { margin-bottom: 18px; }
  .port-header {
    display: flex; align-items: center; gap: 10px; cursor: pointer;
    background: var(--navy-deep); color: #fff; padding: 12px 16px; border-radius: 14px 14px 0 0;
  }
  .port-header .chev { width: 14px; height: 14px; transition: transform .18s ease; flex-shrink: 0; }
  .port-header.collapsed .chev { transform: rotate(-90deg); }
  .port-header .group-name {
    font-size: 14px; font-weight: 700; letter-spacing: .01em; background: transparent;
    border: 1px solid transparent; color: #fff; border-radius: 6px; padding: 2px 6px; font-family: inherit;
  }
  .port-header .group-name:focus { outline: none; border-color: rgba(255,255,255,0.4); background: rgba(255,255,255,0.08); }
  .port-header .group-count { font-size: 11.5px; color: rgba(255,255,255,0.7); font-weight: 500; }
  .group-remove {
    margin-left: auto; background: none; border: none; cursor: pointer;
    font-size: 11px; font-weight: 700; padding: 5px 11px; border-radius: 999px; flex-shrink: 0;
  }
  .port-header .group-remove { color: rgba(255,255,255,0.75); }
  .port-header .group-remove:hover { background: rgba(255,255,255,0.14); color: #fff; }
  .vessel-header .group-remove { color: var(--danger); }
  .vessel-header .group-remove:hover { background: var(--danger-bg); }
  .port-body { border: 1px solid var(--border); border-top: none; border-radius: 0 0 14px 14px; overflow: hidden; background: var(--card); }
  .port-body.collapsed { display: none; }

  .vessel-group { border-bottom: 1px solid var(--border); }
  .vessel-group:last-child { border-bottom: none; }
  .vessel-header {
    display: flex; align-items: center; gap: 10px; cursor: pointer;
    background: color-mix(in srgb, var(--gold) 10%, transparent); padding: 10px 16px;
  }
  .vessel-header .chev { width: 12px; height: 12px; color: var(--gold); transition: transform .18s ease; flex-shrink: 0; }
  .vessel-header.collapsed .chev { transform: rotate(-90deg); }
  .vessel-header .group-name {
    font-size: 13px; font-weight: 700; color: var(--text); background: transparent;
    border: 1px solid transparent; border-radius: 6px; padding: 2px 6px; font-family: inherit;
  }
  .vessel-header .group-name:focus { outline: none; border-color: var(--border); background: var(--card); }
  .vessel-header .group-count { font-size: 11px; color: var(--muted); }
  .vessel-body.collapsed { display: none; }

  /* Table */
  .overflow { overflow-x: auto; }
  table { width: 100%; min-width: 720px; border-collapse: collapse; font-size: 13.5px; table-layout: fixed; }
  th:nth-child(1), td:nth-child(1) { width: 16%; }
  th:nth-child(2), td:nth-child(2) { width: 16%; }
  th:nth-child(3), td:nth-child(3) { width: 16%; }
  th:nth-child(4), td:nth-child(4) { width: 16%; }
  th:nth-child(5), td:nth-child(5) { width: 26%; }
  th:nth-child(6), td:nth-child(6) { width: 10%; }
  th, td { text-align: left; padding: 12px 10px; border-bottom: 1px solid var(--border); overflow: hidden; }
  th {
    color: var(--muted); font-weight: 600; font-size: 10.5px; text-transform: uppercase;
    letter-spacing: .05em; background: color-mix(in srgb, var(--border) 40%, transparent);
  }
  tbody tr { transition: background .12s ease; }
  tbody tr:hover { background: color-mix(in srgb, var(--navy-light) 4%, transparent); }
  tbody tr:last-child td { border-bottom: none; }

  .bl-cell { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; overflow: hidden; }
  .bl-cell b { font-weight: 700; letter-spacing: -0.01em; overflow: hidden; text-overflow: ellipsis; }
  .badge-complete {
    display: inline-flex; align-items: center; gap: 3px;
    background: var(--success-bg); color: var(--success); font-size: 10.5px; font-weight: 700;
    padding: 2px 8px; border-radius: 999px; text-transform: uppercase; letter-spacing: .03em;
  }

  .checkwrap { display: flex; flex-direction: column; gap: 3px; align-items: flex-start; min-height: 34px; justify-content: center; max-width: 100%; }
  .meta { font-size: 10px; color: var(--muted); max-width: 100%; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .remarks-input {
    width: 100%; border: 1px solid transparent; background: transparent; color: var(--text);
    font-size: 12.5px; font-family: inherit; padding: 5px 6px; border-radius: 6px;
  }
  .remarks-input:focus { border-color: var(--border); background: var(--bg); box-shadow: none; }
  .del {
    background: none; color: var(--danger); font-size: 12px; font-weight: 600;
    padding: 5px 10px; border-radius: 999px;
  }
  .del:hover { background: var(--danger-bg); }

  /* Sliding toggle switch */
  .switch { position: relative; display: inline-block; width: 42px; height: 23px; flex-shrink: 0; }
  .switch input { opacity: 0; width: 0; height: 0; }
  .slider {
    position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0;
    background-color: var(--border); transition: background-color .2s ease; border-radius: 24px;
  }
  .slider:before {
    position: absolute; content: ""; height: 17px; width: 17px; left: 3px; bottom: 3px;
    background-color: #fff; transition: transform .2s ease; border-radius: 50%;
    box-shadow: 0 1px 3px rgba(0,0,0,0.3);
  }
  input:checked + .slider { background-color: var(--gold); }
  input:checked + .slider:before { transform: translateX(19px); }

  /* Toast notifications (replace confirm()/alert() popups) */
  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; pointer-events: none; max-width: min(320px, calc(100vw - 40px)); }
  #toastHost .toast { pointer-events: auto; }
  .toast {
    background: var(--navy-deep); color: #fff; padding: 11px 16px; border-radius: 12px; font-size: 13px;
    display: flex; align-items: center; gap: 14px; box-shadow: var(--shadow-md);
    animation: toast-in .18s ease-out; max-width: 320px;
  }
  .toast a { color: var(--gold-light); font-weight: 700; text-decoration: none; cursor: pointer; white-space: nowrap; }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(8px); } }

  @media (max-width: 600px) {
    .stat { min-width: 45%; }
  }
</style>
</head>
<body>
  <div class="topbar">
    <a href="/" class="brand" style="text-decoration:none;">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">DO Tracker</span>
      </div>
    </a>
    <div class="topbar-right">
      <label class="theme-switch" title="Toggle dark mode">
        <input type="checkbox" id="themeToggle" onchange="setTheme(this.checked ? 'dark' : 'light')">
        <span class="theme-track">
          <span class="theme-icon sun">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4 12H2M22 12h-2M5 5l1.4 1.4M17.6 17.6L19 19M19 5l-1.4 1.4M6.4 17.6L5 19"/></svg>
          </span>
          <span class="theme-icon moon">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.8A8.5 8.5 0 1111.2 3a7 7 0 009.8 9.8z"/></svg>
          </span>
          <span class="theme-knob"></span>
        </span>
      </label>
      {% if role == 'admin' %}<a href="/users">Manage Users</a>{% endif %}
      <span class="who">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="card">
    <div class="card-label">Add a manifest</div>
    <div class="tag-fields">
      <div>
        <label for="portField">Discharge Port</label>
        <input type="text" id="portField" list="portDatalist" placeholder="e.g. JEDDAH PORT" style="text-transform:uppercase;"
          oninput="this.value = this.value.toUpperCase(); checkPortSimilar();" onblur="checkPortSimilar()">
        <datalist id="portDatalist"></datalist>
        <div style="font-size:11px; color:var(--muted); margin-top:4px;">Where it's being delivered to - not the Chinese loading port. Pick an existing port from the list so every BL going there lands in the same group.</div>
        <div id="portWarning" style="display:none; font-size:11.5px; color:#9a6b00; background:#fff7e6; border:1px solid #f1d28a; border-radius:6px; padding:6px 9px; margin-top:6px;"></div>
      </div>
      <div>
        <label for="vesselField">Vessel</label>
        <input type="text" id="vesselField" list="vesselDatalist" placeholder="e.g. TAI KNIGHT" style="text-transform:uppercase;" oninput="this.value = this.value.toUpperCase();">
        <datalist id="vesselDatalist"></datalist>
      </div>
    </div>
    <label class="dropzone" id="dropzone" for="manifestFile">
      <div class="dropzone-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
          <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/>
        </svg>
      </div>
      <div>
        <div class="dropzone-text"><b>Click to upload</b> or drag &amp; drop your manifest</div>
        <div class="dropzone-sub">.xlsx, .xls, .csv, .docx or .pdf - the BL Number column is read automatically</div>
        <div class="dropzone-filename" id="dropzoneFilename"></div>
      </div>
      <input type="file" id="manifestFile" accept=".xlsx,.xlsm,.xls,.csv,.docx,.pdf" style="display:none" onchange="stageManifestFile()">
    </label>
    <div class="row" style="margin-top:14px;">
      <button type="button" id="addManifestBtn" onclick="uploadExcel()" disabled>Add to board</button>
    </div>
  </div>

  <div class="summary" id="summary"></div>

  <div class="card">
    <div class="row" style="margin-bottom:14px; flex-wrap:wrap;">
      <input type="text" id="searchBox" placeholder="Search BL number..." oninput="render()" style="flex:1; min-width:180px;">
      <select id="jumpSelect" onchange="jumpToVessel(this.value)" style="min-width:200px;"><option value="">Jump to a vessel...</option></select>
      <button type="button" onclick="setAllGroupsCollapsed(false)" style="background:none; color:var(--text); border:1px solid var(--border);">Expand all</button>
      <button type="button" onclick="setAllGroupsCollapsed(true)" style="background:none; color:var(--text); border:1px solid var(--border);">Collapse all</button>
      <button type="button" id="clearAllBtn" onclick="clearAllRecords()" style="background:none; color:var(--danger); border:1px solid var(--border);">Clear board</button>
    </div>
    <div id="groups"></div>
  </div>

  <div id="toastHost"></div>

<script>
/* ---------- Theme (light/dark, sun/moon toggle) ---------- */
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || 'light'; // default to light for first-time visitors; once they toggle, localStorage remembers it
  document.documentElement.setAttribute('data-theme', mode);
  window.addEventListener('DOMContentLoaded', () => {
    const cb = document.getElementById('themeToggle');
    if (cb) cb.checked = mode === 'dark';
  });
})();
function setTheme(mode) {
  document.documentElement.setAttribute('data-theme', mode);
  try { localStorage.setItem('theme', mode); } catch (e) {}
}

const CURRENT_USER = {{ username|tojson }};
let records = [];
let suppressPollUntil = 0;
let editingCount = 0;
let collapsedGroups = {};

function markEditing(delta) {
  editingCount = Math.max(0, editingCount + delta);
  if (editingCount === 0) render();
}

const MAX_TOASTS = 3;
function showToast(message, opts) {
  opts = opts || {};
  const host = document.getElementById('toastHost');

  // Deleting several BLs in a row used to stack up an ever-growing pile of
  // "Undo" toasts that covered the board and blocked further clicks. Cap
  // how many can be on screen at once - anything older is removed outright
  // (not animated - an animated fade only *schedules* removal, so the
  // count wouldn't actually shrink yet and this loop would spin forever).
  while (host.children.length >= MAX_TOASTS) {
    const oldest = host.firstElementChild;
    if (!oldest) break;
    if (oldest._timer) clearTimeout(oldest._timer);
    oldest.remove();
  }

  const el = document.createElement('div');
  el.className = 'toast';
  const text = document.createElement('span');
  text.textContent = message;
  el.appendChild(text);
  if (opts.actionLabel && typeof opts.onAction === 'function') {
    const a = document.createElement('a');
    a.textContent = opts.actionLabel;
    a.onclick = () => { opts.onAction(); dismiss(); };
    el.appendChild(a);
  }
  host.appendChild(el);
  const duration = opts.duration || 3500;
  const timer = setTimeout(dismiss, duration);
  el._timer = timer;
  function dismiss() {
    clearTimeout(timer);
    el.classList.add('fading');
    setTimeout(() => el.remove(), 220);
  }
}

function naturalCompare(a, b) {
  const re = /(\d+)|(\D+)/g;
  const ax = String(a || '').match(re) || [];
  const bx = String(b || '').match(re) || [];
  const len = Math.max(ax.length, bx.length);
  for (let i = 0; i < len; i++) {
    const av = ax[i] || '', bv = bx[i] || '';
    if (av === bv) continue;
    const an = parseInt(av, 10), bn = parseInt(bv, 10);
    if (!isNaN(an) && !isNaN(bn)) {
      if (an !== bn) return an - bn;
    } else {
      return av < bv ? -1 : 1;
    }
  }
  return 0;
}

async function fetchRecords() {
  if (Date.now() < suppressPollUntil) return;
  const res = await fetch('/api/records');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const fresh = await res.json();
  fresh.forEach(nr => {
    if (remarksTimers[nr.bl_number]) {
      const old = records.find(r => r.bl_number === nr.bl_number);
      if (old) nr.remarks = old.remarks;
    }
  });
  const changed = JSON.stringify(fresh) !== JSON.stringify(records);
  records = fresh;
  if (changed) refreshFieldSuggestions();
  if (editingCount === 0 && changed) render();
}

/* ---------- Port/vessel autocomplete + duplicate-group warning ----------
   With 20-25 manifests a month, retyping the discharge port/vessel by hand
   every time is exactly how "JEDDAH" and "JEDDAH PORT" end up as two
   separate board groups for the same place - a human typo, not a parsing
   bug, but the UI should make it hard to make rather than relying on
   everyone remembering the exact spelling used last time. Two guards:
   1) a <datalist> of every port/vessel already on the board, so picking
      an existing one is a dropdown click instead of retyping it, and
   2) a live warning if what's typed LOOKS like an existing port under a
      different spelling (extra "PORT" word, punctuation, spacing) -
      with a one-click button to adopt the existing spelling exactly. */
// Every discharge port Sea Power regularly handles, so the dropdown offers
// the full list from day one - not just ports a manifest has already been
// uploaded under. Send the real list and this gets hardcoded here instead.
const KNOWN_PORTS = ['JEDDAH', 'DAMMAM', 'JUBAIL', 'YANBU', 'KING ABDULLAH PORT', 'RIYADH DRY PORT'];

function refreshFieldSuggestions() {
  const portField = document.getElementById('portField');
  const vesselField = document.getElementById('vesselField');
  if (!portField || !vesselField) return;

  const ports = [...new Set([...KNOWN_PORTS, ...records.map(r => r.port).filter(Boolean)])].sort(naturalCompare);
  const vessels = [...new Set(records.map(r => r.vessel).filter(Boolean))].sort(naturalCompare);

  document.getElementById('portDatalist').innerHTML = ports.map(p => `<option value="${p.replace(/"/g, '&quot;')}">`).join('');
  document.getElementById('vesselDatalist').innerHTML = vessels.map(v => `<option value="${v.replace(/"/g, '&quot;')}">`).join('');
}

function normPortName(s) {
  // Collapses spelling variants of the "same" port down to one key: splits
  // into alphanumeric tokens and drops the generic token "PORT", so
  // "JEDDAH", "JEDDAH PORT" and "JEDDAH-PORT" all normalize the same way.
  // (Deliberately not a word-boundary regex around PORT - this file is a
  // plain, non-raw Python triple-quoted string, so that particular escape
  // sequence is read by Python as a backspace character before the JS
  // ever reaches the browser, silently breaking the match. Token-splitting
  // sidesteps that whole class of mistake.)
  return String(s || '').toUpperCase().split(/[^A-Z0-9]+/).filter(t => t && t !== 'PORT').join('');
}

function checkPortSimilar() {
  const warnEl = document.getElementById('portWarning');
  const typed = document.getElementById('portField').value.trim();
  if (!typed) { warnEl.style.display = 'none'; return; }

  const existingPorts = [...new Set(records.map(r => r.port).filter(Boolean))];
  const typedNorm = normPortName(typed);
  const match = existingPorts.find(p => p !== typed && normPortName(p) === typedNorm);

  if (match) {
    warnEl.innerHTML = `This looks like <b>${match}</b>, already on the board - use that exact spelling so this manifest joins the same group instead of starting a new one. <button type="button" style="margin-left:4px; font-size:11px; padding:2px 8px;" onclick="document.getElementById('portField').value='${match.replace(/'/g, "\\'")}'; checkPortSimilar();">Use "${match}"</button>`;
    warnEl.style.display = 'block';
  } else {
    warnEl.style.display = 'none';
  }
}

/* ---------- Manifest upload (drag & drop) ----------
   The file is only *staged* here - it does NOT upload right away, so
   there's time to fill in Port/Vessel first. It only actually uploads
   when "Add to board" is clicked (uploadExcel below). */
const dropzone = document.getElementById('dropzone');
['dragenter', 'dragover'].forEach(evt => {
  dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(evt => {
  dropzone.addEventListener(evt, e => { e.preventDefault(); dropzone.classList.remove('dragover'); });
});
dropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (!file) return;
  document.getElementById('manifestFile').files = e.dataTransfer.files;
  stageManifestFile();
});

function stageManifestFile() {
  const fileInput = document.getElementById('manifestFile');
  const file = fileInput.files[0];
  const btn = document.getElementById('addManifestBtn');
  if (!file) {
    document.getElementById('dropzoneFilename').textContent = '';
    btn.disabled = true;
    return;
  }
  document.getElementById('dropzoneFilename').textContent = file.name;
  btn.disabled = false;
}

async function uploadExcel() {
  const fileInput = document.getElementById('manifestFile');
  const file = fileInput.files[0];
  if (!file) { showToast('Choose a manifest file first.'); return; }

  const btn = document.getElementById('addManifestBtn');
  btn.disabled = true;
  const originalLabel = btn.textContent;
  btn.textContent = 'Adding...';

  const port = document.getElementById('portField').value.trim().toUpperCase();
  const vessel = document.getElementById('vesselField').value.trim().toUpperCase();

  const formData = new FormData();
  formData.append('file', file);
  formData.append('port', port);
  formData.append('vessel', vessel);

  const res = await fetch('/api/manifest/upload', { method: 'POST', body: formData });
  const data = await res.json();
  if (data.error) {
    showToast(data.error);
    btn.disabled = false;
    btn.textContent = originalLabel;
    return;
  }

  // Reset the dropzone so the same "Add to board" flow can be repeated
  // for the next manifest without leftover state from this one.
  fileInput.value = '';
  document.getElementById('dropzoneFilename').textContent = '';
  btn.textContent = originalLabel;

  await fetchRecords();
  showToast(data.added + ' new BL record(s) added' + (data.skipped ? `, ${data.skipped} already on the board (skipped)` : '') + '.');
}

function nowLabel() {
  // Stored/compared as a naive UTC string ("YYYY-MM-DD HH:MM"), same as the
  // server - formatLocalTime() below converts it to the viewer's own
  // timezone whenever it's actually displayed.
  const d = new Date();
  return d.toISOString().slice(0, 16).replace('T', ' ');
}

function formatLocalTime(raw) {
  if (!raw) return '';
  // The server stores these as naive UTC ("YYYY-MM-DD HH:MM"); interpret
  // them as UTC explicitly, then let the browser render them in whatever
  // timezone the viewer is actually in.
  const iso = raw.includes('T') ? raw : raw.replace(' ', 'T') + ':00Z';
  const d = new Date(iso);
  if (isNaN(d.getTime())) return raw;
  return d.toLocaleString(undefined, {
    year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
  });
}

function updateToggleUI(bl, field, checked, by, at) {
  const input = document.getElementById(bl + '_' + field);
  if (!input) { render(); return; }
  const wrap = input.closest('.checkwrap');
  let meta = wrap.querySelector('.meta');
  if (checked) {
    const label = (by || '') + ' - ' + formatLocalTime(at);
    if (!meta) {
      meta = document.createElement('span');
      meta.className = 'meta';
      wrap.appendChild(meta);
    }
    meta.textContent = label;
  } else if (meta) {
    meta.remove();
  }
  updateCompleteBadge(bl);
}

function updateCompleteBadge(bl) {
  const rec = records.find(r => r.bl_number === bl);
  if (!rec) return;
  const row = document.getElementById('row_' + cssEscape(bl));
  if (!row) return;
  const cell = row.querySelector('.bl-cell');
  let badge = cell.querySelector('.badge-complete');
  const complete = !!(rec.invoice_issued && rec.approval_received && rec.do_issued);
  if (complete && !badge) {
    badge = document.createElement('span');
    badge.className = 'badge-complete';
    badge.innerHTML = '&check; Complete';
    cell.appendChild(badge);
  } else if (!complete && badge) {
    badge.remove();
  }
}

function cssEscape(s) {
  return String(s).replace(/[^a-zA-Z0-9_-]/g, c => '_' + c.charCodeAt(0) + '_');
}

function updateSummaryOnly() {
  document.getElementById('summary').innerHTML = summaryHtml();
}

function toggle(bl, field, value) {
  const rec = records.find(r => r.bl_number === bl);
  if (rec) {
    rec[field] = value ? 1 : 0;
    const byField = field.replace('_issued', '_by').replace('_received', '_by');
    const atField = field.replace('_issued', '_at').replace('_received', '_at');
    if (value) {
      rec[byField] = CURRENT_USER;
      rec[atField] = nowLabel();
    } else {
      rec[byField] = '';
      rec[atField] = '';
    }
    updateToggleUI(bl, field, value, rec[byField], rec[atField]);
    updateSummaryOnly();
  }
  suppressPollUntil = Date.now() + 1500;
  fetch(`/api/records/${encodeURIComponent(bl)}/toggle`, {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({field, value})
  }).then(() => fetchRecords()).catch(() => { showToast('Could not save that change - retrying...'); fetchRecords(); });
}

let remarksTimers = {};
function onRemarksInput(bl, value) {
  const rec = records.find(r => r.bl_number === bl);
  if (rec) rec.remarks = value;
  clearTimeout(remarksTimers[bl]);
  remarksTimers[bl] = setTimeout(async () => {
    delete remarksTimers[bl];
    await fetch(`/api/records/${encodeURIComponent(bl)}/remarks`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({remarks: value})
    });
  }, 500);
}

function deleteRecord(bl) {
  const idx = records.findIndex(r => r.bl_number === bl);
  if (idx === -1) return;
  const removed = records[idx];
  records.splice(idx, 1);
  render();
  suppressPollUntil = Date.now() + 4000;
  fetch(`/api/records/${encodeURIComponent(bl)}`, {method: 'DELETE'});

  showToast('Removed BL ' + bl + '.', {
    actionLabel: 'Undo',
    duration: 3000,
    onAction: async () => {
      await fetch('/api/records/restore', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(removed)
      });
      await fetchRecords();
      showToast('Restored BL ' + bl + '.');
    }
  });
}

/* ---------- Bulk remove (vessel group / port group / whole board) ----------
   Same staged-confirm + Undo pattern as the single-row delete above, just
   operating on a whole list of records at once via the bulk API so a
   500-BL manifest doesn't fire 500 individual requests. */
function confirmBulkRemove(label, list) {
  if (!list.length) { showToast('Nothing to remove.'); return; }
  showToast(`Remove all ${list.length} BL${list.length === 1 ? '' : 's'}${label ? ' in ' + label : ''}?`, {
    actionLabel: 'Confirm',
    duration: 6000,
    onAction: () => doBulkRemove(list)
  });
}

async function doBulkRemove(list) {
  const blNumbers = list.map(r => r.bl_number);
  const snapshot = list.map(r => ({...r}));

  records = records.filter(r => !blNumbers.includes(r.bl_number));
  render();
  suppressPollUntil = Date.now() + 5000;

  const res = await fetch('/api/records/bulk-delete', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers})
  });
  const data = await res.json();
  const deleted = (data.deleted && data.deleted.length) ? data.deleted : snapshot;

  showToast(`${deleted.length} BL${deleted.length === 1 ? '' : 's'} removed.`, {
    actionLabel: 'Undo',
    duration: 5000,
    onAction: async () => {
      await fetch('/api/records/bulk-restore', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({records: deleted})
      });
      await fetchRecords();
      showToast('Restored.');
    }
  });
  await fetchRecords();
}

function removeVesselGroup(portName, vesselName) {
  const list = records.filter(r => (r.port || 'Unassigned') === portName && (r.vessel || 'Unassigned') === vesselName);
  confirmBulkRemove(vesselName === 'Unassigned' ? null : vesselName, list);
}

function removePortGroup(portName) {
  const list = records.filter(r => (r.port || 'Unassigned') === portName);
  confirmBulkRemove(portName === 'Unassigned' ? null : portName, list);
}

function clearAllRecords() {
  confirmBulkRemove('the whole board', records.slice());
}

async function renameGroup(type, oldPort, oldVessel, newValue, fallbackLabel) {
  const val = newValue.trim() || fallbackLabel;
  await fetch('/api/groups/rename', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({type, old_port: oldPort, old_vessel: oldVessel, new_value: val === fallbackLabel ? '' : val})
  });
  await fetchRecords();
}

function toggleGroup(key) {
  collapsedGroups[key] = !collapsedGroups[key];
  render();
}

function checkbox(bl, field, checked, by, at) {
  const id = bl + '_' + field;
  return `
    <div class="checkwrap">
      <label class="switch">
        <input type="checkbox" id="${id}" ${checked ? 'checked' : ''}
          onchange="toggle('${bl}', '${field}', this.checked)">
        <span class="slider"></span>
      </label>
      ${checked ? `<span class="meta" title="${by || ''} - ${formatLocalTime(at)}">${by || ''} - ${formatLocalTime(at)}</span>` : ''}
    </div>`;
}

function summaryHtml() {
  const total = records.length;
  const invoicePending = records.filter(r => !r.invoice_issued).length;
  const approvalPending = records.filter(r => !r.approval_received).length;
  const doPending = records.filter(r => !r.do_issued).length;
  const complete = records.filter(r => r.invoice_issued && r.approval_received && r.do_issued).length;

  const icons = {
    total: '<svg viewBox="0 0 24 24"><path d="M7 3h7l5 5v13a1 1 0 01-1 1H7a1 1 0 01-1-1V4a1 1 0 011-1z"/><path d="M14 3v5h5"/></svg>',
    invoice: '<svg viewBox="0 0 24 24"><path d="M6 3h12v18l-2.5-1.5L13 21l-2.5-1.5L8 21l-2-1.5V3z"/><path d="M9 8h6M9 12h6M9 16h4"/></svg>',
    approval: '<svg viewBox="0 0 24 24"><path d="M12 2l8 4v6c0 5-3.5 8.5-8 10-4.5-1.5-8-5-8-10V6l8-4z"/><path d="M9 12l2 2 4-4"/></svg>',
    box: '<svg viewBox="0 0 24 24"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8"/><path d="M12 13v8"/></svg>',
    check: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>'
  };

  const remaining = total - complete;

  return `
    <div class="stat"><div class="stat-icon">${icons.total}</div><div><b>${total}</b>Total BLs</div></div>
    <div class="stat ${remaining ? 'gold' : 'done'}"><div class="stat-icon">${icons.check}</div><div><b>${remaining}</b>Remaining</div></div>
    <div class="stat gold"><div class="stat-icon">${icons.invoice}</div><div><b>${invoicePending}</b>Invoice Pending</div></div>
    <div class="stat gold"><div class="stat-icon">${icons.approval}</div><div><b>${approvalPending}</b>Approval Pending</div></div>
    <div class="stat gold"><div class="stat-icon">${icons.box}</div><div><b>${doPending}</b>DO Pending</div></div>
    <div class="stat done"><div class="stat-icon">${icons.check}</div><div><b>${complete}</b>Fully Complete</div></div>
  `;
}

const CHEVRON = '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>';

function rowsHtml(list) {
  return list.map(r => {
    const complete = !!(r.invoice_issued && r.approval_received && r.do_issued);
    return `
    <tr id="row_${cssEscape(r.bl_number)}">
      <td>
        <div class="bl-cell">
          <b>${r.bl_number}</b>
          ${complete ? '<span class="badge-complete">&check; Complete</span>' : ''}
        </div>
      </td>
      <td>${checkbox(r.bl_number, 'invoice_issued', !!r.invoice_issued, r.invoice_by, r.invoice_at)}</td>
      <td>${checkbox(r.bl_number, 'approval_received', !!r.approval_received, r.approval_by, r.approval_at)}</td>
      <td>${checkbox(r.bl_number, 'do_issued', !!r.do_issued, r.do_by, r.do_at)}</td>
      <td><input class="remarks-input" type="text" value="${(r.remarks || '').replace(/"/g,'&quot;')}"
            oninput="onRemarksInput('${r.bl_number}', this.value)"
            onfocus="markEditing(1)" onblur="markEditing(-1)" placeholder="notes..."></td>
      <td><button class="del" onclick="deleteRecord('${r.bl_number}')">Remove</button></td>
    </tr>`;
  }).join('');
}

function tableHtml(list) {
  return `
    <div class="overflow">
      <table>
        <thead>
          <tr>
            <th>BL Number</th>
            <th>Invoice Issued</th>
            <th>Approval Received</th>
            <th>DO Issued</th>
            <th>Remarks</th>
            <th></th>
          </tr>
        </thead>
        <tbody>${rowsHtml(list)}</tbody>
      </table>
    </div>`;
}

function render() {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const filtered = records.filter(r => r.bl_number.toLowerCase().includes(q));

  // Group by Port, then by Vessel within each port.
  const ports = {};
  filtered.forEach(r => {
    const port = r.port || 'Unassigned';
    const vessel = r.vessel || 'Unassigned';
    if (!ports[port]) ports[port] = {};
    if (!ports[port][vessel]) ports[port][vessel] = [];
    ports[port][vessel].push(r);
  });

  const portNames = Object.keys(ports).sort((a, b) => {
    if (a === 'Unassigned') return 1;
    if (b === 'Unassigned') return -1;
    return naturalCompare(a, b);
  });

  const groupsEl = document.getElementById('groups');
  if (portNames.length === 0) {
    groupsEl.innerHTML = '<div style="color:var(--muted); padding:24px 4px;">No BLs on the board yet. Upload an Excel manifest above to get started.</div>';
    updateSummaryOnly();
    return;
  }

  // Vessel/port jump menu - with 20-25 manifests a month, scrolling down
  // the whole board to find one vessel doesn't scale. Built fresh every
  // render so it always reflects what's actually on the board right now.
  const jumpEl = document.getElementById('jumpSelect');
  if (jumpEl) {
    const current = jumpEl.value;
    let options = '<option value="">Jump to a vessel...</option>';
    portNames.forEach(portName => {
      const vessels = ports[portName];
      Object.keys(vessels).sort((a, b) => {
        if (a === 'Unassigned') return 1;
        if (b === 'Unassigned') return -1;
        return naturalCompare(a, b);
      }).forEach(vesselName => {
        const list = vessels[vesselName];
        const left = list.filter(r => !(r.invoice_issued && r.approval_received && r.do_issued)).length;
        const key = 'vessel:' + portName + ':' + vesselName;
        const label = `${vesselName} - ${portName} (${list.length} BL${list.length === 1 ? '' : 's'}${left ? ', ' + left + ' left' : ', done'})`;
        options += `<option value="${key.replace(/"/g, '&quot;')}">${label}</option>`;
      });
    });
    jumpEl.innerHTML = options;
    if ([...jumpEl.options].some(o => o.value === current)) jumpEl.value = current;
  }

  groupsEl.innerHTML = portNames.map(portName => {
    const portKey = 'port:' + portName;
    const vessels = ports[portName];
    const vesselNames = Object.keys(vessels).sort((a, b) => {
      if (a === 'Unassigned') return 1;
      if (b === 'Unassigned') return -1;
      return naturalCompare(a, b);
    });
    const portTotal = vesselNames.reduce((sum, v) => sum + vessels[v].length, 0);
    const portLeft = vesselNames.reduce((sum, v) => sum + vessels[v].filter(r => !(r.invoice_issued && r.approval_received && r.do_issued)).length, 0);
    // Default state (only applies the first time a group is seen - once a
    // person manually expands/collapses it, collapsedGroups remembers
    // their choice and this default is never forced back on them): a
    // port/vessel group with nothing left to do starts collapsed, so a
    // long board folds down to just the groups that still need work
    // instead of everything piled up needing a scroll through finished
    // ones to reach the next open item.
    const portCollapsed = portKey in collapsedGroups ? !!collapsedGroups[portKey] : (portTotal > 0 && portLeft === 0);

    const vesselsHtml = vesselNames.map(vesselName => {
      const vesselKey = 'vessel:' + portName + ':' + vesselName;
      const list = vessels[vesselName]
        .slice()
        .sort((a, b) => naturalCompare(a.bl_number, b.bl_number));
      const left = list.filter(r => !(r.invoice_issued && r.approval_received && r.do_issued)).length;
      const vesselCollapsed = vesselKey in collapsedGroups ? !!collapsedGroups[vesselKey] : (left === 0);
      return `
        <div class="vessel-group" id="group_${cssEscape(vesselKey)}">
          <div class="vessel-header ${vesselCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT') toggleGroup('${vesselKey.replace(/'/g,"\\'")}')">
            ${CHEVRON}
            <input class="group-name" value="${vesselName === 'Unassigned' ? '' : vesselName}" placeholder="Unassigned vessel"
              onclick="event.stopPropagation()"
              onchange="renameGroup('vessel', '${portName.replace(/'/g,"\\'")}', '${vesselName.replace(/'/g,"\\'")}', this.value, 'Unassigned')">
            <span class="group-count">${list.length} BL${list.length === 1 ? '' : 's'}${left ? ` &middot; ${left} left` : ' &middot; done'}</span>
            <button type="button" class="group-remove" onclick="event.stopPropagation(); removeVesselGroup('${portName.replace(/'/g,"\\'")}', '${vesselName.replace(/'/g,"\\'")}')">Remove all</button>
          </div>
          <div class="vessel-body ${vesselCollapsed ? 'collapsed' : ''}">
            ${tableHtml(list)}
          </div>
        </div>`;
    }).join('');

    return `
      <div class="port-group">
        <div class="port-header ${portCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT') toggleGroup('${portKey.replace(/'/g,"\\'")}')">
          ${CHEVRON}
          <input class="group-name" value="${portName === 'Unassigned' ? '' : portName}" placeholder="Unassigned port"
            onclick="event.stopPropagation()"
            onchange="renameGroup('port', '${portName.replace(/'/g,"\\'")}', '', this.value, 'Unassigned')">
          <span class="group-count">${portTotal} BL${portTotal === 1 ? '' : 's'}${portLeft ? ` &middot; ${portLeft} left` : ' &middot; done'}</span>
          <button type="button" class="group-remove" onclick="event.stopPropagation(); removePortGroup('${portName.replace(/'/g,"\\'")}')">Remove all</button>
        </div>
        <div class="port-body ${portCollapsed ? 'collapsed' : ''}">${vesselsHtml}</div>
      </div>`;
  }).join('');

  updateSummaryOnly();
}

function jumpToVessel(key) {
  if (!key) return;
  const parts = key.split(':');
  const portKey = 'port:' + parts[1];
  collapsedGroups[portKey] = false;
  collapsedGroups[key] = false;
  render();
  requestAnimationFrame(() => {
    const el = document.getElementById('group_' + cssEscape(key));
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

function setAllGroupsCollapsed(collapsed) {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  records.filter(r => r.bl_number.toLowerCase().includes(q)).forEach(r => {
    const port = r.port || 'Unassigned';
    const vessel = r.vessel || 'Unassigned';
    collapsedGroups['port:' + port] = collapsed;
    collapsedGroups['vessel:' + port + ':' + vessel] = collapsed;
  });
  render();
}

refreshFieldSuggestions();
fetchRecords();
setInterval(fetchRecords, 4000);
</script>
</body>
</html>
"""

if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
