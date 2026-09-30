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
from flask import Flask, request, jsonify, g, render_template_string, session, redirect, url_for
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
        candidate = str(raw_bl).strip().upper()
        # Skip a trailing "TOTAL:" / "GRAND TOTAL" summary row - manifests
        # commonly have one at the bottom of the same column as the BL
        # numbers, and it isn't a real BL.
        norm_candidate = _normalize_header(candidate)
        if norm_candidate in ("total", "totals", "grandtotal"):
            continue
        out.append(candidate)

    if header_row_index is None:
        sample = out[:5]
        if sample == [str(n) for n in range(1, len(sample) + 1)]:
            return []

    return out


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
        if filename.endswith((".xlsx", ".xlsm")):
            wb = openpyxl.load_workbook(file, data_only=True)
            # Go through every sheet (not just the first/active one) so BLs
            # aren't missed if the file has multiple tabs or was last saved
            # on a different sheet. Only the FIRST sheet gets the "no
            # header -> assume column A" fallback: a real-world manifest is
            # commonly one main BL-list sheet plus per-BL "attachment"
            # sheets (e.g. a vehicle's chassis/engine-number list) that
            # have their own ITEM/serial column but no BL data at all -
            # falling back on those would read "1, 2, 3, 4..." as BL
            # numbers.
            for i, sheet in enumerate(wb.worksheets):
                rows = list(sheet.iter_rows(values_only=True))
                bl_numbers.extend(_extract_bl_numbers_from_rows(rows, allow_no_header_fallback=(i == 0)))

        elif filename.endswith(".xls"):
            book = xlrd.open_workbook(file_contents=file.read())
            for i, sheet in enumerate(book.sheets()):
                rows = [sheet.row_values(r) for r in range(sheet.nrows)]
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
    animation: mark-in .7s .1s cubic-bezier(.34,1.56,.64,1) both;
  }
  @keyframes mark-in {
    from { opacity: 0; transform: scale(.6) rotate(-16deg); }
    to { opacity: 1; transform: scale(1) rotate(0); }
  }
  .brand-mark .co { font-size: 11px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: var(--gold); }
  .brand-mark .tag { font-size: 11.5px; color: var(--muted); margin-top: 2px; }

  h1 { font-size: 19px; margin: 0 0 4px; text-align: center; letter-spacing: -0.01em; }
  .sub { color: var(--muted); font-size: 13px; margin-bottom: 22px; text-align: center; }

  label { font-size: 12px; font-weight: 600; display: block; margin-bottom: 6px; margin-top: 16px; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; }
  input, select {
    width: 100%; padding: 11px 13px; border: 1px solid var(--border); border-radius: 10px;
    font-size: 14.5px; font-family: inherit; background: var(--bg); color: var(--text);
    transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
  }
  input:focus, select:focus {
    outline: none; border-color: var(--navy-light); background: var(--card);
    box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 20%, transparent);
  }
  button {
    width: 100%; background: var(--navy); color: #fff; border: none; border-radius: 999px;
    padding: 13px; font-size: 14px; font-weight: 700; margin-top: 24px; cursor: pointer;
    transition: background .15s ease, transform .08s ease;
  }
  button:hover { background: var(--navy-light); }
  button:active { transform: scale(.98); }
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

  /* Custom cursor - a small dot with a soft trailing glow, on top of the
     drifting blobs. Only on real mouse/trackpad input (hover:hover and
     pointer:fine) - touch and coarse-pointer devices keep the native
     cursor untouched and never see this. */
  @media (hover: hover) and (pointer: fine) {
    body, a, button, input, .theme-switch, .pw-toggle { cursor: none; }
    .cursor-glow {
      position: fixed; left: 0; top: 0; width: 30px; height: 30px; border-radius: 50%;
      background: radial-gradient(circle, var(--gold) 0%, transparent 72%);
      transform: translate(-50%, -50%); pointer-events: none; z-index: 9999;
      opacity: 0; transition: opacity .25s ease, width .2s ease, height .2s ease;
    }
    .cursor-glow.active { opacity: .5; }
    .cursor-glow.hovering { width: 52px; height: 52px; opacity: .65; }
    .cursor-glow.pressed { width: 20px; height: 20px; opacity: .85; }
    .cursor-dot {
      position: fixed; left: 0; top: 0; width: 6px; height: 6px; border-radius: 50%;
      background: var(--gold-light); transform: translate(-50%, -50%); pointer-events: none;
      z-index: 10000; opacity: 0; transition: opacity .25s ease;
      box-shadow: 0 0 8px color-mix(in srgb, var(--gold-light) 70%, transparent);
    }
    .cursor-dot.active { opacity: 1; }
  }
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

/* Custom trailing cursor - a small dot glued to the real pointer, plus a
   soft glow that eases toward it a beat behind. Desktop-only (see the
   hover:hover/pointer:fine CSS gate above); this script itself also bails
   out early on touch/coarse-pointer devices so it costs nothing there. */
(function initCursor() {
  if (!window.matchMedia || !window.matchMedia('(hover: hover) and (pointer: fine)').matches) return;
  window.addEventListener('DOMContentLoaded', () => {
    const glow = document.createElement('div');
    glow.className = 'cursor-glow';
    const dot = document.createElement('div');
    dot.className = 'cursor-dot';
    document.body.appendChild(glow);
    document.body.appendChild(dot);

    let mx = window.innerWidth / 2, my = window.innerHeight / 2;
    let gx = mx, gy = my;
    let active = false;

    document.addEventListener('mousemove', (e) => {
      mx = e.clientX; my = e.clientY;
      dot.style.left = mx + 'px'; dot.style.top = my + 'px';
      if (!active) { active = true; glow.classList.add('active'); dot.classList.add('active'); }
    });
    document.addEventListener('mouseleave', () => {
      active = false; glow.classList.remove('active'); dot.classList.remove('active');
    });
    document.addEventListener('mousedown', () => glow.classList.add('pressed'));
    document.addEventListener('mouseup', () => glow.classList.remove('pressed'));
    document.addEventListener('mouseover', (e) => {
      if (e.target.closest('a, button, input, select, .theme-switch')) glow.classList.add('hovering');
    });
    document.addEventListener('mouseout', (e) => {
      if (e.target.closest('a, button, input, select, .theme-switch')) glow.classList.remove('hovering');
    });

    (function raf() {
      gx += (mx - gx) * 0.16;
      gy += (my - gy) * 0.16;
      glow.style.left = gx + 'px'; glow.style.top = gy + 'px';
      requestAnimationFrame(raf);
    })();
  });
})();
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

    <div class="tile soon">
      <span class="soon-pill">Coming soon</span>
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2v20M2 12h20"/></svg>
      </div>
      <h3>More workspaces</h3>
      <p>New features get added here as tiles, right alongside these.</p>
    </div>
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

  <div class="sub">{% if role == 'admin' %}Admin view - every staff member's records, all in one place. Organized by Port &rarr; Vessel. Updates automatically.{% else %}Your own board - only records you've added. Organized by Port &rarr; Vessel. Updates automatically.{% endif %}</div>

  <div class="card">
    <div class="card-label">Add a manifest</div>
    <div class="tag-fields">
      <div>
        <label for="portField">Port</label>
        <input type="text" id="portField" placeholder="e.g. JEDDAH PORT" style="text-transform:uppercase;" oninput="this.value = this.value.toUpperCase();">
      </div>
      <div>
        <label for="vesselField">Vessel</label>
        <input type="text" id="vesselField" placeholder="e.g. TAI KNIGHT" style="text-transform:uppercase;" oninput="this.value = this.value.toUpperCase();">
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
    <div class="row" style="margin-bottom:14px;">
      <input type="text" id="searchBox" placeholder="Search BL number..." oninput="render()" style="flex:1; min-width:180px;">
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
  if (editingCount === 0 && changed) render();
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

  return `
    <div class="stat"><div class="stat-icon">${icons.total}</div><div><b>${total}</b>Total BLs</div></div>
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

  groupsEl.innerHTML = portNames.map(portName => {
    const portKey = 'port:' + portName;
    const portCollapsed = !!collapsedGroups[portKey];
    const vessels = ports[portName];
    const vesselNames = Object.keys(vessels).sort((a, b) => {
      if (a === 'Unassigned') return 1;
      if (b === 'Unassigned') return -1;
      return naturalCompare(a, b);
    });
    const portTotal = vesselNames.reduce((sum, v) => sum + vessels[v].length, 0);

    const vesselsHtml = vesselNames.map(vesselName => {
      const vesselKey = 'vessel:' + portName + ':' + vesselName;
      const vesselCollapsed = !!collapsedGroups[vesselKey];
      const list = vessels[vesselName]
        .slice()
        .sort((a, b) => naturalCompare(a.bl_number, b.bl_number));
      return `
        <div class="vessel-group">
          <div class="vessel-header ${vesselCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT') toggleGroup('${vesselKey.replace(/'/g,"\\'")}')">
            ${CHEVRON}
            <input class="group-name" value="${vesselName === 'Unassigned' ? '' : vesselName}" placeholder="Unassigned vessel"
              onclick="event.stopPropagation()"
              onchange="renameGroup('vessel', '${portName.replace(/'/g,"\\'")}', '${vesselName.replace(/'/g,"\\'")}', this.value, 'Unassigned')">
            <span class="group-count">${list.length} BL${list.length === 1 ? '' : 's'}</span>
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
          <span class="group-count">${portTotal} BL${portTotal === 1 ? '' : 's'}</span>
          <button type="button" class="group-remove" onclick="event.stopPropagation(); removePortGroup('${portName.replace(/'/g,"\\'")}')">Remove all</button>
        </div>
        <div class="port-body ${portCollapsed ? 'collapsed' : ''}">${vesselsHtml}</div>
      </div>`;
  }).join('');

  updateSummaryOnly();
}

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
