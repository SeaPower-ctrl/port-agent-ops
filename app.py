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
from datetime import datetime
from functools import wraps
from flask import Flask, request, jsonify, g, render_template_string, session, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
import psycopg2
import psycopg2.extras

DATABASE_URL = os.environ.get("DATABASE_URL", "")

app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET_KEY", "change-this-secret-key-later")


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
                (username, generate_password_hash(password), datetime.now().strftime("%Y-%m-%d %H:%M")),
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
    return render_template_string(PAGE_HTML, username=session.get("username"), role=session.get("role"))


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
            (username, generate_password_hash(password), role, datetime.now().strftime("%Y-%m-%d %H:%M")),
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
    rows = db.execute("SELECT * FROM records ORDER BY created_at DESC").fetchall()
    return jsonify([dict(r) for r in rows])


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
            "INSERT INTO records (bl_number, consignee, created_at) VALUES (?, ?, ?)",
            (bl_number, consignee, datetime.now().strftime("%Y-%m-%d %H:%M")),
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


@app.route("/api/manifest/upload", methods=["POST"])
@login_required
def upload_manifest_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file received"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "Please upload an .xlsx Excel file (not .xls or .csv)"}), 400

    # The whole manifest gets tagged with the Port and Vessel it was
    # uploaded for, so the board can be organized Port > Vessel.
    port = request.form.get("port", "").strip()
    vessel = request.form.get("vessel", "").strip()

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
    except Exception:
        return jsonify({"error": "Couldn't read that file. Make sure it's a valid Excel (.xlsx) file."}), 400

    db = get_db()
    added = 0
    skipped = 0

    # Go through every sheet in the workbook (not just the first/active one)
    # so BLs aren't missed if the file has multiple tabs or was last saved
    # on a different sheet.
    for sheet in wb.worksheets:
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            continue

        # Try to find a header row (in the first 5 rows) naming the BL
        # Number column. Only the BL number is tracked - consignee is not
        # collected or stored.
        bl_col = None
        header_row_index = None

        for i, row in enumerate(rows[:5]):
            for col_index, cell in enumerate(row):
                norm = _normalize_header(cell)
                if norm and any(norm == w.replace(" ", "") or norm.startswith(w.replace(" ", "")) for w in BL_HEADER_WORDS):
                    bl_col = col_index
                    header_row_index = i
            if bl_col is not None:
                break

        if bl_col is None:
            # No recognizable header found - fall back to assuming column A is BL number,
            # with no header row.
            bl_col = 0
            data_rows = rows
        else:
            data_rows = rows[header_row_index + 1:]

        for row in data_rows:
            if bl_col >= len(row):
                continue
            raw_bl = row[bl_col]
            if raw_bl is None or str(raw_bl).strip() == "":
                continue
            bl_number = str(raw_bl).strip().upper()

            existing = db.execute("SELECT 1 FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
            if existing:
                skipped += 1
                continue
            db.execute(
                "INSERT INTO records (bl_number, port, vessel, created_at) VALUES (?, ?, ?, ?)",
                (bl_number, port, vessel, datetime.now().strftime("%Y-%m-%d %H:%M")),
            )
            added += 1

    db.commit()
    return jsonify({"added": added, "skipped": skipped})


@app.route("/api/records/<path:bl_number>/toggle", methods=["POST"])
@login_required
def toggle_status(bl_number):
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
    now = datetime.now().strftime("%Y-%m-%d %H:%M") if value else ""
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
    data = request.get_json(force=True)
    remarks = data.get("remarks", "")
    db = get_db()
    db.execute("UPDATE records SET remarks = ? WHERE bl_number = ?", (remarks, bl_number.upper()))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<path:bl_number>", methods=["DELETE"])
@login_required
def delete_record(bl_number):
    db = get_db()
    db.execute("DELETE FROM records WHERE bl_number = ?", (bl_number.upper(),))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/restore", methods=["POST"])
@login_required
def restore_record():
    """Used by the 'Undo' notice after a delete - re-inserts a record with
    all its original fields, rather than a bare fresh row."""
    data = request.get_json(force=True)
    bl_number = str(data.get("bl_number", "")).strip().upper()
    if not bl_number:
        return jsonify({"error": "missing bl_number"}), 400

    db = get_db()
    existing = db.execute("SELECT 1 FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
    if existing:
        return jsonify({"ok": True, "note": "already exists"})

    db.execute(
        """INSERT INTO records
           (bl_number, consignee, port, vessel, invoice_issued, invoice_by, invoice_at,
            approval_received, approval_by, approval_at, do_issued, do_by, do_at,
            remarks, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
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
        ),
    )
    db.commit()
    return jsonify({"ok": True})


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
    if group_type == "port":
        db.execute("UPDATE records SET port = ? WHERE port = ?", (new_value, old_port))
    elif group_type == "vessel":
        old_vessel = data.get("old_vessel", "")
        db.execute(
            "UPDATE records SET vessel = ? WHERE port = ? AND vessel = ?",
            (new_value, old_port, old_vessel),
        )
    else:
        return jsonify({"error": "invalid type"}), 400
    db.commit()
    return jsonify({"ok": True})


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
</style>
<script>
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || ((window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light');
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

SETUP_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Set up - Port Agent Ops</title>""" + AUTH_STYLE + """</head><body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
""" + THEME_TOGGLE_SNIPPET + """
<div class="box">
  <div class="brand-mark">
    <img src="data:image/png;base64,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" alt="Sea Power">
    <span class="co">Sea Power</span>
    <span class="tag">Port Agent Ops</span>
  </div>
  <h1>Welcome aboard</h1>
  <div class="sub">First time here - create the Admin account to get started.</div>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Choose a username</label>
    <input type="text" name="username" required autofocus>
    <label>Choose a password</label>
    <input type="password" name="password" required>
    <button type="submit">Create Admin Account</button>
  </form>
</div>
</body></html>
"""

LOGIN_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Login - Port Agent Ops</title>""" + AUTH_STYLE + """</head><body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
""" + THEME_TOGGLE_SNIPPET + """
<div class="box">
  <div class="brand-mark">
    <img src="data:image/png;base64,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" alt="Sea Power">
    <span class="co">Sea Power</span>
    <span class="tag">Port Agent Ops</span>
  </div>
  <h1>Delivery Order Tracker</h1>
  <div class="sub">Sign in to view the shared board.</div>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Username</label>
    <input type="text" name="username" required autofocus>
    <label>Password</label>
    <input type="password" name="password" required>
    <button type="submit">Sign In</button>
  </form>
</div>
</body></html>
"""

USERS_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Manage Users</title>
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
  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; }
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
  const mode = saved || ((window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light');
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
    <div class="brand">
      <img src="data:image/png;base64,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" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Manage Users</span>
        <span class="app-tag">Sea Power &middot; Port Agent Ops</span>
      </div>
    </div>
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
          <td>{{ u['created_at'] }}</td>
          <td><button class="del" onclick='delUser({{ u["id"] }}, {{ u["username"]|tojson }})'>Remove</button></td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
  <div id="toastHost"></div>
<script>
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

PAGE_HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Port Agent Ops - DO Tracker</title>
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
  table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
  th, td { text-align: left; padding: 12px 10px; border-bottom: 1px solid var(--border); }
  th {
    color: var(--muted); font-weight: 600; font-size: 10.5px; text-transform: uppercase;
    letter-spacing: .05em; background: color-mix(in srgb, var(--border) 40%, transparent);
  }
  tbody tr { transition: background .12s ease; }
  tbody tr:hover { background: color-mix(in srgb, var(--navy-light) 4%, transparent); }
  tbody tr:last-child td { border-bottom: none; }

  .bl-cell { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .bl-cell b { font-weight: 700; letter-spacing: -0.01em; }
  .badge-complete {
    display: inline-flex; align-items: center; gap: 3px;
    background: var(--success-bg); color: var(--success); font-size: 10.5px; font-weight: 700;
    padding: 2px 8px; border-radius: 999px; text-transform: uppercase; letter-spacing: .03em;
  }

  .checkwrap { display: flex; flex-direction: column; gap: 3px; align-items: flex-start; min-height: 34px; justify-content: center; }
  .meta { font-size: 10px; color: var(--muted); }
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
  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; }
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
    <div class="brand">
      <img src="data:image/png;base64,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" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Delivery Order Tracker</span>
        <span class="app-tag">Sea Power &middot; Port Agent Ops</span>
      </div>
    </div>
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

  <div class="sub">Shared board, visible to everyone with a login. Organized by Port &rarr; Vessel. Updates automatically.</div>

  <div class="card">
    <div class="card-label">Add a manifest</div>
    <div class="tag-fields">
      <div>
        <label for="portField">Port</label>
        <input type="text" id="portField" placeholder="e.g. Jeddah Port">
      </div>
      <div>
        <label for="vesselField">Vessel</label>
        <input type="text" id="vesselField" placeholder="e.g. TAI KNIGHT">
      </div>
    </div>
    <label class="dropzone" id="dropzone" for="excelFile">
      <div class="dropzone-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
          <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/>
        </svg>
      </div>
      <div>
        <div class="dropzone-text"><b>Click to upload</b> or drag &amp; drop your Excel manifest</div>
        <div class="dropzone-sub">.xlsx or .xlsm - the BL Number column is read automatically</div>
        <div class="dropzone-filename" id="dropzoneFilename"></div>
      </div>
      <input type="file" id="excelFile" accept=".xlsx,.xlsm" style="display:none" onchange="uploadExcel()">
    </label>
  </div>

  <div class="summary" id="summary"></div>

  <div class="card">
    <div class="row" style="margin-bottom:14px;">
      <input type="text" id="searchBox" placeholder="Search BL number..." oninput="render()">
    </div>
    <div id="groups"></div>
  </div>

  <div id="toastHost"></div>

<script>
/* ---------- Theme (light/dark, sun/moon toggle) ---------- */
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem('theme'); } catch (e) {}
  const mode = saved || ((window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light');
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

function showToast(message, opts) {
  opts = opts || {};
  const host = document.getElementById('toastHost');
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

/* ---------- Excel upload (drag & drop) ---------- */
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
  document.getElementById('excelFile').files = e.dataTransfer.files;
  uploadExcel();
});

async function uploadExcel() {
  const fileInput = document.getElementById('excelFile');
  const file = fileInput.files[0];
  if (!file) { showToast('Choose an Excel file first.'); return; }
  document.getElementById('dropzoneFilename').textContent = file.name;

  const port = document.getElementById('portField').value.trim();
  const vessel = document.getElementById('vesselField').value.trim();

  const formData = new FormData();
  formData.append('file', file);
  formData.append('port', port);
  formData.append('vessel', vessel);

  const res = await fetch('/api/manifest/upload', { method: 'POST', body: formData });
  const data = await res.json();
  if (data.error) { showToast(data.error); return; }

  await fetchRecords();
  showToast(data.added + ' new BL record(s) added' + (data.skipped ? `, ${data.skipped} already on the board (skipped)` : '') + '.');
}

function nowLabel() {
  const d = new Date();
  return d.toISOString().slice(0, 16).replace('T', ' ');
}

function updateToggleUI(bl, field, checked, by, at) {
  const input = document.getElementById(bl + '_' + field);
  if (!input) { render(); return; }
  const wrap = input.closest('.checkwrap');
  let meta = wrap.querySelector('.meta');
  if (checked) {
    const label = (by || '') + ' - ' + (at || '');
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
    duration: 5000,
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
      ${checked ? `<span class="meta">${by || ''} - ${at || ''}</span>` : ''}
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
