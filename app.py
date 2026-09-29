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
        parts = [p.strip() for p in raw.split(",", 1)]
        bl_number = parts[0].upper()
        consignee = _clean_party_name(parts[1]) if len(parts) > 1 else ""
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


# Header names we'll recognize for each column, in the manifest Excel file.
# Matching is case-insensitive and ignores spaces/punctuation.
BL_HEADER_WORDS = ["blnumber", "bl", "billoflading", "billofladingno", "bl no", "blno"]
CONSIGNEE_HEADER_WORDS = ["consignee", "consigneename", "customer", "customername"]
NOTIFY_HEADER_WORDS = ["notifyparty", "notify", "notifypartyname", "notifypartydetails"]


def _normalize_header(text):
    return re.sub(r"[^a-z0-9]", "", str(text or "").lower())


# Words/patterns that mark where an address, phone number, or registration
# detail starts inside a consignee/notify-party cell - everything from the
# earliest of these onward gets cut off, keeping just the company name.
_NAME_STOP_PATTERNS = [
    r"\bADDRESS\b", r"\bADD\s*:", r"\bTEL\b", r"\bFAX\b", r"\bP\.?\s*O\.?\s*BOX\b",
    r"\bC\.?\s*R\.?\s*(NO|NUMBER)?\s*:", r"\bCOMMERCIAL REGISTRATION\b",
    r"\bREGISTRATION NUMBER\b", r"\bVAT\b", r"\bSTREET\b", r"\bDIST\.?\b",
    r"\bKINGDOM OF\b", r"\bKSA\b", r"\bBUILDING\b", r"\bFLOOR\b", r"\bWITH\b",
    r"\d{2,}",  # a run of 2+ digits usually starts a building/street/reg number
]
_BANK_PATTERNS = [r"\bTO\s+(THE\s+)?ORDER\b", r"\bBANK\b"]


def _clean_party_name(text):
    """Take a messy consignee/notify-party cell and return just the company
    name, cutting off address, phone, and registration-number clutter."""
    if not text:
        return ""
    text = str(text)
    # A cell often has the name on its own line, address below - use the
    # first non-empty line as the starting point.
    lines = [l.strip() for l in re.split(r"[\r\n]+", text) if l.strip()]
    if not lines:
        return ""
    candidate = lines[0]

    earliest = len(candidate)
    for pat in _NAME_STOP_PATTERNS:
        m = re.search(pat, candidate, re.IGNORECASE)
        if m and m.start() < earliest:
            earliest = m.start()

    cleaned = candidate[:earliest].strip(" ,.-:;")
    return cleaned if cleaned else candidate.strip()


def _looks_like_bank_or_order(text):
    if not text:
        return False
    return any(re.search(pat, text, re.IGNORECASE) for pat in _BANK_PATTERNS)


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

        # Try to find a header row (in the first 5 rows) naming the BL,
        # Consignee, and Notify Party columns.
        bl_col = None
        consignee_col = None
        notify_col = None
        header_row_index = None

        for i, row in enumerate(rows[:5]):
            for col_index, cell in enumerate(row):
                norm = _normalize_header(cell)
                if norm and any(norm == w.replace(" ", "") or norm.startswith(w.replace(" ", "")) for w in BL_HEADER_WORDS):
                    bl_col = col_index
                    header_row_index = i
                if norm and any(norm == w.replace(" ", "") for w in CONSIGNEE_HEADER_WORDS):
                    consignee_col = col_index
                    header_row_index = i
                if norm and any(norm == w.replace(" ", "") for w in NOTIFY_HEADER_WORDS):
                    notify_col = col_index
                    header_row_index = i
            if bl_col is not None:
                break

        if bl_col is None:
            # No recognizable header found - fall back to assuming column A is BL number,
            # column B is consignee, and there's no header row.
            bl_col = 0
            consignee_col = 1
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

            raw_consignee = ""
            if consignee_col is not None and consignee_col < len(row) and row[consignee_col]:
                raw_consignee = str(row[consignee_col]).strip()
            consignee = _clean_party_name(raw_consignee)

            # If the consignee cell is really a bank / "to order of" clause,
            # the Notify Party is usually the actual receiving company - use
            # that instead when the file has one.
            if _looks_like_bank_or_order(consignee) and notify_col is not None and notify_col < len(row) and row[notify_col]:
                notify_cleaned = _clean_party_name(str(row[notify_col]).strip())
                if notify_cleaned:
                    consignee = notify_cleaned

            existing = db.execute("SELECT 1 FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
            if existing:
                skipped += 1
                continue
            db.execute(
                "INSERT INTO records (bl_number, consignee, created_at) VALUES (?, ?, ?)",
                (bl_number, consignee, datetime.now().strftime("%Y-%m-%d %H:%M")),
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


@app.route("/api/records/<path:bl_number>/consignee", methods=["POST"])
@login_required
def update_consignee(bl_number):
    data = request.get_json(force=True)
    consignee = data.get("consignee", "")
    db = get_db()
    db.execute("UPDATE records SET consignee = ? WHERE bl_number = ?", (consignee, bl_number.upper()))
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
           (bl_number, consignee, invoice_issued, invoice_by, invoice_at,
            approval_received, approval_by, approval_at, do_issued, do_by, do_at,
            remarks, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            bl_number,
            data.get("consignee", ""),
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


# ---------- Templates ----------

AUTH_STYLE = """
<style>
  :root { --bg:#f4f6f8; --card:#fff; --text:#1a2733; --muted:#6b7a89; --border:#dfe6ec; --accent:#1e5f8c; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; background: var(--bg); color: var(--text);
         margin:0; display:flex; align-items:center; justify-content:center; min-height:100vh; padding:20px; }
  .box { background: var(--card); border:1px solid var(--border); border-radius:12px; padding:28px; width:100%; max-width:360px; }
  h1 { font-size:20px; margin:0 0 4px 0; }
  .sub { color:var(--muted); font-size:13px; margin-bottom:20px; }
  label { font-size:13px; font-weight:600; display:block; margin-bottom:4px; margin-top:14px; }
  input { width:100%; padding:10px; border:1px solid var(--border); border-radius:6px; font-size:15px; }
  button { width:100%; background:var(--accent); color:#fff; border:none; border-radius:6px; padding:11px; font-size:14px;
           margin-top:20px; cursor:pointer; }
  .error { background:#fdecec; color:#c0392b; padding:8px 10px; border-radius:6px; font-size:13px; margin-top:14px; }
  select { width:100%; padding:10px; border:1px solid var(--border); border-radius:6px; font-size:15px; margin-top:4px; }
</style>
"""

SETUP_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Set up - Port Agent Ops</title>""" + AUTH_STYLE + """</head><body>
<div class="box">
  <h1>Welcome</h1>
  <div class="sub">This is the first time the app has run. Create your Admin account.</div>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Choose a username</label>
    <input type="text" name="username" required>
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
<div class="box">
  <h1>Port Agent Ops</h1>
  <div class="sub">Sign in to view the DO Tracker</div>
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
  :root { --bg:#f4f6f8; --card:#fff; --text:#1a2733; --muted:#6b7a89; --border:#dfe6ec; --accent:#1e5f8c; }
  * { box-sizing:border-box; }
  body { font-family:-apple-system, Segoe UI, Roboto, Arial, sans-serif; background:var(--bg); color:var(--text); margin:0; padding:20px; }
  .topbar { display:flex; justify-content:space-between; align-items:center; margin-bottom:16px; flex-wrap:wrap; gap:8px; }
  a { color:var(--accent); text-decoration:none; font-size:13px; }
  .card { background:var(--card); border:1px solid var(--border); border-radius:10px; padding:16px; margin-bottom:16px; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th, td { text-align:left; padding:8px 6px; border-bottom:1px solid var(--border); }
  input, select { border:1px solid var(--border); border-radius:6px; padding:8px; font-size:13px; }
  button { background:var(--accent); color:#fff; border:none; border-radius:6px; padding:8px 14px; font-size:13px; cursor:pointer; }
  .del { background:none; color:#c0392b; }
  .row { display:flex; gap:8px; flex-wrap:wrap; }
  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; }
  .toast { background: #1a2733; color: #fff; padding: 10px 14px; border-radius: 8px; font-size: 13px;
           display: flex; align-items: center; gap: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.2);
           animation: toast-in .15s ease-out; max-width: 320px; }
  .toast a { color: #7fc8ff; font-weight: 600; text-decoration: none; cursor: pointer; white-space: nowrap; }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(6px); } }
</style></head><body>
  <div class="topbar">
    <h2 style="margin:0;font-size:18px;">Manage Users</h2>
    <div><a href="/">&larr; Back to board</a> &nbsp;|&nbsp; Signed in as {{ username }} &nbsp;|&nbsp; <a href="/logout">Log out</a></div>
  </div>
  <div class="card">
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
          <td>{{ u['role'] }}</td>
          <td>{{ u['created_at'] }}</td>
          <td><button class="del" onclick="delUser({{ u['id'] }}, {{ u['username']|tojson }})">Remove</button></td>
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
async function addUser() {
  const username = document.getElementById('newUsername').value.trim();
  const password = document.getElementById('newPassword').value;
  const role = document.getElementById('newRole').value;
  if (!username || !password) { showToast('Fill in username and password'); return; }
  const res = await fetch('/api/users', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({username, password, role})});
  const data = await res.json();
  if (data.error) { showToast(data.error); return; }
  location.reload();
}
async function delUser(id, username) {
  await fetch('/api/users/' + id, {method:'DELETE'});
  showToast('Removed user ' + (username || '') + '.');
  setTimeout(() => location.reload(), 600);
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
    --bg: #f4f6f8; --card: #ffffff; --text: #1a2733; --muted: #6b7a89;
    --border: #dfe6ec; --accent: #1e5f8c; --green: #1f9d55; --amber: #c9891a;
  }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 16px; }
  .topbar { display:flex; justify-content:space-between; align-items:center; margin-bottom: 10px; flex-wrap: wrap; gap: 8px; }
  .topbar a { color: var(--accent); text-decoration: none; font-size: 13px; }
  h1 { font-size: 19px; margin: 0; }
  .sub { color: var(--muted); font-size: 13px; margin-bottom: 14px; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 14px; margin-bottom: 14px; }
  .row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
  textarea, input[type=text] { border: 1px solid var(--border); border-radius: 6px; padding: 9px; font-size: 14px; font-family: inherit; width: 100%; }
  textarea { min-height: 70px; resize: vertical; }
  button { background: var(--accent); color: #fff; border: none; border-radius: 6px; padding: 9px 14px; font-size: 13px; cursor: pointer; }
  button:hover { opacity: 0.9; }
  .summary { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; }
  .stat { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 8px 14px; font-size: 12px; flex: 1; min-width: 90px; }
  .stat b { display: block; font-size: 18px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { text-align: left; padding: 8px 6px; border-bottom: 1px solid var(--border); }
  th { color: var(--muted); font-weight: 600; font-size: 11px; text-transform: uppercase; }
  .checkwrap { display: flex; flex-direction: column; gap: 2px; align-items: flex-start; }
  .meta { font-size: 10px; color: var(--muted); }
  .remarks-input, .consignee-input { width: 100%; border: 1px solid transparent; background: transparent; font-size: 12px; font-family: inherit; padding: 3px 4px; border-radius: 4px; }
  .remarks-input:focus, .consignee-input:focus { border-color: var(--border); background: #fff; }
  .del { background: none; color: #c0392b; font-size: 12px; padding: 2px 6px; }
  .overflow { overflow-x: auto; }

  /* Sliding toggle switch */
  .switch { position: relative; display: inline-block; width: 42px; height: 23px; flex-shrink: 0; }
  .switch input { opacity: 0; width: 0; height: 0; }
  .slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0;
            background-color: #dfe6ec; transition: background-color .2s ease; border-radius: 24px; }
  .slider:before { position: absolute; content: ""; height: 17px; width: 17px; left: 3px; bottom: 3px;
                   background-color: #fff; transition: transform .2s ease; border-radius: 50%;
                   box-shadow: 0 1px 3px rgba(0,0,0,0.3); }
  input:checked + .slider { background-color: var(--green); }
  input:checked + .slider:before { transform: translateX(19px); }

  /* Toast notifications (replace confirm()/alert() popups) */
  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; }
  .toast { background: #1a2733; color: #fff; padding: 10px 14px; border-radius: 8px; font-size: 13px;
           display: flex; align-items: center; gap: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.2);
           animation: toast-in .15s ease-out; max-width: 320px; }
  .toast a { color: #7fc8ff; font-weight: 600; text-decoration: none; cursor: pointer; white-space: nowrap; }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(6px); } }

  @media (max-width: 600px) {
    .stat { min-width: 45%; }
  }
</style>
</head>
<body>
  <div class="topbar">
    <h1>Delivery Order Tracker</h1>
    <div>
      {% if role == 'admin' %}<a href="/users">Manage Users</a> &nbsp;|&nbsp; {% endif %}
      {{ username }} &nbsp;|&nbsp; <a href="/logout">Log out</a>
    </div>
  </div>
  <div class="sub">Shared board - visible to everyone with a login. Updates every few seconds.</div>

  <div class="card">
    <div class="row" style="align-items:flex-start;">
      <textarea id="manifestInput" placeholder="Paste manifest: one BL per line, e.g.&#10;MSCU1234567, ABC Trading Co&#10;COSU9876543, Al Fahad Trading"></textarea>
      <button onclick="submitManifest()">Add to Board</button>
    </div>
    <div class="row" style="margin-top:10px; align-items:center;">
      <span style="font-size:12px; color:var(--muted);">Or upload an Excel manifest (.xlsx):</span>
      <input type="file" id="excelFile" accept=".xlsx,.xlsm">
      <button onclick="uploadExcel()" class="secondary" style="background:#eef2f5; color:var(--text);">Upload Excel</button>
    </div>
  </div>

  <div class="summary" id="summary"></div>

  <div class="card">
    <div class="row" style="margin-bottom:12px;">
      <input type="text" id="searchBox" placeholder="Search BL number or consignee..." oninput="render()">
    </div>
    <div class="overflow">
      <table>
        <thead>
          <tr>
            <th>BL Number</th>
            <th>Consignee</th>
            <th>Invoice Issued</th>
            <th>Approval Received</th>
            <th>DO Issued</th>
            <th>Remarks</th>
            <th></th>
          </tr>
        </thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
  </div>

  <div id="toastHost"></div>

<script>
const CURRENT_USER = {{ username|tojson }};
let records = [];
let pollTimer = null;
let suppressPollUntil = 0;

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

async function fetchRecords() {
  if (Date.now() < suppressPollUntil) return;
  const res = await fetch('/api/records');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  records = await res.json();
  render();
}

async function submitManifest() {
  const lines = document.getElementById('manifestInput').value;
  if (!lines.trim()) return;
  const res = await fetch('/api/manifest', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({lines})
  });
  const data = await res.json();
  document.getElementById('manifestInput').value = '';
  await fetchRecords();
  showToast(data.added + ' new BL record(s) added.');
}

async function uploadExcel() {
  const fileInput = document.getElementById('excelFile');
  const file = fileInput.files[0];
  if (!file) { showToast('Choose an Excel file first.'); return; }

  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch('/api/manifest/upload', { method: 'POST', body: formData });
  const data = await res.json();
  if (data.error) { showToast(data.error); return; }

  fileInput.value = '';
  await fetchRecords();
  showToast(data.added + ' new BL record(s) added' + (data.skipped ? `, ${data.skipped} already on the board (skipped)` : '') + '.');
}

function nowLabel() {
  const d = new Date();
  return d.toISOString().slice(0, 16).replace('T', ' ');
}

function toggle(bl, field, value) {
  // Optimistic update: reflect the change instantly, no waiting on the server.
  const rec = records.find(r => r.bl_number === bl);
  if (rec) {
    rec[field] = value ? 1 : 0;
    const byField = field.replace('_issued', '_by').replace('_received', '_by');
    const atField = field.replace('_issued', '_at').replace('_received', '_at');
    if (value) {
      rec[byField] = CURRENT_USER;
      rec[atField] = nowLabel();
    }
    render();
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
    await fetch(`/api/records/${encodeURIComponent(bl)}/remarks`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({remarks: value})
    });
  }, 500);
}

let consigneeTimers = {};
function onConsigneeInput(bl, value) {
  const rec = records.find(r => r.bl_number === bl);
  if (rec) rec.consignee = value;
  clearTimeout(consigneeTimers[bl]);
  consigneeTimers[bl] = setTimeout(async () => {
    await fetch(`/api/records/${encodeURIComponent(bl)}/consignee`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({consignee: value})
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

function render() {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const filtered = records.filter(r =>
    r.bl_number.toLowerCase().includes(q) || (r.consignee || '').toLowerCase().includes(q)
  );

  const tbody = document.getElementById('tbody');
  tbody.innerHTML = filtered.map(r => `
    <tr>
      <td><b>${r.bl_number}</b></td>
      <td><input class="consignee-input" type="text" value="${(r.consignee || '').replace(/"/g,'&quot;')}"
            oninput="onConsigneeInput('${r.bl_number}', this.value)" placeholder="consignee name..."></td>
      <td>${checkbox(r.bl_number, 'invoice_issued', !!r.invoice_issued, r.invoice_by, r.invoice_at)}</td>
      <td>${checkbox(r.bl_number, 'approval_received', !!r.approval_received, r.approval_by, r.approval_at)}</td>
      <td>${checkbox(r.bl_number, 'do_issued', !!r.do_issued, r.do_by, r.do_at)}</td>
      <td><input class="remarks-input" type="text" value="${(r.remarks || '').replace(/"/g,'&quot;')}"
            oninput="onRemarksInput('${r.bl_number}', this.value)" placeholder="notes..."></td>
      <td><button class="del" onclick="deleteRecord('${r.bl_number}')">Remove</button></td>
    </tr>
  `).join('') || '<tr><td colspan="7" style="color:#888;">No BLs on the board yet. Paste a manifest above to get started.</td></tr>';

  const total = records.length;
  const invoicePending = records.filter(r => !r.invoice_issued).length;
  const approvalPending = records.filter(r => !r.approval_received).length;
  const doPending = records.filter(r => !r.do_issued).length;
  const complete = records.filter(r => r.invoice_issued && r.approval_received && r.do_issued).length;

  document.getElementById('summary').innerHTML = `
    <div class="stat"><b>${total}</b>Total BLs</div>
    <div class="stat"><b>${invoicePending}</b>Invoice Pending</div>
    <div class="stat"><b>${approvalPending}</b>Approval Pending</div>
    <div class="stat"><b>${doPending}</b>DO Pending</div>
    <div class="stat"><b>${complete}</b>Fully Complete</div>
  `;
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
