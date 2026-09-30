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
                "INSERT INTO records (bl_number, created_at) VALUES (?, ?)",
                (bl_number, datetime.now().strftime("%Y-%m-%d %H:%M")),
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
    padding: 14px 16px; margin: 0 -16px 20px;
    background: color-mix(in srgb, var(--bg) 86%, transparent);
    backdrop-filter: saturate(180%) blur(14px);
    -webkit-backdrop-filter: saturate(180%) blur(14px);
    border-bottom: 1px solid var(--border);
  }
  .brand { display: flex; align-items: center; gap: 10px; }
  .brand img { height: 34px; width: auto; display: block; border-radius: 8px; box-shadow: var(--shadow-sm); }
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
  .row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
  :root[data-theme="dark"] .card-label { color: var(--navy-light); }

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
      <img src="data:image/jpeg;base64,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" alt="Sea Power">
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

  <div class="sub">Shared board, visible to everyone with a login. Updates automatically.</div>

  <div class="card">
    <div class="card-label">Add BL numbers</div>
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
        <tbody id="tbody"></tbody>
      </table>
    </div>
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
  // Don't let a poll stomp remarks the user is still typing (unsaved edits).
  fresh.forEach(nr => {
    if (remarksTimers[nr.bl_number]) {
      const old = records.find(r => r.bl_number === nr.bl_number);
      if (old) nr.remarks = old.remarks;
    }
  });
  // Skip the rebuild entirely if nothing actually changed - this is what
  // was causing the periodic flicker (the whole table used to redraw
  // every 4 seconds even when nothing was different).
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

  const formData = new FormData();
  formData.append('file', file);

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

/* Surgical DOM update for a single toggle - avoids rebuilding the whole
   table (which used to interrupt the slide animation and cause the
   flicker/glitch the switches had). */
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
  const row = document.getElementById('row_' + bl);
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

function updateSummaryOnly() {
  document.getElementById('summary').innerHTML = summaryHtml();
}

function toggle(bl, field, value) {
  // Optimistic update: reflect the change instantly, no waiting on the server,
  // and without rebuilding the whole table (that's what used to cause the
  // slider to glitch/flicker instead of sliding smoothly).
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

function render() {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const filtered = records
    .filter(r => r.bl_number.toLowerCase().includes(q))
    .sort((a, b) => naturalCompare(a.bl_number, b.bl_number));

  const tbody = document.getElementById('tbody');
  tbody.innerHTML = filtered.map(r => {
    const complete = !!(r.invoice_issued && r.approval_received && r.do_issued);
    return `
    <tr id="row_${r.bl_number}">
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
    </tr>
  `;
  }).join('') || '<tr><td colspan="6" style="color:var(--muted); padding:24px 10px;">No BLs on the board yet. Upload an Excel manifest above to get started.</td></tr>';

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
