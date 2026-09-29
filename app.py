#!/usr/bin/env python3
"""
Port Agent Ops - DO Tracker (v2, with logins & Neon DB support)
----------------------------------------------------------------
Shared Delivery Order board configured for production on Render + Neon DB.
"""

import os
import re
import time
from datetime import datetime
from functools import wraps
from flask import Flask, request, jsonify, g, render_template_string, session, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
import psycopg2
import psycopg2.extras
from psycopg2.pool import ThreadedConnectionPool
from psycopg2 import sql

# Get Database URL and normalize schema prefix for psycopg2/SQLAlchemy compatibility
DATABASE_URL = os.environ.get("DATABASE_URL", "")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

# Ensure SSL mode is enabled for Neon DB
if DATABASE_URL and "sslmode" not in DATABASE_URL:
    delimiter = "&" if "?" in DATABASE_URL else "?"
    DATABASE_URL += f"{delimiter}sslmode=require"

app = Flask(__name__)

# Security & Session Configuration
app.secret_key = os.environ.get("APP_SECRET_KEY", "change-this-secret-key-later")
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    SESSION_COOKIE_SECURE=os.environ.get("FLASK_ENV") == "production"
)

# Initialize Connection Pool
db_pool = ThreadedConnectionPool(1, 10, dsn=DATABASE_URL) if DATABASE_URL else None


class DBWrapper:
    """Wrapper so the app uses SQLite-style '?' placeholders while executing against Postgres."""

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
        conn = db_pool.getconn()
        # Ping check: verify connection is alive (handles Neon DB auto-suspend wakeups)
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
        except (psycopg2.OperationalError, psycopg2.InterfaceError):
            db_pool.putconn(conn, close=True)
            conn = db_pool.getconn()

        g.db = DBWrapper(conn)
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db_pool.putconn(db.conn)


def init_db():
    if not DATABASE_URL:
        return
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
    cur.execute("CREATE INDEX IF NOT EXISTS idx_records_created_at ON records (created_at DESC);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_records_consignee ON records (consignee);")
    conn.commit()
    cur.close()
    conn.close()


# Auto-initialize database tables on app startup
try:
    init_db()
except Exception as e:
    print(f"Database initialization warning: {e}")


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

LOGIN_ATTEMPTS = {}


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
        ip = request.remote_addr
        now = time.time()

        # Brute force protection: max 5 attempts per minute per IP
        attempts = [t for t in LOGIN_ATTEMPTS.get(ip, []) if now - t < 60]
        if len(attempts) >= 5:
            return render_template_string(LOGIN_HTML, error="Too many failed attempts. Please try again in a minute.")

        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()

        if user and check_password_hash(user["password_hash"], password):
            LOGIN_ATTEMPTS.pop(ip, None)
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("index"))

        attempts.append(now)
        LOGIN_ATTEMPTS[ip] = attempts
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
    rows = db.execute("SELECT * FROM records ORDER BY created_at DESC LIMIT 500").fetchall()
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
        consignee = parts[1] if len(parts) > 1 else ""
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


BL_HEADER_WORDS = ["blnumber", "bl", "billoflading", "billofladingno", "bl no", "blno"]
CONSIGNEE_HEADER_WORDS = ["consignee", "consigneename", "customer", "customername"]


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
        sheet = wb.active
    except Exception:
        return jsonify({"error": "Couldn't read that file. Make sure it's a valid Excel (.xlsx) file."}), 400

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return jsonify({"error": "That file looks empty."}), 400

    bl_col = None
    consignee_col = None
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
        if bl_col is not None:
            break

    if bl_col is None:
        bl_col = 0
        consignee_col = 1
        data_rows = rows
    else:
        data_rows = rows[header_row_index + 1:]

    db = get_db()
    added = 0
    skipped = 0
    for row in data_rows:
        if bl_col >= len(row):
            continue
        raw_bl = row[bl_col]
        if raw_bl is None or str(raw_bl).strip() == "":
            continue
        bl_number = str(raw_bl).strip().upper()
        consignee = ""
        if consignee_col is not None and consignee_col < len(row) and row[consignee_col]:
            consignee = str(row[consignee_col]).strip()

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


@app.route("/api/records/<bl_number>/toggle", methods=["POST"])
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
    query = sql.SQL("UPDATE records SET {} = %s, {} = %s, {} = %s WHERE bl_number = %s").format(
        sql.Identifier(field),
        sql.Identifier(by_field),
        sql.Identifier(at_field)
    )
    
    cur = db.conn.cursor()
    cur.execute(query, (value, by_val, now, bl_number.upper()))
    db.commit()
    cur.close()
    return jsonify({"ok": True})


@app.route("/api/records/<bl_number>/remarks", methods=["POST"])
@login_required
def update_remarks(bl_number):
    data = request.get_json(force=True)
    remarks = data.get("remarks", "")
    db = get_db()
    db.execute("UPDATE records SET remarks = ? WHERE bl_number = ?", (remarks, bl_number.upper()))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<bl_number>", methods=["DELETE"])
@login_required
def delete_record(bl_number):
    db = get_db()
    db.execute("DELETE FROM records WHERE bl_number = ?", (bl_number.upper(),))
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
          <td><button class="del" onclick="delUser({{ u['id'] }})">Remove</button></td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
<script>
async function addUser() {
  const username = document.getElementById('newUsername').value.trim();
  const password = document.getElementById('newPassword').value;
  const role = document.getElementById('newRole').value;
  if (!username || !password) { alert('Fill in username and password'); return; }
  const res = await fetch('/api/users', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({username, password, role})});
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  location.reload();
}
async function delUser(id) {
  if (!confirm('Remove this user?')) return;
  await fetch('/api/users/' + id, {method:'DELETE'});
  location.reload();
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
  .pill { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
  .pill.yes { background: #e3f6ea; color: var(--green); }
  .pill.no { background: #fdecec; color: #c0392b; }
  .checkwrap { display: flex; flex-direction: column; gap: 2px; align-items: flex-start; }
  .meta { font-size: 10px; color: var(--muted); }
  .remarks-input { width: 100%; border: 1px solid transparent; background: transparent; font-size: 12px; }
  .remarks-input:focus { border-color: var(--border); background: #fff; }
  .del { background: none; color: #c0392b; font-size: 12px; padding: 2px 6px; }
  .overflow { overflow-x: auto; }
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

<script>
let records = [];

async function fetchRecords() {
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
  alert(data.added + ' new BL record(s) added.');
}

async function uploadExcel() {
  const fileInput = document.getElementById('excelFile');
  const file = fileInput.files[0];
  if (!file) { alert('Choose an Excel file first.'); return; }

  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch('/api/manifest/upload', { method: 'POST', body: formData });
  const data = await res.json();
  if (data.error) { alert(data.error); return; }

  fileInput.value = '';
  await fetchRecords();
  alert(data.added + ' new BL record(s) added' + (data.skipped ? `, ${data.skipped} already on the board (skipped)` : '') + '.');
}

async function toggle(bl, field, value) {
  await fetch(`/api/records/${bl}/toggle`, {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({field, value})
  });
  await fetchRecords();
}

let remarksTimers = {};
function onRemarksInput(bl, value) {
  clearTimeout(remarksTimers[bl]);
  remarksTimers[bl] = setTimeout(async () => {
    await fetch(`/api/records/${bl}/remarks`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({remarks: value})
    });
  }, 500);
}

async function deleteRecord(bl) {
  if (!confirm('Remove BL ' + bl + ' from the board?')) return;
  await fetch(`/api/records/${bl}`, {method: 'DELETE'});
  await fetchRecords();
}

function checkbox(bl, field, checked, by, at) {
  const id = bl + '_' + field;
  return `
    <div class="checkwrap">
      <label>
        <input type="checkbox" id="${id}" ${checked ? 'checked' : ''}
          onchange="toggle('${bl}', '${field}', this.checked)">
        <span class="pill ${checked ? 'yes' : 'no'}">${checked ? 'Yes' : 'No'}</span>
      </label>
      ${checked ? `<span class="meta">${by} -${at}</span>` : ''}
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
      <td>${r.consignee || '-'}</td>
      <td>${checkbox(r.bl_number, 'invoice_issued', !!r.invoice_issued, r.invoice_by, r.invoice_at)}</td>
      <td>${checkbox(r.bl_number, 'approval_received', !!r.approval_received, r.approval_by, r.approval_at)}</td>
      <td>${checkbox(r.bl_number, 'do_issued', !!r.do_issued, r.do_by, r.do_at)}</td>
      <td><input class="remarks-input" type="text" value="${(r.remarks || '').replace(/"/g,'&quot;')}"
            onchange="onRemarksInput('${r.bl_number}', this.value)" placeholder="notes..."></td>
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
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
