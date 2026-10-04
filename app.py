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
import time
import secrets
import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, date
from functools import wraps
from flask import Flask, request, jsonify, g, render_template_string, session, redirect, url_for, send_file, Response
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
import xlrd
import psycopg2
import psycopg2.extras
from fpdf import FPDF
from fpdf.enums import XPos, YPos

DATABASE_URL = os.environ.get("DATABASE_URL", "")

app = Flask(__name__)
# A hardcoded fallback secret here would be a real vulnerability, not a
# theoretical one: Flask signs session cookies with this key, so anyone who
# read this source (and it has been pasted/shared plenty) could forge a
# valid session for any username/role, admin included, on any deployment
# that forgot to set APP_SECRET_KEY. Generating a random key instead means
# a missing env var only costs everyone a re-login after a restart - a
# safe failure instead of a silent one - and it's logged loudly so a
# misconfigured deployment doesn't go unnoticed.
_configured_secret = os.environ.get("APP_SECRET_KEY", "")
if _configured_secret:
    app.secret_key = _configured_secret
else:
    app.secret_key = secrets.token_hex(32)
    print(
        "WARNING: APP_SECRET_KEY is not set in the environment. Using a "
        "random key generated for this process only - every login session "
        "will be invalidated on the next restart/redeploy, and this "
        "deployment should not be relied on for real use until "
        "APP_SECRET_KEY is set to a long random value (e.g. in Render's "
        "Environment tab)."
    )

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


# Field-for-field match to the company's actual Statement of Facts Word
# template (vessel/voyage/port through to the master/agent sign-off),
# shared between the DB schema, the API and the PDF builder so the three
# never drift out of sync with each other.
SOF_COLUMNS = [
    "vessel", "voyage", "port", "berth", "owners", "charterer",
    "end_of_sea_passage", "customs_cleared",
    "nor_tendered", "commenced_discharge",
    "nor_accepted", "completed_discharge",
    "anchored", "documents_on_board",
    "left_anchorage", "clearance_delivered",
    "pilot_boarded_arrival", "pilot_boarded_departure",
    "first_line_to_shore", "left_berth",
    "berthed_all_fast",
    "cargo_discharge_mtons",
    "rob_arrival_ifo", "rob_arrival_mdo", "rob_arrival_lubs", "rob_arrival_fwater",
    "rob_departure_ifo", "rob_departure_mdo", "rob_departure_lubs", "rob_departure_fwater",
    "arrival_draft_fwd", "arrival_draft_aft", "departure_draft_fwd", "departure_draft_aft",
    "delays_remarks", "masters_remarks",
]

# (left_column, left_label, right_column, right_label) - the paired
# two-column timeline exactly as laid out in the source template.
SOF_TIMELINE_PAIRS = [
    ("end_of_sea_passage", "End of Sea Passage", "customs_cleared", "Customs Cleared"),
    ("nor_tendered", "NOR Tendered", "commenced_discharge", "Commenced Discharge"),
    ("nor_accepted", "NOR Accepted", "completed_discharge", "Completed Discharge"),
    ("anchored", "Anchored", "documents_on_board", "Documents on Board"),
    ("left_anchorage", "Left Anchorage", "clearance_delivered", "Clearance Delivered"),
    ("pilot_boarded_arrival", "Pilot Boarded (Arrival)", "pilot_boarded_departure", "Pilot Boarded (Departure)"),
    ("first_line_to_shore", "First Line to Shore", "left_berth", "Left Berth"),
]

# Invoice / DO file attachments on a DO Tracker record - kind -> display
# label, shared between the upload/download routes and the UI.
ATTACHMENT_KINDS = {"invoice": "Invoice", "do": "Delivery Order"}
MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024  # 10MB - comfortably more than a scanned invoice PDF needs


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
    # Set per-vessel (same value on every BL in that vessel's group), not
    # per-BL - lets the board sort "arriving soonest first" instead of
    # alphabetically, and surfaces an ETA without a separate vessels table.
    cur.execute("ALTER TABLE records ADD COLUMN IF NOT EXISTS eta TEXT DEFAULT ''")
    # A vessel group that's fully complete and old can be archived off the
    # main board (manually, from the UI) without deleting its data - it's
    # still searchable/exportable, just out of the day-to-day view.
    cur.execute("ALTER TABLE records ADD COLUMN IF NOT EXISTS archived INTEGER DEFAULT 0")
    cur.execute(
        """CREATE TABLE IF NOT EXISTS audit_log (
            id SERIAL PRIMARY KEY,
            bl_number TEXT NOT NULL,
            action TEXT NOT NULL,
            field TEXT DEFAULT '',
            old_value TEXT DEFAULT '',
            new_value TEXT DEFAULT '',
            by_user TEXT DEFAULT '',
            at TEXT DEFAULT ''
        )"""
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_audit_log_bl ON audit_log (bl_number)")
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

    # Invoice / Delivery Order file attachments - one PDF per BL per kind
    # ('invoice' or 'do'), stored as bytea rather than on disk because
    # Render's app filesystem isn't persistent across deploys/restarts, but
    # Postgres already is. ON CONFLICT (bl_number, kind) lets a re-upload
    # simply replace the previous file (correcting a mistake) instead of
    # piling up duplicates. CASCADE so deleting a BL cleans up its files too.
    cur.execute(
        """CREATE TABLE IF NOT EXISTS record_attachments (
            id SERIAL PRIMARY KEY,
            bl_number TEXT NOT NULL REFERENCES records(bl_number) ON DELETE CASCADE,
            kind TEXT NOT NULL,
            filename TEXT DEFAULT '',
            content_type TEXT DEFAULT 'application/pdf',
            data BYTEA NOT NULL,
            file_size INTEGER DEFAULT 0,
            uploaded_by TEXT DEFAULT '',
            uploaded_at TEXT DEFAULT '',
            UNIQUE (bl_number, kind)
        )"""
    )

    # PDA / FDA (Proforma / Final Disbursement Account) - a per-port charge
    # template (port dues, pilotage, towage, agency fee, ...) that pre-fills
    # a new PDA for a vessel call; the agent adjusts amounts per vessel.
    # When the vessel sails, the same document is "finalized" into an FDA by
    # filling in actual amounts alongside the original estimate - one record
    # carries both, rather than two documents that can drift apart.
    cur.execute(
        """CREATE TABLE IF NOT EXISTS pda_templates (
            id SERIAL PRIMARY KEY,
            port TEXT NOT NULL,
            name TEXT NOT NULL,
            default_amount NUMERIC DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT ''
        )"""
    )
    cur.execute(
        """CREATE TABLE IF NOT EXISTS pda_documents (
            id SERIAL PRIMARY KEY,
            port TEXT DEFAULT '',
            vessel TEXT DEFAULT '',
            reference TEXT DEFAULT '',
            currency TEXT DEFAULT 'SAR',
            status TEXT DEFAULT 'draft',
            notes TEXT DEFAULT '',
            created_by TEXT DEFAULT '',
            created_at TEXT DEFAULT '',
            sent_at TEXT DEFAULT '',
            finalized_by TEXT DEFAULT '',
            finalized_at TEXT DEFAULT ''
        )"""
    )
    cur.execute(
        """CREATE TABLE IF NOT EXISTS pda_line_items (
            id SERIAL PRIMARY KEY,
            pda_id INTEGER NOT NULL REFERENCES pda_documents(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            estimated_amount NUMERIC DEFAULT 0,
            actual_amount NUMERIC,
            sort_order INTEGER DEFAULT 0
        )"""
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_pda_line_items_pda ON pda_line_items (pda_id)")

    # Small generic key/value store - currently just the overdue-ETA alert
    # settings (recipient list + on/off), so it doesn't need its own table
    # and its own migration every time a new setting shows up.
    cur.execute(
        """CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT DEFAULT ''
        )"""
    )

    # Statement of Facts - one row per vessel call, matching the company's
    # actual SOF template field-for-field (see SOF_COLUMNS) rather than a
    # generic event log, so the PDF this produces is a drop-in replacement
    # for the Word template, not an approximation of it. Every field is
    # plain text (not a real timestamp column) because the source document
    # itself writes times as free text ("26.06.26 AT 0648 HRS") and is
    # routinely saved with some of them still blank while the port call is
    # in progress - a strict datetime type would reject exactly the
    # half-filled state this form normally sits in.
    sof_cols_sql = ",\n            ".join(f"{col} TEXT DEFAULT ''" for col in SOF_COLUMNS)
    cur.execute(
        f"""CREATE TABLE IF NOT EXISTS sof_documents (
            id SERIAL PRIMARY KEY,
            {sof_cols_sql},
            created_by TEXT DEFAULT '',
            created_at TEXT DEFAULT '',
            updated_at TEXT DEFAULT ''
        )"""
    )
    conn.commit()
    cur.close()
    conn.close()


def fmt_money(value):
    try:
        return "{:,.2f}".format(float(value or 0))
    except (TypeError, ValueError):
        return "0.00"


def build_pda_pdf(doc, items):
    """Renders a PDA (while draft/sent) or FDA (once finalized) as a PDF,
    reusing the Sea Power logo already embedded in the app. Finalized
    documents get an extra Actual + Variance column so the agent can see
    at a glance where the final cost diverged from the estimate."""
    is_fda = doc["status"] == "finalized"
    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    try:
        logo_bytes = base64.b64decode(LOGO_B64)
        pdf.image(io.BytesIO(logo_bytes), x=15, y=12, w=20)
    except Exception:
        pass

    pdf.set_xy(40, 14)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 7, "Sea Power Marine Services Co. Ltd", ln=1)
    pdf.set_x(40)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(110, 120, 130)
    pdf.cell(0, 5, "Compass - Disbursement Account", ln=1)
    pdf.set_text_color(0, 0, 0)

    pdf.ln(10)
    pdf.set_font("Helvetica", "B", 16)
    title = "FINAL DISBURSEMENT ACCOUNT (FDA)" if is_fda else "PROFORMA DISBURSEMENT ACCOUNT (PDA)"
    pdf.cell(0, 9, title, ln=1)

    pdf.set_font("Helvetica", "", 10.5)
    pdf.ln(2)
    meta_rows = [
        ("Port", doc.get("port") or "-"),
        ("Vessel", doc.get("vessel") or "-"),
        ("Reference", doc.get("reference") or "-"),
        ("Currency", doc.get("currency") or "SAR"),
        ("Status", (doc.get("status") or "draft").capitalize()),
        ("Prepared by", doc.get("created_by") or "-"),
        ("Date", doc.get("created_at") or "-"),
    ]
    if is_fda:
        meta_rows.append(("Finalized by", doc.get("finalized_by") or "-"))
        meta_rows.append(("Finalized", doc.get("finalized_at") or "-"))
    for label, value in meta_rows:
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(38, 6.5, label + ":", border=0)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6.5, str(value), ln=1)

    pdf.ln(4)
    currency = doc.get("currency") or "SAR"
    if is_fda:
        col_w = [84, 32, 32, 32]
        headers = ["Charge", "Estimate", "Actual", "Variance"]
    else:
        col_w = [116, 64]
        headers = ["Charge", "Estimate"]

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(18, 58, 86)
    pdf.set_text_color(255, 255, 255)
    for w, h in zip(col_w, headers):
        align = "L" if h == "Charge" else "R"
        pdf.cell(w, 8, h, border=1, align=align, fill=True)
    pdf.ln()
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 10)

    est_total = 0.0
    act_total = 0.0
    fill = False
    for item in items:
        est = float(item.get("estimated_amount") or 0)
        est_total += est
        pdf.set_fill_color(246, 248, 250)
        pdf.cell(col_w[0], 7.5, str(item.get("name") or ""), border=1, align="L", fill=fill)
        pdf.cell(col_w[1], 7.5, fmt_money(est), border=1, align="R", fill=fill)
        if is_fda:
            act = item.get("actual_amount")
            act = float(act) if act is not None else est
            act_total += act
            variance = act - est
            pdf.cell(col_w[2], 7.5, fmt_money(act), border=1, align="R", fill=fill)
            pdf.cell(col_w[3], 7.5, ("+" if variance > 0 else "") + fmt_money(variance), border=1, align="R", fill=fill)
        pdf.ln()
        fill = not fill

    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(col_w[0], 8, "Total (" + currency + ")", border=1, align="L")
    pdf.cell(col_w[1], 8, fmt_money(est_total), border=1, align="R")
    if is_fda:
        variance_total = act_total - est_total
        pdf.cell(col_w[2], 8, fmt_money(act_total), border=1, align="R")
        pdf.cell(col_w[3], 8, ("+" if variance_total > 0 else "") + fmt_money(variance_total), border=1, align="R")
    pdf.ln(12)

    if doc.get("notes"):
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 6, "Notes", ln=1)
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 6, str(doc.get("notes")))

    out = pdf.output(dest="S")
    return bytes(out)


def _sof_val(doc, key):
    v = (doc.get(key) or "").strip()
    return v if v else "-"


def build_sof_pdf(doc):
    """Mirrors the company's own Statement of Facts template field-for-
    field (see SOF_COLUMNS/SOF_TIMELINE_PAIRS), with the header restyled
    to match the Daily Vessel Line-Up report's look - logo + bold navy
    company name + muted subtitle + a right-aligned date block, under a
    gold divider - rather than the plain centered letterhead the original
    .doc used.

    The body is a real bordered table (fixed column widths, shaded label
    cells), not loose label/value text - an earlier label-left-value-right
    version let label width vary per field ("Anchored" vs. "Pilot Boarded
    (Departure)"), so values never lined up from one row to the next and
    long values (e.g. Owners) could run straight into the next column.
    A bordered grid makes every column's width explicit, so nothing drifts
    regardless of how long any one label or value happens to be."""
    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=16)
    pdf.add_page()
    pdf.set_margins(15, 12, 15)
    PAGE_L, PAGE_R = 15, 195
    CONTENT_W = PAGE_R - PAGE_L  # 180mm

    NAVY = (18, 58, 86)
    GOLD = (201, 162, 39)
    MUTED = (110, 120, 130)
    LABEL_FILL = (238, 242, 246)
    LINE = (210, 216, 222)

    try:
        logo_bytes = base64.b64decode(LOGO_B64)
        pdf.image(io.BytesIO(logo_bytes), x=15, y=12, w=18)
    except Exception:
        pass

    pdf.set_xy(37, 13)
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(*NAVY)
    pdf.cell(130, 7, "SEA POWER FOR MARINE SERVICES CO LTD", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(37)
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.set_text_color(*MUTED)
    pdf.cell(130, 5.5, "Statement of Facts", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(0, 0, 0)

    pdf.set_xy(150, 13)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*NAVY)
    pdf.cell(45, 5, "Prepared", align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(150)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(45, 5, str(doc.get("created_at") or "-"), align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_y(32)
    pdf.set_draw_color(*GOLD)
    pdf.set_line_width(0.6)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(7)

    ROW_H = 7.2

    def cell(w, text, bold=False, fill=False, align="L", size=9.5):
        pdf.set_font("Helvetica", "B" if bold else "", size)
        pdf.set_text_color(0, 0, 0)
        if fill:
            pdf.set_fill_color(*LABEL_FILL)
        pdf.set_draw_color(*LINE)
        pdf.cell(w, ROW_H, ("  " + text) if align == "L" else text, border=1, align=align, fill=fill)

    def section_heading(text):
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 11.5)
        pdf.set_text_color(*NAVY)
        pdf.cell(0, 7, text, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(0.5)

    # ---------- Vessel Particulars ----------
    section_heading("Vessel Particulars")
    LW, VW = 32, (CONTENT_W / 2) - 32
    for pair in (
        [("Vessel", _sof_val(doc, "vessel")), ("Voyage", _sof_val(doc, "voyage"))],
        [("Port", _sof_val(doc, "port")), ("Berth", _sof_val(doc, "berth"))],
    ):
        for label, value in pair:
            cell(LW, label, bold=True, fill=True)
            cell(VW, value)
        pdf.ln(ROW_H)
    # Owners/Charterer get the full row width - registered company names
    # routinely run longer than a shared two-column cell can hold.
    for label, value in [("Owners", _sof_val(doc, "owners")), ("Charterer", _sof_val(doc, "charterer"))]:
        cell(LW, label, bold=True, fill=True)
        cell(CONTENT_W - LW, value)
        pdf.ln(ROW_H)

    # ---------- Event Timeline ----------
    section_heading("Event Timeline")
    TLW, TVW = 44, (CONTENT_W / 2) - 44
    for left_col, left_label, right_col, right_label in SOF_TIMELINE_PAIRS:
        cell(TLW, left_label, bold=True, fill=True)
        cell(TVW, _sof_val(doc, left_col))
        cell(TLW, right_label, bold=True, fill=True)
        cell(TVW, _sof_val(doc, right_col))
        pdf.ln(ROW_H)
    cell(TLW, "Berthed (All Fast)", bold=True, fill=True)
    cell(TVW, _sof_val(doc, "berthed_all_fast"))
    cell(TLW, "Cargo Discharged (Final)", bold=True, fill=True)
    cell(TVW, _sof_val(doc, "cargo_discharge_mtons"))
    pdf.ln(ROW_H)

    # ---------- Remaining On Board - a header row instead of cramming
    # four readings onto one line of "IFO: x  MDO: y  ..." text ----------
    section_heading("Remaining On Board (MT)")
    RLW = 34
    RCW = (CONTENT_W - RLW) / 4
    cell(RLW, "", bold=True, fill=True)
    for h in ["IFO", "MDO", "LUBS", "F/Water"]:
        cell(RCW, h, bold=True, fill=True, align="C")
    pdf.ln(ROW_H)
    cell(RLW, "ROB Arrival", bold=True, fill=True)
    for key in ["rob_arrival_ifo", "rob_arrival_mdo", "rob_arrival_lubs", "rob_arrival_fwater"]:
        cell(RCW, _sof_val(doc, key), align="C")
    pdf.ln(ROW_H)
    cell(RLW, "ROB Departure", bold=True, fill=True)
    for key in ["rob_departure_ifo", "rob_departure_mdo", "rob_departure_lubs", "rob_departure_fwater"]:
        cell(RCW, _sof_val(doc, key), align="C")
    pdf.ln(ROW_H)

    # ---------- Draft ----------
    section_heading("Draft (M)")
    DLW = 34
    DCW = (CONTENT_W - DLW) / 2
    cell(DLW, "", bold=True, fill=True)
    for h in ["Forward", "Aft"]:
        cell(DCW, h, bold=True, fill=True, align="C")
    pdf.ln(ROW_H)
    cell(DLW, "Arrival Draft", bold=True, fill=True)
    cell(DCW, _sof_val(doc, "arrival_draft_fwd"), align="C")
    cell(DCW, _sof_val(doc, "arrival_draft_aft"), align="C")
    pdf.ln(ROW_H)
    cell(DLW, "Departure Draft", bold=True, fill=True)
    cell(DCW, _sof_val(doc, "departure_draft_fwd"), align="C")
    cell(DCW, _sof_val(doc, "departure_draft_aft"), align="C")
    pdf.ln(ROW_H)

    if (doc.get("delays_remarks") or "").strip():
        section_heading("Delays / Remarks")
        pdf.set_font("Helvetica", "", 9.5)
        pdf.multi_cell(0, 6, doc.get("delays_remarks"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(1)

    if (doc.get("masters_remarks") or "").strip():
        section_heading("Master's Remarks")
        pdf.set_font("Helvetica", "", 9.5)
        pdf.multi_cell(0, 6, doc.get("masters_remarks"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(2)

    pdf.ln(8)
    y = pdf.get_y()
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_draw_color(*GOLD)
    pdf.set_line_width(0.6)
    pdf.set_xy(15, y)
    pdf.cell(80, 6, "MASTER", border="T")
    pdf.set_xy(130, y)
    pdf.cell(65, 6, "AGENT", border="T", align="R")

    out = pdf.output(dest="S")
    return bytes(out)


def send_email(to_addrs, subject, body):
    """Sends through SMTP creds in the environment (SMTP_HOST/PORT/USER/
    PASSWORD/FROM) - nothing is hardcoded here, same pattern as
    DATABASE_URL/APP_SECRET_KEY. Returns (ok, error_message) instead of
    raising, so a missing/misconfigured mail account degrades to "alert
    not sent" rather than a 500 on whatever triggered it."""
    host = os.environ.get("SMTP_HOST", "")
    port = os.environ.get("SMTP_PORT", "587")
    user = os.environ.get("SMTP_USER", "")
    password = os.environ.get("SMTP_PASSWORD", "")
    sender = os.environ.get("SMTP_FROM", user)
    if not host or not user or not password:
        return False, "Email isn't configured yet (SMTP_HOST/SMTP_USER/SMTP_PASSWORD missing)."
    if not to_addrs:
        return False, "No recipient configured."
    try:
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = ", ".join(to_addrs)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))
        with smtplib.SMTP(host, int(port), timeout=20) as server:
            server.starttls()
            server.login(user, password)
            server.sendmail(sender, to_addrs, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)


def get_setting(key, default=""):
    db = get_db()
    row = db.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else default


def set_setting(key, value):
    db = get_db()
    db.execute(
        "INSERT INTO app_settings (key, value) VALUES (?, ?) ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value",
        (key, value),
    )
    db.commit()


def find_overdue_vessel_groups():
    """A vessel group is overdue once its ETA has passed and at least one
    of its BLs still isn't fully through invoice/approval/DO - the same
    "left" count already shown on the DO Tracker board, just filtered to
    ETA < today and rolled up per port/vessel instead of per BL."""
    db = get_db()
    rows = db.execute(
        "SELECT * FROM records WHERE archived = 0 AND eta != '' ORDER BY port, vessel"
    ).fetchall()
    today = date.today().isoformat()
    groups = {}
    for r in rows:
        if r["eta"] >= today:
            continue
        key = (r["port"], r["vessel"])
        g_ = groups.setdefault(key, {"port": r["port"], "vessel": r["vessel"], "eta": r["eta"], "total": 0, "left": 0})
        g_["total"] += 1
        if not (r["invoice_issued"] and r["approval_received"] and r["do_issued"]):
            g_["left"] += 1
    return [g_ for g_ in groups.values() if g_["left"] > 0]


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


# ---------- Login rate limiting ----------
# A lightweight guard against password-guessing: tracks failed login
# attempts per source IP in memory and locks out further tries for a
# cooldown period once too many pile up. This is a single Flask process
# (no gunicorn/multi-worker setup here), so an in-memory dict is actually
# shared across every request rather than being per-worker and useless -
# it would need a shared store (Redis etc.) if this ever ran as more than
# one instance/process. Not a substitute for a real WAF, but it turns
# "try a password list all night" into "wait 15 minutes," which closes
# off the main risk of a plain username+password login with no 2FA.
_LOGIN_MAX_ATTEMPTS = 8
_LOGIN_WINDOW_SECONDS = 900  # 15 minutes
_login_failures = {}  # ip -> [timestamp, ...] of recent failed attempts


def _client_ip():
    # Render terminates TLS and proxies requests, so the real client IP
    # arrives in X-Forwarded-For (first hop) rather than as the direct
    # socket peer.
    fwd = request.headers.get("X-Forwarded-For", "")
    if fwd:
        return fwd.split(",")[0].strip()
    return request.remote_addr or "unknown"


def _login_locked_out(ip):
    now = time.time()
    recent = [t for t in _login_failures.get(ip, []) if now - t < _LOGIN_WINDOW_SECONDS]
    _login_failures[ip] = recent
    return len(recent) >= _LOGIN_MAX_ATTEMPTS


def _record_login_failure(ip):
    _login_failures.setdefault(ip, []).append(time.time())


@app.route("/login", methods=["GET", "POST"])
def login():
    if not any_users_exist():
        return redirect(url_for("setup"))
    error = None
    if request.method == "POST":
        ip = _client_ip()
        if _login_locked_out(ip):
            error = "Too many failed attempts. Please wait 15 minutes and try again."
            return render_template_string(LOGIN_HTML, error=error)
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            session["role"] = user["role"]
            return redirect(url_for("index"))
        _record_login_failure(ip)
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


@app.route("/pda")
@login_required
def pda_page():
    return render_template_string(PDA_HTML, username=session.get("username"), role=session.get("role"))


@app.route("/sof")
@login_required
def sof_page():
    return render_template_string(SOF_HTML, username=session.get("username"), role=session.get("role"))


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

ATTACHMENT_FLAGS_SQL = """,
    EXISTS(SELECT 1 FROM record_attachments a WHERE a.bl_number = r.bl_number AND a.kind = 'invoice') AS has_invoice_file,
    EXISTS(SELECT 1 FROM record_attachments a WHERE a.bl_number = r.bl_number AND a.kind = 'do') AS has_do_file"""


@app.route("/api/records", methods=["GET"])
@login_required
def list_records():
    db = get_db()
    if session.get("role") == "admin":
        rows = db.execute(f"SELECT r.*{ATTACHMENT_FLAGS_SQL} FROM records r ORDER BY created_at DESC").fetchall()
    else:
        rows = db.execute(
            f"SELECT r.*{ATTACHMENT_FLAGS_SQL} FROM records r WHERE created_by = ? ORDER BY created_at DESC",
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


def _log_audit(bl_number, action, field="", old_value="", new_value=""):
    """Appends one row to audit_log - who did what, when, to which BL.
    Shares the caller's transaction (no commit here), so it only actually
    lands if the caller's own db.commit() goes through right after."""
    db = get_db()
    db.execute(
        "INSERT INTO audit_log (bl_number, action, field, old_value, new_value, by_user, at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (
            bl_number, action, field, "" if old_value is None else str(old_value),
            "" if new_value is None else str(new_value),
            session.get("username", ""), datetime.utcnow().strftime("%Y-%m-%d %H:%M"),
        ),
    )


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
    # A BL number is unique board-wide (it's the primary key), so a manifest
    # that re-lists a BL already on the board just gets silently skipped as
    # "already there" - correct when it's the same shipment uploaded twice,
    # but if that existing row is sitting under a DIFFERENT vessel, this is
    # actually a real collision worth a human looking at (either a genuine
    # data error, or a BL number a shipper has reused for a new shipment),
    # not a routine duplicate. Collected separately and surfaced in the
    # response instead of disappearing into the plain "skipped" count.
    duplicate_elsewhere = []
    for bl_number in bl_numbers:
        if not bl_number:
            continue
        existing = db.execute("SELECT vessel, port FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
        if existing:
            skipped += 1
            if (existing["vessel"] or "") != vessel or (existing["port"] or "") != port:
                duplicate_elsewhere.append({
                    "bl_number": bl_number,
                    "existing_port": existing["port"], "existing_vessel": existing["vessel"],
                })
            continue
        db.execute(
            "INSERT INTO records (bl_number, port, vessel, created_at, created_by) VALUES (?, ?, ?, ?, ?)",
            (bl_number, port, vessel, datetime.utcnow().strftime("%Y-%m-%d %H:%M"), session.get("username")),
        )
        _log_audit(bl_number, "added", "", "", f"{port} / {vessel}")
        added += 1

    db.commit()
    return jsonify({"added": added, "skipped": skipped, "duplicate_elsewhere": duplicate_elsewhere})


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


@app.route("/api/direct-delivery/restore", methods=["POST"])
@login_required
def restore_direct_delivery():
    """Used by the 'Undo' toast after a row is removed from the Direct
    Delivery results/review list - re-inserts it with its original
    classification fields, same pattern as DO Tracker's record restore.
    A no-op (not an error) if that BL already exists, so Undo stays safe
    to click more than once."""
    data = request.get_json(force=True)
    bl_number = str(data.get("bl_number", "")).strip()
    if not bl_number:
        return jsonify({"error": "missing bl_number"}), 400
    db = get_db()
    existing = db.execute("SELECT 1 FROM direct_delivery WHERE bl_number = ?", (bl_number,)).fetchone()
    if existing:
        return jsonify({"ok": True, "note": "already exists"})
    db.execute(
        """INSERT INTO direct_delivery
           (bl_number, is_direct, reason, classified_by, classified_at, needs_review, review_note)
           VALUES (?, ?, ?, ?, ?, ?, ?)""",
        (
            bl_number,
            1 if data.get("is_direct") else 0,
            data.get("reason", ""),
            data.get("classified_by") or session.get("username"),
            data.get("classified_at", ""),
            1 if data.get("needs_review") else 0,
            data.get("review_note", ""),
        ),
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
    _log_audit(bl_number.upper(), "toggle", field, "" if value else "1", "1" if value else "")
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
    old = db.execute("SELECT remarks FROM records WHERE bl_number = ?", (bl_number.upper(),)).fetchone()
    db.execute("UPDATE records SET remarks = ? WHERE bl_number = ?", (remarks, bl_number.upper()))
    _log_audit(bl_number.upper(), "remarks", "remarks", old["remarks"] if old else "", remarks)
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<path:bl_number>", methods=["DELETE"])
@login_required
def delete_record(bl_number):
    if not _owns_record(bl_number.upper()):
        return "Not your record.", 403
    db = get_db()
    db.execute("DELETE FROM records WHERE bl_number = ?", (bl_number.upper(),))
    _log_audit(bl_number.upper(), "deleted")
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
    if inserted:
        _log_audit(str(data.get("bl_number", "")).strip().upper(), "restored")
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
        _log_audit(bl_number, "deleted", "", "", "bulk")
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
            _log_audit(str(item.get("bl_number", "")).strip().upper(), "restored", "", "", "bulk")
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


@app.route("/api/records/bulk-toggle", methods=["POST"])
@login_required
def bulk_toggle_records():
    """Sets one status field (invoice/approval/DO) for many BLs at once -
    the board's row-select checkboxes + action bar use this so clearing a
    whole lot that came in together is one click instead of one toggle per
    row. Staff can only touch their own records, same as every other write
    here; rows they don't own are silently skipped rather than failing the
    whole batch."""
    data = request.get_json(force=True)
    bl_numbers = [str(b).strip().upper() for b in data.get("bl_numbers", []) if str(b).strip()]
    field = data.get("field")
    value = 1 if data.get("value") else 0
    allowed = {
        "invoice_issued": ("invoice_by", "invoice_at"),
        "approval_received": ("approval_by", "approval_at"),
        "do_issued": ("do_by", "do_at"),
    }
    if field not in allowed or not bl_numbers:
        return jsonify({"error": "invalid request"}), 400

    by_field, at_field = allowed[field]
    user = session.get("username", "Unknown")
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M") if value else ""
    by_val = user if value else ""

    db = get_db()
    updated = []
    for bl_number in bl_numbers:
        if not _owns_record(bl_number):
            continue
        db.execute(
            f"UPDATE records SET {field} = ?, {by_field} = ?, {at_field} = ? WHERE bl_number = ?",
            (value, by_val, now, bl_number),
        )
        _log_audit(bl_number, "toggle", field, "" if value else "1", "1" if value else "")
        updated.append(bl_number)
    db.commit()
    return jsonify({"updated": updated})


@app.route("/api/vessel/eta", methods=["POST"])
@login_required
def set_vessel_eta():
    """Sets the expected-arrival date for a whole vessel group at once -
    it's a property of the vessel's call, not of any one BL, so it's stored
    the same way on every row in the group rather than needing a separate
    vessels table. Takes an explicit bl_numbers list (like bulk-delete)
    rather than matching on the port/vessel text - a blank/"Unassigned"
    port or vessel is stored as '' in the database but shown as the literal
    word "Unassigned" in the UI, so matching by that text would silently
    match nothing for any unassigned group."""
    data = request.get_json(force=True)
    bl_numbers = [str(b).strip().upper() for b in data.get("bl_numbers", []) if str(b).strip()]
    eta = str(data.get("eta", "")).strip()
    if not bl_numbers:
        return jsonify({"error": "No BL numbers given"}), 400
    db = get_db()
    updated = 0
    for bl_number in bl_numbers:
        if not _owns_record(bl_number):
            continue
        db.execute("UPDATE records SET eta = ? WHERE bl_number = ?", (eta, bl_number))
        updated += 1
    db.commit()
    return jsonify({"updated": updated})


@app.route("/api/vessel/archive", methods=["POST"])
@login_required
def set_vessel_archived():
    """Archives (or restores) a whole vessel group - takes a finished
    vessel off the main board without deleting its data; archived records
    are still searchable/exportable, just filtered out of the day-to-day
    view by default. Same explicit bl_numbers-list approach as the ETA
    route above, for the same reason."""
    data = request.get_json(force=True)
    bl_numbers = [str(b).strip().upper() for b in data.get("bl_numbers", []) if str(b).strip()]
    archived = 1 if data.get("archived") else 0
    if not bl_numbers:
        return jsonify({"error": "No BL numbers given"}), 400
    db = get_db()
    updated = 0
    for bl_number in bl_numbers:
        if not _owns_record(bl_number):
            continue
        db.execute("UPDATE records SET archived = ? WHERE bl_number = ?", (archived, bl_number))
        updated += 1
    db.commit()
    return jsonify({"updated": updated})


@app.route("/api/records/<path:bl_number>/history", methods=["GET"])
@login_required
def record_history(bl_number):
    if session.get("role") != "admin":
        return "Admins only.", 403
    if not _owns_record(bl_number.upper()):
        return "Not your record.", 403
    db = get_db()
    rows = db.execute(
        "SELECT * FROM audit_log WHERE bl_number = ? ORDER BY id DESC", (bl_number.upper(),)
    ).fetchall()
    return jsonify([dict(r) for r in rows])


# ---------- Invoice / DO file attachments ----------

def _extract_pdf_text(data):
    """All pages' text, concatenated. Fasah's Invoice and Delivery Order
    PDFs are text-layer PDFs (not scans), so pdfplumber's default
    extract_text() - the same approach already used for manifest/packing
    list PDFs elsewhere in this app - reads them cleanly without needing
    OCR."""
    import pdfplumber
    parts = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)


def _detect_fasah_doc_kind(text):
    """Fasah's Invoice and Delivery Order documents are both fixed
    templates - the user confirmed only the BL number/vessel/amounts
    change between documents of the same kind - so a couple of fixed
    anchor phrases from the real templates reliably tell them apart
    without needing to parse the whole layout. Returns 'invoice', 'do',
    or None if neither anchor is found (an unrecognized/different kind
    of PDF, which gets held for manual review rather than guessed at)."""
    upper = text.upper()
    if "DELIVERY ORDER NUMBER" in upper or "DELIVERY ORDER SERIAL NUMBER" in upper:
        return "do"
    if "FASAH PAY INVOICE" in upper or "INVOICE REF" in upper:
        return "invoice"
    return None


def _find_bl_in_text(text, known_bls):
    """Which of the caller's own BL numbers appear in this document's
    text - matched as a whole token (not a bare substring) so one BL
    number that happens to be a prefix of another (e.g. "BO123" inside
    "BO1234") doesn't produce a false match. Returns the list of matches;
    the caller treats exactly one as a confident auto-match and
    zero-or-many as needing a human to pick."""
    upper = text.upper()
    found = []
    for bl in known_bls:
        bl_u = (bl or "").strip().upper()
        if not bl_u:
            continue
        if re.search(r"(?<![A-Z0-9])" + re.escape(bl_u) + r"(?![A-Z0-9])", upper):
            found.append(bl)
    return found


@app.route("/api/attachments/detect", methods=["POST"])
@login_required
def detect_attachment():
    """Reads one dropped PDF and reports what it probably is, without
    saving anything - the frontend's "drop a batch of Invoices/DOs"
    flow calls this once per file, then either auto-uploads it (via the
    existing upload_attachment route, unique kind + unique BL match) or
    shows it for the user to confirm/correct by hand."""
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file received."}), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported."}), 400
    data = file.read()
    if not data:
        return jsonify({"error": "That file is empty."}), 400
    if len(data) > MAX_ATTACHMENT_BYTES:
        return jsonify({"error": "That file is larger than 10MB."}), 400

    try:
        text = _extract_pdf_text(data)
    except Exception:
        return jsonify({"error": "Could not read this PDF - it may be a scanned image rather than a text document."}), 400

    kind = _detect_fasah_doc_kind(text)

    # Only match against BLs this user could actually upload to anyway
    # (same scope _owns_record would allow) - no point surfacing a match
    # the uploader isn't permitted to attach to.
    db = get_db()
    if session.get("role") == "admin":
        bl_rows = db.execute("SELECT bl_number FROM records").fetchall()
    else:
        bl_rows = db.execute(
            "SELECT bl_number FROM records WHERE created_by = ?", (session.get("username"),)
        ).fetchall()
    known_bls = [r["bl_number"] for r in bl_rows]
    matches = _find_bl_in_text(text, known_bls)

    return jsonify({
        "kind": kind,
        "matched_bl": matches[0] if len(matches) == 1 else None,
        "candidates": matches if len(matches) > 1 else [],
    })


@app.route("/api/records/<path:bl_number>/attachment/<kind>", methods=["POST"])
@login_required
def upload_attachment(bl_number, kind):
    bl_number = bl_number.upper()
    if kind not in ATTACHMENT_KINDS:
        return jsonify({"error": "Invalid attachment kind."}), 400
    if not _owns_record(bl_number):
        return jsonify({"error": "Not your record."}), 403
    file = request.files.get("file")
    if not file or not file.filename:
        return jsonify({"error": "No file was selected."}), 400
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are accepted."}), 400
    data = file.read()
    if not data:
        return jsonify({"error": "That file is empty."}), 400
    if len(data) > MAX_ATTACHMENT_BYTES:
        return jsonify({"error": "That file is larger than 10MB."}), 400

    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    user = session.get("username", "Unknown")
    db.execute(
        """INSERT INTO record_attachments (bl_number, kind, filename, content_type, data, file_size, uploaded_by, uploaded_at)
           VALUES (?, ?, ?, 'application/pdf', ?, ?, ?, ?)
           ON CONFLICT (bl_number, kind) DO UPDATE SET
             filename = EXCLUDED.filename, data = EXCLUDED.data, file_size = EXCLUDED.file_size,
             uploaded_by = EXCLUDED.uploaded_by, uploaded_at = EXCLUDED.uploaded_at""",
        (bl_number, kind, file.filename, psycopg2.Binary(data), len(data), user, now),
    )
    _log_audit(bl_number, "attachment", kind, "", file.filename)

    # Attaching the file *is* the real-world signal that the invoice/DO was
    # actually issued - no reason to also make someone flip the slider by
    # hand afterward. Only flips the status 0 -> 1 though: if it's already
    # marked issued and this upload is just a Replace (a corrected file),
    # that's not a fresh issuance, so the original by/at stays as-is rather
    # than being silently rewritten to whoever happened to replace the file.
    status_field = {"invoice": "invoice_issued", "do": "do_issued"}.get(kind)
    status_by_at = {"invoice_issued": ("invoice_by", "invoice_at"), "do_issued": ("do_by", "do_at")}
    auto_issued = False
    if status_field:
        by_field, at_field = status_by_at[status_field]
        current = db.execute(f"SELECT {status_field} FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
        if current and not current[status_field]:
            db.execute(
                f"UPDATE records SET {status_field} = 1, {by_field} = ?, {at_field} = ? WHERE bl_number = ?",
                (user, now, bl_number),
            )
            _log_audit(bl_number, "toggle", status_field, "", "1")
            auto_issued = True

    db.commit()
    return jsonify({
        "ok": True, "filename": file.filename, "uploaded_by": user, "uploaded_at": now,
        "auto_issued_field": status_field if auto_issued else None,
    })


@app.route("/api/records/<path:bl_number>/attachment/<kind>", methods=["GET"])
@login_required
def download_attachment(bl_number, kind):
    """Deliberately NOT gated by _owns_record - the whole point of this
    feature is the handoff between two different people (whoever issues
    the invoice/DO isn't necessarily who forwards it to the customs
    broker), so any signed-in user who already knows the BL number can
    pull the file, same as the /lookup route below."""
    bl_number = bl_number.upper()
    if kind not in ATTACHMENT_KINDS:
        return "Invalid attachment kind.", 400
    db = get_db()
    row = db.execute(
        "SELECT filename, content_type, data FROM record_attachments WHERE bl_number = ? AND kind = ?",
        (bl_number, kind),
    ).fetchone()
    if not row:
        return "No file attached yet.", 404
    filename = row["filename"] or f"{kind}_{bl_number}.pdf"
    return Response(
        bytes(row["data"]),
        mimetype=row["content_type"] or "application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.route("/api/records/<path:bl_number>/attachment/<kind>", methods=["DELETE"])
@login_required
def delete_attachment(bl_number, kind):
    bl_number = bl_number.upper()
    if kind not in ATTACHMENT_KINDS:
        return jsonify({"error": "Invalid attachment kind."}), 400
    if not _owns_record(bl_number):
        return jsonify({"error": "Not your record."}), 403
    db = get_db()
    db.execute("DELETE FROM record_attachments WHERE bl_number = ? AND kind = ?", (bl_number, kind))
    _log_audit(bl_number, "attachment_removed", kind)
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/records/<path:bl_number>/lookup", methods=["GET"])
@login_required
def lookup_record(bl_number):
    """Looks up one BL by its exact number regardless of who created it.
    Used by the Documents modal (the doc chips next to a BL's own row) -
    not scoped to the viewer's own BLs since a record's docs are still
    shown to an admin or anyone else who can already see that row."""
    bl_number = bl_number.strip().upper()
    if not bl_number:
        return jsonify({"error": "Enter a BL number."}), 400
    db = get_db()
    row = db.execute(
        f"""SELECT bl_number, port, vessel, created_by, invoice_issued, invoice_by, invoice_at,
                   do_issued, do_by, do_at{ATTACHMENT_FLAGS_SQL}
            FROM records r WHERE bl_number = ?""",
        (bl_number,),
    ).fetchone()
    if not row:
        return jsonify({"error": f'No BL found matching "{bl_number}".'}), 404
    result = dict(row)
    # Attach filename/uploaded_by/uploaded_at per kind too - the has_invoice_file/
    # has_do_file flags above are enough for a status dot, but the Documents
    # modal and the Find-a-BL lookup both want to show who uploaded what and
    # when, not just whether something's there.
    atts = db.execute(
        "SELECT kind, filename, uploaded_by, uploaded_at FROM record_attachments WHERE bl_number = ?",
        (bl_number,),
    ).fetchall()
    result["attachments"] = {a["kind"]: dict(a) for a in atts}
    return jsonify(result)


@app.route("/api/export", methods=["GET"])
@login_required
def export_records():
    """Downloads one vessel group's BLs as an .xlsx, via ?port=&vessel= -
    for handing a status report to the principal or management without
    them needing a login. Vessel-wise only (there's no board- or
    port-level export anymore): different vessels can have different
    cargo owners, so a combined export doesn't make sense here."""
    port = request.args.get("port", "")
    vessel = request.args.get("vessel", "")
    db = get_db()
    is_admin = session.get("role") == "admin"
    sql = "SELECT * FROM records WHERE port = ? AND vessel = ?"
    params = [port, vessel]
    if not is_admin:
        sql += " AND created_by = ?"
        params.append(session.get("username"))
    sql += " ORDER BY bl_number"
    rows = [dict(r) for r in db.execute(sql, tuple(params)).fetchall()]

    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DO Tracker"
    headers = ["BL Number", "DO Issued"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for r in rows:
        ws.append([
            r.get("bl_number", ""),
            "Yes" if r.get("do_issued") else "No",
        ])
    for col_cells in ws.columns:
        width = max((len(str(c.value)) for c in col_cells if c.value is not None), default=8)
        ws.column_dimensions[col_cells[0].column_letter].width = min(width + 2, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = (vessel.strip().upper() or "UNASSIGNED") + ".xlsx"
    return Response(
        buf.read(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


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


@app.route("/api/vessels/restore", methods=["POST"])
@login_required
def restore_vessel():
    """Used by the 'Undo' toast after a vessel is removed from the
    tracker - re-inserts it with its original MMSI/operator, same pattern
    as the DO Tracker record restore. A no-op (not an error) if a vessel
    with that name already exists, so Undo stays safe to click more than
    once."""
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"error": "missing name"}), 400
    db = get_db()
    existing = db.execute("SELECT 1 FROM vessels WHERE name = ?", (name,)).fetchone()
    if existing:
        return jsonify({"ok": True, "note": "already exists"})
    db.execute(
        "INSERT INTO vessels (name, mmsi, operator, updated_by, updated_at) VALUES (?, ?, ?, ?, ?)",
        (
            name,
            data.get("mmsi", ""),
            data.get("operator", ""),
            data.get("updated_by") or session.get("username"),
            data.get("updated_at", ""),
        ),
    )
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


# ---------- PDA / FDA (Disbursement Accounts) ----------

def _num_row(row, fields):
    d = dict(row)
    for k in fields:
        if d.get(k) is not None:
            d[k] = float(d[k])
    return d


@app.route("/api/pda/templates", methods=["GET"])
@login_required
def list_pda_templates():
    db = get_db()
    rows = db.execute("SELECT * FROM pda_templates ORDER BY port, sort_order, id").fetchall()
    templates = {}
    for r in rows:
        templates.setdefault(r["port"], []).append(_num_row(r, ["default_amount"]))
    return jsonify(templates)


@app.route("/api/pda/templates", methods=["POST"])
@login_required
@admin_required
def add_pda_template():
    data = request.get_json(force=True)
    port = (data.get("port") or "").strip()
    name = (data.get("name") or "").strip()
    if not port or not name:
        return jsonify({"error": "Port and charge name are required."}), 400
    try:
        amount = float(data.get("default_amount") or 0)
    except (TypeError, ValueError):
        return jsonify({"error": "Default amount must be a number."}), 400
    db = get_db()
    order_row = db.execute("SELECT COALESCE(MAX(sort_order), -1) + 1 AS n FROM pda_templates WHERE port = ?", (port,)).fetchone()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    row = db.execute(
        "INSERT INTO pda_templates (port, name, default_amount, sort_order, created_at) VALUES (?, ?, ?, ?, ?) RETURNING *",
        (port, name, amount, order_row["n"], now),
    ).fetchone()
    db.commit()
    return jsonify(_num_row(row, ["default_amount"]))


@app.route("/api/pda/templates/<int:template_id>", methods=["PUT"])
@login_required
@admin_required
def update_pda_template(template_id):
    data = request.get_json(force=True)
    db = get_db()
    existing = db.execute("SELECT * FROM pda_templates WHERE id = ?", (template_id,)).fetchone()
    if not existing:
        return jsonify({"error": "Not found."}), 404
    name = existing["name"]
    if "name" in data:
        name = (data.get("name") or "").strip()
        if not name:
            return jsonify({"error": "Charge name can't be empty."}), 400
    amount = existing["default_amount"]
    if "default_amount" in data:
        try:
            amount = float(data.get("default_amount") or 0)
        except (TypeError, ValueError):
            return jsonify({"error": "Default amount must be a number."}), 400
    db.execute("UPDATE pda_templates SET name = ?, default_amount = ? WHERE id = ?", (name, amount, template_id))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/templates/<int:template_id>", methods=["DELETE"])
@login_required
@admin_required
def delete_pda_template(template_id):
    db = get_db()
    db.execute("DELETE FROM pda_templates WHERE id = ?", (template_id,))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/documents", methods=["GET"])
@login_required
def list_pda_documents():
    db = get_db()
    rows = db.execute("SELECT * FROM pda_documents ORDER BY created_at DESC, id DESC").fetchall()
    docs = []
    for r in rows:
        d = dict(r)
        items = db.execute(
            "SELECT estimated_amount, actual_amount FROM pda_line_items WHERE pda_id = ?", (r["id"],)
        ).fetchall()
        d["estimated_total"] = round(sum(float(i["estimated_amount"] or 0) for i in items), 2)
        d["actual_total"] = (
            round(sum(float(i["actual_amount"]) if i["actual_amount"] is not None else float(i["estimated_amount"] or 0) for i in items), 2)
            if d["status"] == "finalized" else None
        )
        d["line_item_count"] = len(items)
        docs.append(d)
    return jsonify(docs)


@app.route("/api/pda/documents", methods=["POST"])
@login_required
def create_pda_document():
    data = request.get_json(force=True)
    port = (data.get("port") or "").strip()
    vessel = (data.get("vessel") or "").strip()
    if not port or not vessel:
        return jsonify({"error": "Port and vessel are required."}), 400
    reference = (data.get("reference") or "").strip()
    currency = (data.get("currency") or "SAR").strip() or "SAR"
    notes = data.get("notes") or ""
    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    doc = db.execute(
        """INSERT INTO pda_documents (port, vessel, reference, currency, status, notes, created_by, created_at)
           VALUES (?, ?, ?, ?, 'draft', ?, ?, ?) RETURNING *""",
        (port, vessel, reference, currency, notes, session.get("username"), now),
    ).fetchone()
    templates = db.execute("SELECT * FROM pda_templates WHERE port = ? ORDER BY sort_order, id", (port,)).fetchall()
    for t in templates:
        db.execute(
            "INSERT INTO pda_line_items (pda_id, name, estimated_amount, sort_order) VALUES (?, ?, ?, ?)",
            (doc["id"], t["name"], t["default_amount"], t["sort_order"]),
        )
    db.commit()
    return jsonify({"ok": True, "id": doc["id"]})


@app.route("/api/pda/documents/<int:pda_id>", methods=["GET"])
@login_required
def get_pda_document(pda_id):
    db = get_db()
    doc = db.execute("SELECT * FROM pda_documents WHERE id = ?", (pda_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    items = db.execute("SELECT * FROM pda_line_items WHERE pda_id = ? ORDER BY sort_order, id", (pda_id,)).fetchall()
    d = dict(doc)
    d["items"] = [_num_row(i, ["estimated_amount", "actual_amount"]) for i in items]
    return jsonify(d)


@app.route("/api/pda/documents/<int:pda_id>", methods=["PUT"])
@login_required
def update_pda_document(pda_id):
    data = request.get_json(force=True)
    db = get_db()
    doc = db.execute("SELECT * FROM pda_documents WHERE id = ?", (pda_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    reference = data.get("reference", doc["reference"])
    currency = (data.get("currency", doc["currency"]) or doc["currency"])
    notes = data.get("notes", doc["notes"])
    vessel = (data.get("vessel", doc["vessel"]) or doc["vessel"])
    port = (data.get("port", doc["port"]) or doc["port"])
    db.execute(
        "UPDATE pda_documents SET port = ?, vessel = ?, reference = ?, currency = ?, notes = ? WHERE id = ?",
        (port, vessel, reference, currency, notes, pda_id),
    )
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/documents/<int:pda_id>/mark-sent", methods=["POST"])
@login_required
def mark_pda_sent(pda_id):
    db = get_db()
    doc = db.execute("SELECT * FROM pda_documents WHERE id = ?", (pda_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    if doc["status"] != "draft":
        return jsonify({"error": "Only a draft PDA can be marked as sent."}), 400
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    db.execute("UPDATE pda_documents SET status = 'sent', sent_at = ? WHERE id = ?", (now, pda_id))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/documents/<int:pda_id>/finalize", methods=["POST"])
@login_required
def finalize_pda_document(pda_id):
    db = get_db()
    doc = db.execute("SELECT * FROM pda_documents WHERE id = ?", (pda_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    if doc["status"] == "finalized":
        return jsonify({"error": "This document is already finalized as an FDA."}), 400
    items = db.execute("SELECT * FROM pda_line_items WHERE pda_id = ?", (pda_id,)).fetchall()
    for item in items:
        if item["actual_amount"] is None:
            db.execute("UPDATE pda_line_items SET actual_amount = ? WHERE id = ?", (item["estimated_amount"], item["id"]))
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    db.execute(
        "UPDATE pda_documents SET status = 'finalized', finalized_by = ?, finalized_at = ? WHERE id = ?",
        (session.get("username"), now, pda_id),
    )
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/documents/<int:pda_id>", methods=["DELETE"])
@login_required
def delete_pda_document(pda_id):
    db = get_db()
    db.execute("DELETE FROM pda_documents WHERE id = ?", (pda_id,))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/documents/<int:pda_id>/line-items", methods=["POST"])
@login_required
def add_pda_line_item(pda_id):
    data = request.get_json(force=True)
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"error": "Charge name is required."}), 400
    db = get_db()
    doc = db.execute("SELECT id FROM pda_documents WHERE id = ?", (pda_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    try:
        amount = float(data.get("estimated_amount") or 0)
    except (TypeError, ValueError):
        return jsonify({"error": "Amount must be a number."}), 400
    order_row = db.execute("SELECT COALESCE(MAX(sort_order), -1) + 1 AS n FROM pda_line_items WHERE pda_id = ?", (pda_id,)).fetchone()
    item = db.execute(
        "INSERT INTO pda_line_items (pda_id, name, estimated_amount, sort_order) VALUES (?, ?, ?, ?) RETURNING *",
        (pda_id, name, amount, order_row["n"]),
    ).fetchone()
    db.commit()
    return jsonify(_num_row(item, ["estimated_amount", "actual_amount"]))


@app.route("/api/pda/line-items/<int:item_id>", methods=["PUT"])
@login_required
def update_pda_line_item(item_id):
    data = request.get_json(force=True)
    db = get_db()
    item = db.execute("SELECT * FROM pda_line_items WHERE id = ?", (item_id,)).fetchone()
    if not item:
        return jsonify({"error": "Not found."}), 404
    updates = {}
    if "name" in data:
        name = (data.get("name") or "").strip()
        if not name:
            return jsonify({"error": "Charge name can't be empty."}), 400
        updates["name"] = name
    for field in ("estimated_amount", "actual_amount"):
        if field in data:
            val = data.get(field)
            if val is None or val == "":
                updates[field] = None
            else:
                try:
                    updates[field] = float(val)
                except (TypeError, ValueError):
                    return jsonify({"error": field + " must be a number."}), 400
    if not updates:
        return jsonify({"ok": True})
    set_clause = ", ".join(k + " = ?" for k in updates)
    db.execute(f"UPDATE pda_line_items SET {set_clause} WHERE id = ?", (*updates.values(), item_id))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/line-items/<int:item_id>", methods=["DELETE"])
@login_required
def delete_pda_line_item(item_id):
    db = get_db()
    db.execute("DELETE FROM pda_line_items WHERE id = ?", (item_id,))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/pda/documents/<int:pda_id>/pdf", methods=["GET"])
@login_required
def export_pda_pdf(pda_id):
    db = get_db()
    doc = db.execute("SELECT * FROM pda_documents WHERE id = ?", (pda_id,)).fetchone()
    if not doc:
        return "Not found.", 404
    items = db.execute("SELECT * FROM pda_line_items WHERE pda_id = ? ORDER BY sort_order, id", (pda_id,)).fetchall()
    pdf_bytes = build_pda_pdf(dict(doc), [dict(i) for i in items])
    kind = "FDA" if doc["status"] == "finalized" else "PDA"
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{kind}_{doc['port']}_{doc['vessel']}_{doc['id']}")
    return Response(
        pdf_bytes,
        mimetype="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{safe}.pdf"'},
    )


# ---------- Alerts (ETA-overdue email notifications) ----------

@app.route("/api/settings/alerts", methods=["GET"])
@login_required
@admin_required
def get_alert_settings():
    return jsonify({
        "enabled": get_setting("alerts_enabled", "") == "1",
        "recipients": get_setting("alerts_recipients", ""),
        "mail_configured": bool(os.environ.get("SMTP_HOST") and os.environ.get("SMTP_USER") and os.environ.get("SMTP_PASSWORD")),
    })


@app.route("/api/settings/alerts", methods=["POST"])
@login_required
@admin_required
def update_alert_settings():
    data = request.get_json(force=True)
    if "enabled" in data:
        set_setting("alerts_enabled", "1" if data.get("enabled") else "0")
    if "recipients" in data:
        set_setting("alerts_recipients", (data.get("recipients") or "").strip())
    return jsonify({"ok": True})


@app.route("/api/alerts/check-overdue", methods=["POST"])
@login_required
@admin_required
def check_overdue_now():
    overdue = find_overdue_vessel_groups()
    recipients = [a.strip() for a in get_setting("alerts_recipients", "").split(",") if a.strip()]
    if not overdue:
        return jsonify({"ok": True, "overdue_count": 0, "sent": False, "note": "Nothing overdue right now."})
    if not recipients:
        return jsonify({"ok": True, "overdue_count": len(overdue), "sent": False, "note": "No alert recipients configured yet."})
    lines = [f"{len(overdue)} vessel group(s) have an ETA that's passed with BLs still pending:", ""]
    for g_ in overdue:
        lines.append(f"- {g_['port']} / {g_['vessel']}: ETA {g_['eta']}, {g_['left']} of {g_['total']} BL(s) still pending")
    lines.append("")
    lines.append("- Compass (Sea Power DO Tracker)")
    ok, err = send_email(recipients, f"Compass: {len(overdue)} vessel(s) overdue on ETA", "\n".join(lines))
    return jsonify({"ok": ok, "overdue_count": len(overdue), "sent": ok, "error": err})


# ---------- SOF (Statement of Facts) ----------

@app.route("/api/sof/documents", methods=["GET"])
@login_required
def list_sof_documents():
    db = get_db()
    rows = db.execute(
        "SELECT id, vessel, voyage, port, berth, created_by, created_at, updated_at FROM sof_documents ORDER BY created_at DESC, id DESC"
    ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/sof/documents", methods=["POST"])
@login_required
def create_sof_document():
    data = request.get_json(force=True)
    vessel = (data.get("vessel") or "").strip()
    port = (data.get("port") or "").strip()
    if not vessel or not port:
        return jsonify({"error": "Vessel and port are required."}), 400
    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    cols = list(SOF_COLUMNS) + ["created_by", "created_at", "updated_at"]
    vals = [(data.get(col) or "").strip() if isinstance(data.get(col), str) else (data.get(col) or "") for col in SOF_COLUMNS]
    vals += [session.get("username"), now, now]
    col_sql = ", ".join(cols)
    placeholders = ", ".join("?" for _ in cols)
    row = db.execute(
        f"INSERT INTO sof_documents ({col_sql}) VALUES ({placeholders}) RETURNING id",
        tuple(vals),
    ).fetchone()
    db.commit()
    return jsonify({"ok": True, "id": row["id"]})


@app.route("/api/sof/documents/<int:doc_id>", methods=["GET"])
@login_required
def get_sof_document(doc_id):
    db = get_db()
    doc = db.execute("SELECT * FROM sof_documents WHERE id = ?", (doc_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    return jsonify(dict(doc))


@app.route("/api/sof/documents/<int:doc_id>", methods=["PUT"])
@login_required
def update_sof_document(doc_id):
    data = request.get_json(force=True)
    db = get_db()
    doc = db.execute("SELECT * FROM sof_documents WHERE id = ?", (doc_id,)).fetchone()
    if not doc:
        return jsonify({"error": "Not found."}), 404
    updates = {}
    for col in SOF_COLUMNS:
        if col in data:
            val = data.get(col)
            updates[col] = val.strip() if isinstance(val, str) else (val or "")
    if "vessel" in updates and not updates["vessel"]:
        return jsonify({"error": "Vessel can't be empty."}), 400
    if "port" in updates and not updates["port"]:
        return jsonify({"error": "Port can't be empty."}), 400
    if not updates:
        return jsonify({"ok": True})
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    updates["updated_at"] = now
    set_clause = ", ".join(k + " = ?" for k in updates)
    db.execute(f"UPDATE sof_documents SET {set_clause} WHERE id = ?", (*updates.values(), doc_id))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/sof/documents/<int:doc_id>", methods=["DELETE"])
@login_required
def delete_sof_document(doc_id):
    db = get_db()
    db.execute("DELETE FROM sof_documents WHERE id = ?", (doc_id,))
    db.commit()
    return jsonify({"ok": True})


@app.route("/api/sof/documents/<int:doc_id>/pdf", methods=["GET"])
@login_required
def export_sof_pdf(doc_id):
    db = get_db()
    doc = db.execute("SELECT * FROM sof_documents WHERE id = ?", (doc_id,)).fetchone()
    if not doc:
        return "Not found.", 404
    pdf_bytes = build_sof_pdf(dict(doc))
    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", f"SOF_{doc['port']}_{doc['vessel']}_{doc['id']}")
    return Response(
        pdf_bytes,
        mimetype="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{safe}.pdf"'},
    )


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

    <a class="tile" href="/pda">
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16v16H4z"/><path d="M8 9h8M8 13h8M8 17h4"/></svg>
      </div>
      <h3>Disbursement Accounts</h3>
      <p>Build a PDA from a per-port charge template, then finalize it into an FDA once actual costs are known.</p>
      <span class="tile-go">Open <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 6l6 6-6 6"/></svg></span>
    </a>

    <a class="tile" href="/sof">
      <div class="tile-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l8 4v6c0 5-3.5 8.5-8 10-4.5-1.5-8-5-8-10V6l8-4z"/><path d="M9 12l2 2 4-4"/></svg>
      </div>
      <h3>Statement of Facts</h3>
      <p>Log a vessel call's event timeline field-by-field and export it as a signed-off SOF.</p>
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
  .vessel-row.active { background: color-mix(in srgb, var(--navy) 12%, transparent); }
  :root[data-theme="dark"] .vessel-row.active { background: color-mix(in srgb, var(--navy-light) 20%, transparent); }
  .vname-wrap { overflow: hidden; min-width: 0; }
  .vessel-row .vname { display: block; font-size: 13.5px; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .vessel-row .voperator { display: block; font-size: 11px; color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
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
  .toast a { color: var(--gold-light); font-weight: 700; text-decoration: none; cursor: pointer; flex-shrink: 0; }
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
    <p>Real-time positions for the vessels you're tracking, pulled straight from MarineTraffic.</p>
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
      '</div>';
  }).join('');
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

document.getElementById('vesselList').addEventListener('click', (e) => {
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
  // Instant delete + Undo toast, same pattern as DO Tracker - no blocking
  // confirm() dialog.
  const info = vessels[name] || { mmsi: '', operator: '' };
  const removed = { name, mmsi: info.mmsi || '', operator: info.operator || '' };
  await fetch('/api/vessels/' + encodeURIComponent(name), {method: 'DELETE'});
  if (selected === name) selected = null;
  await loadData();
  showToast('Removed ' + name + '.', {
    actionLabel: 'Undo',
    duration: 5000,
    onAction: async () => {
      await fetch('/api/vessels/restore', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(removed)
      });
      await loadData();
      showToast('Restored ' + name + '.');
    }
  });
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

function showToast(message, opts) {
  opts = opts || {};
  const host = document.getElementById('toastHost');
  const el = document.createElement('div');
  el.className = 'toast' + (opts.error ? ' error' : '');
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
  const duration = opts.duration || 3200;
  const timer = setTimeout(dismiss, duration);
  function dismiss() {
    clearTimeout(timer);
    el.classList.add('fading');
    setTimeout(() => el.remove(), 220);
  }
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

  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; pointer-events: none; max-width: min(320px, calc(100vw - 40px)); }
  #toastHost .toast { pointer-events: auto; }
  .toast { background: var(--navy-deep); color: #fff; padding: 11px 16px; border-radius: 12px; font-size: 13px; display: flex; align-items: center; gap: 14px; box-shadow: 0 10px 30px rgba(0,0,0,0.25); animation: toast-in .18s ease-out; max-width: 320px; }
  .toast a { color: var(--gold-light); font-weight: 700; text-decoration: none; cursor: pointer; flex-shrink: 0; }
  .toast.error { background: var(--danger); }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(8px); } }
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
    <p>Upload a cargo packing list and every BL over 30MT or 12m gets flagged as Direct Delivery - unless it's wheeled or a coil, in which case it doesn't need a low-bed trailer.</p>
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

  <div id="toastHost"></div>

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

let lastResultsRows = [];
let lastReviewRows = [];

function showToast(message, opts) {
  opts = opts || {};
  const host = document.getElementById('toastHost');
  const el = document.createElement('div');
  el.className = 'toast' + (opts.error ? ' error' : '');
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

async function loadResults() {
  const res = await fetch('/api/direct-delivery');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const rows = await res.json();
  lastResultsRows = rows;
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
  lastReviewRows = rows;
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
  const removed = (fromReview ? lastReviewRows : lastResultsRows).find(r => r.bl_number === bl)
    || lastResultsRows.find(r => r.bl_number === bl) || lastReviewRows.find(r => r.bl_number === bl)
    || { bl_number: bl };

  // Instant delete + Undo toast, same pattern as DO Tracker - no blocking
  // confirm() dialog.
  try {
    await fetch(`/api/direct-delivery/${encodeURIComponent(bl)}`, {method: 'DELETE'});
  } catch (e) {}
  await loadResults();
  await loadReview();

  showToast('Removed BL ' + bl + '.', {
    actionLabel: 'Undo',
    duration: 5000,
    onAction: async () => {
      await fetch('/api/direct-delivery/restore', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(removed)
      });
      await loadResults();
      await loadReview();
      showToast('Restored BL ' + bl + '.');
    }
  });
}

loadResults();
loadReview();
</script>
</body></html>
"""

PDA_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Disbursement Accounts</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --danger: #d1483f; --danger-bg: #fbeceb; --ok: #1c8a5a; --ok-bg: #e7f5ee;
    --warn: #8a6d1f; --warn-bg: rgba(212,160,23,0.16);
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05); --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --danger: #e2685f; --danger-bg: #3a2220; --ok: #3ecb8e; --ok-bg: #163329;
    --warn: var(--gold-light); --warn-bg: rgba(227,187,76,0.16);
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25); --shadow-md: 0 10px 30px rgba(0,0,0,0.35);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 40px;
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
  .page-head p { color: var(--muted); margin: 0; font-size: 13.5px; max-width: 680px; }

  .panel { background: var(--card); border: 1px solid var(--border); border-radius: 18px; box-shadow: var(--shadow-sm); padding: 18px 20px; margin-bottom: 18px; }
  .panel h2 { font-size: 15px; margin: 0 0 2px; }
  .panel .panel-sub { font-size: 12px; color: var(--muted); margin: 0 0 14px; }

  label.field-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; color: var(--muted); display: block; margin-bottom: 5px; }
  input[type=text], input[type=number], input[type=email], select, textarea {
    width: 100%; padding: 9px 11px; border: 1px solid var(--border); border-radius: 9px;
    font-size: 13.5px; font-family: inherit; background: var(--bg); color: var(--text);
  }
  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--navy-light); }
  textarea { resize: vertical; min-height: 56px; }

  .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 12px; }
  .form-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }

  .btn { background: var(--navy); color: #fff; border: none; border-radius: 999px; padding: 9px 17px; font-size: 12.5px; font-weight: 600; cursor: pointer; transition: background .15s ease, transform .08s ease; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
  .btn:hover { background: var(--navy-light); }
  .btn:active { transform: scale(.97); }
  .btn:disabled { opacity: .5; cursor: default; }
  .btn.ghost { background: none; color: var(--navy); border: 1px solid var(--border); }
  :root[data-theme="dark"] .btn.ghost { color: var(--navy-light); }
  .btn.ghost:hover { background: var(--border); }
  .btn.ghost.danger { color: var(--danger); }
  .btn.ghost.danger:hover { background: var(--danger-bg); }
  .btn.small { padding: 5px 11px; font-size: 11.5px; }

  table.doc-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  table.doc-table th { text-align: left; font-size: 10.5px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); padding: 8px 10px; border-bottom: 1px solid var(--border); }
  table.doc-table td { padding: 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }
  table.doc-table tr:last-child td { border-bottom: none; }
  table.doc-table tr.doc-row { cursor: pointer; }
  table.doc-table tr.doc-row:hover td { background: color-mix(in srgb, var(--navy-light) 5%, transparent); }

  .badge { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 10.5px; font-weight: 700; letter-spacing: .3px; text-transform: uppercase; }
  .badge.draft { background: color-mix(in srgb, var(--muted) 16%, transparent); color: var(--muted); }
  .badge.sent { background: var(--warn-bg); color: var(--warn); }
  .badge.finalized { background: var(--ok-bg); color: var(--ok); }

  .empty-note { color: var(--muted); font-size: 13px; padding: 10px 2px; }

  #docDetail { display: none; }
  .detail-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 14px; }
  .detail-title { font-size: 16px; font-weight: 700; margin: 0 0 2px; }
  .detail-sub { font-size: 12px; color: var(--muted); }
  .detail-actions { display: flex; gap: 8px; flex-wrap: wrap; }

  table.items-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 6px; }
  table.items-table th { text-align: left; font-size: 10.5px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); padding: 7px 8px; border-bottom: 1px solid var(--border); }
  table.items-table td { padding: 6px 8px; border-bottom: 1px solid var(--border); }
  table.items-table td input { text-align: right; }
  table.items-table td:first-child input { text-align: left; }
  table.items-table tr.total-row td { font-weight: 700; border-top: 2px solid var(--border); border-bottom: none; padding-top: 10px; }
  .variance-pos { color: var(--danger); }
  .variance-neg { color: var(--ok); }
  .row-remove { background: none; border: none; color: var(--muted); cursor: pointer; font-size: 15px; padding: 2px 6px; border-radius: 6px; }
  .row-remove:hover { background: var(--danger-bg); color: var(--danger); }

  .tmpl-port-group { margin-bottom: 16px; }
  .tmpl-port-group h3 { font-size: 13px; margin: 0 0 8px; color: var(--navy-light); }
  .tmpl-add-row { display: flex; gap: 8px; margin-top: 8px; }
  .tmpl-add-row input[type=text] { flex: 2; }
  .tmpl-add-row input[type=number] { flex: 1; }

  .alert-row { display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }
  .switch-sm { position: relative; display: inline-flex; width: 38px; height: 21px; cursor: pointer; flex-shrink: 0; }
  .switch-sm input { opacity: 0; width: 0; height: 0; position: absolute; }
  .switch-track-sm { position: absolute; inset: 0; border-radius: 999px; background: var(--border); transition: background .2s ease; }
  .switch-knob-sm { position: absolute; top: 2px; left: 2px; width: 17px; height: 17px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,0.3); transition: transform .2s ease; }
  input:checked + .switch-track-sm { background: var(--ok); }
  input:checked + .switch-track-sm .switch-knob-sm { transform: translateX(17px); }
  .mail-status { font-size: 11.5px; color: var(--muted); margin-top: 4px; }
  .mail-status.warn { color: var(--warn); }
  .mail-status.ok { color: var(--ok); }

  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; pointer-events: none; max-width: min(340px, calc(100vw - 40px)); }
  #toastHost .toast { pointer-events: auto; }
  .toast { background: var(--navy-deep); color: #fff; padding: 11px 16px; border-radius: 12px; font-size: 13px; display: flex; align-items: center; gap: 14px; box-shadow: 0 10px 30px rgba(0,0,0,0.25); animation: toast-in .18s ease-out; max-width: 340px; }
  .toast.error { background: var(--danger); }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(8px); } }
</style>
</head>
<body>
  <div class="topbar">
    <a href="/" class="brand">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Disbursement Accounts</span>
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
    <h1>Disbursement Accounts</h1>
    <p>Build a Proforma Disbursement Account (PDA) per port call from a per-port charge template, then finalize it into an FDA once the real costs are known - estimate and actual stay on the same document so the variance is never a separate reconciliation step.</p>
  </div>

  <div id="toastHost"></div>

  <div class="panel" id="listPanel">
    <h2>New disbursement account</h2>
    <p class="panel-sub">Pick the port and vessel - charges from that port's template are added automatically, ready to adjust.</p>
    <div class="form-grid">
      <div>
        <label class="field-label" for="newPort">Port</label>
        <select id="newPort">
          <option value="DAMMAM PORT">Dammam Port</option>
          <option value="JUBAIL COMMERCIAL PORT">Jubail Commercial Port</option>
          <option value="JEDDAH PORT">Jeddah Port</option>
          <option value="YANBU COMMERCIAL PORT">Yanbu Commercial Port</option>
          <option value="YANBU INDUSTRIAL PORT">Yanbu Industrial Port</option>
          <option value="KAP">KAP</option>
        </select>
      </div>
      <div>
        <label class="field-label" for="newVessel">Vessel</label>
        <input type="text" id="newVessel" placeholder="e.g. TAI KNIGHT">
      </div>
      <div>
        <label class="field-label" for="newReference">Reference (optional)</label>
        <input type="text" id="newReference" placeholder="Voyage no. / call ref">
      </div>
      <div>
        <label class="field-label" for="newCurrency">Currency</label>
        <input type="text" id="newCurrency" value="SAR">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn" onclick="createDocument()">Create PDA</button>
    </div>
  </div>

  <div class="panel" id="docsListPanel">
    <h2>Documents</h2>
    <p class="panel-sub" id="docsSub">Loading...</p>
    <div id="docsBody"></div>
  </div>

  <div class="panel" id="docDetail">
    <div class="detail-head">
      <div>
        <button class="btn ghost small" onclick="closeDocument()" style="margin-bottom:8px;">&larr; All documents</button>
        <div class="detail-title" id="detailTitle"></div>
        <div class="detail-sub" id="detailSub"></div>
      </div>
      <div class="detail-actions" id="detailActions"></div>
    </div>

    <div class="form-grid">
      <div>
        <label class="field-label" for="detRef">Reference</label>
        <input type="text" id="detRef" onchange="saveDocField('reference', this.value)">
      </div>
      <div>
        <label class="field-label" for="detCurrency">Currency</label>
        <input type="text" id="detCurrency" onchange="saveDocField('currency', this.value)">
      </div>
    </div>
    <div style="margin-bottom:14px;">
      <label class="field-label" for="detNotes">Notes</label>
      <textarea id="detNotes" onchange="saveDocField('notes', this.value)" placeholder="Anything worth noting on this account..."></textarea>
    </div>

    <table class="items-table" id="itemsTable">
      <thead><tr><th>Charge</th><th style="text-align:right;">Estimate</th><th style="text-align:right;" id="actualHeader">Actual</th><th style="text-align:right;" id="varianceHeader">Variance</th><th></th></tr></thead>
      <tbody id="itemsBody"></tbody>
    </table>
    <div class="form-actions" style="margin-top:10px;">
      <button class="btn ghost small" onclick="addLineItemRow()">+ Add charge</button>
    </div>
  </div>

  {% if role == 'admin' %}
  <div class="panel" id="templatesPanel">
    <h2>Port charge templates</h2>
    <p class="panel-sub">Default charges that pre-fill a new PDA for each port. Editing a template doesn't change documents already created from it.</p>
    <div id="templatesBody">Loading...</div>
  </div>

  <div class="panel" id="alertsPanel">
    <h2>Overdue ETA alerts</h2>
    <p class="panel-sub">When a vessel's ETA has passed with BLs still pending on the DO Tracker board, send a digest email listing them.</p>
    <div class="alert-row">
      <label class="switch-sm">
        <input type="checkbox" id="alertsEnabled" onchange="saveAlertSettings()">
        <span class="switch-track-sm"><span class="switch-knob-sm"></span></span>
      </label>
      <span style="font-size:13px;">Enable overdue-ETA alerts</span>
    </div>
    <label class="field-label" for="alertsRecipients">Recipient email(s)</label>
    <input type="text" id="alertsRecipients" placeholder="name@seapower.com, name2@seapower.com" onchange="saveAlertSettings()">
    <div class="mail-status" id="mailStatus"></div>
    <div class="form-actions">
      <button class="btn ghost small" onclick="checkOverdueNow()">Check now</button>
    </div>
  </div>
  {% endif %}

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

function fmtMoney(n) {
  n = Number(n || 0);
  return n.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function escHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

let currentDocId = null;
let docsCache = [];

async function loadDocuments() {
  const res = await fetch('/api/pda/documents');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const rows = await res.json();
  docsCache = rows;
  const sub = document.getElementById('docsSub');
  const body = document.getElementById('docsBody');
  if (!rows.length) {
    sub.textContent = 'No disbursement accounts yet.';
    body.innerHTML = '<div class="empty-note">Create one above once you have a port and vessel to work from.</div>';
    return;
  }
  sub.textContent = rows.length + ' document(s).';
  body.innerHTML = `
    <table class="doc-table">
      <thead><tr><th>Port</th><th>Vessel</th><th>Reference</th><th>Status</th><th style="text-align:right;">Estimate</th><th style="text-align:right;">Actual</th><th>Created</th><th></th></tr></thead>
      <tbody>
        ${rows.map(d => `
          <tr class="doc-row" onclick="openDocument(${d.id})">
            <td>${escHtml(d.port)}</td>
            <td>${escHtml(d.vessel)}</td>
            <td style="color:var(--muted);">${escHtml(d.reference) || '-'}</td>
            <td><span class="badge ${d.status}">${d.status === 'finalized' ? 'FDA' : (d.status === 'sent' ? 'Sent' : 'Draft')}</span></td>
            <td style="text-align:right;">${d.currency} ${fmtMoney(d.estimated_total)}</td>
            <td style="text-align:right;">${d.actual_total !== null ? d.currency + ' ' + fmtMoney(d.actual_total) : '-'}</td>
            <td style="color:var(--muted);font-size:12px;">${escHtml(d.created_by)}${d.created_at ? ' - ' + escHtml(d.created_at) : ''}</td>
            <td><button class="btn ghost danger small" onclick="event.stopPropagation(); deleteDocument(${d.id})">Delete</button></td>
          </tr>`).join('')}
      </tbody>
    </table>`;
}

async function createDocument() {
  const port = document.getElementById('newPort').value;
  const vessel = document.getElementById('newVessel').value.trim();
  const reference = document.getElementById('newReference').value.trim();
  const currency = document.getElementById('newCurrency').value.trim() || 'SAR';
  if (!vessel) { showToast('Enter a vessel name.', {error:true}); return; }
  const res = await fetch('/api/pda/documents', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({port, vessel, reference, currency})
  });
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not create that document.', {error:true}); return; }
  document.getElementById('newVessel').value = '';
  document.getElementById('newReference').value = '';
  await loadDocuments();
  openDocument(data.id);
}

async function deleteDocument(id) {
  if (!confirm('Delete this disbursement account? This cannot be undone.')) return;
  await fetch('/api/pda/documents/' + id, {method: 'DELETE'});
  if (currentDocId === id) closeDocument();
  await loadDocuments();
  showToast('Document deleted.');
}

let currentDoc = null;
let currentItems = [];

async function openDocument(id) {
  const res = await fetch('/api/pda/documents/' + id);
  if (!res.ok) { showToast('Could not load that document.', {error:true}); return; }
  const doc = await res.json();
  currentDocId = id;
  currentDoc = doc;
  currentItems = doc.items || [];
  document.getElementById('listPanel').style.display = 'none';
  document.getElementById('docsListPanel').style.display = 'none';
  document.getElementById('docDetail').style.display = 'block';
  renderDetail();
}

function closeDocument() {
  currentDocId = null;
  document.getElementById('docDetail').style.display = 'none';
  document.getElementById('listPanel').style.display = '';
  document.getElementById('docsListPanel').style.display = '';
}

function renderDetail() {
  const doc = currentDoc;
  const isFda = doc.status === 'finalized';
  document.getElementById('detailTitle').textContent = (isFda ? 'FDA' : 'PDA') + ' - ' + doc.port + ' / ' + doc.vessel;
  const statusLabel = isFda ? 'Finalized (FDA)' : (doc.status === 'sent' ? 'Sent' : 'Draft');
  document.getElementById('detailSub').textContent = statusLabel + ' - prepared by ' + (doc.created_by || '-') + (doc.created_at ? ' on ' + doc.created_at : '');
  document.getElementById('detRef').value = doc.reference || '';
  document.getElementById('detCurrency').value = doc.currency || 'SAR';
  document.getElementById('detNotes').value = doc.notes || '';

  const actions = [];
  if (doc.status === 'draft') {
    actions.push('<button class="btn ghost small" onclick="markSent()">Mark as sent</button>');
  }
  if (doc.status !== 'finalized') {
    actions.push('<button class="btn small" onclick="finalizeDoc()">Finalize as FDA</button>');
  }
  actions.push('<a class="btn ghost small" href="/api/pda/documents/' + doc.id + '/pdf">Export PDF</a>');
  document.getElementById('detailActions').innerHTML = actions.join('');

  document.getElementById('actualHeader').style.display = '';
  document.getElementById('varianceHeader').style.display = isFda ? '' : 'none';
  renderItems();
}

function renderItems() {
  const isFda = currentDoc.status === 'finalized';
  const body = document.getElementById('itemsBody');
  let estTotal = 0, actTotal = 0;
  const rows = currentItems.map(item => {
    const est = Number(item.estimated_amount || 0);
    estTotal += est;
    const hasActual = item.actual_amount !== null && item.actual_amount !== undefined;
    const act = hasActual ? Number(item.actual_amount) : null;
    if (isFda) actTotal += (act !== null ? act : est);
    const variance = (act !== null) ? (act - est) : null;
    const varianceHtml = (isFda && variance !== null)
      ? `<span class="${variance > 0 ? 'variance-pos' : (variance < 0 ? 'variance-neg' : '')}">${variance > 0 ? '+' : ''}${fmtMoney(variance)}</span>`
      : '';
    return `<tr>
      <td><input type="text" value="${escHtml(item.name)}" onchange="updateLineItem(${item.id}, 'name', this.value)"></td>
      <td><input type="number" step="0.01" value="${est}" onchange="updateLineItem(${item.id}, 'estimated_amount', this.value)"></td>
      <td><input type="number" step="0.01" value="${act !== null ? act : ''}" placeholder="-" onchange="updateLineItem(${item.id}, 'actual_amount', this.value)"></td>
      <td style="text-align:right;">${varianceHtml}</td>
      <td><button class="row-remove" title="Remove charge" onclick="deleteLineItem(${item.id})">&times;</button></td>
    </tr>`;
  }).join('');
  const varianceTotal = isFda ? (actTotal - estTotal) : null;
  const totalRow = `<tr class="total-row">
    <td>Total (${currentDoc.currency || 'SAR'})</td>
    <td style="text-align:right;">${fmtMoney(estTotal)}</td>
    <td style="text-align:right;">${isFda ? fmtMoney(actTotal) : ''}</td>
    <td style="text-align:right;">${isFda ? `<span class="${varianceTotal > 0 ? 'variance-pos' : (varianceTotal < 0 ? 'variance-neg' : '')}">${varianceTotal > 0 ? '+' : ''}${fmtMoney(varianceTotal)}</span>` : ''}</td>
    <td></td>
  </tr>`;
  body.innerHTML = rows + totalRow;
}

async function saveDocField(field, value) {
  if (!currentDocId) return;
  await fetch('/api/pda/documents/' + currentDocId, {
    method: 'PUT', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({[field]: value})
  });
  currentDoc[field] = value;
  if (field === 'currency') renderDetail();
}

async function addLineItemRow() {
  const res = await fetch('/api/pda/documents/' + currentDocId + '/line-items', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({name: 'New charge', estimated_amount: 0})
  });
  const item = await res.json();
  if (!res.ok || item.error) { showToast(item.error || 'Could not add that charge.', {error:true}); return; }
  currentItems.push(item);
  renderItems();
}

async function updateLineItem(itemId, field, value) {
  const payload = {};
  if (field === 'name') {
    if (!value.trim()) { showToast('Charge name can\\'t be empty.', {error:true}); renderItems(); return; }
    payload.name = value;
  } else {
    payload[field] = value === '' ? null : value;
  }
  const res = await fetch('/api/pda/line-items/' + itemId, {
    method: 'PUT', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload)
  });
  if (!res.ok) { showToast('Could not save that change.', {error:true}); return; }
  const item = currentItems.find(i => i.id === itemId);
  if (item) item[field] = payload[field] === null ? null : (field === 'name' ? value : Number(value));
  renderItems();
}

async function deleteLineItem(itemId) {
  await fetch('/api/pda/line-items/' + itemId, {method: 'DELETE'});
  currentItems = currentItems.filter(i => i.id !== itemId);
  renderItems();
}

async function markSent() {
  const res = await fetch('/api/pda/documents/' + currentDocId + '/mark-sent', {method: 'POST'});
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not update status.', {error:true}); return; }
  currentDoc.status = 'sent';
  renderDetail();
  loadDocuments();
  showToast('Marked as sent.');
}

async function finalizeDoc() {
  if (!confirm('Finalize this as an FDA? Any charge without an actual amount yet will use its estimate. This can still be edited afterward, but the document moves out of draft/sent.')) return;
  const res = await fetch('/api/pda/documents/' + currentDocId + '/finalize', {method: 'POST'});
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not finalize.', {error:true}); return; }
  await openDocument(currentDocId);
  loadDocuments();
  showToast('Finalized as FDA.');
}

{% if role == 'admin' %}
let templatesCache = {};
const TEMPLATE_PORTS = ['DAMMAM PORT', 'JUBAIL COMMERCIAL PORT', 'JEDDAH PORT', 'YANBU COMMERCIAL PORT', 'YANBU INDUSTRIAL PORT', 'KAP'];

async function loadTemplates() {
  const res = await fetch('/api/pda/templates');
  templatesCache = await res.json();
  renderTemplates();
}

function renderTemplates() {
  const body = document.getElementById('templatesBody');
  body.innerHTML = TEMPLATE_PORTS.map(port => {
    const items = templatesCache[port] || [];
    const rows = items.map(t => `
      <tr>
        <td><input type="text" value="${escHtml(t.name)}" onchange="updateTemplateItem(${t.id}, 'name', this.value)"></td>
        <td><input type="number" step="0.01" value="${t.default_amount}" onchange="updateTemplateItem(${t.id}, 'default_amount', this.value)"></td>
        <td><button class="row-remove" title="Remove" onclick="deleteTemplateItem(${t.id})">&times;</button></td>
      </tr>`).join('');
    return `<div class="tmpl-port-group">
      <h3>${port.replace(/\\w\\S*/g, w => w.charAt(0) + w.slice(1).toLowerCase())}</h3>
      <table class="items-table"><tbody>${rows || '<tr><td colspan="3" style="color:var(--muted);">No charges yet.</td></tr>'}</tbody></table>
      <div class="tmpl-add-row">
        <input type="text" id="tmplName_${port.replace(/[^A-Za-z0-9]/g, '')}" placeholder="Charge name">
        <input type="number" step="0.01" id="tmplAmount_${port.replace(/[^A-Za-z0-9]/g, '')}" placeholder="0.00">
        <button class="btn ghost small" onclick="addTemplateItem('${port.replace(/'/g, "\\\\'")}')">Add</button>
      </div>
    </div>`;
  }).join('');
}

async function addTemplateItem(port) {
  const key = port.replace(/[^A-Za-z0-9]/g, '');
  const nameEl = document.getElementById('tmplName_' + key);
  const amountEl = document.getElementById('tmplAmount_' + key);
  const name = nameEl.value.trim();
  if (!name) { showToast('Enter a charge name.', {error:true}); return; }
  const res = await fetch('/api/pda/templates', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({port, name, default_amount: amountEl.value || 0})
  });
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not add that charge.', {error:true}); return; }
  nameEl.value = ''; amountEl.value = '';
  await loadTemplates();
}

async function updateTemplateItem(id, field, value) {
  await fetch('/api/pda/templates/' + id, {
    method: 'PUT', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({[field]: value})
  });
  await loadTemplates();
}

async function deleteTemplateItem(id) {
  await fetch('/api/pda/templates/' + id, {method: 'DELETE'});
  await loadTemplates();
}

async function loadAlertSettings() {
  const res = await fetch('/api/settings/alerts');
  const data = await res.json();
  document.getElementById('alertsEnabled').checked = !!data.enabled;
  document.getElementById('alertsRecipients').value = data.recipients || '';
  const status = document.getElementById('mailStatus');
  status.textContent = data.mail_configured
    ? 'Email sending is configured.'
    : 'Email isn\\'t configured on the server yet (SMTP_HOST / SMTP_USER / SMTP_PASSWORD) - alerts will be tracked but not sent until that\\'s set up.';
  status.className = 'mail-status ' + (data.mail_configured ? 'ok' : 'warn');
}

async function saveAlertSettings() {
  const enabled = document.getElementById('alertsEnabled').checked;
  const recipients = document.getElementById('alertsRecipients').value.trim();
  await fetch('/api/settings/alerts', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({enabled, recipients})
  });
}

async function checkOverdueNow() {
  const res = await fetch('/api/alerts/check-overdue', {method: 'POST'});
  const data = await res.json();
  if (!res.ok) { showToast(data.error || 'Could not run the check.', {error:true}); return; }
  if (data.overdue_count === 0) { showToast('Nothing overdue right now.'); return; }
  if (data.sent) { showToast(data.overdue_count + ' vessel(s) overdue - alert emailed.'); return; }
  showToast(data.note || data.error || (data.overdue_count + ' vessel(s) overdue, but the alert could not be sent.'), {error:true, duration: 6000});
}

loadTemplates();
loadAlertSettings();
{% endif %}

loadDocuments();
</script>
</body></html>
"""

SOF_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Statement of Facts</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<style>
  :root {
    --bg: #f2f4f7; --card: #ffffff; --text: #1c2b3a; --muted: #7a8794; --border: #e6e9ed;
    --navy: #123a56; --navy-deep: #0b2740; --navy-light: #1f5c85; --gold: #c9a227; --gold-light: #e0bd53;
    --danger: #d1483f; --danger-bg: #fbeceb; --ok: #1c8a5a; --ok-bg: #e7f5ee;
    --warn: #8a6d1f; --warn-bg: rgba(212,160,23,0.16);
    --shadow-sm: 0 1px 2px rgba(18,58,86,0.05); --shadow-md: 0 10px 30px rgba(18,58,86,0.10);
    color-scheme: light;
  }
  :root[data-theme="dark"] {
    --bg: #131a23; --card: #1a232f; --text: #e9eef3; --muted: #93a1b1; --border: #29323f;
    --navy: #3f86ba; --navy-deep: #274a67; --navy-light: #5aa2d1; --gold: #e3bb4c; --gold-light: #f0cf72;
    --danger: #e2685f; --danger-bg: #3a2220; --ok: #3ecb8e; --ok-bg: #163329;
    --warn: var(--gold-light); --warn-bg: rgba(227,187,76,0.16);
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.25); --shadow-md: 0 10px 30px rgba(0,0,0,0.35);
    color-scheme: dark;
  }
  * { box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, Arial, sans-serif;
    background: var(--bg); color: var(--text); margin: 0; padding: 0 16px 40px;
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
  .page-head p { color: var(--muted); margin: 0; font-size: 13.5px; max-width: 680px; }

  .panel { background: var(--card); border: 1px solid var(--border); border-radius: 18px; box-shadow: var(--shadow-sm); padding: 18px 20px; margin-bottom: 18px; }
  .panel h2 { font-size: 15px; margin: 0 0 2px; }
  .panel .panel-sub { font-size: 12px; color: var(--muted); margin: 0 0 14px; }

  label.field-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .03em; color: var(--muted); display: block; margin-bottom: 5px; }
  input[type=text], input[type=number], input[type=email], select, textarea {
    width: 100%; padding: 9px 11px; border: 1px solid var(--border); border-radius: 9px;
    font-size: 13.5px; font-family: inherit; background: var(--bg); color: var(--text);
  }
  input:focus, select:focus, textarea:focus { outline: none; border-color: var(--navy-light); }
  textarea { resize: vertical; min-height: 56px; }

  .form-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 12px; margin-bottom: 12px; }
  .form-actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }

  .btn { background: var(--navy); color: #fff; border: none; border-radius: 999px; padding: 9px 17px; font-size: 12.5px; font-weight: 600; cursor: pointer; transition: background .15s ease, transform .08s ease; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
  .btn:hover { background: var(--navy-light); }
  .btn:active { transform: scale(.97); }
  .btn:disabled { opacity: .5; cursor: default; }
  .btn.ghost { background: none; color: var(--navy); border: 1px solid var(--border); }
  :root[data-theme="dark"] .btn.ghost { color: var(--navy-light); }
  .btn.ghost:hover { background: var(--border); }
  .btn.ghost.danger { color: var(--danger); }
  .btn.ghost.danger:hover { background: var(--danger-bg); }
  .btn.small { padding: 5px 11px; font-size: 11.5px; }

  table.doc-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  table.doc-table th { text-align: left; font-size: 10.5px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); padding: 8px 10px; border-bottom: 1px solid var(--border); }
  table.doc-table td { padding: 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }
  table.doc-table tr:last-child td { border-bottom: none; }
  table.doc-table tr.doc-row { cursor: pointer; }
  table.doc-table tr.doc-row:hover td { background: color-mix(in srgb, var(--navy-light) 5%, transparent); }

  .empty-note { color: var(--muted); font-size: 13px; padding: 10px 2px; }

  #docDetail { display: none; }
  .detail-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 14px; }
  .detail-title { font-size: 16px; font-weight: 700; margin: 0 0 2px; }
  .detail-sub { font-size: 12px; color: var(--muted); }
  .detail-actions { display: flex; gap: 8px; flex-wrap: wrap; }

  .field-section { margin-top: 20px; margin-bottom: 6px; }
  .field-section h3 { font-size: 13px; margin: 0 0 10px; color: var(--navy-light); text-transform: uppercase; letter-spacing: .04em; }
  .field-section:first-of-type { margin-top: 0; }

  .timeline-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px 16px; margin-bottom: 4px; }
  @media (max-width: 560px) { .timeline-grid { grid-template-columns: 1fr; } }

  #toastHost { position: fixed; bottom: 20px; right: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1000; pointer-events: none; max-width: min(340px, calc(100vw - 40px)); }
  #toastHost .toast { pointer-events: auto; }
  .toast { background: var(--navy-deep); color: #fff; padding: 11px 16px; border-radius: 12px; font-size: 13px; display: flex; align-items: center; gap: 14px; box-shadow: 0 10px 30px rgba(0,0,0,0.25); animation: toast-in .18s ease-out; max-width: 340px; }
  .toast.error { background: var(--danger); }
  .toast.fading { animation: toast-out .2s ease-in forwards; }
  @keyframes toast-in { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes toast-out { to { opacity: 0; transform: translateY(8px); } }
</style>
</head>
<body>
  <div class="topbar">
    <a href="/" class="brand">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag">Statement of Facts</span>
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
      <a href="/pda">Disbursement Accounts</a>
      <a href="/do-tracker">DO Tracker</a>
      {% if role == 'admin' %}<a href="/users">Manage Users</a>{% endif %}
      <span style="padding:6px 4px;">Signed in as <b>{{ username }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="page-head">
    <div class="eyebrow">Compass</div>
    <h1>Statement of Facts</h1>
    <p>Record a vessel call's event timeline field-by-field, the same way the paper SOF is filled in over the course of the call, then export it as a finished document once the call is complete.</p>
  </div>

  <div id="toastHost"></div>

  <div class="panel" id="listPanel">
    <h2>New Statement of Facts</h2>
    <p class="panel-sub">Start with what you know now - vessel and port are required, everything else (including the whole timeline) can be filled in as the call progresses.</p>
    <div class="form-grid">
      <div>
        <label class="field-label" for="newVessel">Vessel</label>
        <input type="text" id="newVessel" placeholder="e.g. M.V. RICH GLORY">
      </div>
      <div>
        <label class="field-label" for="newVoyage">Voyage</label>
        <input type="text" id="newVoyage" placeholder="e.g. MAC015">
      </div>
      <div>
        <label class="field-label" for="newPort">Port</label>
        <select id="newPort">
          <option value="DAMMAM PORT">Dammam Port</option>
          <option value="JUBAIL COMMERCIAL PORT">Jubail Commercial Port</option>
          <option value="JEDDAH PORT">Jeddah Port</option>
          <option value="YANBU COMMERCIAL PORT">Yanbu Commercial Port</option>
          <option value="YANBU INDUSTRIAL PORT">Yanbu Industrial Port</option>
          <option value="KAP">KAP</option>
        </select>
      </div>
      <div>
        <label class="field-label" for="newBerth">Berth</label>
        <input type="text" id="newBerth" placeholder="e.g. Berth 22">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn" onclick="createDocument()">Create SOF</button>
    </div>
  </div>

  <div class="panel" id="docsListPanel">
    <h2>Documents</h2>
    <p class="panel-sub" id="docsSub">Loading...</p>
    <div id="docsBody"></div>
  </div>

  <div class="panel" id="docDetail">
    <div class="detail-head">
      <div>
        <button class="btn ghost small" onclick="closeDocument()" style="margin-bottom:8px;">&larr; All documents</button>
        <div class="detail-title" id="detailTitle"></div>
        <div class="detail-sub" id="detailSub"></div>
      </div>
      <div class="detail-actions" id="detailActions"></div>
    </div>

    <div class="field-section">
      <h3>Vessel Particulars</h3>
      <div class="form-grid" id="particularsGrid"></div>
    </div>

    <div class="field-section">
      <h3>Event Timeline</h3>
      <div class="timeline-grid" id="timelineGrid"></div>
    </div>

    <div class="field-section">
      <h3>Remaining On Board</h3>
      <div class="form-grid" id="robGrid"></div>
    </div>

    <div class="field-section">
      <h3>Draft</h3>
      <div class="form-grid" id="draftGrid"></div>
    </div>

    <div class="field-section">
      <h3>Delays / Remarks</h3>
      <textarea id="detDelays" onchange="saveField('delays_remarks', this.value)" placeholder="Any delays worth recording, with reasons and durations..."></textarea>
    </div>

    <div class="field-section">
      <h3>Master's Remarks</h3>
      <textarea id="detMastersRemarks" onchange="saveField('masters_remarks', this.value)" placeholder="Remarks for the master's signature section..."></textarea>
    </div>
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

function escHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const SOF_PARTICULARS = [
  ['vessel', 'Vessel', ''], ['voyage', 'Voyage', ''],
  ['berth', 'Berth', ''], ['owners', 'Owners', ''], ['charterer', 'Charterer', '']
];
const SOF_TIMELINE_FIELDS = [
  ['end_of_sea_passage', 'End of Sea Passage', 'DD.MM.YY AT HHMM HRS'],
  ['customs_cleared', 'Customs Cleared', 'DD.MM.YY AT HHMM HRS'],
  ['nor_tendered', 'NOR Tendered', 'DD.MM.YY AT HHMM HRS'],
  ['commenced_discharge', 'Commenced Discharge', 'DD.MM.YY AT HHMM HRS'],
  ['nor_accepted', 'NOR Accepted', 'DD.MM.YY AT HHMM HRS'],
  ['completed_discharge', 'Completed Discharge', 'DD.MM.YY AT HHMM HRS'],
  ['anchored', 'Anchored', 'DD.MM.YY AT HHMM HRS'],
  ['documents_on_board', 'Documents on Board', 'DD.MM.YY AT HHMM HRS'],
  ['left_anchorage', 'Left Anchorage', 'DD.MM.YY AT HHMM HRS'],
  ['clearance_delivered', 'Clearance Delivered', 'DD.MM.YY AT HHMM HRS'],
  ['pilot_boarded_arrival', 'Pilot Boarded (Arrival)', 'DD.MM.YY AT HHMM HRS'],
  ['pilot_boarded_departure', 'Pilot Boarded (Departure)', 'DD.MM.YY AT HHMM HRS'],
  ['first_line_to_shore', 'First Line to Shore', 'DD.MM.YY AT HHMM HRS'],
  ['left_berth', 'Left Berth', 'DD.MM.YY AT HHMM HRS'],
  ['berthed_all_fast', 'Berthed (All Fast)', 'DD.MM.YY AT HHMM HRS'],
  ['cargo_discharge_mtons', 'Cargo Discharged (M.Tons)', 'e.g. 12,500.00']
];
const SOF_ROB_FIELDS = [
  ['rob_arrival_ifo', 'ROB Arrival - IFO', 'e.g. 180.5 MT'], ['rob_arrival_mdo', 'ROB Arrival - MDO', 'e.g. 45.0 MT'],
  ['rob_arrival_lubs', 'ROB Arrival - LUBS', 'e.g. 8.2 MT'], ['rob_arrival_fwater', 'ROB Arrival - F.Water', 'e.g. 60.0 MT'],
  ['rob_departure_ifo', 'ROB Departure - IFO', 'e.g. 170.0 MT'], ['rob_departure_mdo', 'ROB Departure - MDO', 'e.g. 43.0 MT'],
  ['rob_departure_lubs', 'ROB Departure - LUBS', 'e.g. 8.0 MT'], ['rob_departure_fwater', 'ROB Departure - F.Water', 'e.g. 55.0 MT']
];
const SOF_DRAFT_FIELDS = [
  ['arrival_draft_fwd', 'Arrival Draft - FWD', 'e.g. 8.20 M'], ['arrival_draft_aft', 'Arrival Draft - AFT', 'e.g. 9.10 M'],
  ['departure_draft_fwd', 'Departure Draft - FWD', 'e.g. 7.50 M'], ['departure_draft_aft', 'Departure Draft - AFT', 'e.g. 8.40 M']
];

function fieldBlock(col, label, placeholder) {
  return `<div>
    <label class="field-label" for="f_${col}">${label}</label>
    <input type="text" id="f_${col}" onchange="saveField('${col}', this.value)" placeholder="${placeholder || ''}">
  </div>`;
}

function portFieldBlock() {
  return `<div>
    <label class="field-label" for="f_port">Port</label>
    <select id="f_port" onchange="saveField('port', this.value)">
      <option value="DAMMAM PORT">Dammam Port</option>
      <option value="JUBAIL COMMERCIAL PORT">Jubail Commercial Port</option>
      <option value="JEDDAH PORT">Jeddah Port</option>
      <option value="YANBU COMMERCIAL PORT">Yanbu Commercial Port</option>
      <option value="YANBU INDUSTRIAL PORT">Yanbu Industrial Port</option>
      <option value="KAP">KAP</option>
    </select>
  </div>`;
}

document.getElementById('particularsGrid').innerHTML =
  fieldBlock('vessel', 'Vessel') + fieldBlock('voyage', 'Voyage') + portFieldBlock() +
  fieldBlock('berth', 'Berth') + fieldBlock('owners', 'Owners') + fieldBlock('charterer', 'Charterer');
document.getElementById('timelineGrid').innerHTML = SOF_TIMELINE_FIELDS.map(([c, l, p]) => fieldBlock(c, l, p)).join('');
document.getElementById('robGrid').innerHTML = SOF_ROB_FIELDS.map(([c, l, p]) => fieldBlock(c, l, p)).join('');
document.getElementById('draftGrid').innerHTML = SOF_DRAFT_FIELDS.map(([c, l, p]) => fieldBlock(c, l, p)).join('');

const SOF_ALL_FIELDS = ['vessel', 'voyage', 'port', 'berth', 'owners', 'charterer']
  .concat(SOF_TIMELINE_FIELDS.map(f => f[0]))
  .concat(SOF_ROB_FIELDS.map(f => f[0]))
  .concat(SOF_DRAFT_FIELDS.map(f => f[0]));

let currentDocId = null;
let currentDoc = null;

async function loadDocuments() {
  const res = await fetch('/api/sof/documents');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const rows = await res.json();
  const sub = document.getElementById('docsSub');
  const body = document.getElementById('docsBody');
  if (!rows.length) {
    sub.textContent = 'No Statements of Facts yet.';
    body.innerHTML = '<div class="empty-note">Create one above once you have a vessel and port to work from.</div>';
    return;
  }
  sub.textContent = rows.length + ' document(s).';
  body.innerHTML = `
    <table class="doc-table">
      <thead><tr><th>Vessel</th><th>Voyage</th><th>Port</th><th>Berth</th><th>Created</th><th></th></tr></thead>
      <tbody>
        ${rows.map(d => `
          <tr class="doc-row" onclick="openDocument(${d.id})">
            <td>${escHtml(d.vessel)}</td>
            <td style="color:var(--muted);">${escHtml(d.voyage) || '-'}</td>
            <td>${escHtml(d.port)}</td>
            <td style="color:var(--muted);">${escHtml(d.berth) || '-'}</td>
            <td style="color:var(--muted);font-size:12px;">${escHtml(d.created_by)}${d.created_at ? ' - ' + escHtml(d.created_at) : ''}</td>
            <td><button class="btn ghost danger small" onclick="event.stopPropagation(); deleteDocument(${d.id})">Delete</button></td>
          </tr>`).join('')}
      </tbody>
    </table>`;
}

async function createDocument() {
  const vessel = document.getElementById('newVessel').value.trim();
  const voyage = document.getElementById('newVoyage').value.trim();
  const port = document.getElementById('newPort').value;
  const berth = document.getElementById('newBerth').value.trim();
  if (!vessel) { showToast('Enter a vessel name.', {error:true}); return; }
  const res = await fetch('/api/sof/documents', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({vessel, voyage, port, berth})
  });
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not create that document.', {error:true}); return; }
  document.getElementById('newVessel').value = '';
  document.getElementById('newVoyage').value = '';
  document.getElementById('newBerth').value = '';
  await loadDocuments();
  openDocument(data.id);
}

async function deleteDocument(id) {
  if (!confirm('Delete this Statement of Facts? This cannot be undone.')) return;
  await fetch('/api/sof/documents/' + id, {method: 'DELETE'});
  if (currentDocId === id) closeDocument();
  await loadDocuments();
  showToast('Document deleted.');
}

async function openDocument(id) {
  const res = await fetch('/api/sof/documents/' + id);
  if (!res.ok) { showToast('Could not load that document.', {error:true}); return; }
  const doc = await res.json();
  currentDocId = id;
  currentDoc = doc;
  document.getElementById('listPanel').style.display = 'none';
  document.getElementById('docsListPanel').style.display = 'none';
  document.getElementById('docDetail').style.display = 'block';
  renderDetail();
}

function closeDocument() {
  currentDocId = null;
  document.getElementById('docDetail').style.display = 'none';
  document.getElementById('listPanel').style.display = '';
  document.getElementById('docsListPanel').style.display = '';
}

function renderDetail() {
  const doc = currentDoc;
  document.getElementById('detailTitle').textContent = 'SOF - ' + (doc.vessel || '-') + (doc.port ? ' / ' + doc.port : '');
  document.getElementById('detailSub').textContent =
    'Prepared by ' + (doc.created_by || '-') + (doc.created_at ? ' on ' + doc.created_at : '') +
    (doc.updated_at && doc.updated_at !== doc.created_at ? ' - last updated ' + doc.updated_at : '');
  document.getElementById('detailActions').innerHTML =
    '<a class="btn ghost small" href="/api/sof/documents/' + doc.id + '/pdf">Export PDF</a>' +
    '<button class="btn ghost danger small" onclick="deleteDocument(' + doc.id + ')">Delete</button>';

  SOF_ALL_FIELDS.forEach(c => {
    const el = document.getElementById('f_' + c);
    if (el) el.value = doc[c] || '';
  });
  document.getElementById('detDelays').value = doc.delays_remarks || '';
  document.getElementById('detMastersRemarks').value = doc.masters_remarks || '';
}

async function saveField(col, value) {
  if (!currentDocId) return;
  const res = await fetch('/api/sof/documents/' + currentDocId, {
    method: 'PUT', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({[col]: value})
  });
  const data = await res.json();
  if (!res.ok || data.error) { showToast((data && data.error) || 'Could not save that change.', {error:true}); return; }
  currentDoc[col] = value;
  if (col === 'vessel' || col === 'port' || col === 'voyage' || col === 'berth') loadDocuments();
}

loadDocuments();
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
    --topbar-h: 64px;
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
  /* Native <select> elements largely ignore/flatten backdrop-filter and
     translucent backgrounds (esp. Chromium), painting an opaque native
     control instead. So the frosted-glass look lives on this wrapper div
     (which DOES get backdrop-filter applied by the browser), and the
     select itself sits on top with a transparent background so the glass
     behind it shows through. */
  .glass-select-wrap {
    position: relative; display: inline-block; width: 100%;
    border-radius: 10px;
    overflow: hidden;
    border: 1.5px solid color-mix(in srgb, var(--navy-light) 70%, transparent);
    background:
      linear-gradient(160deg, color-mix(in srgb, #ffffff 65%, var(--navy-light) 10%) 0%, color-mix(in srgb, var(--navy-light) 22%, white) 55%, color-mix(in srgb, #ffffff 70%, var(--navy-light) 14%) 100%);
    backdrop-filter: blur(14px) saturate(200%);
    -webkit-backdrop-filter: blur(14px) saturate(200%);
    box-shadow:
      0 10px 22px -6px color-mix(in srgb, var(--navy) 42%, transparent),
      0 2px 6px color-mix(in srgb, var(--navy) 22%, transparent),
      inset 0 1.5px 0 rgba(255,255,255,0.95),
      inset 0 -1.5px 0 color-mix(in srgb, var(--navy) 22%, transparent);
    transition: border-color .15s ease, background .15s ease, box-shadow .15s ease, transform .15s ease;
  }
  .glass-select-wrap::before {
    content: ""; position: absolute; inset: 0; z-index: -1; pointer-events: none;
    background: linear-gradient(115deg, rgba(255,255,255,0.9) 0%, rgba(255,255,255,0.05) 30%, rgba(255,255,255,0.05) 60%, rgba(255,255,255,0.6) 100%);
    mix-blend-mode: overlay;
  }
  .glass-select-wrap:hover {
    border-color: color-mix(in srgb, var(--navy-light) 85%, transparent);
    transform: translateY(-1px);
    box-shadow:
      0 14px 26px -6px color-mix(in srgb, var(--navy) 50%, transparent),
      0 3px 8px color-mix(in srgb, var(--navy) 26%, transparent),
      inset 0 1.5px 0 rgba(255,255,255,0.95),
      inset 0 -1.5px 0 color-mix(in srgb, var(--navy) 24%, transparent);
  }
  .glass-select-wrap:focus-within {
    border-color: var(--navy-light);
    box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 30%, transparent),
      0 14px 30px -8px color-mix(in srgb, var(--navy) 50%, transparent),
      inset 0 1.5px 0 rgba(255,255,255,1);
  }
  :root[data-theme="dark"] .glass-select-wrap {
    border-color: color-mix(in srgb, #ffffff 30%, transparent);
    background:
      linear-gradient(165deg, rgba(255,255,255,0.22) 0%, rgba(255,255,255,0.02) 40%, rgba(255,255,255,0.10) 100%),
      linear-gradient(135deg, color-mix(in srgb, var(--navy-light) 42%, var(--card)), color-mix(in srgb, var(--card) 60%, black) 80%);
    box-shadow:
      0 14px 32px -6px rgba(0,0,0,0.6),
      0 3px 10px rgba(0,0,0,0.35),
      inset 0 1.5px 0 rgba(255,255,255,0.22),
      inset 0 -1.5px 0 rgba(0,0,0,0.35);
  }
  :root[data-theme="dark"] .glass-select-wrap::before {
    background: linear-gradient(115deg, rgba(255,255,255,0.22) 0%, rgba(255,255,255,0) 24%, rgba(255,255,255,0) 64%, rgba(255,255,255,0.12) 100%);
  }
  :root[data-theme="dark"] .glass-select-wrap:hover {
    border-color: color-mix(in srgb, #ffffff 42%, transparent);
    transform: translateY(-1px);
    box-shadow:
      0 18px 36px -6px rgba(0,0,0,0.65),
      0 4px 12px rgba(0,0,0,0.4),
      inset 0 1.5px 0 rgba(255,255,255,0.28),
      inset 0 -1.5px 0 rgba(0,0,0,0.4);
  }
  :root[data-theme="dark"] .glass-select-wrap:focus-within {
    border-color: var(--navy-light);
    box-shadow: 0 0 0 3px color-mix(in srgb, var(--navy-light) 35%, transparent),
      0 14px 32px -6px rgba(0,0,0,0.6),
      inset 0 1.5px 0 rgba(255,255,255,0.26);
  }
  select.nice-select {
    appearance: none; -webkit-appearance: none; -moz-appearance: none;
    border: none; border-radius: 10px; padding: 10px 34px 10px 12px;
    font-size: 14px; font-family: inherit; width: 100%;
    background-color: transparent;
    color: var(--text); cursor: pointer;
    background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%237a8794' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 9l6 6 6-6'/></svg>");
    background-repeat: no-repeat; background-position: right 12px center; background-size: 14px;
  }
  select.nice-select:focus { outline: none; }
  :root[data-theme="dark"] select.nice-select {
    background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%2393a1b1' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><path d='M6 9l6 6 6-6'/></svg>");
  }
  select.nice-select option { background: var(--card); color: var(--text); }

  /* ---------- Custom glass dropdown (replaces the native <select> popup,
     which cannot be styled in any browser) for portField/jumpSelect/
     operatorFilter. The real <select> stays in the DOM (visually hidden)
     so all existing .value reads/writes and onchange handlers keep
     working unchanged - see initGlassSelects() in the script below. ---------- */
  select.nice-select.cs-native-hidden {
    opacity: 0; pointer-events: none;
  }
  .cs-trigger {
    position: absolute; inset: 0; z-index: 1;
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    padding: 10px 12px; font-size: 14px; font-family: inherit;
    color: var(--text); cursor: pointer; user-select: none; outline: none;
  }
  .cs-trigger-label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .cs-trigger .cs-chevron {
    width: 14px; height: 14px; flex-shrink: 0; color: #7a8794;
    transition: transform .15s ease;
  }
  .cs-trigger.cs-open .cs-chevron { transform: rotate(180deg); }
  :root[data-theme="dark"] .cs-trigger .cs-chevron { color: #93a1b1; }
  .cs-panel {
    position: fixed; z-index: 3000; min-width: 160px;
    border-radius: 12px;
    border: 1.5px solid color-mix(in srgb, var(--navy-light) 70%, transparent);
    background:
      linear-gradient(160deg, color-mix(in srgb, #ffffff 90%, var(--navy-light) 6%) 0%, color-mix(in srgb, var(--navy-light) 14%, white) 60%, color-mix(in srgb, #ffffff 92%, var(--navy-light) 8%) 100%);
    backdrop-filter: blur(18px) saturate(200%);
    -webkit-backdrop-filter: blur(18px) saturate(200%);
    box-shadow:
      0 18px 40px -10px color-mix(in srgb, var(--navy) 45%, transparent),
      0 4px 14px color-mix(in srgb, var(--navy) 25%, transparent),
      inset 0 1.5px 0 rgba(255,255,255,0.9);
    padding: 6px; max-height: 280px; overflow-y: auto;
    display: none;
    scrollbar-width: thin; scrollbar-color: color-mix(in srgb, var(--navy-light) 55%, transparent) transparent;
  }
  .cs-panel.cs-open { display: block; }
  .cs-panel::-webkit-scrollbar { width: 8px; }
  .cs-panel::-webkit-scrollbar-thumb { background: color-mix(in srgb, var(--navy-light) 55%, transparent); border-radius: 8px; }
  :root[data-theme="dark"] .cs-panel {
    border-color: color-mix(in srgb, #ffffff 30%, transparent);
    background:
      linear-gradient(165deg, rgba(255,255,255,0.16) 0%, rgba(255,255,255,0.02) 40%, rgba(255,255,255,0.09) 100%),
      linear-gradient(135deg, color-mix(in srgb, var(--navy-light) 38%, var(--card)), color-mix(in srgb, var(--card) 65%, black) 85%);
    box-shadow:
      0 20px 44px -8px rgba(0,0,0,0.65),
      0 4px 14px rgba(0,0,0,0.4),
      inset 0 1.5px 0 rgba(255,255,255,0.2);
  }
  .cs-option {
    display: flex; align-items: center; justify-content: space-between; gap: 8px;
    padding: 9px 10px; border-radius: 8px; font-size: 14px; color: var(--text);
    cursor: pointer;
  }
  .cs-option:hover, .cs-option.cs-highlight { background: color-mix(in srgb, var(--navy-light) 22%, transparent); }
  :root[data-theme="dark"] .cs-option:hover, :root[data-theme="dark"] .cs-option.cs-highlight { background: rgba(255,255,255,0.12); }
  .cs-option.cs-selected { font-weight: 600; }
  .cs-option .cs-check { width: 14px; height: 14px; flex-shrink: 0; opacity: 0; color: var(--navy-light); }
  .cs-option.cs-selected .cs-check { opacity: 1; }
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

  /* Toolbar button hierarchy: the plain navy `button` above is for
     primary/constructive actions (Add to board, Mark buttons). Neutral and
     destructive toolbar actions get their own consistent outline styles
     instead of one-off inline styles, so every "remove"-type control in the
     app (toolbar, bulk bar, group header, row) looks the same. */
  .btn-neutral { background: none; color: var(--text); border: 1px solid var(--border); }
  .btn-neutral:hover { background: var(--border); }
  .btn-danger {
    background: none; color: var(--danger);
    border: 1px solid color-mix(in srgb, var(--danger) 45%, var(--border));
  }
  .btn-danger:hover { background: var(--danger-bg); }

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
    padding: 14px 16px; font-size: 12px; color: var(--muted); flex: 1 1 200px; max-width: 260px; min-width: 130px;
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

  /* Port landing nav - sits above #groups. In the "All ports" state
     (selectedPortTab === '') it's a clickable card grid, one card per
     port, each a drill-in into that port's vessels; once a port is
     selected it becomes a small breadcrumb/back control instead, and
     #groups shows just that port. Built fresh in render() from whatever
     ports currently exist (same pattern as jumpSelect/operatorFilter),
     so it never needs its own data fetch. */
  .port-card-grid {
    display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 14px; margin-bottom: 14px;
  }
  .port-card {
    display: flex; flex-direction: column; align-items: flex-start; gap: 8px;
    text-align: left; font-family: inherit; cursor: pointer;
    background: var(--card); color: var(--text); border: 1px solid var(--border); border-radius: 14px;
    padding: 16px; transition: border-color .15s ease, box-shadow .15s ease, transform .15s ease;
  }
  .port-card:hover { background: var(--card); color: var(--text); border-color: var(--navy-light); box-shadow: var(--shadow-sm); transform: translateY(-1px); }
  .port-card-icon {
    width: 34px; height: 34px; border-radius: 10px; flex-shrink: 0;
    display: flex; align-items: center; justify-content: center;
    background: color-mix(in srgb, var(--navy-light) 12%, transparent); color: var(--navy-light);
  }
  .port-card-icon svg { width: 18px; height: 18px; stroke: currentColor; fill: none; stroke-width: 1.8; stroke-linecap: round; stroke-linejoin: round; }
  .port-card-name { font-size: 14.5px; font-weight: 700; color: var(--text); letter-spacing: -0.01em; }
  .port-card-meta { font-size: 12px; color: var(--muted); margin-top: -4px; }
  .port-card-progress { width: 100%; height: 6px; }
  .port-card-pct { font-size: 11.5px; font-weight: 600; color: var(--navy-light); }
  .port-card-pct.done { color: var(--success); }

  .port-breadcrumb { display: flex; align-items: center; gap: 12px; margin-bottom: 14px; flex-wrap: wrap; }
  .port-back-btn {
    display: inline-flex; align-items: center; gap: 4px;
    font-size: 12.5px; font-weight: 700; font-family: inherit;
    padding: 7px 14px; border-radius: 999px; cursor: pointer;
    background: var(--card); color: var(--navy-light); border: 1px solid var(--border);
    transition: background-color .15s ease;
  }
  .port-back-btn:hover { background: var(--border); color: var(--navy-light); }
  .port-breadcrumb-heading { font-size: 15px; font-weight: 700; color: var(--text); margin: 0; }

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
    background: none; border: 1px solid transparent; cursor: pointer;
    font-size: 11px; font-weight: 700; padding: 5px 11px; border-radius: 999px; flex-shrink: 0;
  }
  .port-header .group-remove { color: rgba(255,255,255,0.75); }
  .port-header .group-remove:hover { background: color-mix(in srgb, var(--danger) 55%, transparent); color: #fff; }
  /* Destructive ("Remove all") gets the same red outline as every other
     destructive control in the app (.del, bulk Remove) - see the
     button-hierarchy notes near .btn-danger below. */
  .vessel-header .group-remove {
    color: var(--danger); border-color: color-mix(in srgb, var(--danger) 45%, var(--border));
  }
  .vessel-header .group-remove:hover { background: var(--danger-bg); }
  /* Neutral/secondary group-header action (Archive / Unarchive) - same pill
     sizing as .group-remove/.group-export, but colored like the toolbar's
     neutral buttons (Expand all / Collapse all) so it reads as "secondary",
     not destructive and not primary. */
  .group-neutral {
    background: none; border: 1px solid var(--border); cursor: pointer;
    font-size: 11px; font-weight: 700; padding: 5px 11px; border-radius: 999px; flex-shrink: 0;
    color: var(--text);
  }
  .group-neutral:hover { background: color-mix(in srgb, var(--border) 70%, transparent); }
  .group-export {
    margin-left: auto; font-size: 11px; font-weight: 700; padding: 5px 11px;
    border-radius: 999px; flex-shrink: 0; text-decoration: none;
  }
  .port-header .group-export { color: rgba(255,255,255,0.75); }
  .port-header .group-export:hover { background: rgba(255,255,255,0.14); color: #fff; }
  .vessel-header .group-export { color: var(--navy-light); }
  .vessel-header .group-export:hover { background: color-mix(in srgb, var(--navy-light) 14%, transparent); }
  .eta-wrap { display: flex; align-items: center; gap: 5px; flex-shrink: 0; }
  .eta-label {
    font-size: 10px; font-weight: 700; color: var(--muted); text-transform: uppercase;
    letter-spacing: .04em;
  }
  .eta-input {
    font-size: 11.5px; font-family: inherit; border: 1px solid var(--border); border-radius: 6px;
    padding: 3px 6px; background: var(--card); color: var(--text); flex-shrink: 0;
  }
  .eta-input.eta-unset { border-style: dashed; border-color: var(--muted); }
  .eta-unset-hint { font-size: 10px; color: var(--muted); font-style: italic; }
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

  /* Per-vessel progress bar - % of this vessel's BLs fully complete, next
     to the existing "N BLs - N left" text. Recomputed every render(), same
     cadence as that text, so it updates on the same toggles already do. */
  .vessel-progress {
    position: relative; width: 64px; height: 6px; border-radius: 999px;
    background: color-mix(in srgb, var(--border) 85%, transparent); overflow: hidden; flex-shrink: 0;
  }
  .vessel-progress-fill {
    position: absolute; inset: 0; width: 0; border-radius: 999px;
    background: linear-gradient(90deg, var(--gold), var(--navy-light));
    transition: width .25s ease;
  }
  .vessel-progress.done .vessel-progress-fill { background: var(--success); }

  /* Table */
  /* This is the real scroll container for a vessel's table: bounded height
     + overflow-y:auto on purpose, so a long BL list scrolls inside its own
     card instead of the whole page, and the sticky <th> below sticks to
     THIS box's scrollport (its nearest actual scrolling ancestor), which
     works reliably everywhere. (Earlier this only had overflow-x:auto for
     horizontal scroll on narrow screens; that alone forces the browser to
     also treat overflow-y as a scroll container per the CSS overflow spec,
     but with no bounded height it never actually scrolled - so the sticky
     header had no real scrollport to stick within and just scrolled away
     with the page. Giving it a real max-height fixes that at the root.)
     A short vessel (fits within max-height) just renders in full with no
     scrollbar, exactly as before. */
  .overflow { overflow-x: auto; overflow-y: auto; max-height: 65vh; }
  table { width: 100%; min-width: 760px; border-collapse: collapse; font-size: 13.5px; table-layout: fixed; }
  th:nth-child(1), td:nth-child(1) { width: 4%; }
  th:nth-child(2), td:nth-child(2) { width: 15%; }
  th:nth-child(3), td:nth-child(3) { width: 14%; }
  th:nth-child(4), td:nth-child(4) { width: 14%; }
  th:nth-child(5), td:nth-child(5) { width: 14%; }
  th:nth-child(6), td:nth-child(6) { width: 23%; }
  th:nth-child(7), td:nth-child(7) { width: 6%; }
  th:nth-child(8), td:nth-child(8) { width: 10%; }
  th, td { text-align: left; padding: 12px 10px; border-bottom: 1px solid var(--border); overflow: hidden; }
  .select-col { text-align: center; }
  th {
    color: var(--muted); font-weight: 600; font-size: 10.5px; text-transform: uppercase;
    letter-spacing: .05em; background: color-mix(in srgb, var(--border) 40%, var(--card));
    /* Sticky column header: .overflow (its scrolling parent) now has a
       bounded max-height + overflow-y:auto, so it's the real scroll
       container a long vessel's rows scroll inside - this sticks to ITS
       top, not the page. Needs an opaque-ish background (above, mixed onto
       --card instead of transparent) so rows don't show through as they
       scroll underneath it. */
    position: sticky; top: 0; z-index: 10;
  }
  tbody tr { transition: background .12s ease, opacity .15s ease; }
  tbody tr:hover { background: color-mix(in srgb, var(--navy-light) 4%, transparent); }
  tbody tr:last-child td { border-bottom: none; }

  /* De-emphasize fully-complete rows so the eye skips them while scanning -
     subtle, not celebratory: slightly muted + a thin green accent on the BL
     number cell, not a full highlight. */
  tbody tr.row-complete { opacity: .6; }
  tbody tr.row-complete:hover { opacity: .85; }
  tbody tr.row-complete td { color: var(--muted); }
  tbody tr.row-complete td:nth-child(2) { box-shadow: inset 3px 0 0 var(--success); }

  /* Column, not row: the BL number sits on its own line and the doc chips
     always sit on the line below it. A single flex-wrap row here wrapped
     inconsistently depending on how long the BL number text happened to
     be - short numbers left room for the chips to tuck in beside them,
     long ones pushed the chips down, so the same row shape looked
     different BL to BL. Forcing two rows keeps it identical everywhere. */
  .bl-cell { display: flex; flex-direction: column; align-items: flex-start; gap: 4px; overflow: hidden; }
  .bl-cell b { font-weight: 700; letter-spacing: -0.01em; overflow: hidden; text-overflow: ellipsis; max-width: 100%; }
  .bl-cell-chips { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .badge-complete {
    display: inline-flex; align-items: center; gap: 3px;
    background: var(--success-bg); color: var(--success); font-size: 10.5px; font-weight: 700;
    padding: 2px 8px; border-radius: 999px; text-transform: uppercase; letter-spacing: .03em;
  }
  .hist-btn {
    background: none; color: var(--muted); font-size: 11.5px; font-weight: 600;
    padding: 5px 10px; border-radius: 999px; line-height: 1;
    border: 1px solid var(--border); white-space: nowrap;
  }
  .hist-btn:hover { background: var(--border); color: var(--text); }

  .bulk-bar {
    display: none; align-items: center; gap: 8px; flex-wrap: wrap;
    padding: 10px 14px; margin: 0 0 1px;
    background: color-mix(in srgb, var(--gold) 10%, transparent);
    border: 1px solid color-mix(in srgb, var(--gold) 30%, var(--border));
    border-radius: 10px; font-size: 12.5px;
  }
  .bulk-bar.active { display: flex; }
  .bulk-bar .bulk-count {
    font-weight: 700; color: var(--navy); background: color-mix(in srgb, var(--gold) 22%, transparent);
    padding: 3px 10px; border-radius: 999px; margin-right: 2px;
  }
  :root[data-theme="dark"] .bulk-bar .bulk-count { color: var(--gold-light); }
  .bulk-bar button {
    font-size: 11.5px; padding: 6px 13px; background: var(--navy); color: #fff;
  }
  .bulk-bar button:hover { background: var(--navy-light); }
  .bulk-bar .bulk-group {
    display: flex; align-items: center; gap: 6px; flex-wrap: wrap;
    padding: 3px 6px 3px 3px; border-radius: 999px;
    background: color-mix(in srgb, var(--card) 60%, transparent);
    border: 1px solid color-mix(in srgb, var(--border) 70%, transparent);
  }
  .bulk-bar .bulk-divider {
    width: 1px; align-self: stretch; margin: 2px 0;
    background: color-mix(in srgb, var(--gold) 35%, var(--border));
  }
  .bulk-bar button.bulk-unmark-btn {
    background: none; color: var(--navy); border: 1px solid var(--border);
  }
  :root[data-theme="dark"] .bulk-bar button.bulk-unmark-btn { color: var(--navy-light); }
  .bulk-bar button.bulk-unmark-btn:hover { background: color-mix(in srgb, var(--navy-light) 16%, transparent); }
  .bulk-bar button.bulk-remove-btn {
    background: none; color: var(--danger); margin-left: auto;
    border: 1px solid color-mix(in srgb, var(--danger) 45%, var(--border));
  }
  .bulk-bar button.bulk-remove-btn:hover { background: var(--danger-bg); }

  .history-overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 60;
    display: flex; align-items: center; justify-content: center; padding: 20px;
  }
  .history-modal {
    background: var(--card); border-radius: 14px; max-width: 480px; width: 100%;
    max-height: 70vh; display: flex; flex-direction: column; overflow: hidden;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
  }
  .history-modal-head {
    display: flex; align-items: center; gap: 10px; padding: 14px 18px;
    border-bottom: 1px solid var(--border);
  }
  .history-modal-body { padding: 10px 18px; overflow-y: auto; }
  .history-row { padding: 9px 0; border-bottom: 1px solid var(--border); font-size: 12.5px; }
  .history-row:last-child { border-bottom: none; }
  .history-row .when { color: var(--muted); font-size: 11px; }

  /* Doc chips next to the BL number - small pills showing whether an
     Invoice / DO file is attached. The native title attribute gives a real
     hover tooltip on desktop; clicking (works on both desktop and mobile,
     where hover doesn't exist) opens the Documents modal. */
  .doc-chip {
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 10px; font-weight: 700; letter-spacing: 0.02em;
    padding: 1px 6px; border-radius: 5px; cursor: pointer; border: 1px solid transparent;
    line-height: 1.5;
  }
  .doc-chip.has-file {
    background: color-mix(in srgb, var(--success, #1f9d55) 16%, transparent);
    color: var(--success, #1f9d55);
    border-color: color-mix(in srgb, var(--success, #1f9d55) 35%, transparent);
  }
  .doc-chip.no-file {
    background: var(--muted-bg, rgba(120,130,140,0.12)); color: var(--muted);
    border-color: var(--border);
  }
  .doc-chip:hover { filter: brightness(0.95); }

  .docs-section { padding: 12px 0; border-bottom: 1px solid var(--border); }
  .docs-section:last-child { border-bottom: none; }
  .docs-section .docs-section-title { font-weight: 700; font-size: 13px; margin-bottom: 6px; }
  .docs-section .docs-status { font-size: 12px; color: var(--muted); margin-bottom: 8px; }
  .docs-section .docs-actions { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  .docs-section .docs-actions label.btn-upload {
    display: inline-flex; align-items: center; gap: 6px; cursor: pointer;
    padding: 6px 12px; border-radius: 8px; background: var(--navy-light, #12405e);
    color: #fff; font-size: 12.5px; font-weight: 600;
  }
  .docs-section .docs-actions input[type=file] { display: none; }
  .docs-find-row { display: flex; gap: 8px; margin-bottom: 10px; }
  .docs-find-row input { flex: 1; }

  /* "Attach documents" auto-match batch rows - one per dropped file, while
     it's being read/matched, once it's auto-attached, or (when the BL
     couldn't be pinned down automatically) while it waits for the user to
     pick the right one by hand. */
  .match-row {
    display: flex; align-items: center; gap: 10px; padding: 9px 0;
    border-bottom: 1px solid var(--border); font-size: 12.5px;
  }
  .match-row:last-child { border-bottom: none; }
  .match-row .match-file { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .match-row .match-status { font-size: 11.5px; color: var(--muted); }
  .match-row.ok .match-status { color: var(--success, #1f9d55); }
  .match-row.review { flex-wrap: wrap; }
  .match-row.review .match-review-controls { display: flex; gap: 6px; align-items: center; width: 100%; margin-top: 4px; }
  .match-row.review .match-review-controls select,
  .match-row.review .match-review-controls input { flex: 1; min-width: 0; }

  /* Mobile - below this width, each row becomes a stacked card instead of
     a table row (a wide table just forces sideways scrolling on a phone,
     which is exactly what you don't want checking a BL at the port). */
  @media (max-width: 700px) {
    table, thead, tbody, th, td, tr { display: block; width: 100% !important; min-width: 0 !important; }
    thead { display: none; }
    tbody tr {
      border: 1px solid var(--border); border-radius: 10px; margin-bottom: 10px; padding: 8px 10px;
    }
    tbody tr td { border-bottom: none; padding: 7px 2px; }
    tbody tr td[data-label]::before {
      content: attr(data-label); display: block; font-size: 10px; font-weight: 700;
      text-transform: uppercase; letter-spacing: .04em; color: var(--muted); margin-bottom: 3px;
    }
    .select-col { display: flex; justify-content: flex-end; }
    .tag-fields { flex-direction: column; }
    /* Group headers pack a name, ETA, counts and three buttons into one
       row - fine on a desktop width, but forced onto one line on a phone
       it pushes the page wider than the screen. Let them wrap instead. */
    .port-header, .vessel-header { flex-wrap: wrap; row-gap: 6px; }
    .group-export { margin-left: 0; }
    .group-remove:last-of-type { margin-left: auto; }
    .row { flex-wrap: wrap; }
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
    border: 1px solid color-mix(in srgb, var(--danger) 40%, var(--border));
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

  /* Floating "back to top" bubble - jumps back up to the Discharge Port /
     Vessel fields from anywhere on a long board. Sits above #toastHost's
     own resting position (bottom:20px) so a toast popping in doesn't land
     right on top of it. */
  .scroll-top-btn {
    position: fixed; bottom: 86px; right: 24px; z-index: 900;
    width: 46px; height: 46px; border-radius: 50%; padding: 0;
    background: var(--navy-deep); color: #fff; border: none;
    display: flex; align-items: center; justify-content: center;
    box-shadow: var(--shadow-md); cursor: pointer;
  }
  .scroll-top-btn:hover { background: var(--navy-light); }
  .scroll-top-btn svg { width: 20px; height: 20px; }

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

  <div class="card" id="manifestCard">
    <div class="card-label">Add a manifest</div>
    <div class="tag-fields">
      <div>
        <label for="portField">Discharge Port</label>
        <div class="glass-select-wrap">
          <select id="portField" class="nice-select">
            <option value="">Select a port...</option>
            <option value="DAMMAM PORT">Dammam Port</option>
            <option value="JUBAIL COMMERCIAL PORT">Jubail Commercial Port</option>
            <option value="JEDDAH PORT">Jeddah Port</option>
            <option value="YANBU COMMERCIAL PORT">Yanbu Commercial Port</option>
            <option value="YANBU INDUSTRIAL PORT">Yanbu Industrial Port</option>
            <option value="KAP">KAP</option>
          </select>
        </div>
      </div>
      <div>
        <label for="vesselField">Vessel</label>
        <input type="text" id="vesselField" placeholder="e.g. TAI KNIGHT" style="text-transform:uppercase;" oninput="this.value = this.value.toUpperCase();"
          onkeydown="if(event.key==='Enter'){ event.preventDefault(); uploadExcel(); }">
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

  <div class="card">
    <div class="card-label">Attach documents</div>
    <div style="font-size:12.5px; color:var(--muted); margin-bottom:12px;">
      Drop Invoice / Delivery Order PDFs here - each one is read and matched to its BL automatically, same as the manifest upload above.
      The <span class="doc-chip has-file" style="cursor:default;">INV</span> / <span class="doc-chip has-file" style="cursor:default;">DO</span> chips next to a BL number below show what's already attached.
    </div>
    <label class="dropzone" id="autoMatchDropzone" for="autoMatchFile">
      <div class="dropzone-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
          <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/>
        </svg>
      </div>
      <div>
        <div class="dropzone-text"><b>Click to upload</b> or drag &amp; drop Invoice/DO PDFs</div>
        <div class="dropzone-sub">Drop as many at once as you like - each is matched to its BL automatically</div>
      </div>
      <input type="file" id="autoMatchFile" accept=".pdf" multiple style="display:none" onchange="handleAutoMatchFiles(this.files)">
    </label>
    <div id="autoMatchSummary" style="font-size:12.5px; color:var(--muted); margin-top:10px;"></div>
    <div id="autoMatchList" style="margin-top:6px;"></div>
  </div>

  <div class="summary" id="summary"></div>

  <div class="card">
    <div class="row" style="margin-bottom:14px; flex-wrap:wrap;">
      <input type="text" id="searchBox" placeholder="Search BL number..." oninput="render()" style="flex:1; min-width:180px;">
      <div class="glass-select-wrap" style="width:auto; min-width:200px;">
        <select id="jumpSelect" class="nice-select" onchange="jumpToVessel(this.value)"><option value="">Select a vessel to view</option></select>
      </div>
      {% if role == 'admin' %}
      <div class="glass-select-wrap" style="width:auto; min-width:140px;">
        <select id="operatorFilter" class="nice-select" onchange="render()"><option value="">All operators</option></select>
      </div>
      {% endif %}
      <button type="button" class="btn-neutral" onclick="setAllGroupsCollapsed(true)">Collapse all</button>
    </div>
    <div id="portTabs"></div>
    <div id="groups"></div>
  </div>

  <div class="card" id="archivedCard" style="display:none;">
    <div class="row" style="margin-bottom:14px; cursor:pointer;" onclick="archivedSectionOpen = !archivedSectionOpen; render();">
      <b style="flex:1;">Archived vessels</b>
      <span class="group-count" id="archivedCount"></span>
    </div>
    <div id="archivedGroups"></div>
  </div>

  <div id="toastHost"></div>
  <button type="button" id="scrollTopBtn" class="scroll-top-btn" title="Back to Discharge Port / Vessel" onclick="scrollToManifestForm()">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2.2"><path d="M12 19V5M5 12l7-7 7 7"/></svg>
  </button>
  <div id="historyOverlay" class="history-overlay" style="display:none;" onclick="if(event.target===this) closeHistory()">
    <div class="history-modal">
      <div class="history-modal-head">
        <b id="historyTitle"></b>
        <button type="button" onclick="closeHistory()" style="background:none; color:var(--text); padding:4px 10px;">&times;</button>
      </div>
      <div id="historyBody" class="history-modal-body"></div>
    </div>
  </div>

  <div id="docsOverlay" class="history-overlay" style="display:none;" onclick="if(event.target===this) closeDocs()">
    <div class="history-modal">
      <div class="history-modal-head">
        <b id="docsTitle">Documents</b>
        <button type="button" onclick="closeDocs()" style="background:none; color:var(--text); padding:4px 10px;">&times;</button>
      </div>
      <div id="docsBody" class="history-modal-body"></div>
    </div>
  </div>

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
const IS_ADMIN = {{ (role == 'admin')|tojson }};
let records = [];
let suppressPollUntil = 0;
let editingCount = 0;
let collapsedGroups = {};
let archivedSectionOpen = false;
let selectedPortTab = '';

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

  // A BL that's already on the board under a DIFFERENT vessel than the one
  // just uploaded is worth a second look - either this file re-lists a BL
  // that's really a different shipment (a shipper reusing a number), or a
  // genuine mistake. Either way, silently skipping it like an ordinary
  // repeat-upload duplicate would hide it.
  if (data.duplicate_elsewhere && data.duplicate_elsewhere.length) {
    const lines = data.duplicate_elsewhere.slice(0, 5).map(d =>
      `${d.bl_number} (already under ${d.existing_vessel || 'Unassigned'} / ${d.existing_port || 'Unassigned'})`
    ).join('; ');
    const more = data.duplicate_elsewhere.length > 5 ? ` and ${data.duplicate_elsewhere.length - 5} more` : '';
    showToast(`Heads up - ${data.duplicate_elsewhere.length} BL(s) already exist under a different vessel: ${lines}${more}.`, {duration: 9000});
  }
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
  const chips = cell.querySelector('.bl-cell-chips') || cell;
  let badge = cell.querySelector('.badge-complete');
  const complete = !!(rec.invoice_issued && rec.approval_received && rec.do_issued);
  if (complete && !badge) {
    badge = document.createElement('span');
    badge.className = 'badge-complete';
    badge.innerHTML = '&check; Complete';
    chips.appendChild(badge);
  } else if (!complete && badge) {
    badge.remove();
  }
  row.classList.toggle('row-complete', complete);
}

function cssEscape(s) {
  return String(s).replace(/[^a-zA-Z0-9_-]/g, c => '_' + c.charCodeAt(0) + '_');
}

function updateSummaryOnly() {
  document.getElementById('summary').innerHTML = summaryHtml();
}

// Rapid clicking on the same slider used to be able to "undo" an earlier
// click: toggle() fired its own POST immediately on every call, so a quick
// burst of clicks put several requests for the same bl+field in flight at
// once, with nothing guaranteeing they reached (or were processed by) the
// server in the same order they were sent - whichever one the server
// happened to finish last would win, which wasn't necessarily the one
// matching the final click. Debouncing the actual network send below - so
// a burst of clicks on the same bl+field within a short window results in
// exactly one request, carrying whatever the latest click's value was -
// removes that race instead of trying to patch it up after the fact. The
// on-screen state is still updated on every single click, so it never
// feels laggy; only the save to the server is coalesced.
let toggleSendTimers = {};

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
  suppressPollUntil = Date.now() + 2000;

  const key = bl + '::' + field;
  clearTimeout(toggleSendTimers[key]);
  toggleSendTimers[key] = setTimeout(() => {
    delete toggleSendTimers[key];
    fetch(`/api/records/${encodeURIComponent(bl)}/toggle`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({field, value})
    }).then(() => fetchRecords()).catch(() => { showToast('Could not save that change - retrying...'); fetchRecords(); });
  }, 350);
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

/* ---------- Per-BL history ---------- */
const AUDIT_ACTION_LABELS = {
  added: 'Added to board', deleted: 'Removed', restored: 'Restored',
  toggle: 'status changed', remarks: 'Remarks edited',
};

function historyFieldLabel(field) {
  return {invoice_issued: 'Invoice Issued', approval_received: 'Approval Received', do_issued: 'DO Issued', remarks: 'Remarks'}[field] || field;
}

async function showHistory(bl) {
  const overlay = document.getElementById('historyOverlay');
  const body = document.getElementById('historyBody');
  document.getElementById('historyTitle').textContent = 'History - ' + bl;
  body.innerHTML = '<div style="color:var(--muted); padding:10px 0;">Loading...</div>';
  overlay.style.display = 'flex';

  const res = await fetch(`/api/records/${encodeURIComponent(bl)}/history`);
  if (!res.ok) { body.innerHTML = '<div style="color:var(--muted); padding:10px 0;">Could not load history.</div>'; return; }
  const entries = await res.json();
  if (!entries.length) {
    body.innerHTML = '<div style="color:var(--muted); padding:10px 0;">No history recorded yet.</div>';
    return;
  }
  body.innerHTML = entries.map(e => {
    let line;
    if (e.action === 'toggle') {
      line = `<b>${e.by_user || 'Unknown'}</b> set ${historyFieldLabel(e.field)} to ${e.new_value ? 'Yes' : 'No'}`;
    } else if (e.action === 'remarks') {
      line = `<b>${e.by_user || 'Unknown'}</b> edited remarks${e.new_value ? ': "' + e.new_value + '"' : ' (cleared)'}`;
    } else if (e.action === 'added') {
      line = `<b>${e.by_user || 'Unknown'}</b> added this BL${e.new_value ? ' (' + e.new_value + ')' : ''}`;
    } else {
      line = `<b>${e.by_user || 'Unknown'}</b> ${AUDIT_ACTION_LABELS[e.action] || e.action}`;
    }
    return `<div class="history-row"><div>${line}</div><div class="when">${formatLocalTime(e.at)}</div></div>`;
  }).join('');
}

function closeHistory() {
  document.getElementById('historyOverlay').style.display = 'none';
}

/* ---------- Invoice / DO file attachments ----------
   Small "doc chip" badges next to the BL number show at a glance whether an
   Invoice/DO file has been attached (native title= gives a real hover
   tooltip on desktop, and the click handler covers mobile where hover
   doesn't exist). Clicking either chip - or any cell in the row that opens
   it - shows the Documents modal. Download (and the /lookup it's built on)
   deliberately isn't limited to the BL's creator, even though upload/
   replace/remove still are - but note this only helps someone who can
   already see the row: the board itself (and search) stays scoped to each
   staff member's own BLs. */
const DOC_KINDS = [['invoice', 'Invoice'], ['do', 'Delivery Order']];

function docChip(bl, kind, hasFile) {
  const label = kind === 'invoice' ? 'INV' : 'DO';
  const full = kind === 'invoice' ? 'Invoice' : 'Delivery Order';
  const title = hasFile ? `${full} attached - click to view` : `${full} not attached yet - click to upload`;
  return `<span class="doc-chip ${hasFile ? 'has-file' : 'no-file'}" title="${title}"
            onclick="event.stopPropagation(); showDocs('${bl}')">${label}</span>`;
}

function renderDocsSections(data, canManage) {
  const atts = data.attachments || {};
  return DOC_KINDS.map(([kind, label]) => {
    const att = atts[kind];
    let status, actions;
    if (att) {
      status = `Attached: <b>${att.filename || (kind + '.pdf')}</b><br>by ${att.uploaded_by || 'Unknown'} - ${formatLocalTime(att.uploaded_at)}`;
      actions = `
        <a href="/api/records/${encodeURIComponent(data.bl_number)}/attachment/${kind}" target="_blank" rel="noopener">
          <button type="button">Download</button>
        </a>
        ${canManage ? `
          <label class="btn-upload">Replace<input type="file" accept=".pdf,application/pdf" onchange="uploadAttachment('${data.bl_number}', '${kind}', this)"></label>
          <button type="button" class="btn-danger" onclick="removeAttachment('${data.bl_number}', '${kind}')">Remove</button>
        ` : ''}`;
    } else {
      status = canManage ? 'Not attached yet.' : 'Not attached yet - waiting on the issuing staff member.';
      actions = canManage ? `
        <label class="btn-upload">Upload PDF<input type="file" accept=".pdf,application/pdf" onchange="uploadAttachment('${data.bl_number}', '${kind}', this)"></label>
      ` : '';
    }
    return `
      <div class="docs-section">
        <div class="docs-section-title">${label}</div>
        <div class="docs-status">${status}</div>
        <div class="docs-actions">${actions}</div>
      </div>`;
  }).join('');
}

async function showDocs(bl) {
  const overlay = document.getElementById('docsOverlay');
  const body = document.getElementById('docsBody');
  document.getElementById('docsTitle').textContent = 'Documents - ' + bl;
  body.innerHTML = '<div style="color:var(--muted); padding:10px 0;">Loading...</div>';
  overlay.style.display = 'flex';
  overlay.dataset.bl = bl;

  const res = await fetch(`/api/records/${encodeURIComponent(bl)}/lookup`);
  const data = await res.json();
  if (!res.ok) { body.innerHTML = `<div style="color:var(--muted); padding:10px 0;">${data.error || 'Could not load this BL.'}</div>`; return; }
  const canManage = IS_ADMIN || data.created_by === CURRENT_USER;
  body.innerHTML = renderDocsSections(data, canManage);
}

function closeDocs() {
  document.getElementById('docsOverlay').style.display = 'none';
}

async function submitAttachmentFile(bl, kind, file) {
  // Shared by the single-file Documents-modal upload and the batch
  // auto-match flow below - just the raw POST, no UI side effects, so
  // each caller decides how to react to the result.
  const form = new FormData();
  form.append('file', file);
  const res = await fetch(`/api/records/${encodeURIComponent(bl)}/attachment/${kind}`, {method: 'POST', body: form});
  const data = await res.json();
  return {ok: res.ok, data};
}

async function uploadAttachment(bl, kind, input) {
  const file = input.files && input.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pdf')) { showToast('Only PDF files are accepted.'); input.value = ''; return; }
  if (file.size > 10 * 1024 * 1024) { showToast('That file is larger than 10MB.'); input.value = ''; return; }

  const {ok, data} = await submitAttachmentFile(bl, kind, file);
  if (!ok) { showToast(data.error || 'Upload failed.'); return; }
  const label = kind === 'invoice' ? 'Invoice' : 'Delivery Order';
  // Attaching the file auto-flips the matching Issued slider server-side
  // (see upload_attachment) - say so, so it's obvious the status change
  // wasn't a separate click someone forgot to make.
  showToast(data.auto_issued_field ? `${label} uploaded - marked as issued.` : `${label} uploaded.`);
  await fetchRecords();
  if (document.getElementById('docsOverlay').style.display !== 'none') await showDocs(bl);
}

/* ---------- Attach documents (batch auto-match) ----------
   Drop a pile of Fasah Invoice/DO PDFs at once; each is sent to
   /api/attachments/detect (reads the PDF, figures out Invoice vs DO from
   fixed template anchors, and checks which of this user's own BLs appears
   in it). A clean single match uploads immediately via the same route the
   Documents modal uses; anything else (no match, more than one candidate,
   or an unrecognized document) is left in the list for the user to assign
   by hand rather than guessed at. */
const autoMatchDropzone = document.getElementById('autoMatchDropzone');
['dragenter', 'dragover'].forEach(evt => {
  autoMatchDropzone.addEventListener(evt, e => { e.preventDefault(); autoMatchDropzone.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(evt => {
  autoMatchDropzone.addEventListener(evt, e => { e.preventDefault(); autoMatchDropzone.classList.remove('dragover'); });
});
autoMatchDropzone.addEventListener('drop', e => {
  if (e.dataTransfer.files && e.dataTransfer.files.length) handleAutoMatchFiles(e.dataTransfer.files);
});

function autoMatchKindLabel(kind) {
  return kind === 'invoice' ? 'Invoice' : kind === 'do' ? 'Delivery Order' : 'Unrecognized document';
}

async function handleAutoMatchFiles(fileList) {
  const files = Array.from(fileList || []).filter(f => f.name.toLowerCase().endsWith('.pdf'));
  if (!files.length) { showToast('Drop PDF files only.'); return; }

  const listEl = document.getElementById('autoMatchList');
  const summaryEl = document.getElementById('autoMatchSummary');
  let attached = 0, needsReview = 0;
  const updateSummary = () => {
    summaryEl.textContent = `${attached} attached automatically` + (needsReview ? `, ${needsReview} need your input` : '') +
      ((attached + needsReview) < files.length ? ` (processing ${files.length - attached - needsReview} more...)` : '.');
  };
  updateSummary();

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const rowId = `matchrow_${Date.now()}_${i}`;
    const row = document.createElement('div');
    row.className = 'match-row';
    row.id = rowId;
    row.innerHTML = `<div class="match-file" title="${file.name}">${file.name}</div><div class="match-status">Reading...</div>`;
    listEl.appendChild(row);

    if (file.size > 10 * 1024 * 1024) {
      row.querySelector('.match-status').textContent = 'Too large (over 10MB) - skipped.';
      needsReview++; updateSummary();
      continue;
    }

    let detect;
    try {
      const form = new FormData();
      form.append('file', file);
      const res = await fetch('/api/attachments/detect', {method: 'POST', body: form});
      detect = await res.json();
      if (!res.ok) throw new Error(detect.error || 'Could not read this file.');
    } catch (err) {
      row.className = 'match-row review';
      row.innerHTML = `<div class="match-file" title="${file.name}">${file.name}</div><div class="match-status">${err.message}</div>`;
      needsReview++; updateSummary();
      continue;
    }

    if (detect.kind && detect.matched_bl) {
      const {ok, data} = await submitAttachmentFile(detect.matched_bl, detect.kind, file);
      if (ok) {
        row.className = 'match-row ok';
        const issuedNote = data.auto_issued_field ? ' &middot; marked issued' : '';
        row.innerHTML = `<div class="match-file" title="${file.name}">${file.name}</div><div class="match-status">&check; ${detect.matched_bl} - ${autoMatchKindLabel(detect.kind)}${issuedNote}</div>`;
        attached++; updateSummary();
        continue;
      }
      // Fall through to manual review if the upload itself was rejected
      // (e.g. not the owner of that BL after all) - rare, since the
      // candidate list was already scoped server-side, but don't just
      // drop the file silently if it happens.
      row.className = 'match-row review';
      renderAutoMatchReviewRow(row, file, detect, data.error);
      needsReview++; updateSummary();
      continue;
    }

    row.className = 'match-row review';
    renderAutoMatchReviewRow(row, file, detect, null);
    needsReview++; updateSummary();
  }

  await fetchRecords();
}

function renderAutoMatchReviewRow(row, file, detect, errorMsg) {
  const blOptions = records.map(r => r.bl_number).sort();
  const candidates = (detect.candidates && detect.candidates.length) ? detect.candidates : blOptions;
  const statusText = errorMsg ? errorMsg
    : detect.candidates && detect.candidates.length ? `Matches more than one BL - pick the right one`
    : !detect.kind ? `Couldn't tell Invoice from Delivery Order`
    : `No BL on your board matched this document`;

  row.innerHTML = `
    <div class="match-file" title="${file.name}">${file.name}</div>
    <div class="match-status">${statusText}</div>
    <div class="match-review-controls">
      <select class="review-bl">
        <option value="">Select BL...</option>
        ${candidates.map(bl => `<option value="${bl}" ${bl === detect.matched_bl ? 'selected' : ''}>${bl}</option>`).join('')}
      </select>
      <select class="review-kind">
        <option value="">Kind...</option>
        <option value="invoice" ${detect.kind === 'invoice' ? 'selected' : ''}>Invoice</option>
        <option value="do" ${detect.kind === 'do' ? 'selected' : ''}>Delivery Order</option>
      </select>
      <button type="button" class="review-attach-btn">Attach</button>
    </div>`;

  row.querySelector('.review-attach-btn').onclick = async () => {
    const bl = row.querySelector('.review-bl').value;
    const kind = row.querySelector('.review-kind').value;
    if (!bl || !kind) { showToast('Pick both a BL number and a document kind.'); return; }
    const btn = row.querySelector('.review-attach-btn');
    btn.disabled = true;
    btn.textContent = 'Attaching...';
    const {ok, data} = await submitAttachmentFile(bl, kind, file);
    if (!ok) { showToast(data.error || 'Attach failed.'); btn.disabled = false; btn.textContent = 'Attach'; return; }
    row.className = 'match-row ok';
    const issuedNote = data.auto_issued_field ? ' &middot; marked issued' : '';
    row.innerHTML = `<div class="match-file" title="${file.name}">${file.name}</div><div class="match-status">&check; ${bl} - ${autoMatchKindLabel(kind)}${issuedNote}</div>`;
    await fetchRecords();
  };
}

async function removeAttachment(bl, kind) {
  const res = await fetch(`/api/records/${encodeURIComponent(bl)}/attachment/${kind}`, {method: 'DELETE'});
  if (!res.ok) { showToast('Could not remove the file.'); return; }
  showToast(`${kind === 'invoice' ? 'Invoice' : 'Delivery Order'} removed.`);
  await fetchRecords();
  if (document.getElementById('docsOverlay').style.display !== 'none') await showDocs(bl);
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
  const complete = records.filter(r => r.invoice_issued && r.approval_received && r.do_issued).length;

  const icons = {
    total: '<svg viewBox="0 0 24 24"><path d="M7 3h7l5 5v13a1 1 0 01-1 1H7a1 1 0 01-1-1V4a1 1 0 011-1z"/><path d="M14 3v5h5"/></svg>',
    check: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/></svg>'
  };

  const remaining = total - complete;

  return `
    <div class="stat"><div class="stat-icon">${icons.total}</div><div><b>${total}</b>Total BLs</div></div>
    <div class="stat ${remaining ? 'gold' : 'done'}"><div class="stat-icon">${icons.check}</div><div><b>${remaining}</b>Remaining</div></div>
    <div class="stat done"><div class="stat-icon">${icons.check}</div><div><b>${complete}</b>Fully Complete</div></div>
  `;
}

const CHEVRON = '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>';

let selectedBLs = new Set();

function rowsHtml(list) {
  return list.map(r => {
    const complete = !!(r.invoice_issued && r.approval_received && r.do_issued);
    return `
    <tr id="row_${cssEscape(r.bl_number)}" class="${complete ? 'row-complete' : ''}">
      <td class="select-col"><input type="checkbox" class="row-select" ${selectedBLs.has(r.bl_number) ? 'checked' : ''}
            onchange="toggleRowSelect('${r.bl_number}', this.checked)"></td>
      <td>
        <div class="bl-cell">
          <b>${r.bl_number}</b>
          <div class="bl-cell-chips">
            ${docChip(r.bl_number, 'invoice', r.has_invoice_file)}
            ${docChip(r.bl_number, 'do', r.has_do_file)}
            ${complete ? '<span class="badge-complete">&check; Complete</span>' : ''}
          </div>
        </div>
      </td>
      <td data-label="Invoice Issued">${checkbox(r.bl_number, 'invoice_issued', !!r.invoice_issued, r.invoice_by, r.invoice_at)}</td>
      <td data-label="Approval Received">${checkbox(r.bl_number, 'approval_received', !!r.approval_received, r.approval_by, r.approval_at)}</td>
      <td data-label="DO Issued">${checkbox(r.bl_number, 'do_issued', !!r.do_issued, r.do_by, r.do_at)}</td>
      <td data-label="Remarks"><input class="remarks-input" type="text" value="${(r.remarks || '').replace(/"/g,'&quot;')}"
            oninput="onRemarksInput('${r.bl_number}', this.value)"
            onfocus="markEditing(1)" onblur="markEditing(-1)" placeholder="notes..."></td>
      <td>{% if role == 'admin' %}<button type="button" class="hist-btn" title="History" onclick="showHistory('${r.bl_number}')">History</button>{% endif %}</td>
      <td><button class="del" onclick="deleteRecord('${r.bl_number}')">Remove</button></td>
    </tr>`;
  }).join('');
}

function toggleRowSelect(bl, checked) {
  if (checked) selectedBLs.add(bl); else selectedBLs.delete(bl);
  updateBulkBars();
}

function updateBulkBars() {
  document.querySelectorAll('.bulk-bar').forEach(bar => {
    const scope = (bar.dataset.bls || '').split('|').filter(Boolean);
    const count = scope.filter(bl => selectedBLs.has(bl)).length;
    const label = bar.querySelector('.bulk-count');
    if (label) label.textContent = count ? `${count} selected` : '';
    bar.classList.toggle('active', count > 0);
  });
  // Keep each vessel group's "select all" header checkbox in sync with the
  // actual selection: unchecked when none selected, checked when every BL
  // in the group is selected, indeterminate when only some are - so the
  // same control always does the obvious next thing (select all / clear all)
  // instead of only ever being able to check itself.
  document.querySelectorAll('.select-all-vessel').forEach(cb => {
    const key = cb.dataset.barKey;
    const bar = document.querySelector(`.bulk-bar[data-bar-key="${CSS.escape(key)}"]`);
    const scope = bar ? (bar.dataset.bls || '').split('|').filter(Boolean) : [];
    const count = scope.filter(bl => selectedBLs.has(bl)).length;
    cb.checked = scope.length > 0 && count === scope.length;
    cb.indeterminate = count > 0 && count < scope.length;
  });
}

async function bulkSetField(vesselKey, field, value) {
  const listEl = document.querySelector(`[data-bar-key="${CSS.escape(vesselKey)}"]`);
  const scope = listEl ? (listEl.dataset.bls || '').split('|').filter(Boolean) : [];
  const blNumbers = scope.filter(bl => selectedBLs.has(bl));
  if (!blNumbers.length) { showToast('Select at least one BL first.'); return; }

  blNumbers.forEach(bl => {
    const rec = records.find(r => r.bl_number === bl);
    if (rec) {
      rec[field] = value ? 1 : 0;
      const byField = field.replace('_issued', '_by').replace('_received', '_by');
      const atField = field.replace('_issued', '_at').replace('_received', '_at');
      rec[byField] = value ? CURRENT_USER : '';
      rec[atField] = value ? nowLabel() : '';
    }
  });
  suppressPollUntil = Date.now() + 2000;
  render();

  await fetch('/api/records/bulk-toggle', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers, field, value})
  });
  showToast(`${blNumbers.length} BL(s) updated.`);
  await fetchRecords();
}

function bulkRemoveSelected(vesselKey) {
  // Scoped to only the checked BLs in this vessel's bar - NOT the whole
  // vessel group (that's what the header's "Remove all" button is for) -
  // and passes an explicit bl_numbers list to the backend, same as every
  // other per-group bulk action here.
  const listEl = document.querySelector(`[data-bar-key="${CSS.escape(vesselKey)}"]`);
  const scope = listEl ? (listEl.dataset.bls || '').split('|').filter(Boolean) : [];
  const blNumbers = scope.filter(bl => selectedBLs.has(bl));
  if (!blNumbers.length) { showToast('Select at least one BL first.'); return; }
  const list = blNumbers.map(bl => records.find(r => r.bl_number === bl)).filter(Boolean);
  confirmBulkRemove(null, list);
}

function bulkBarHtml(vesselKey, list) {
  const key = vesselKey.replace(/"/g, '&quot;');
  const blsAttr = list.map(r => r.bl_number).join('|');
  return `
    <div class="bulk-bar" data-bar-key="${key}" data-bls="${blsAttr}">
      <span class="bulk-count"></span>
      <span class="bulk-group">
        <button type="button" onclick="bulkSetField('${vesselKey.replace(/'/g,"\\'")}', 'invoice_issued', true)">Mark Invoice Issued</button>
        <button type="button" class="bulk-unmark-btn" onclick="bulkSetField('${vesselKey.replace(/'/g,"\\'")}', 'invoice_issued', false)">Unmark</button>
      </span>
      <span class="bulk-group">
        <button type="button" onclick="bulkSetField('${vesselKey.replace(/'/g,"\\'")}', 'approval_received', true)">Mark Approval Received</button>
        <button type="button" class="bulk-unmark-btn" onclick="bulkSetField('${vesselKey.replace(/'/g,"\\'")}', 'approval_received', false)">Unmark</button>
      </span>
      <span class="bulk-group">
        <button type="button" onclick="bulkSetField('${vesselKey.replace(/'/g,"\\'")}', 'do_issued', true)">Mark DO Issued</button>
        <button type="button" class="bulk-unmark-btn" onclick="bulkSetField('${vesselKey.replace(/'/g,"\\'")}', 'do_issued', false)">Unmark</button>
      </span>
      <button type="button" class="bulk-remove-btn" onclick="bulkRemoveSelected('${vesselKey.replace(/'/g,"\\'")}')">Remove</button>
    </div>`;
}

function tableHtml(list, vesselKey) {
  const key = vesselKey.replace(/"/g, '&quot;');
  return `
    ${bulkBarHtml(vesselKey, list)}
    <div class="overflow">
      <table>
        <thead>
          <tr>
            <th class="select-col"><input type="checkbox" class="select-all-vessel" data-bar-key="${key}" title="Select/deselect all in this vessel" onchange="list_selectAllVessel('${vesselKey.replace(/'/g,"\\'")}', this.checked)"></th>
            <th>BL Number</th>
            <th>Invoice Issued</th>
            <th>Approval Received</th>
            <th>DO Issued</th>
            <th>Remarks</th>
            <th></th>
            <th></th>
          </tr>
        </thead>
        <tbody>${rowsHtml(list)}</tbody>
      </table>
    </div>`;
}

function list_selectAllVessel(vesselKey, checked) {
  const bar = document.querySelector(`[data-bar-key="${CSS.escape(vesselKey)}"]`);
  const bls = bar ? (bar.dataset.bls || '').split('|').filter(Boolean) : [];
  bls.forEach(bl => { if (checked) selectedBLs.add(bl); else selectedBLs.delete(bl); });
  render();
}

function groupRecordsByPortVessel(list) {
  const ports = {};
  list.forEach(r => {
    const port = r.port || 'Unassigned';
    const vessel = r.vessel || 'Unassigned';
    if (!ports[port]) ports[port] = {};
    if (!ports[port][vessel]) ports[port][vessel] = [];
    ports[port][vessel].push(r);
  });
  return ports;
}

function sortedPortNames(ports) {
  return Object.keys(ports).sort((a, b) => {
    if (a === 'Unassigned') return 1;
    if (b === 'Unassigned') return -1;
    return naturalCompare(a, b);
  });
}

function sortedVesselNames(vessels) {
  // A vessel with an ETA set sorts soonest-first (the next ship in is the
  // one you'd actually work first); vessels with no ETA fall after,
  // alphabetically.
  return Object.keys(vessels).sort((a, b) => {
    if (a === 'Unassigned') return 1;
    if (b === 'Unassigned') return -1;
    const etaA = (vessels[a][0] && vessels[a][0].eta) || '';
    const etaB = (vessels[b][0] && vessels[b][0].eta) || '';
    if (etaA && etaB && etaA !== etaB) return etaA < etaB ? -1 : 1;
    if (etaA && !etaB) return -1;
    if (!etaA && etaB) return 1;
    return naturalCompare(a, b);
  });
}

function vesselGroupHtml(portName, vesselName, list, archivedView) {
  const vesselKey = 'vessel:' + portName + ':' + vesselName;
  const collapseKey = archivedView ? vesselKey + ':archived' : vesselKey;
  const sortedList = list.slice().sort((a, b) => naturalCompare(a.bl_number, b.bl_number));
  const left = sortedList.filter(r => !(r.invoice_issued && r.approval_received && r.do_issued)).length;
  const pct = sortedList.length ? Math.round(((sortedList.length - left) / sortedList.length) * 100) : 0;
  const vesselCollapsed = collapseKey in collapsedGroups ? !!collapsedGroups[collapseKey] : (archivedView ? true : left === 0);
  const eta = (sortedList[0] && sortedList[0].eta) || '';
  const rawPort = (sortedList[0] && sortedList[0].port) || '';
  const rawVessel = (sortedList[0] && sortedList[0].vessel) || '';
  const blList = sortedList.map(r => r.bl_number);
  const pEsc = portName.replace(/'/g, "\\'");
  const vEsc = vesselName.replace(/'/g, "\\'");
  const blsJson = JSON.stringify(blList).replace(/'/g, "&#39;");
  return `
    <div class="vessel-group" id="group_${cssEscape(vesselKey)}">
      <div class="vessel-header ${vesselCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT' && event.target.tagName!=='BUTTON' && event.target.tagName!=='A') toggleGroup('${collapseKey.replace(/'/g,"\\'")}')">
        ${CHEVRON}
        <input class="group-name" value="${vesselName === 'Unassigned' ? '' : vesselName}" placeholder="Unassigned vessel"
          onclick="event.stopPropagation()"
          onchange="renameGroup('vessel', '${pEsc}', '${vEsc}', this.value, 'Unassigned')">
        ${archivedView ? '' : `<span class="eta-wrap" onclick="event.stopPropagation()">
          <span class="eta-label">ETA</span>
          <input type="date" class="eta-input${eta ? '' : ' eta-unset'}" value="${eta}" title="Expected arrival"
            onchange="this.classList.toggle('eta-unset', !this.value); setVesselEta(JSON.parse(this.dataset.bls), this.value)" data-bls='${blsJson}'>
          ${eta ? '' : '<span class="eta-unset-hint">Not set</span>'}
        </span>`}
        <span class="group-count">${sortedList.length} BL${sortedList.length === 1 ? '' : 's'}${left ? ` &middot; ${left} left` : ' &middot; done'}</span>
        <span class="vessel-progress ${pct >= 100 ? 'done' : ''}" title="${pct}% complete"><span class="vessel-progress-fill" style="width:${pct}%"></span></span>
        <a onclick="event.stopPropagation()" href="/api/export?port=${encodeURIComponent(rawPort)}&vessel=${encodeURIComponent(rawVessel)}" class="group-export" title="Export this vessel to Excel">Export</a>
        <button type="button" class="group-neutral" onclick='event.stopPropagation(); setVesselArchived(${blsJson}, ${archivedView ? 'false' : 'true'})'>${archivedView ? 'Unarchive' : 'Archive'}</button>
        <button type="button" class="group-remove" onclick="event.stopPropagation(); removeVesselGroup('${pEsc}', '${vEsc}')">Remove all</button>
      </div>
      <div class="vessel-body ${vesselCollapsed ? 'collapsed' : ''}">
        ${tableHtml(sortedList, vesselKey)}
      </div>
    </div>`;
}

function portGroupHtml(portName, vesselNames, vessels, archivedView, suppressHeader) {
  const portKey = archivedView ? 'port:' + portName + ':archived' : 'port:' + portName;
  const vesselsHtml = vesselNames.map(vesselName => vesselGroupHtml(portName, vesselName, vessels[vesselName], archivedView)).join('');
  // When a single port is already drilled into (the breadcrumb above the
  // board names it), this group's own dark "<PORT> ▾" header would just be
  // repeating that same name a few pixels below it with nothing new to
  // collapse into - so skip the header/collapse chrome entirely and show
  // the vessel groups directly. Still used (header shown) for the "all
  // ports" search view and the archived-vessels section, where more than
  // one port can appear at once and the header is the only thing naming
  // which port a group belongs to.
  if (suppressHeader) {
    return `<div class="port-group port-group-flat">${vesselsHtml}</div>`;
  }
  const portTotal = vesselNames.reduce((sum, v) => sum + vessels[v].length, 0);
  const portLeft = vesselNames.reduce((sum, v) => sum + vessels[v].filter(r => !(r.invoice_issued && r.approval_received && r.do_issued)).length, 0);
  // Default state (only applies the first time a group is seen - once a
  // person manually expands/collapses it, collapsedGroups remembers their
  // choice and this default is never forced back on them): a port group
  // with nothing left to do starts collapsed, so a long board folds down
  // to just the groups that still need work. Archived ports always start
  // collapsed - that section is for reference, not day-to-day work.
  const portCollapsed = portKey in collapsedGroups ? !!collapsedGroups[portKey] : (archivedView ? true : (portTotal > 0 && portLeft === 0));
  const pEsc = portName.replace(/'/g, "\\'");
  return `
    <div class="port-group">
      <div class="port-header ${portCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT' && event.target.tagName!=='A') toggleGroup('${portKey.replace(/'/g,"\\'")}')">
        ${CHEVRON}
        <input class="group-name" value="${portName === 'Unassigned' ? '' : portName}" placeholder="Unassigned port"
          onclick="event.stopPropagation()"
          onchange="renameGroup('port', '${pEsc}', '', this.value, 'Unassigned')">
      </div>
      <div class="port-body ${portCollapsed ? 'collapsed' : ''}">${vesselsHtml}</div>
    </div>`;
}

function render() {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const operatorFilterEl = document.getElementById('operatorFilter');
  if (operatorFilterEl) {
    const operators = [...new Set(records.map(r => r.created_by).filter(Boolean))].sort();
    const current = operatorFilterEl.value;
    operatorFilterEl.innerHTML = '<option value="">All operators</option>' + operators.map(a => `<option value="${a.replace(/"/g,'&quot;')}">${a}</option>`).join('');
    if (operators.includes(current)) operatorFilterEl.value = current;
    syncGlassSelectLabel('operatorFilter');
  }
  const operatorFilter = operatorFilterEl ? operatorFilterEl.value : '';

  const base = records.filter(r => r.bl_number.toLowerCase().includes(q) && (!operatorFilter || r.created_by === operatorFilter));
  const activeList = base.filter(r => !r.archived);
  const archivedList = base.filter(r => !!r.archived);

  const ports = groupRecordsByPortVessel(activeList);
  const portNames = sortedPortNames(ports);

  // Port landing nav - rebuilt fresh every render from whatever ports are
  // currently on the board (post search/operator filter), same pattern
  // as jumpSelect/operatorFilter just above. If the previously-selected
  // port disappeared (renamed away, last BL removed, etc.) fall back to
  // "All ports" rather than showing an empty board.
  if (selectedPortTab && !portNames.includes(selectedPortTab)) selectedPortTab = '';
  const portTabsEl = document.getElementById('portTabs');
  if (portTabsEl) {
    if (portNames.length === 0) {
      portTabsEl.innerHTML = '';
    } else if (selectedPortTab === '') {
      // Landing state: one clickable card per port, each summarizing that
      // port's vessel/BL counts and completion progress.
      const anchorIcon = '<svg viewBox="0 0 24 24"><circle cx="12" cy="5" r="3"></circle><line x1="12" y1="22" x2="12" y2="8"></line><path d="M5 12H2a10 10 0 0020 0h-3"></path></svg>';
      const cards = portNames.map(portName => {
        const vessels = ports[portName];
        const vesselCount = Object.keys(vessels).length;
        const portRecords = Object.values(vessels).reduce((all, list) => all.concat(list), []);
        const blCount = portRecords.length;
        const completeCount = portRecords.filter(r => r.invoice_issued && r.approval_received && r.do_issued).length;
        const pct = blCount ? Math.round(completeCount / blCount * 100) : 0;
        const done = pct >= 100;
        const label = portName.replace(/</g, '&lt;').replace(/>/g, '&gt;');
        const pEsc = portName.replace(/'/g, "\\'");
        return `<button type="button" class="port-card" onclick="selectPortTab('${pEsc}')">
          <div class="port-card-icon">${anchorIcon}</div>
          <div class="port-card-name">${label}</div>
          <div class="port-card-meta">${vesselCount} vessel${vesselCount === 1 ? '' : 's'} &middot; ${blCount} BL${blCount === 1 ? '' : 's'}</div>
          <span class="vessel-progress port-card-progress ${done ? 'done' : ''}"><span class="vessel-progress-fill" style="width:${pct}%"></span></span>
          <div class="port-card-pct ${done ? 'done' : ''}">${done ? 'All done' : pct + '% complete'}</div>
        </button>`;
      }).join('');
      portTabsEl.innerHTML = `<div class="port-card-grid">${cards}</div>`;
    } else {
      // Drilled-in state: a port is selected, so the card grid is hidden
      // and replaced by a small breadcrumb/back control + heading above
      // that one port's (unchanged) groups.
      const label = selectedPortTab.replace(/</g, '&lt;').replace(/>/g, '&gt;');
      portTabsEl.innerHTML = `<div class="port-breadcrumb">
        <button type="button" class="port-back-btn" onclick="selectPortTab('')">&larr; All Ports</button>
        <h2 class="port-breadcrumb-heading">${label}</h2>
      </div>`;
    }
  }
  // Landing state (no port selected) shows only the card grid above - the
  // whole point of drilling in is that the board isn't also dumped below
  // it. The one exception is an active search: if the person is searching
  // for a BL, the grid cards don't show BL numbers, so bypass the
  // grid-only restriction and surface matching results across all ports
  // (same ports/groups the grid itself was just filtered down to above).
  const displayPortNames = selectedPortTab ? portNames.filter(p => p === selectedPortTab) : (q ? portNames : []);

  const groupsEl = document.getElementById('groups');
  if (portNames.length === 0) {
    groupsEl.innerHTML = '<div style="color:var(--muted); padding:24px 4px;">No BLs on the board yet. Upload an Excel manifest above to get started.</div>';
  } else {
    // Vessel/port jump menu - with 20-25 manifests a month, scrolling down
    // the whole board to find one vessel doesn't scale. Built fresh every
    // render so it always reflects what's actually on the board right now.
    const jumpEl = document.getElementById('jumpSelect');
    if (jumpEl) {
      const current = jumpEl.value;
      let options = '<option value="">Select a vessel to view</option>';
      portNames.forEach(portName => {
        const vessels = ports[portName];
        sortedVesselNames(vessels).forEach(vesselName => {
          const key = 'vessel:' + portName + ':' + vesselName;
          options += `<option value="${key.replace(/"/g, '&quot;')}">${vesselName.replace(/"/g, '&quot;')}</option>`;
        });
      });
      jumpEl.innerHTML = options;
      if ([...jumpEl.options].some(o => o.value === current)) jumpEl.value = current;
      syncGlassSelectLabel('jumpSelect');
    }

    groupsEl.innerHTML = displayPortNames.map(portName => {
      const vessels = ports[portName];
      return portGroupHtml(portName, sortedVesselNames(vessels), vessels, false, !!selectedPortTab);
    }).join('');
  }

  const archivedCard = document.getElementById('archivedCard');
  if (archivedCard) {
    if (archivedList.length === 0) {
      archivedCard.style.display = 'none';
    } else {
      archivedCard.style.display = '';
      document.getElementById('archivedCount').textContent = archivedList.length + ' BL' + (archivedList.length === 1 ? '' : 's');
      const archivedGroupsEl = document.getElementById('archivedGroups');
      archivedGroupsEl.style.display = archivedSectionOpen ? '' : 'none';
      if (archivedSectionOpen) {
        const archivedPorts = groupRecordsByPortVessel(archivedList);
        archivedGroupsEl.innerHTML = sortedPortNames(archivedPorts).map(portName => {
          const vessels = archivedPorts[portName];
          return portGroupHtml(portName, sortedVesselNames(vessels), vessels, true);
        }).join('');
      }
    }
  }

  updateSummaryOnly();
  updateBulkBars();
}

async function setVesselEta(blNumbers, eta) {
  blNumbers.forEach(bl => {
    const rec = records.find(r => r.bl_number === bl);
    if (rec) rec.eta = eta;
  });
  suppressPollUntil = Date.now() + 1500;
  render();
  await fetch('/api/vessel/eta', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers, eta})
  });
  await fetchRecords();
}

async function setVesselArchived(blNumbers, archived) {
  blNumbers.forEach(bl => {
    const rec = records.find(r => r.bl_number === bl);
    if (rec) rec.archived = archived ? 1 : 0;
  });
  suppressPollUntil = Date.now() + 1500;
  render();
  await fetch('/api/vessel/archive', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers, archived})
  });
  showToast(archived ? 'Vessel archived - find it under "Archived vessels" below.' : 'Vessel restored to the board.');
  await fetchRecords();
}

function selectPortTab(port) {
  selectedPortTab = port;
  render();
  // The grid and the drilled-in groups are never shown at once any more,
  // so bring whichever one is now visible (the breadcrumb+groups, or back
  // up to the card grid) into view.
  requestAnimationFrame(() => {
    const el = document.getElementById('portTabs');
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

function jumpToVessel(key) {
  if (!key) return;
  const parts = key.split(':');
  const portKey = 'port:' + parts[1];
  // Jumping to a vessel should land on that vessel's own groups, which
  // only render when its port is the selected (drilled-in) one - so
  // switch to that vessel's actual port rather than resetting to the
  // "All ports" grid, which wouldn't show the vessel at all.
  selectedPortTab = parts[1];
  collapsedGroups[portKey] = false;
  collapsedGroups[key] = false;
  // Reset the dropdown's value before re-rendering (which would otherwise
  // restore it to this same key) - a <select> only fires 'change' when its
  // value actually changes, so without this, picking the same vessel twice
  // in a row would do nothing the second time.
  const jumpEl = document.getElementById('jumpSelect');
  if (jumpEl) jumpEl.value = '';
  syncGlassSelectLabel('jumpSelect');
  render();
  requestAnimationFrame(() => {
    const el = document.getElementById('group_' + cssEscape(key));
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

function scrollToManifestForm() {
  const el = document.getElementById('manifestCard');
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
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

/* ---------- Custom glass dropdowns ----------
   Native <select> popups (the list that appears when you click a <select>)
   cannot be styled with CSS in any browser - the OS/browser paints its own
   plain popup no matter what. So for portField/jumpSelect/operatorFilter we
   build a custom trigger+panel on top of the real <select>, which stays in
   the DOM (just visually hidden via the cs-native-hidden class) so every
   existing .value read/write, onchange="..." handler, and the dynamic
   innerHTML option-rebuilding for jumpSelect/operatorFilter in render()
   keep working completely unchanged. The custom panel's rows are rebuilt
   from the real select's current <option> list every time it's opened, so
   they can never go stale relative to whatever render() last put there. */
const customSelects = {};

function initGlassSelects() {
  ['portField', 'jumpSelect', 'operatorFilter'].forEach(id => {
    const select = document.getElementById(id);
    if (!select || customSelects[id]) return;
    const wrap = select.closest('.glass-select-wrap');
    if (!wrap) return;

    select.classList.add('cs-native-hidden');
    select.setAttribute('tabindex', '-1');

    const trigger = document.createElement('div');
    trigger.className = 'cs-trigger';
    trigger.tabIndex = 0;
    trigger.setAttribute('role', 'button');
    trigger.setAttribute('aria-haspopup', 'listbox');
    trigger.setAttribute('aria-expanded', 'false');
    trigger.innerHTML =
      '<span class="cs-trigger-label"></span>' +
      '<svg class="cs-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" ' +
      'stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>';
    wrap.appendChild(trigger);

    const panel = document.createElement('div');
    panel.className = 'cs-panel';
    panel.setAttribute('role', 'listbox');
    panel.dataset.for = id;
    document.body.appendChild(panel);

    const state = { select, wrap, trigger, panel, highlight: -1, open: false };
    customSelects[id] = state;

    trigger.addEventListener('click', () => toggleGlassSelect(id));
    trigger.addEventListener('keydown', e => onGlassTriggerKeydown(id, e));

    syncGlassSelectLabel(id);
  });
}

function syncGlassSelectLabel(id) {
  const state = customSelects[id];
  if (!state) return;
  const { select, trigger } = state;
  const opt = select.options[select.selectedIndex];
  trigger.querySelector('.cs-trigger-label').textContent = opt ? opt.textContent : '';
}

function positionGlassPanel(id) {
  const { wrap, panel } = customSelects[id];
  const r = wrap.getBoundingClientRect();
  panel.style.left = r.left + 'px';
  panel.style.top = (r.bottom + 6) + 'px';
  panel.style.width = Math.max(r.width, 160) + 'px';
  // Flip above the trigger if there isn't room below (e.g. jumpSelect's
  // panel can be long and the search row sits mid-page).
  const estHeight = Math.min(panel.scrollHeight || 280, 280);
  if (r.bottom + 6 + estHeight > window.innerHeight && r.top - 6 - estHeight > 0) {
    panel.style.top = (r.top - 6 - estHeight) + 'px';
    panel.style.transform = 'translateY(-100%)';
  } else {
    panel.style.transform = 'none';
  }
}

function buildGlassPanelOptions(id) {
  const state = customSelects[id];
  const { select, panel } = state;
  panel.innerHTML = '';
  state.highlight = -1;
  [...select.options].forEach((opt, i) => {
    const row = document.createElement('div');
    row.className = 'cs-option';
    row.setAttribute('role', 'option');
    row.dataset.value = opt.value;
    const selected = i === select.selectedIndex;
    row.setAttribute('aria-selected', selected ? 'true' : 'false');
    if (selected) { row.classList.add('cs-selected'); state.highlight = i; }
    row.innerHTML =
      '<span class="cs-option-label"></span>' +
      '<svg class="cs-check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" ' +
      'stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>';
    row.querySelector('.cs-option-label').textContent = opt.textContent;
    row.addEventListener('mousedown', e => e.preventDefault()); // keep focus on trigger, not the row
    row.addEventListener('click', () => chooseGlassOption(id, i));
    row.addEventListener('mouseenter', () => setGlassHighlight(id, i));
    panel.appendChild(row);
  });
}

function setGlassHighlight(id, index) {
  const { panel } = customSelects[id];
  customSelects[id].highlight = index;
  [...panel.children].forEach((row, i) => row.classList.toggle('cs-highlight', i === index));
  const row = panel.children[index];
  if (row) row.scrollIntoView({ block: 'nearest' });
}

function chooseGlassOption(id, index) {
  const { select } = customSelects[id];
  const opt = select.options[index];
  if (!opt) return;
  select.value = opt.value;
  select.dispatchEvent(new Event('change'));
  syncGlassSelectLabel(id);
  closeGlassSelect(id);
  customSelects[id].trigger.focus();
}

function openGlassSelect(id) {
  const state = customSelects[id];
  if (!state || state.open) return;
  Object.keys(customSelects).forEach(other => { if (other !== id) closeGlassSelect(other); });
  buildGlassPanelOptions(id);
  state.panel.classList.add('cs-open');
  positionGlassPanel(id);
  // scrollHeight is only known once it's visible, so position once more now that it's rendered
  positionGlassPanel(id);
  state.trigger.classList.add('cs-open');
  state.trigger.setAttribute('aria-expanded', 'true');
  state.open = true;
  if (state.highlight >= 0) setGlassHighlight(id, state.highlight);
  document.addEventListener('mousedown', glassOutsideHandler, true);
  window.addEventListener('scroll', glassScrollHandler, true);
  window.addEventListener('resize', glassScrollHandler, true);
}

function closeGlassSelect(id) {
  const state = customSelects[id];
  if (!state || !state.open) return;
  state.panel.classList.remove('cs-open');
  state.trigger.classList.remove('cs-open');
  state.trigger.setAttribute('aria-expanded', 'false');
  state.open = false;
  document.removeEventListener('mousedown', glassOutsideHandler, true);
  window.removeEventListener('scroll', glassScrollHandler, true);
  window.removeEventListener('resize', glassScrollHandler, true);
}

function toggleGlassSelect(id) {
  if (customSelects[id] && customSelects[id].open) closeGlassSelect(id); else openGlassSelect(id);
}

function glassOutsideHandler(e) {
  Object.keys(customSelects).forEach(id => {
    const state = customSelects[id];
    if (state.open && !state.trigger.contains(e.target) && !state.panel.contains(e.target)) closeGlassSelect(id);
  });
}

function glassScrollHandler() {
  Object.keys(customSelects).forEach(id => { if (customSelects[id].open) positionGlassPanel(id); });
}

function onGlassTriggerKeydown(id, e) {
  const state = customSelects[id];
  if (!state) return;
  if (e.key === 'Escape') { if (state.open) { e.preventDefault(); e.stopPropagation(); closeGlassSelect(id); } return; }
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    if (!state.open) { openGlassSelect(id); return; }
    if (state.highlight >= 0) chooseGlassOption(id, state.highlight);
    return;
  }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    if (!state.open) { openGlassSelect(id); return; }
    const count = state.panel.children.length;
    if (!count) return;
    let next = state.highlight + (e.key === 'ArrowDown' ? 1 : -1);
    next = Math.max(0, Math.min(count - 1, next));
    setGlassHighlight(id, next);
    return;
  }
  if (e.key === 'Tab' && state.open) closeGlassSelect(id);
}

// Keyboard shortcuts: "/" focuses search, Escape collapses all groups -
// both no-ops while the person is actually typing, and Escape defers to
// whatever more specific thing (a glass-dropdown panel, the history modal)
// is already open, so it never fights with that control's own handling.
document.addEventListener('keydown', (e) => {
  const active = document.activeElement;
  const tag = active ? active.tagName : '';
  const isTyping = tag === 'INPUT' || tag === 'TEXTAREA' || (active && active.isContentEditable);

  if (e.key === '/' && !isTyping) {
    e.preventDefault();
    const box = document.getElementById('searchBox');
    if (box) box.focus();
    return;
  }

  if (e.key === 'Escape') {
    if (isTyping) return;
    // A glass-select panel being open means its own trigger keydown
    // handler (onGlassTriggerKeydown) already closed it and stopped this
    // keystroke from bubbling here - this is just a safety net.
    if (Object.values(customSelects).some(s => s.open)) return;
    const historyOverlay = document.getElementById('historyOverlay');
    if (historyOverlay && historyOverlay.style.display !== 'none') { closeHistory(); return; }
    if (document.querySelector('.bulk-bar.active')) return;
    setAllGroupsCollapsed(true);
  }
});

initGlassSelects();

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
