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
import zipfile
from urllib.parse import urlparse
import csv
import unicodedata
import html.parser
import io
import time
import secrets
import base64
import hashlib
import hmac
import json
import struct
import smtplib
import socket
import ssl
from email.message import EmailMessage
from email.utils import formataddr
from datetime import datetime, date, timedelta
from functools import wraps
from urllib.parse import quote as url_quote
import segno
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

# Session cookie hardening + "keep me signed in". Without the tick the
# cookie dies when the browser closes; with it, it lasts 30 days.
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = bool(os.environ.get("RENDER"))  # HTTPS on Render; plain http locally
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=30)

# Shared Sea Power logo (used in-page and as the browser-tab favicon).
LOGO_B64 = "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"


# Shared English/Arabic translation dictionary + helper JS, injected into
# every page's <script> block (via `""" + I18N_JS + """`) so every page
# draws from the same terms - a BL stays "Bill of Lading" the same way on
# every screen instead of getting re-translated slightly differently each
# time. Keys are generic (not DO-Tracker-specific) on purpose so the same
# dictionary grows to cover the other pages without restructuring.
#
# How a page uses this:
#   - Static HTML text: add data-i18n="key" (textContent), data-i18n-ph="key"
#     (placeholder) or data-i18n-title="key" (title) to the element; applyI18n()
#     fills it in from I18N[currentLang].
#   - JS-generated HTML (template literals in render()-type functions):
#     call t('key') or t('key', {var: value}) directly instead of hardcoding
#     English text, so it already comes out right on every re-render.
#   - The language itself is NOT saved to the account - it's a per-browser
#     localStorage choice (same mechanism as the existing light/dark toggle),
#     so it doesn't follow a person to a different device or session.
I18N_JS = """
const I18N = {
  en: {
    // Topbar
    app_tagline: "DO Tracker",
    manage_users: "Manage Users",
    signed_in_as: "Signed in as",
    log_out: "Log out",
    toggle_dark_mode: "Toggle dark mode",
    switch_language: "Switch language",
    // Manifest card
    add_a_manifest: "Add a manifest",
    discharge_port: "Discharge Port",
    select_a_port: "Select a port...",
    port_dammam: "Dammam Port",
    port_jubail: "Jubail Commercial Port",
    port_jeddah: "Jeddah Port",
    port_yanbu_commercial: "Yanbu Commercial Port",
    port_yanbu_industrial: "Yanbu Industrial Port",
    port_kap: "KAP",
    vessel_label: "Vessel",
    click_to_upload: "Click to upload",
    or_drag_drop_manifest: "or drag & drop your manifest (one or several files)",
    n_files_selected: "{n} files: {names}",
    adding_progress: "Adding {i} of {n}...",
    file_failed: "{name}: {error}",
    manifest_dropzone_sub: ".xlsx, .xls, .csv, .docx or .pdf - the BL numbers are read automatically",
    add_to_board: "Add to board",
    vessel_placeholder: "e.g. TAI KNIGHT",
    adding_ellipsis: "Adding...",
    // Attach documents card
    attach_documents: "Attach documents",
    attach_results_title: "Attaching documents",
    attach_batch_done: "Documents: {summary}",
    attach_docs_help: "Drop Invoice / Delivery Order PDFs here - each one is read and matched to its BL automatically, same as the manifest upload above.",
    attach_docs_help2: "chips next to a BL number below show what's already attached.",
    or_drag_drop_docs: "or drag & drop Invoice/DO PDFs",
    auto_match_dropzone_sub: "Drop as many at once as you like - each is matched to its BL automatically",
    // Toolbar / search
    search_bl_placeholder: "Search BL number...",
    search_shortcut_title: "Shortcut: press / from anywhere to search. Enter jumps to the BL, Esc clears.",
    no_bl_match: 'No BL matches "{q}".',
    many_bls_match: "{n} BLs match - keep typing to narrow it down.",
    select_vessel_to_view: "Select a vessel to view",
    all_operators: "All operators",
    collapse_all: "Collapse all",
    archived_vessels: "Archived vessels",
    scroll_top_title: "Back to Discharge Port / Vessel",
    // Stats
    total_bls: "Total BLs",
    remaining: "Remaining",
    fully_complete: "Fully Complete",
    // Port landing / breadcrumb
    all_ports_back: "All Ports",
    all_done: "All done",
    pct_complete: "{pct}% complete",
    port_card_meta: "{vesselCount} vessel{vp} · {blCount} BL{bp}",
    no_bls_yet: "No BLs on the board yet. Upload an Excel manifest above to get started.",
    // Table headers
    th_bl_number: "BL Number",
    th_invoice_issued: "Invoice Issued",
    th_approval_received: "Approval Received",
    th_do_issued: "DO Issued",
    th_remarks: "Remarks",
    // Row actions
    remove: "Remove",
    history: "History",
    notes_placeholder: "notes...",
    unassigned_vessel_ph: "Unassigned vessel",
    unassigned_port_ph: "Unassigned port",
    complete_badge: "Complete",
    // Vessel group header
    eta_label: "ETA",
    not_set: "Not set",
    expected_arrival_title: "Expected arrival",
    export: "Export",
    export_vessel_title: "Export this vessel to Excel",
    archive: "Archive",
    unarchive: "Unarchive",
    remove_all: "Remove all",
    select_deselect_vessel_title: "Select/deselect all in this vessel",
    bl_count: "{n} BL{p}",
    left_suffix: " · {n} left",
    done_suffix: " · done",
    // Bulk bar
    mark_invoice_issued: "Mark Invoice Issued",
    mark_approval_received: "Mark Approval Received",
    mark_do_issued: "Mark DO Issued",
    unmark: "Unmark",
    selected_count: "{n} selected",
    // Documents modal / doc chips
    documents_title: "Documents",
    documents_for: "Documents - {bl}",
    loading: "Loading...",
    could_not_load_bl: "Could not load this BL.",
    doc_invoice: "Invoice",
    doc_delivery_order: "Delivery Order",
    doc_chip_inv: "INV",
    doc_chip_do: "DO",
    attached_label: "Attached: {filename}",
    by_at: "by {user} - {time}",
    not_attached_yet: "Not attached yet.",
    not_attached_waiting: "Not attached yet - waiting on the issuing staff member.",
    download: "Download",
    replace: "Replace",
    upload_pdf: "Upload PDF",
    attached_tooltip: "{full} attached - click to view",
    not_attached_tooltip: "{full} not attached yet - click to upload",
    // Auto-match batch
    select_bl_ellipsis: "Select BL...",
    kind_ellipsis: "Kind...",
    attach_btn: "Attach",
    attaching_ellipsis: "Attaching...",
    reading_ellipsis: "Reading...",
    too_large_skipped: "Too large (over 10MB) - skipped.",
    could_not_read_file: "Could not read this file.",
    matches_multiple: "Matches more than one BL - pick the right one",
    could_not_tell_kind: "Couldn't tell Invoice from Delivery Order",
    no_bl_matched: "No BL on your board matched this document",
    unrecognized_document: "Unrecognized document",
    marked_issued_suffix: " · marked issued",
    attached_automatically: "{n} attached automatically",
    need_your_input: ", {n} need your input",
    processing_more: " (processing {n} more...)",
    drop_pdf_only: "Drop PDF files only.",
    pick_bl_and_kind: "Pick both a BL number and a document kind.",
    could_not_remove_file: "Could not remove the file.",
    doc_removed: "{kind} removed.",
    upload_failed: "Upload failed.",
    attach_failed: "Attach failed.",
    // History modal
    history_for: "History - {bl}",
    could_not_load_history: "Could not load history.",
    no_history_yet: "No history recorded yet.",
    unknown_user: "Unknown",
    action_added: "Added to board",
    action_deleted: "Removed",
    action_restored: "Restored",
    action_toggle: "status changed",
    action_remarks: "Remarks edited",
    action_attachment: "attached a document",
    action_attachment_removed: "removed a document",
    history_customer_download: "Customer downloaded the invoice through the tracking link",
    yes: "Yes",
    no: "No",
    history_set_field: "set {field} to {value}",
    history_edited_remarks: 'edited remarks: "{value}"',
    history_cleared_remarks: "edited remarks (cleared)",
    history_added_this_bl: "added this BL ({value})",
    history_added_bl_plain: "added this BL",
    // Toasts
    choose_manifest_first: "Choose a manifest file first.",
    bl_records_added: "{added} new BL record(s) added",
    already_on_board_skipped: ", {skipped} already on the board (skipped)",
    heads_up_duplicate: "Heads up - {n} BL(s) already exist under a different vessel: {lines}{more}.",
    already_under: "{bl} (already under {vessel} / {port})",
    and_n_more: " and {n} more",
    could_not_save_retry: "Could not save that change - please try again.",
    // Manifest preview (check before adding)
    mp_title: "Check before adding",
    mp_reading: "Reading the manifest...",
    mp_read_failed: "Could not read the file(s) - please try again.",
    mp_destination: "Adding to {port} · {vessel}",
    mp_no_vessel: "No vessel typed in - these will go under Unassigned.",
    mp_no_port: "No discharge port selected - these will go under Unassigned.",
    mp_found: "{n} BL(s) found",
    mp_summary: "{sel} of {total} selected",
    mp_add_n: "Add {n} BL(s)",
    mp_nothing_new: "Nothing new to add",
    mp_adding: "Adding...",
    mp_cancel: "Cancel",
    mp_toggle_file: "Select / clear all in this file",
    mp_status_new: "New",
    mp_status_on_board: "Already on this vessel",
    mp_status_elsewhere: "Already under {vessel} / {port}",
    mp_status_repeat: "Listed twice - added once",
    mp_note_hidden_row: "Hidden row in Excel - tick to add",
    mp_note_hidden_sheet: "On a hidden sheet - tick to add",
    mp_note_struck: "Crossed out in the file - tick to add",
    mp_err_unsupported: "Not a supported file type (.xlsx, .xls, .csv, .docx, .pdf).",
    mp_err_unreadable: "Could not open this file - it may be damaged or password-protected.",
    mp_err_no_bls: "No BL numbers found - the file needs a B/L NO. column.",
    mp_err_too_large: "This file is larger than 25MB.",
    mp_err_empty: "This file is empty.",
    mp_contact_title: "E-mail: {email} · Phone: {phone}",
    already_has_file_confirm: "Matches {bl}, which already has a {kind} attached - pick it below to replace it",
    member_match_confirm: "Names one of the BLs in the combined entry {bl} - check it and attach",
    mp_status_taken: "Already on the board (another user's)",
    mp_status_fills: "Already on the board - adds missing contact details",
    mp_save_contacts_n: "Save contact details for {n} BL(s)",
    mp_summary_fills: " · contact details for {n} existing",
    mp_warn_truncated: "Very large file - only the first 20,000 rows were read.",
    contacts_updated: "Contact details added to {n} BL(s) already on the board.",
    left_out_hidden: "{n} BL(s) in hidden or crossed-out rows were left out.",
    only_pdf_accepted: "Only PDF files are accepted.",
    file_too_large: "That file is larger than 10MB.",
    uploaded_marked_issued: "{label} uploaded - marked as issued.",
    uploaded: "{label} uploaded.",
    removed_bl: "Removed BL {bl}.",
    restored_bl: "Restored BL {bl}.",
    nothing_to_remove: "Nothing to remove.",
    confirm_remove_all: "Remove all {n} BL{p}{inLabel}?",
    in_label: " in {label}",
    bls_removed: "{n} BL{p} removed.",
    restored: "Restored.",
    select_at_least_one: "Select at least one BL first.",
    bls_updated: "{n} BL(s) updated.",
    vessel_archived: 'Vessel archived - find it under "Archived vessels" below.',
    vessel_restored: "Vessel restored to the board.",
    undo: "Undo",
    confirm: "Confirm",
    // Customer sharing (Documents popup)
    customer_contacts: "Customer contacts",
    consignee_name: "Consignee",
    consignee_email: "Consignee email",
    consignee_phone: "Consignee phone",
    broker_email: "Customs broker email",
    broker_phone: "Customs broker phone",
    save_contacts: "Save contacts",
    contacts_saved: "Contacts saved.",
    no_contacts: "No contact details yet.",
    share_with_customer: "Tracking link",
    share_help: "A private page where the customer can check this BL's status and download the invoice - no login.",
    create_link: "Create tracking link",
    copy_link: "Copy link",
    link_copied: "Tracking link copied.",
    show_qr: "QR code",
    download_qr: "Download QR (PNG)",
    new_link: "New link",
    new_link_title: "Replace this link - the old one stops working",
    confirm_new_link: "Replace the tracking link? The old link will stop working.",
    link_replaced: "New link created - the old link no longer works.",
    notify_customer: "Notify customer",
    email_do_to: "Email the DO to:",
    recipient_consignee: "Consignee",
    recipient_broker: "Broker",
    send_do_email: "Email the DO",
    sending_ellipsis: "Sending...",
    do_emailed: "DO emailed to {to}.",
    attach_do_first: "Attach the Delivery Order PDF to email it.",
    email_not_setup: "Email isn't set up yet - an admin needs to add the company mailbox.",
    whatsapp_consignee: "WhatsApp consignee",
    whatsapp_broker: "WhatsApp broker",
    whatsapp_help: "Opens WhatsApp with the message and tracking link ready - you press Send.",
    err_bad_email: "That doesn't look like an email address.",
    err_bad_phone: "That doesn't look like a phone number.",
    err_no_recipient: "Tick at least one recipient that has an email address.",
    err_no_do: "Attach the Delivery Order PDF to email it.",
    err_email_not_configured: "Email isn't set up yet - an admin needs to add the company mailbox.",
    err_send_failed: "The email could not be sent. Check the mailbox settings and try again.",
    err_email_unreachable: "Couldn't reach the mail server. Check the mail host and port, or whether the hosting plan allows outgoing email.",
    err_email_auth: "The mailbox rejected the login. Check the mailbox username and password with IT.",
    err_email_recipient: "The recipient's email address was refused. Check the address and try again.",
    err_email_timeout: "The email is taking too long, so it was stopped. Nothing was confirmed as sent. Try again in a moment.",
    err_no_phone: "No phone number saved for that contact.",
    contacts_found: " · contact details found for {n}",
    action_contacts: "updated the contact details",
    action_link_reset: "replaced the tracking link",
    history_notified_email: "emailed the DO to {value}",
    history_notified_whatsapp: "opened WhatsApp to {value}",
  },
  ar: {
    app_tagline: "متتبع أوامر التسليم",
    manage_users: "إدارة المستخدمين",
    signed_in_as: "مسجّل الدخول باسم",
    log_out: "تسجيل الخروج",
    toggle_dark_mode: "تبديل الوضع الداكن",
    switch_language: "تبديل اللغة",
    add_a_manifest: "إضافة بيان شحن",
    discharge_port: "ميناء التفريغ",
    select_a_port: "اختر ميناء...",
    port_dammam: "ميناء الدمام",
    port_jubail: "ميناء الجبيل التجاري",
    port_jeddah: "ميناء جدة",
    port_yanbu_commercial: "ميناء ينبع التجاري",
    port_yanbu_industrial: "ميناء ينبع الصناعي",
    port_kap: "KAP",
    vessel_label: "السفينة",
    click_to_upload: "اضغط للرفع",
    or_drag_drop_manifest: "أو اسحب وأفلت بيان الشحن (ملف واحد أو عدة ملفات)",
    n_files_selected: "{n} ملفات: {names}",
    adding_progress: "جارٍ إضافة {i} من {n}...",
    file_failed: "{name}: {error}",
    manifest_dropzone_sub: "\\u2066.xlsx, .xls, .csv, .docx, .pdf\\u2069 - تُقرأ أرقام البوالص تلقائيًا",
    add_to_board: "إضافة إلى اللوحة",
    vessel_placeholder: "مثال: TAI KNIGHT",
    adding_ellipsis: "جارٍ الإضافة...",
    attach_documents: "إرفاق المستندات",
    attach_results_title: "نتائج إرفاق المستندات",
    attach_batch_done: "المستندات: {summary}",
    attach_docs_help: "أسقط ملفات الفاتورة / أمر التسليم (PDF) هنا - تتم قراءة كل ملف ومطابقته تلقائيًا مع رقم البوليصة، بنفس طريقة رفع بيان الشحن أعلاه.",
    attach_docs_help2: "الشارات الظاهرة بجانب رقم البوليصة أدناه توضح المستندات المرفقة بالفعل.",
    or_drag_drop_docs: "أو اسحب وأفلت ملفات الفاتورة / أمر التسليم",
    auto_match_dropzone_sub: "أسقط أي عدد من الملفات دفعة واحدة - تتم مطابقة كل ملف تلقائيًا مع رقم البوليصة الخاص به",
    search_bl_placeholder: "ابحث برقم البوليصة...",
    search_shortcut_title: "اختصار: اضغط / من أي مكان للبحث. Enter للانتقال إلى البوليصة، وEsc للمسح.",
    no_bl_match: 'لا توجد بوليصة تطابق "{q}".',
    many_bls_match: "{n} بوليصة تطابق البحث - تابع الكتابة لتضييق النتائج.",
    select_vessel_to_view: "اختر سفينة للعرض",
    all_operators: "جميع الموظفين",
    collapse_all: "طي الكل",
    archived_vessels: "السفن المؤرشفة",
    scroll_top_title: "الرجوع إلى ميناء التفريغ / السفينة",
    total_bls: "إجمالي البوالص",
    remaining: "المتبقي",
    fully_complete: "مكتمل بالكامل",
    all_ports_back: "جميع الموانئ",
    all_done: "اكتمل الكل",
    pct_complete: "{pct}% مكتمل",
    port_card_meta: "{vesselCount} سفينة · {blCount} بوليصة",
    no_bls_yet: "لا توجد بوالص على اللوحة بعد. قم برفع بيان شحن (إكسل) أعلاه للبدء.",
    th_bl_number: "رقم البوليصة",
    th_invoice_issued: "صدور الفاتورة",
    th_approval_received: "استلام الموافقة",
    th_do_issued: "صدور أمر التسليم",
    th_remarks: "ملاحظات",
    remove: "إزالة",
    history: "السجل",
    notes_placeholder: "ملاحظات...",
    unassigned_vessel_ph: "سفينة غير محددة",
    unassigned_port_ph: "ميناء غير محدد",
    complete_badge: "مكتمل",
    eta_label: "الوصول المتوقع",
    not_set: "غير محدد",
    expected_arrival_title: "تاريخ الوصول المتوقع",
    export: "تصدير",
    export_vessel_title: "تصدير بيانات هذه السفينة إلى إكسل",
    archive: "أرشفة",
    unarchive: "إلغاء الأرشفة",
    remove_all: "إزالة الكل",
    select_deselect_vessel_title: "تحديد/إلغاء تحديد الكل في هذه السفينة",
    bl_count: "{n} بوليصة",
    left_suffix: " · متبقي {n}",
    done_suffix: " · مكتمل",
    mark_invoice_issued: "تمييز: صدور الفاتورة",
    mark_approval_received: "تمييز: استلام الموافقة",
    mark_do_issued: "تمييز: صدور أمر التسليم",
    unmark: "إلغاء التمييز",
    selected_count: "تم تحديد {n}",
    documents_title: "المستندات",
    documents_for: "المستندات - {bl}",
    loading: "جارٍ التحميل...",
    could_not_load_bl: "تعذر تحميل بيانات هذه البوليصة.",
    doc_invoice: "الفاتورة",
    doc_delivery_order: "أمر التسليم",
    doc_chip_inv: "فاتورة",
    doc_chip_do: "تسليم",
    attached_label: "مرفق: {filename}",
    by_at: "بواسطة {user} - {time}",
    not_attached_yet: "لم يتم الإرفاق بعد.",
    not_attached_waiting: "لم يتم الإرفاق بعد - بانتظار الموظف المصدر.",
    download: "تحميل",
    replace: "استبدال",
    upload_pdf: "رفع ملف PDF",
    attached_tooltip: "{full} مرفقة - اضغط للعرض",
    not_attached_tooltip: "{full} غير مرفقة بعد - اضغط للرفع",
    select_bl_ellipsis: "اختر رقم البوليصة...",
    kind_ellipsis: "النوع...",
    attach_btn: "إرفاق",
    attaching_ellipsis: "جارٍ الإرفاق...",
    reading_ellipsis: "جارٍ القراءة...",
    too_large_skipped: "الحجم كبير جدًا (أكثر من 10 ميجابايت) - تم التخطي.",
    could_not_read_file: "تعذرت قراءة هذا الملف.",
    matches_multiple: "يطابق أكثر من بوليصة - يرجى اختيار الصحيحة",
    could_not_tell_kind: "تعذر تحديد ما إذا كان فاتورة أو أمر تسليم",
    no_bl_matched: "لا توجد بوليصة على لوحتك تطابق هذا المستند",
    unrecognized_document: "مستند غير معروف",
    marked_issued_suffix: " · تم التمييز كصادرة",
    attached_automatically: "{n} تم إرفاقها تلقائيًا",
    need_your_input: "، {n} بحاجة إلى إدخال يدوي",
    processing_more: " (جارٍ معالجة {n} أخرى...)",
    drop_pdf_only: "يرجى إسقاط ملفات PDF فقط.",
    pick_bl_and_kind: "يرجى اختيار رقم البوليصة ونوع المستند معًا.",
    could_not_remove_file: "تعذرت إزالة الملف.",
    doc_removed: "تمت إزالة {kind}.",
    upload_failed: "فشل الرفع.",
    attach_failed: "فشل الإرفاق.",
    history_for: "السجل - {bl}",
    could_not_load_history: "تعذر تحميل السجل.",
    no_history_yet: "لا يوجد سجل مسجل بعد.",
    unknown_user: "غير معروف",
    action_added: "تمت الإضافة إلى اللوحة",
    action_deleted: "تمت الإزالة",
    action_restored: "تمت الاستعادة",
    action_toggle: "تم تغيير الحالة",
    action_remarks: "تم تعديل الملاحظات",
    action_attachment: "أرفق مستندًا",
    action_attachment_removed: "أزال مستندًا",
    history_customer_download: "قام العميل بتحميل الفاتورة عبر رابط التتبع",
    yes: "نعم",
    no: "لا",
    history_set_field: "قام بتعيين {field} إلى {value}",
    history_edited_remarks: 'قام بتعديل الملاحظات: "{value}"',
    history_cleared_remarks: "قام بمسح الملاحظات",
    history_added_this_bl: "أضاف هذه البوليصة ({value})",
    history_added_bl_plain: "أضاف هذه البوليصة",
    choose_manifest_first: "يرجى اختيار ملف بيان الشحن أولًا.",
    bl_records_added: "تمت إضافة {added} بوليصة جديدة",
    already_on_board_skipped: "، و{skipped} موجودة مسبقًا على اللوحة (تم تخطيها)",
    heads_up_duplicate: "تنبيه - توجد {n} بوليصة مسجلة مسبقًا تحت سفينة مختلفة: {lines}{more}.",
    already_under: "{bl} (مسجلة تحت {vessel} / {port})",
    and_n_more: "، و{n} أخرى",
    could_not_save_retry: "تعذر حفظ هذا التغيير - يرجى المحاولة مرة أخرى.",
    mp_title: "راجِع قبل الإضافة",
    mp_reading: "جارٍ قراءة بيان الشحن...",
    mp_read_failed: "تعذرت قراءة الملفات - يرجى المحاولة مرة أخرى.",
    mp_destination: "ستُضاف إلى {port} · {vessel}",
    mp_no_vessel: "لم يُكتب اسم السفينة - ستُضاف تحت «غير محدد».",
    mp_no_port: "لم يُختر ميناء التفريغ - ستُضاف تحت «غير محدد».",
    mp_found: "تم العثور على {n} بوليصة",
    mp_summary: "تم تحديد {sel} من {total}",
    mp_add_n: "إضافة {n} بوليصة",
    mp_nothing_new: "لا يوجد جديد للإضافة",
    mp_adding: "جارٍ الإضافة...",
    mp_cancel: "إلغاء",
    mp_toggle_file: "تحديد / إلغاء تحديد الكل في هذا الملف",
    mp_status_new: "جديدة",
    mp_status_on_board: "موجودة على هذه السفينة",
    mp_status_elsewhere: "مسجلة تحت {vessel} / {port}",
    mp_status_repeat: "مكررة - تُضاف مرة واحدة",
    mp_note_hidden_row: "صف مخفي في الإكسل - حدّدها لإضافتها",
    mp_note_hidden_sheet: "في ورقة مخفية - حدّدها لإضافتها",
    mp_note_struck: "مشطوبة في الملف - حدّدها لإضافتها",
    mp_err_unsupported: "نوع الملف غير مدعوم (\\u2066.xlsx, .xls, .csv, .docx, .pdf\\u2069).",
    mp_err_unreadable: "تعذر فتح هذا الملف - قد يكون تالفًا أو محميًا بكلمة مرور.",
    mp_err_no_bls: "لم يتم العثور على أرقام بوالص - يجب أن يحتوي الملف على عمود رقم البوليصة.",
    mp_err_too_large: "حجم هذا الملف أكبر من 25 ميجابايت.",
    mp_err_empty: "هذا الملف فارغ.",
    mp_contact_title: "البريد: {email} · الهاتف: {phone}",
    already_has_file_confirm: "يطابق {bl}، ومرفق بها {kind} مسبقًا - اخترها أدناه لاستبداله",
    member_match_confirm: "يذكر إحدى البوالص ضمن الإدخال المجمّع {bl} - تحقّق ثم أرفِق",
    mp_status_taken: "موجودة على اللوحة لدى مستخدم آخر",
    mp_status_fills: "موجودة - تُضاف بيانات الاتصال الناقصة",
    mp_save_contacts_n: "حفظ بيانات الاتصال لـ {n} بوليصة",
    mp_summary_fills: " · بيانات اتصال لـ {n} موجودة",
    mp_warn_truncated: "ملف كبير جدًا - تمت قراءة أول 20,000 صف فقط.",
    contacts_updated: "تمت إضافة بيانات الاتصال إلى {n} بوليصة موجودة على اللوحة.",
    left_out_hidden: "تم استبعاد {n} بوليصة في صفوف مخفية أو مشطوبة.",
    only_pdf_accepted: "يُقبل فقط ملفات PDF.",
    file_too_large: "حجم هذا الملف أكبر من 10 ميجابايت.",
    uploaded_marked_issued: "تم رفع {label} - وتم تمييزها كصادرة.",
    uploaded: "تم رفع {label}.",
    removed_bl: "تمت إزالة البوليصة {bl}.",
    restored_bl: "تمت استعادة البوليصة {bl}.",
    nothing_to_remove: "لا يوجد ما يمكن إزالته.",
    confirm_remove_all: "هل تريد إزالة جميع البوالص البالغ عددها {n}{inLabel}؟",
    in_label: " في {label}",
    bls_removed: "تمت إزالة {n} بوليصة.",
    restored: "تمت الاستعادة.",
    select_at_least_one: "يرجى تحديد بوليصة واحدة على الأقل أولًا.",
    bls_updated: "تم تحديث {n} بوليصة.",
    vessel_archived: 'تمت أرشفة السفينة - يمكنك إيجادها ضمن "السفن المؤرشفة" أدناه.',
    vessel_restored: "تمت استعادة السفينة إلى اللوحة.",
    undo: "تراجع",
    confirm: "تأكيد",
    customer_contacts: "جهات اتصال العميل",
    consignee_name: "المرسل إليه",
    consignee_email: "بريد المرسل إليه",
    consignee_phone: "هاتف المرسل إليه",
    broker_email: "بريد المخلص الجمركي",
    broker_phone: "هاتف المخلص الجمركي",
    save_contacts: "حفظ جهات الاتصال",
    contacts_saved: "تم حفظ جهات الاتصال.",
    no_contacts: "لا توجد بيانات اتصال بعد.",
    share_with_customer: "رابط التتبع",
    share_help: "صفحة خاصة يتابع منها العميل حالة هذه البوليصة ويحمّل الفاتورة دون تسجيل دخول.",
    create_link: "إنشاء رابط التتبع",
    copy_link: "نسخ الرابط",
    link_copied: "تم نسخ رابط التتبع.",
    show_qr: "رمز QR",
    download_qr: "تنزيل رمز QR (PNG)",
    new_link: "رابط جديد",
    new_link_title: "استبدال الرابط - سيتوقف الرابط القديم عن العمل",
    confirm_new_link: "استبدال رابط التتبع؟ سيتوقف الرابط القديم عن العمل.",
    link_replaced: "تم إنشاء رابط جديد - الرابط القديم لم يعد يعمل.",
    notify_customer: "إشعار العميل",
    email_do_to: "إرسال أمر التسليم بالبريد إلى:",
    recipient_consignee: "المرسل إليه",
    recipient_broker: "المخلص الجمركي",
    send_do_email: "إرسال أمر التسليم",
    sending_ellipsis: "جارٍ الإرسال...",
    do_emailed: "تم إرسال أمر التسليم إلى {to}.",
    attach_do_first: "أرفق ملف أمر التسليم (PDF) لإرساله بالبريد.",
    email_not_setup: "لم يتم إعداد البريد بعد - يلزم أن يضيف المسؤول بريد الشركة.",
    whatsapp_consignee: "واتساب المرسل إليه",
    whatsapp_broker: "واتساب المخلص",
    whatsapp_help: "يفتح واتساب مع الرسالة ورابط التتبع جاهزين - اضغط إرسال.",
    err_bad_email: "صيغة البريد الإلكتروني غير صحيحة.",
    err_bad_phone: "صيغة رقم الهاتف غير صحيحة.",
    err_no_recipient: "اختر مستلمًا واحدًا على الأقل لديه بريد إلكتروني.",
    err_no_do: "أرفق ملف أمر التسليم (PDF) لإرساله بالبريد.",
    err_email_not_configured: "لم يتم إعداد البريد بعد - يلزم أن يضيف المسؤول بريد الشركة.",
    err_send_failed: "تعذر إرسال البريد. تحقق من إعدادات البريد وحاول مرة أخرى.",
    err_email_unreachable: "تعذر الوصول إلى خادم البريد. تحقق من عنوان الخادم والمنفذ، أو من أن خطة الاستضافة تسمح بإرسال البريد.",
    err_email_auth: "رفض صندوق البريد بيانات الدخول. تحقق من اسم المستخدم وكلمة المرور مع قسم تقنية المعلومات.",
    err_email_recipient: "تم رفض عنوان البريد الإلكتروني للمستلم. تحقق من العنوان وحاول مرة أخرى.",
    err_email_timeout: "استغرق إرسال البريد وقتاً طويلاً فتم إيقافه. لم يتأكد إرسال الرسالة. حاول مرة أخرى بعد قليل.",
    err_no_phone: "لا يوجد رقم هاتف محفوظ لجهة الاتصال هذه.",
    contacts_found: " · تم العثور على بيانات اتصال لـ {n}",
    action_contacts: "حدّث بيانات الاتصال",
    action_link_reset: "استبدل رابط التتبع",
    history_notified_email: "أرسل أمر التسليم بالبريد إلى {value}",
    history_notified_whatsapp: "فتح واتساب إلى {value}",
  },
};

let currentLang = 'en';
try { currentLang = localStorage.getItem('lang') || 'en'; } catch (e) {}

function t(key, vars) {
  const dict = I18N[currentLang] || I18N.en;
  let str = (key in dict) ? dict[key] : (I18N.en[key] !== undefined ? I18N.en[key] : key);
  if (vars) {
    str = str.replace(/\\{(\\w+)\\}/g, (_, k) => (vars[k] !== undefined ? vars[k] : ''));
  }
  return str;
}

function applyI18n() {
  document.documentElement.lang = currentLang === 'ar' ? 'ar' : 'en';
  document.documentElement.dir = currentLang === 'ar' ? 'rtl' : 'ltr';
  document.querySelectorAll('[data-i18n]').forEach(el => { el.textContent = t(el.getAttribute('data-i18n')); });
  document.querySelectorAll('[data-i18n-ph]').forEach(el => { el.placeholder = t(el.getAttribute('data-i18n-ph')); });
  document.querySelectorAll('[data-i18n-title]').forEach(el => { el.title = t(el.getAttribute('data-i18n-title')); });
  document.querySelectorAll('.lang-opt').forEach(b => b.classList.toggle('active', b.dataset.langBtn === currentLang));
}

function setLang(lang) {
  if (lang === currentLang) return;
  currentLang = lang;
  try { localStorage.setItem('lang', lang); } catch (e) {}
  applyI18n();
  if (typeof render === 'function') render();
  try { document.dispatchEvent(new Event('langchange')); } catch (e) {}
}
"""


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
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024   # nothing legitimate here is bigger; stops giant uploads early


@app.errorhandler(413)
def _too_big(e):
    return jsonify({"error": "That file is too large.", "error_code": "too_large"}), 413


MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024  # 10MB - comfortably more than a scanned invoice PDF needs


def _content_disposition(filename, fallback="download"):
    """Builds a Content-Disposition header that works for any filename.
    HTTP headers can only carry Latin-1 text, so a raw Arabic (or Chinese)
    filename in the header crashed the response and the file could never
    be downloaded. Standard fix (RFC 6266/5987): an ASCII-only filename=
    for old clients plus a UTF-8 filename*= that every modern browser
    prefers, so the user still gets the real name."""
    name = str(filename or "").replace("\r", " ").replace("\n", " ").strip() or fallback
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    fb_stem, fb_dot, fb_ext = fallback.rpartition(".")
    if not fb_dot:
        fb_stem, fb_ext = fallback, ""
    ascii_stem = re.sub(r"[^A-Za-z0-9_ -]+", "_", stem).strip(" _") or fb_stem
    ascii_ext = re.sub(r"[^A-Za-z0-9]+", "", ext) or fb_ext
    ascii_name = ascii_stem + ("." + ascii_ext if ascii_ext else "")
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{url_quote(name, safe='')}"


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
    # Company-email accounts + two-step verification (Google/Microsoft Authenticator).
    for _col, _ddl in (("email", "TEXT DEFAULT ''"), ("full_name", "TEXT DEFAULT ''"), ("totp_secret", "TEXT DEFAULT ''"),
                       ("totp_enabled", "INTEGER DEFAULT 0"), ("totp_last_step", "BIGINT DEFAULT 0"), ("recovery_codes", "TEXT DEFAULT ''")):
        cur.execute(f"ALTER TABLE users ADD COLUMN IF NOT EXISTS {_col} {_ddl}")
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_users_email ON users (LOWER(email)) WHERE email <> ''")
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
    # Customer sharing: who to notify about this BL (consignee and/or their
    # customs broker - which one varies per BL), and the secret code behind
    # its public tracking link (/t/<token>). The token is random, never the
    # BL number, so nobody can look up someone else's cargo by guessing.
    for col in ("consignee_email", "consignee_phone", "broker_email", "broker_phone", "track_token"):
        cur.execute(f"ALTER TABLE records ADD COLUMN IF NOT EXISTS {col} TEXT DEFAULT ''")
    cur.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_records_track_token ON records (track_token) "
        "WHERE track_token IS NOT NULL AND track_token <> ''"
    )
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
    # Holding area for a removed BL's Invoice/DO files. Removing a BL
    # cascade-deletes its attachments, so "Undo" used to bring the row back
    # with its files gone for good (while still showing Invoice/DO as
    # issued). The files are now copied here first and moved back on Undo;
    # anything left unclaimed is purged after 7 days. No foreign key on
    # purpose - the parent record no longer exists while files sit here.
    cur.execute(
        """CREATE TABLE IF NOT EXISTS record_attachments_trash (
            id SERIAL PRIMARY KEY,
            bl_number TEXT NOT NULL,
            kind TEXT NOT NULL,
            filename TEXT DEFAULT '',
            content_type TEXT DEFAULT 'application/pdf',
            data BYTEA NOT NULL,
            file_size INTEGER DEFAULT 0,
            uploaded_by TEXT DEFAULT '',
            uploaded_at TEXT DEFAULT '',
            deleted_by TEXT DEFAULT '',
            deleted_at TEXT DEFAULT ''
        )"""
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_attachments_trash_bl ON record_attachments_trash (bl_number)")

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


class SafeFPDF(FPDF):
    """FPDF that never crashes on text its built-in font can't draw.

    The PDA/SOF PDFs use fpdf's built-in Helvetica, which only covers
    basic Latin. Plain FPDF raised an exception - and the whole download
    failed - on a single en dash or curly quote pasted from Word, or any
    Arabic word. Two changes:
      - windows-1252 encoding instead of latin-1, so the usual Word
        punctuation (– — ‘ ’ “ ” … • €) prints correctly as-is;
      - anything still outside that (e.g. Arabic) is swapped for "?"
        rather than aborting the document.
    Printing real Arabic would need a bundled Arabic font file."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.core_fonts_encoding = "windows-1252"

    def normalize_text(self, text):
        if not self.is_ttf_font and self.core_fonts_encoding:
            text = str(text).encode(self.core_fonts_encoding, errors="replace").decode(self.core_fonts_encoding)
        return super().normalize_text(text)


def build_pda_pdf(doc, items):
    """Renders a PDA (while draft/sent) or FDA (once finalized) as a PDF,
    reusing the Sea Power logo already embedded in the app. Finalized
    documents get an extra Actual + Variance column so the agent can see
    at a glance where the final cost diverged from the estimate."""
    is_fda = doc["status"] == "finalized"
    pdf = SafeFPDF(format="A4")
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
    pdf = SafeFPDF(format="A4")
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


def email_configured():
    return bool(os.environ.get("SMTP_HOST") and os.environ.get("SMTP_USER") and os.environ.get("SMTP_PASSWORD"))


def send_email(to_addrs, subject, body, attachments=None):
    """Sends through SMTP creds in the environment (SMTP_HOST/PORT/USER/
    PASSWORD/FROM, optional SMTP_FROM_NAME) - nothing is hardcoded here,
    same pattern as DATABASE_URL/APP_SECRET_KEY. Returns (ok, error_message)
    instead of raising, so a missing/misconfigured mail account degrades to
    "not sent" rather than a 500 on whatever triggered it.

    attachments: optional list of (filename, bytes, mime_type).
    Port 465 uses implicit TLS; anything else (587, Office 365 / Google
    Workspace) upgrades with STARTTLS. Either way the server's certificate
    is now VERIFIED (ssl.create_default_context) - the old bare
    starttls() encrypted the connection but accepted any certificate, so
    someone on the network could have posed as the mail server and
    collected the mailbox password."""
    host = os.environ.get("SMTP_HOST", "")
    port = int(os.environ.get("SMTP_PORT", "587") or 587)
    user = os.environ.get("SMTP_USER", "")
    password = os.environ.get("SMTP_PASSWORD", "")
    sender = os.environ.get("SMTP_FROM", user)
    sender_name = os.environ.get("SMTP_FROM_NAME", "Sea Power Marine Services")
    if not email_configured():
        return False, "Email isn't configured yet (SMTP_HOST/SMTP_USER/SMTP_PASSWORD missing)."
    if not to_addrs:
        return False, "No recipient configured."
    try:
        msg = EmailMessage()  # handles UTF-8 (Arabic) subjects, bodies and filenames
        msg["From"] = formataddr((sender_name, sender)) if sender_name else sender
        msg["To"] = ", ".join(to_addrs)
        msg["Subject"] = subject
        msg.set_content(body)
        for fname, data, mime in (attachments or []):
            maintype, _, subtype = (mime or "application/octet-stream").partition("/")
            msg.add_attachment(data, maintype=maintype, subtype=subtype or "octet-stream", filename=fname)
        ctx = ssl.create_default_context()
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=12, context=ctx) as server:
                server.login(user, password)
                server.send_message(msg, from_addr=sender, to_addrs=to_addrs)
        else:
            with smtplib.SMTP(host, port, timeout=12) as server:
                server.starttls(context=ctx)
                server.login(user, password)
                server.send_message(msg, from_addr=sender, to_addrs=to_addrs)
        return True, None
    except smtplib.SMTPAuthenticationError:
        return False, "auth"
    except smtplib.SMTPRecipientsRefused:
        return False, "recipient"
    except ssl.SSLError as e:
        return False, f"secure connection failed ({e.reason or e})"
    except (socket.timeout, TimeoutError, ConnectionError, socket.gaierror, OSError):
        # Couldn't reach the mail server at all - wrong host/port, or the
        # hosting plan blocks outgoing mail (Render's free plan does).
        return False, "unreachable"
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
            if request.path.startswith("/api/"):
                return jsonify({"error": "Please sign in again.", "error_code": "session_expired"}), 401
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

# ---------- Accounts, two-step verification, inactivity ----------
COMPANY_EMAIL_DOMAIN = os.environ.get("COMPANY_EMAIL_DOMAIN", "seapower.com.sa").strip().lower().lstrip("@")
REQUIRE_2FA = os.environ.get("REQUIRE_2FA", "1") != "0"
try:
    IDLE_MINUTES = max(2, int(os.environ.get("IDLE_MINUTES", "30") or 30))
except ValueError:
    IDLE_MINUTES = 30
_EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")


def _check_company_email(email):
    """Returns (clean_email, error). Empty is allowed (older accounts)."""
    email = (email or "").strip().lower()
    if not email:
        return "", None
    if not _EMAIL_RE.match(email):
        return "", "That doesn't look like an email address."
    if COMPANY_EMAIL_DOMAIN and not email.endswith("@" + COMPANY_EMAIL_DOMAIN):
        return "", f"Use the company email (name@{COMPANY_EMAIL_DOMAIN})."
    return email, None


# ---------- Password rules ----------
# At least 10 characters, not a common password (or a common word with a few digits
# stuck on, like "Seapower123"), not made of the person's own name/email/username,
# not all digits, not the same password as before. Length counts more than symbols,
# so there is no forced "must contain a symbol" and no forced 90-day change.
# PASSWORD_RULES=0 in the environment switches the checks off (used only by the
# automated tests).
PASSWORD_RULES = os.environ.get("PASSWORD_RULES", "1") != "0"
PASSWORD_MIN_LENGTH = 10
_COMMON_BASES = sorted({
    "password", "passw", "pass", "passcode", "letmein", "welcome", "welcome1", "admin", "administrator", "root", "user", "guest",
    "login", "signin", "qwerty", "qwertyuiop", "qwertyui", "asdfghjkl", "asdfgh", "zxcvbnm", "qazwsx", "qazwsxedc", "azerty",
    "abc", "abcd", "abcde", "abcdef", "abcdefg", "abcdefgh", "iloveyou", "monkey", "dragon", "master", "football", "baseball",
    "superman", "batman", "shadow", "sunshine", "princess", "freedom", "whatever", "trustno", "changeme", "change", "temp",
    "temporary", "temppass", "tempo", "test", "testing", "testtest", "secret", "default", "demo", "sample", "hello", "hellohello",
    "seapower", "seapowermarine", "marine", "compass", "compassapp", "jeddah", "jedda", "dammam", "riyadh", "saudi", "saudia", "ksa",
    "makkah", "madinah", "shipping", "shipment", "cargo", "vessel", "port", "ports", "agent", "agency", "operations", "ops", "ahmed",
    "mohammed", "muhammad", "mohammad", "allah", "bismillah", "alhamdulillah", "inshallah", "ramadan", "arabic", "company", "office",
    "work", "worker", "staff", "employee", "manager", "supervisor", "outlook", "microsoft", "google", "windows", "computer", "internet",
    "summer", "winter", "spring", "autumn", "january", "february", "march", "april", "may", "june", "july", "august", "september",
    "october", "november", "december", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday",
})


def _password_problem(pw, username="", email="", full_name="", old_hash=None):
    """Returns a short, plain message if the password is not acceptable, else None."""
    if not PASSWORD_RULES:
        return None
    pw = pw or ""
    if len(pw) < PASSWORD_MIN_LENGTH:
        return f"Use at least {PASSWORD_MIN_LENGTH} characters."
    low = pw.lower()
    if low.isdigit():
        return "Use letters as well as numbers."
    if len(set(low)) <= 3:
        return "Avoid repeating the same few characters."
    letters = re.sub(r"[^a-z]", "", low)
    leftovers = len(low) - len(letters)
    if letters in _COMMON_BASES and leftovers <= 6:
        return "That password is too common. Pick something harder to guess."
    personal = set()
    for piece in (username, (email or "").split("@")[0]):
        piece = re.sub(r"[^a-z0-9]", "", (piece or "").lower())
        if len(piece) >= 3:
            personal.add(piece)
    for token in re.split(r"[^a-z0-9]+", (full_name or "").lower()):
        if len(token) >= 4:
            personal.add(token)
    if any(p in re.sub(r"[^a-z0-9]", "", low) for p in personal):
        return "Don't use your name, email or username inside the password."
    if old_hash and check_password_hash(old_hash, pw):
        return "Choose a new password - it can't be the same as the old one."
    return None


def _unique_username(db, base):
    base = re.sub(r"[^a-z0-9._\-]", "", (base or "").lower()) or "user"
    name, n = base, 1
    while db.execute("SELECT 1 FROM users WHERE LOWER(username) = LOWER(?)", (name,)).fetchone():
        n += 1
        name = f"{base}{n}"
    return name


# TOTP (RFC 6238) - the same 6-digit code Google Authenticator and Microsoft
# Authenticator show. Written with the standard library, so no new package.
def _totp_new_secret():
    return base64.b32encode(secrets.token_bytes(20)).decode().rstrip("=")


def _totp_at(secret, step):
    key = base64.b32decode(secret + "=" * (-len(secret) % 8))
    d = hmac.new(key, struct.pack(">Q", step), hashlib.sha1).digest()
    o = d[-1] & 15
    return f"{(struct.unpack('>I', d[o:o + 4])[0] & 0x7fffffff) % 10 ** 6:06d}"


def _totp_check(secret, code, last_step=0):
    """Returns the matched time-step, or None. A code can't be used twice
    (step must be newer than the last accepted one); +/-30 s of clock drift is allowed."""
    code = re.sub(r"\s", "", code or "")
    if not secret or not re.fullmatch(r"\d{6}", code):
        return None
    now = int(time.time() // 30)
    for step in (now, now - 1, now + 1):
        if step > (last_step or 0) and hmac.compare_digest(_totp_at(secret, step), code):
            return step
    return None


def _new_recovery_codes():
    alphabet = "abcdefghjkmnpqrstuvwxyz23456789"
    return ["".join(secrets.choice(alphabet) for _ in range(4)) + "-" + "".join(secrets.choice(alphabet) for _ in range(4)) for _ in range(8)]


def _hash_recovery(code):
    return hashlib.sha256(re.sub(r"[^a-z0-9]", "", (code or "").lower()).encode()).hexdigest()


def _use_recovery_code(db, user, code):
    try:
        hashes = json.loads(user["recovery_codes"] or "[]")
    except ValueError:
        hashes = []
    h = _hash_recovery(code)
    for x in hashes:
        if hmac.compare_digest(x, h):
            hashes.remove(x)
            db.execute("UPDATE users SET recovery_codes = ? WHERE id = ?", (json.dumps(hashes), user["id"]))
            db.commit()
            return True
    return False


def _complete_login(user, remember):
    session.clear()
    session["user_id"] = user["id"]
    session["username"] = user["username"]          # the saved spelling - used for record ownership
    session["display_name"] = (user["full_name"] or "").strip() or user["username"]
    session["role"] = user["role"]
    session["remember"] = bool(remember)
    session["last_seen"] = int(time.time())
    session["chk"] = int(time.time())
    session.permanent = bool(remember)
    if REQUIRE_2FA and not user["totp_enabled"]:
        session["needs_2fa_setup"] = True
        return redirect(url_for("account_security"))
    return redirect(url_for("index"))


@app.before_request
def _same_origin_only():
    """Browsers always say where a form/fetch came from. A write that claims
    to come from another website is refused (a second wall behind SameSite)."""
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        src = request.headers.get("Origin") or ""
        if not src:
            ref = request.headers.get("Referer") or ""
            m = re.match(r"^(https?://[^/]+)", ref)
            src = m.group(1) if m else ""
        if src and src != "null" and urlparse(src).netloc.lower() != request.host.lower():
            if request.path.startswith("/api/"):
                return jsonify({"error": "Request refused (it came from another site).", "error_code": "bad_origin"}), 403
            return "Request refused (it came from another site).", 403
    return None


@app.before_request
def _session_guards():
    if "user_id" not in session:
        return None
    p = request.path
    now = int(time.time())
    expired_json = lambda: (jsonify({"error": "Your session ended. Please sign in again.", "error_code": "session_expired"}), 401)
    # 1. inactivity. Background refreshes (GET /api/...) don't count as activity - only
    # clicks/typing do (the page pings /api/ping while the person is active).
    if not session.get("remember"):
        if now - session.get("last_seen", now) > IDLE_MINUTES * 60 + 30:
            session.clear()
            return expired_json() if p.startswith("/api/") else redirect(url_for("login", expired=1))
        if request.method != "GET" or not p.startswith("/api/") or p == "/api/ping":
            session["last_seen"] = now
    # 2. every few minutes re-check the account still exists (removed staff lose access
    # even if their cookie is still valid) and pick up name/role changes.
    if now - session.get("chk", 0) > 300:
        row = get_db().execute("SELECT username, full_name, role FROM users WHERE id = ?", (session["user_id"],)).fetchone()
        if not row:
            session.clear()
            return expired_json() if p.startswith("/api/") else redirect(url_for("login"))
        session["role"] = row["role"]
        session["username"] = row["username"]
        session["display_name"] = (row["full_name"] or "").strip() or row["username"]
        session["chk"] = now
    # 3. someone who still has to set up their authenticator can't use anything else yet
    if session.get("needs_2fa_setup") and p not in ("/account/security", "/logout", "/api/ping"):
        if p.startswith("/api/"):
            return jsonify({"error": "Set up two-step verification first.", "error_code": "2fa_setup_required"}), 403
        return redirect(url_for("account_security"))
    return None


@app.context_processor
def _inject_names():
    out = {"pw_config": {"on": PASSWORD_RULES, "min": PASSWORD_MIN_LENGTH, "common": _COMMON_BASES}}
    if "user_id" in session:
        dn = session.get("display_name") or session.get("username") or ""
        out.update({"display_name": dn, "first_name": (dn.split() or [""])[0]})
    return out


_IDLE_SNIPPET = """<script>
(function () {
  var IDLE = __IDLE__ * 1000, WARN = 60000, USE_TIMER = __TIMER__;
  var lang = 'en'; try { lang = localStorage.getItem('lang') || 'en'; } catch (e) {}
  var AR = lang === 'ar';
  var T = AR ? {h: 'هل ما زلت هنا؟', b: 'سيتم تسجيل خروجك خلال', s: 'بسبب عدم النشاط، لحماية بياناتك.', stay: 'البقاء متصلاً', out: 'تسجيل الخروج'}
             : {h: 'Still there?', b: 'You will be signed out in', s: 'because of inactivity, to keep your data safe.', stay: 'Stay signed in', out: 'Sign out'};
  // a finished session (removed account, expired cookie) -> straight to the sign-in page
  var _fetch = window.fetch;
  window.fetch = function () {
    return _fetch.apply(this, arguments).then(function (r) {
      if (r.status === 401) { try { if (new URL(r.url).pathname.indexOf('/api/') === 0) location.href = '/login?expired=1'; } catch (e) {} }
      return r;
    });
  };
  if (!USE_TIMER) return;
  var last = Date.now(), lastPing = Date.now(), shown = false, box, tick;
  function mark() { last = Date.now(); try { localStorage.setItem('compass_active', String(last)); } catch (e) {} }
  ['mousedown', 'keydown', 'touchstart', 'scroll', 'wheel', 'pointerdown'].forEach(function (ev) { window.addEventListener(ev, function () { if (!shown) mark(); }, {passive: true, capture: true}); });
  var mt = 0; window.addEventListener('mousemove', function () { var n = Date.now(); if (n - mt > 5000 && !shown) { mt = n; mark(); } }, {passive: true});
  function shared() { try { return parseInt(localStorage.getItem('compass_active') || '0', 10) || 0; } catch (e) { return 0; } }
  function ping() { lastPing = Date.now(); return _fetch('/api/ping', {method: 'POST', credentials: 'same-origin'}).then(function (r) { if (r.status === 401) location.href = '/login?expired=1'; }).catch(function () {}); }
  function fmt(ms) { var s = Math.max(0, Math.ceil(ms / 1000)); return Math.floor(s / 60) + ':' + ('0' + (s % 60)).slice(-2); }
  function build() {
    box = document.createElement('div');
    box.setAttribute('role', 'alertdialog'); box.setAttribute('aria-modal', 'true');
    box.style.cssText = 'position:fixed;inset:0;z-index:2147483000;display:flex;align-items:center;justify-content:center;padding:20px;background:rgba(8,20,32,.55);backdrop-filter:blur(4px);-webkit-backdrop-filter:blur(4px);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;';
    if (AR) box.setAttribute('dir', 'rtl');
    box.innerHTML = '<div style="background:var(--card,#fff);color:var(--text,#1c2b3a);border:1px solid var(--border,#e6e9ed);border-radius:20px;padding:28px 26px 22px;max-width:360px;width:100%;text-align:center;box-shadow:0 24px 70px rgba(0,0,0,.35);">' +
      '<div style="width:46px;height:46px;border-radius:50%;margin:0 auto 14px;display:flex;align-items:center;justify-content:center;background:rgba(201,162,39,.16);color:#c9a227;"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg></div>' +
      '<div style="font-size:18px;font-weight:700;margin-bottom:6px;">' + T.h + '</div>' +
      '<div style="font-size:13.5px;color:var(--muted,#7a8794);line-height:1.5;">' + T.b + '</div>' +
      '<div id="idleCount" style="font-size:34px;font-weight:700;letter-spacing:.02em;margin:6px 0 4px;color:var(--navy,#123a56);font-variant-numeric:tabular-nums;"></div>' +
      '<div style="font-size:12.5px;color:var(--muted,#7a8794);margin-bottom:18px;">' + T.s + '</div>' +
      '<button id="idleStay" style="width:100%;border:0;border-radius:999px;padding:12px;font-size:14px;font-weight:700;background:var(--navy,#123a56);color:#fff;cursor:pointer;">' + T.stay + '</button>' +
      '<button id="idleOut" style="width:100%;border:0;background:none;margin-top:6px;padding:9px;font-size:13px;font-weight:600;color:var(--muted,#7a8794);cursor:pointer;">' + T.out + '</button></div>';
    document.body.appendChild(box);
    box.querySelector('#idleStay').onclick = stay;
    box.querySelector('#idleOut').onclick = function () { location.href = '/logout'; };
    box.querySelector('#idleStay').focus();
  }
  function stay() { shown = false; if (box) { box.remove(); box = null; } mark(); ping(); }
  tick = setInterval(function () {
    var now = Date.now(), act = Math.max(last, shared()), idle = now - act;
    if (shown) {
      if (act > last) { last = act; }
      if (idle < IDLE - WARN) { stay(); return; }       // activity in another tab
    }
    if (idle >= IDLE) { location.href = '/logout?idle=1'; return; }
    if (idle >= IDLE - WARN && !shown) { shown = true; build(); }
    if (shown && box) { box.querySelector('#idleCount').textContent = fmt(IDLE - idle); }
    if (!shown && idle < 60000 && now - lastPing > 55000) { ping(); }   // keep the server's clock in step while the person is active
  }, 1000);
})();
</script>"""


@app.after_request
def _security_headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    resp.headers.setdefault("Content-Security-Policy", "frame-ancestors 'none'; base-uri 'self'; form-action 'self'")
    resp.headers.setdefault("Referrer-Policy", "same-origin" if not request.path.startswith("/t/") else "no-referrer")
    resp.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=()")
    if os.environ.get("RENDER"):
        resp.headers.setdefault("Strict-Transport-Security", "max-age=31536000")
    if not request.path.startswith("/static/"):
        resp.headers["Cache-Control"] = "no-store"
    if request.path.startswith("/t/"):
        resp.headers.setdefault("X-Robots-Tag", "noindex, nofollow")
    return resp


@app.after_request
def _add_idle_script(resp):
    try:
        if ("user_id" in session and resp.status_code == 200 and resp.mimetype == "text/html"
                and not resp.direct_passthrough and request.path not in ("/login", "/setup")):
            body = resp.get_data(as_text=True)
            i = body.rfind("</body>")
            if i != -1:
                timer = "false" if session.get("remember") else "true"
                snippet = _IDLE_SNIPPET.replace("__IDLE__", str(IDLE_MINUTES * 60)).replace("__TIMER__", timer)
                resp.set_data(body[:i] + snippet + body[i:])
    except Exception:
        pass
    return resp


@app.route("/api/ping", methods=["POST"])
@login_required
def api_ping():
    return jsonify({"ok": True})


@app.route("/setup", methods=["GET", "POST"])
def setup():
    if any_users_exist():
        return redirect(url_for("login"))
    error = None
    if request.method == "POST":
        full_name = request.form.get("full_name", "").strip()
        email, email_err = _check_company_email(request.form.get("email", ""))
        username = request.form.get("username", "").strip() or (email.split("@")[0] if email else "")
        password = request.form.get("password", "")
        if email_err:
            error = email_err
        elif not username or not password:
            error = "Please fill in all the fields."
        elif _password_problem(password, username, email, full_name):
            error = _password_problem(password, username, email, full_name)
        else:
            db = get_db()
            db.execute(
                "INSERT INTO users (username, password_hash, role, created_at, full_name, email) VALUES (?, ?, 'admin', ?, ?, ?)",
                (username, generate_password_hash(password), datetime.utcnow().strftime("%Y-%m-%d %H:%M"), full_name, email),
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
#
# Second, per-ACCOUNT limit: the per-IP limit depends on the client IP that
# Render reports in X-Forwarded-For, and a client can put fake addresses in
# that header itself - so on its own, rotating fake IPs could keep guessing
# forever. Capping failures per username as well bounds guessing against
# any one account no matter what IP is reported. It's set higher than the
# per-IP limit so a colleague mistyping from the same office doesn't lock
# the account; the trade-off is that someone deliberately spamming wrong
# passwords can lock an account for 15 minutes.
_LOGIN_MAX_ATTEMPTS = 8
_LOGIN_MAX_ATTEMPTS_PER_USER = 20
_LOGIN_WINDOW_SECONDS = 900  # 15 minutes
_login_failures = {}  # ip -> [timestamp, ...] of recent failed attempts
_login_failures_user = {}  # lowercased username -> [timestamp, ...]


def _client_ip():
    # Render terminates TLS and proxies requests, so the real client IP
    # arrives in X-Forwarded-For rather than as the direct socket peer.
    # Render states it puts the real client IP first in that list.
    fwd = request.headers.get("X-Forwarded-For", "")
    if fwd:
        return fwd.split(",")[0].strip()
    return request.remote_addr or "unknown"


def _recent(store, key):
    now = time.time()
    recent = [t for t in store.get(key, []) if now - t < _LOGIN_WINDOW_SECONDS]
    if recent:
        store[key] = recent
    else:
        store.pop(key, None)  # don't let the dict grow forever with stale keys
    return recent


def _login_locked_out(ip, username=""):
    if len(_recent(_login_failures, ip)) >= _LOGIN_MAX_ATTEMPTS:
        return True
    user_key = (username or "").strip().lower()
    return bool(user_key) and len(_recent(_login_failures_user, user_key)) >= _LOGIN_MAX_ATTEMPTS_PER_USER


def _record_login_failure(ip, username=""):
    _login_failures.setdefault(ip, []).append(time.time())
    user_key = (username or "").strip().lower()
    if user_key:
        _login_failures_user.setdefault(user_key, []).append(time.time())


_DUMMY_HASH = generate_password_hash("not-a-real-password")


@app.route("/login", methods=["GET", "POST"])
def login():
    if not any_users_exist():
        return redirect(url_for("setup"))
    error = None
    typed = ""
    notice = "You were signed out after a period of inactivity. Please sign in again." if request.args.get("expired") else None
    if request.method == "POST":
        ip = _client_ip()
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        typed = username
        if _login_locked_out(ip, username):
            error = "Too many failed attempts. Please wait 15 minutes and try again."
            return render_template_string(LOGIN_HTML, error=error, typed=typed, notice=None)
        db = get_db()
        # Sign in with the company email or the username, in any capitalisation.
        # If two accounts somehow differ only by case (made before this rule),
        # the exact spelling is tried first and the password decides.
        candidates = db.execute(
            "SELECT * FROM users WHERE LOWER(username) = LOWER(?) OR (email <> '' AND LOWER(email) = LOWER(?)) "
            "ORDER BY (username = ?) DESC, id",
            (username, username, username),
        ).fetchall()
        user = next((u for u in candidates if check_password_hash(u["password_hash"], password)), None)
        if not candidates:
            check_password_hash(_DUMMY_HASH, password)   # same work for unknown names, so timing doesn't reveal who has an account
        if user:
            remember = bool(request.form.get("remember"))
            if user["totp_enabled"]:
                session.clear()
                session["pre2fa"] = {"id": user["id"], "remember": remember, "t": int(time.time())}
                return redirect(url_for("login_verify"))
            return _complete_login(user, remember)
        _record_login_failure(ip, username)
        error = "Wrong email/username or password."
    return render_template_string(LOGIN_HTML, error=error, typed=typed, notice=notice)


@app.route("/login/verify", methods=["GET", "POST"])
def login_verify():
    pre = session.get("pre2fa")
    if not pre or int(time.time()) - pre.get("t", 0) > 300:
        session.pop("pre2fa", None)
        return redirect(url_for("login"))
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ?", (pre["id"],)).fetchone()
    if not user:
        session.clear()
        return redirect(url_for("login"))
    error = None
    if request.method == "POST":
        ip = _client_ip()
        if _login_locked_out(ip, user["username"]):
            return render_template_string(VERIFY_HTML, error="Too many failed attempts. Please wait 15 minutes and try again.", name=user["full_name"] or user["username"])
        code = request.form.get("code", "")
        step = _totp_check(user["totp_secret"], code, user["totp_last_step"])
        if step:
            db.execute("UPDATE users SET totp_last_step = ? WHERE id = ?", (step, user["id"]))
            db.commit()
            return _complete_login(user, pre.get("remember"))
        if re.search(r"[a-zA-Z]", code or "") and _use_recovery_code(db, user, code):
            return _complete_login(user, pre.get("remember"))
        _record_login_failure(ip, user["username"])
        error = "That code isn't right. Check the code in your authenticator app and try again."
    return render_template_string(VERIFY_HTML, error=error, name=user["full_name"] or user["username"])


@app.route("/account/security", methods=["GET", "POST"])
@login_required
def account_security():
    db = get_db()
    user = db.execute("SELECT * FROM users WHERE id = ?", (session["user_id"],)).fetchone()
    if not user:
        session.clear()
        return redirect(url_for("login"))
    name = user["full_name"] or user["username"]
    if user["totp_enabled"] and not session.get("show_codes"):
        return render_template_string(SECURITY_HTML, mode="on", name=name, error=None, codes=None, qr=None, secret=None, email=user["email"])
    if session.get("show_codes"):
        codes = session.pop("show_codes")
        session.pop("needs_2fa_setup", None)
        return render_template_string(SECURITY_HTML, mode="codes", name=name, error=None, codes=codes, qr=None, secret=None, email=user["email"])
    error = None
    if request.method == "POST":
        pending = session.get("totp_pending")
        step = _totp_check(pending, request.form.get("code", ""), 0) if pending else None
        if step:
            codes = _new_recovery_codes()
            db.execute(
                "UPDATE users SET totp_secret = ?, totp_enabled = 1, totp_last_step = ?, recovery_codes = ? WHERE id = ?",
                (pending, step, json.dumps([_hash_recovery(c) for c in codes]), user["id"]),
            )
            db.commit()
            session.pop("totp_pending", None)
            session["show_codes"] = codes
            return redirect(url_for("account_security"))
        error = "That code isn't right. Check the code in your authenticator app and try again."
    secret = session.get("totp_pending")
    if not secret:
        secret = session["totp_pending"] = _totp_new_secret()
    label = user["email"] or user["username"]
    uri = f"otpauth://totp/Compass:{url_quote(label)}?secret={secret}&issuer=Compass&algorithm=SHA1&digits=6&period=30"
    qr = segno.make(uri, error="m").svg_data_uri(scale=5, border=2, dark="#123a56")
    pretty = " ".join(secret[i:i + 4] for i in range(0, len(secret), 4))
    return render_template_string(SECURITY_HTML, mode="setup", name=name, error=error, codes=None, qr=qr, secret=pretty, email=user["email"])


@app.route("/logout")
def logout():
    idle = request.args.get("idle")
    session.clear()
    return redirect(url_for("login", expired=1) if idle else url_for("login"))


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
    users = db.execute("SELECT id, username, full_name, email, role, created_at, totp_enabled FROM users ORDER BY created_at, id").fetchall()
    return render_template_string(USERS_HTML, users=users, username=session.get("username"))


@app.route("/api/users", methods=["POST"])
@login_required
@admin_required
def add_user():
    data = request.get_json(force=True)
    full_name = (data.get("full_name") or "").strip()
    email, email_err = _check_company_email(data.get("email"))
    if email_err:
        return jsonify({"error": email_err}), 400
    username = (data.get("username") or "").strip()
    password = data.get("password", "")
    role = data.get("role", "staff")
    if role not in ("admin", "staff"):
        role = "staff"
    db = get_db()
    if not username and email:
        username = _unique_username(db, email.split("@")[0])
    if not username or not password:
        return jsonify({"error": "Missing fields"}), 400
    problem = _password_problem(password, username, email, full_name)
    if problem:
        return jsonify({"error": problem, "error_code": "weak_password"}), 400
    if db.execute("SELECT 1 FROM users WHERE LOWER(username) = LOWER(?)", (username,)).fetchone():
        return jsonify({"error": "Username already exists"}), 400
    if email and db.execute("SELECT 1 FROM users WHERE LOWER(email) = LOWER(?)", (email,)).fetchone():
        return jsonify({"error": "That email already has an account."}), 400
    try:
        db.execute(
            "INSERT INTO users (username, password_hash, role, created_at, full_name, email) VALUES (?, ?, ?, ?, ?, ?)",
            (username, generate_password_hash(password), role, datetime.utcnow().strftime("%Y-%m-%d %H:%M"), full_name, email),
        )
        db.commit()
    except psycopg2.IntegrityError:
        db.conn.rollback()
        return jsonify({"error": "Username already exists"}), 400
    return jsonify({"ok": True, "username": username})


@app.route("/api/users/<int:user_id>", methods=["PATCH"])
@login_required
@admin_required
def edit_user(user_id):
    data = request.get_json(force=True) or {}
    db = get_db()
    u = db.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    if not u:
        return jsonify({"error": "User not found."}), 404
    sets, vals = [], []
    if "full_name" in data:
        sets.append("full_name = ?"); vals.append((data.get("full_name") or "").strip())
    if "email" in data:
        email, err = _check_company_email(data.get("email"))
        if err:
            return jsonify({"error": err}), 400
        if email and db.execute("SELECT 1 FROM users WHERE LOWER(email) = LOWER(?) AND id <> ?", (email, user_id)).fetchone():
            return jsonify({"error": "That email already has an account."}), 400
        sets.append("email = ?"); vals.append(email)
    if data.get("password"):
        problem = _password_problem(
            data["password"], u["username"],
            (data.get("email") if "email" in data else u["email"]) or "",
            (data.get("full_name") if "full_name" in data else u["full_name"]) or "",
            u["password_hash"],
        )
        if problem:
            return jsonify({"error": problem, "error_code": "weak_password"}), 400
        sets.append("password_hash = ?"); vals.append(generate_password_hash(data["password"]))
    if data.get("role") in ("admin", "staff"):
        if user_id == session.get("user_id") and data["role"] != u["role"]:
            return jsonify({"error": "You can't change your own role."}), 400
        sets.append("role = ?"); vals.append(data["role"])
    if data.get("reset_2fa"):
        sets += ["totp_secret = ''", "totp_enabled = 0", "totp_last_step = 0", "recovery_codes = ''"]
    if not sets:
        return jsonify({"ok": True})
    db.execute(f"UPDATE users SET {', '.join(sets)} WHERE id = ?", (*vals, user_id))
    db.commit()
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
    out = []
    for r in rows:
        rec = dict(r)
        # The individual B/Ls a combined entry stands for, so searching the
        # board for "...002" finds "...001-003".
        members = bl_members(rec["bl_number"])[1:]
        if members:
            rec["bl_members"] = members
        out.append(rec)
    return jsonify(out)


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

            # Archived vessels are finished/parked by definition - counting
            # their unticked BLs as live backlog inflated the "pending"
            # numbers. Their timings still feed the turnaround averages.
            if r.get("archived"):
                continue
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
    """Older "paste a list of BLs" endpoint (not used by the board itself
    any more). One BL per line; "BL, something" keeps just the BL. Lines
    that aren't a B/L number are ignored, and BLs go through the same add
    path as a manifest upload."""
    data = request.get_json(force=True, silent=True) or {}
    items = []
    for raw in str(data.get("lines", "")).splitlines():
        for bl in _mf_bls_from_line(raw.split(",", 1)[0], allow_numeric=True)[:1]:
            items.append({"bl": bl})
    if not items:
        return jsonify({"added": 0})
    return jsonify({"added": _manifest_add("", "", items)["added"]})


# ---------- Consignee contact details (for customer notifications) ----------

_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
_EMAIL_FULL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def _normalize_phone(raw):
    """Phone -> international '+<digits>' form (what WhatsApp needs), or ''.
    Saudi local formats get +966: '0501234567', '501234567',
    '00966 50 123 4567', '966-50-1234567' all -> '+966501234567';
    '011 230 8888' -> '+966112308888'. A number typed with '+' or '00'
    keeps its own country code. The local trunk 0 that people often keep
    after the country code ('+966 0501234567', '+0549928955') is dropped."""
    s = str(raw or "").strip()
    if not s:
        return ""
    d = re.sub(r"\D", "", s)
    if s.startswith("+") and not s.startswith("+0"):
        pass
    elif d.startswith("00"):
        d = d[2:]
    elif d.startswith("0") and len(d) == 10:      # 05x xxx xxxx / 01x xxx xxxx
        d = "966" + d[1:]
    elif d.startswith("5") and len(d) == 9:       # 5x xxx xxxx
        d = "966" + d
    if d.startswith("9660") and len(d) == 13:     # +966 0 5x xxx xxxx
        d = "966" + d[4:]
    if not (8 <= len(d) <= 15):
        return ""
    return "+" + d


def _clean_email(raw):
    e = str(raw or "").strip()
    return e if e and len(e) <= 254 and _EMAIL_FULL_RE.match(e) else ""


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


# ---------- Manifest reading (DO Tracker "Add a manifest") ----------
#
# Every load port / forwarder sends its manifest in its own layout, so this
# never trusts one fixed template. What it relies on instead:
#
#  * The B/L column is found by its HEADER ("B/L NO.", "BL Number",
#    "Bill of Lading No.", "B/L Nr.", "提单号码" ...) anywhere in the first 40
#    rows - not just the first 5, and never by assuming column A. Only that
#    column is read, below the header.
#  * Every value read from it must actually LOOK like a B/L number (letters
#    and digits, no spaces/sentences/phone numbers/dates). Notes typed under
#    a BL in the same cell ("no have bl draft", "according to MR",
#    "(050-055+165)", "will switch bl"), TOTAL rows, repeated header rows and
#    blank lines are dropped.
#  * A sheet/table with NO B/L header is only used when it is clearly the
#    manifest itself or its continuation (page 2 of a Word/PDF table): one
#    column holding nothing but distinct B/L-shaped values of one series
#    (e.g. JYM2603BYQJD2..). Attachment sheets (VIN / chassis / engine lists,
#    "B/L ATTACHMENT") never qualify, so their numbers can't become BLs.
#  * A tab with a B/L column but no cargo columns (a "CONTACT INFORMATION"
#    or remarks tab) is used only for contact details, never to add BLs.
#  * Rows hidden in Excel (e.g. other discharge ports filtered out), hidden
#    sheets, and B/L cells that are crossed out (strikethrough = cancelled /
#    shut out) are reported separately and left unticked in the preview
#    instead of being added silently.
#
# The result is shown in a preview before anything is added, so even a
# layout never seen before can't put junk on the board without someone
# seeing it first.

_MF_MAX_ROWS = 20000
_MF_MAX_COLS = 60
_MF_INVISIBLE_RE = re.compile(r"[​-‏‪-‮⁦-⁩﻿]")
_MF_SEP_TRANSLATE = str.maketrans({
    "、": "/", ",": "/", "&": "/", "+": "/", "\\": "/",
    "—": "-", "–": "-", "‐": "-", "‑": "-", "~": "-", "−": "-",
})


def _mf_norm(text):
    """Upper-cased, NFKC-normalized (full-width Ａ１／：（ -> A1/:( ),
    invisible direction marks removed. Used for B/L values and headers."""
    s = unicodedata.normalize("NFKC", str(text or ""))
    s = _MF_INVISIBLE_RE.sub("", s)
    return s.upper()


def _mf_cell_text(v):
    """One spreadsheet/table cell as text. Excel stores numbers as floats, so
    a whole number comes back as 2015.0 from .xls files - shown as 2015."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    return str(v)


# A header cell naming the B/L column. "MBL"/"HBL"/"HOUSE B/L" are accepted
# too, but a plain B/L column is preferred when a sheet has both.
_MF_BL_HEADER_RE = re.compile(
    r"^(?P<pre>(?:HOUSE|MASTER|OCEAN|H|M)\s*[./-]?\s*)?"
    r"(?:B\s*[/\\.-]?\s*L|BILL\s+OF\s+LADING|BOL)\.?"
    r"(?:\s*(?:NO|NOS|NR|NUM|NUMBER|NUMBERS|N°|#))?\.?\s*[:#]?\s*$"
)
_MF_BL_HEADER_CN_RE = re.compile(r"^(?:海运)?(?:提单|提單|运单|運單)(?:号码|號碼|号|號|编号|編號)?$")

# Other column headers that only a real cargo manifest has. A table whose
# header has the B/L column plus at least one of these is a manifest table;
# one with only a B/L column and e.g. "CONTACT INFORMATION" is reference info.
_MF_CARGO_HEADER_WORDS = (
    "MARK", "DESCRIPTION", "PACKAGE", "PKG", "QTY", "QUANTITY", "WEIGHT", "MEASUREMENT", "CBM",
    "VOLUME", "SHIPPER", "CONSIGNEE", "NOTIFY", "CARGO", "GOODS", "COMMODITY", "CONTAINER",
    "唛头", "唛", "货名", "品名", "件数", "包装", "毛重", "重量", "尺码", "体积", "货主", "发货人", "收货人", "通知人",
)


def _mf_header_rank(cell):
    """0 = this cell is a plain B/L-number column header, 1 = a house/master
    B/L header, None = not a B/L header. A label that already carries a value
    ("B/L NO.CH26207YJED301", "B/L NO.: MX26188LJED025" - the title line of an
    attachment sheet) is not a column header."""
    lines = [ln.strip() for ln in str(cell or "").splitlines() if ln.strip()]
    if not lines or len(lines) > 3:
        return None
    best = None
    for ln in lines:
        n = _mf_norm(ln).strip()
        m = _MF_BL_HEADER_RE.match(n)
        if m:
            r = 1 if m.group("pre") else 0
        elif _MF_BL_HEADER_CN_RE.match(re.sub(r"\s+", "", n)):
            r = 0
        else:
            if any(ch.isdigit() for ch in n):
                return None
            continue
        best = r if best is None else min(best, r)
    return best


def _mf_row_has_cargo_words(row, skip=None):
    for j, c in enumerate(row):
        if j == skip:
            continue
        n = re.sub(r"\s+", "", _mf_norm(c))
        if n and any(w in n for w in _MF_CARGO_HEADER_WORDS):
            return True
    return False


def _mf_header_col(row):
    """Column of the B/L header if this row is a table header row, else None.
    A lone "B/L" word in a title/letterhead row doesn't count: the row must
    also name a cargo column, or be a short (1-2 cell) header like
    "BL Number" / "BL NO. | CONTACT INFORMATION"."""
    found = [(rk, j) for j, c in enumerate(row) if (rk := _mf_header_rank(c)) is not None]
    if not found:
        return None
    rk, j = min(found)  # plain B/L before house/master; leftmost first
    nonempty = sum(1 for c in row if str(c or "").strip())
    if nonempty > 2 and not _mf_row_has_cargo_words(row, skip=j):
        return None
    return j


def _mf_find_header(rows, limit=40):
    """(row_index, col_index) of the B/L column header, or (None, None)."""
    for i in range(min(len(rows), limit)):
        j = _mf_header_col(rows[i])
        if j is not None:
            return i, j
    return None, None


_MF_BL_CHARS_RE = re.compile(r"^[A-Z0-9](?:[A-Z0-9/.\-_#]*[A-Z0-9])?$")
_MF_TOTAL_RE = re.compile(r"TOTAL|合计|总计|汇总|小计")

_VIN_VALUES = dict(zip("0123456789ABCDEFGHJKLMNPRSTUVWXYZ",
                       [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2, 3, 4, 5, 6, 7, 8, 1, 2, 3, 4, 5, 7, 9, 2, 3, 4, 5, 6, 7, 8, 9]))
_VIN_WEIGHTS = (8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2)


def _mf_is_vin(tok):
    """A 17-character vehicle VIN with a valid check digit (the chassis
    numbers listed on a B/L's attachment sheet)."""
    if len(tok) != 17 or not re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", tok):
        return False
    if re.fullmatch(r"[A-Z]{4,}\d{6,}", tok):
        return False        # carrier-style B/L (SCAC + port + serial, e.g. HLCUSHA2401234567), never a VIN
    check = sum(_VIN_VALUES[c] * w for c, w in zip(tok, _VIN_WEIGHTS)) % 11
    return tok[8] == ("X" if check == 10 else str(check))


def _mf_shape_ok(tok, allow_numeric=False, strict=False):
    """Does this single token look like a B/L number? Letters + digits, no
    spaces, 5-40 characters, nothing but / - . _ # as separators. A purely
    numeric B/L is only accepted directly under a B/L header (allow_numeric).
    strict (used where there is no header to go by): must start with a
    letter, have 2+ letters and 2+ digits, and must not be a vehicle VIN."""
    if not (5 <= len(tok) <= 40) or not _MF_BL_CHARS_RE.match(tok):
        return False
    letters = sum(c.isalpha() for c in tok)
    digits = sum(c.isdigit() for c in tok)
    if digits == 0 or _MF_TOTAL_RE.search(tok):
        return False
    if letters == 0:
        # Some carriers' B/Ls are all digits - accepted only under a B/L
        # header, and never when it is plainly a Saudi phone number.
        return (allow_numeric and tok.isdigit() and 6 <= len(tok) <= 20
                and not re.fullmatch(r"(?:00)?966[15]\d{8}|0[15]\d{8}", tok))
    if strict and (letters < 2 or digits < 2 or not tok[0].isalpha() or _mf_is_vin(tok)):
        return False
    return True


def _mf_family(tok):
    """The series a B/L belongs to: 'JYM2603BYQJD' for JYM2603BYQJD201,
    'QCLYGJD' for QCLYGJD31A/B. Used to recognise continuation tables and
    to keep stray tokens out of free-text reads."""
    first = re.split(r"[/-]", tok, maxsplit=1)[0]
    m = re.match(r"^(.*[A-Z])(\d+)[A-Z]{0,2}$", first)
    if not m:
        return None
    fam = m.group(1)
    return fam if len(fam) >= 2 and fam[0].isalpha() else None


# What may follow a separator inside one B/L entry: more numbers of the same
# series ("-003", "/27", "/002A"), a suffix letter ("31A/B"), a short code
# ("-HK", "-TS1") or the full next B/L of the same series. A word after a
# dash/slash ("QCLYGJD29 - CANCELLED", "COSU6123456780 + VIN LIST") is a note,
# not part of the number, and is cut off there.
_MF_TAIL_SEGMENT_RE = re.compile(r"^(?:\d{1,6}[A-Z]{0,2}|[A-Z]{1,2}|[A-Z]{1,3}\d{1,4}[A-Z]?)$")


def _mf_trim_tail(tok):
    parts = re.split(r"([/-])", tok)
    fam = _mf_family(parts[0])
    if len(parts) == 1 or not fam:
        return tok          # nothing joined on, or no series to judge by (e.g. "SEAPWR-2026-JED-0021")
    out = parts[0]
    for i in range(1, len(parts) - 1, 2):
        sep, seg = parts[i], parts[i + 1]
        if not (_MF_TAIL_SEGMENT_RE.match(seg) or _mf_family(seg) == fam):
            break
        out += sep + seg
    return out


def _mf_cell_lines(cell):
    """A cell's lines, with a B/L entry that wrapped onto the next line put
    back together ("QCLYGJD26/27/" + "28", "BO26215XJED001" + "-003")."""
    out = []
    for ln in str(cell or "").splitlines():
        ln = ln.strip()
        if not ln:
            continue
        n = _mf_norm(ln)
        if out and (re.search(r"[/\-、,&+~]$", _mf_norm(out[-1])) or re.match(r"^[/\-、,&+~]\s*\d", n)):
            out[-1] = out[-1] + ln
        else:
            out.append(ln)
    return out


_MF_LABEL_RE = re.compile(r"^(?:B\s*/\s*L|B\s*L|BILL\s+OF\s+LADING)\s*(?:(?:NO|NUMBER|NR)\b\.?\s*[:#]?|[:#])\s*")


def _mf_bls_from_line(line, allow_numeric=False, strict=False):
    """B/L number(s) on one line of a B/L cell, normalized ('CH26207YJED502、503'
    -> 'CH26207YJED502/503', 'BO26215XJED001 - 003' -> 'BO26215XJED001-003').
    A note after the number ('YZ26228XJED086-88 will switch bl',
    'QCLYGJD29 (see attachment)') is cut off; a line that is only a note
    returns []. Two numbers of the same series on one line ('QCNJJD01
    QCNJJD02') come back as two."""
    s = _mf_norm(line).strip()
    if not s:
        return []
    s = _MF_LABEL_RE.sub("", s)
    if s.startswith("+"):
        return []                       # a phone number typed in the B/L column
    s = re.split(r"[(\[{<]", s, maxsplit=1)[0]
    s = s.translate(_MF_SEP_TRANSLATE)
    s = re.sub(r"\s*([/-])\s*", r"\1", s)
    s = re.sub(r"-{2,}", "-", s)
    s = re.sub(r"/{2,}", "/", s)
    toks = [_mf_trim_tail(t.strip(".,;:/-#")) for t in s.split()]
    toks = [t for t in toks if t]
    if not toks:
        return []
    first = toks[0]
    if not _mf_shape_ok(first, allow_numeric, strict):
        return []
    out = [first]
    fam = _mf_family(first)
    for t in toks[1:]:
        if fam and _mf_shape_ok(t, False, strict) and _mf_family(t) == fam and t not in out:
            out.append(t)
    return out


def _mf_bls_from_cell(cell, allow_numeric=False, strict=False):
    """All B/L numbers in one B/L cell. Several stacked on separate lines
    are separate BLs; lines that are notes are dropped."""
    out = []
    for ln in _mf_cell_lines(cell):
        for b in _mf_bls_from_line(ln, allow_numeric, strict):
            if b not in out:
                out.append(b)
    return out


def bl_match_key(bl):
    """How two B/L strings are compared for "already on the board": case,
    spaces and separator style don't matter, so a BL saved earlier as
    'CH26207YJED502、503' is the same as 'CH26207YJED502/503'."""
    s = _mf_norm(bl).translate(_MF_SEP_TRANSLATE)
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"-{2,}", "-", s)
    return re.sub(r"/{2,}", "/", s)


def bl_members(bl):
    """The individual B/L numbers one manifest entry stands for. Manifests
    combine B/Ls on one line in many ways; each of these is understood:
      BO26215XJED001-003   -> ...001, ...002, ...003
      QCLYGJD26/27/28      -> QCLYGJD26, QCLYGJD27, QCLYGJD28
      BO26215XJED052-3/118 -> ...052, ...053, ...118
      MX26188BJED011/17-18 -> ...011, ...017, ...018
      QCLYGJD31A/B         -> QCLYGJD31A, QCLYGJD31B
    The entry itself always comes first. Anything that can't be read with
    certainty adds only the first number (never a guessed range)."""
    whole = bl_match_key(bl)
    out = [whole]
    parts = re.split(r"([/-])", whole)
    m = re.match(r"^(.*[A-Z])(\d+)([A-Z]{0,2})$", parts[0])
    if not m:
        return out
    prefix, num, _suf = m.groups()
    width = len(num)
    members = [parts[0]]
    last_s = num
    ok = True
    i = 1
    while ok and i < len(parts) - 1:
        sep, tok = parts[i], parts[i + 1]
        i += 2
        mm = re.match(r"^" + re.escape(prefix) + r"(\d+)([A-Z]{0,2})$", tok) or re.match(r"^(\d*)([A-Z]{0,2})$", tok)
        if not mm or not (mm.group(1) or mm.group(2)):
            ok = False
            break
        d, sfx = mm.groups()
        if not d:                                   # "31A/B": same number, next letter
            if sep != "/":
                ok = False
                break
            members.append(prefix + last_s + sfx)
            continue
        if len(d) < len(last_s):                    # "010-12" -> 012, "052-3" -> 053
            d = last_s[:len(last_s) - len(d)] + d
        if sep == "-":
            a, b = int(last_s), int(d)
            if sfx or b <= a or b - a > 300:
                ok = False
                break
            members.extend(prefix + str(n).zfill(width) for n in range(a + 1, b + 1))
        else:
            members.append(prefix + d + sfx)
        last_s = d
    for b in (members if ok else members[:1]):
        if b not in out:
            out.append(b)
    return out


# ----- reading the file into tables -----

def _mf_tables_from_excel(raw, filename):
    """[{name, rows, hidden_rows, hidden_sheet, struck}] for every sheet.
    Formatting is read too: which rows are hidden (filtered out / hidden by
    hand) and which cells are crossed out."""
    tables = []
    if raw[:4] == _ZIP_MAGIC:
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        for ws in wb.worksheets:
            rows, struck = [], set()
            truncated = (ws.max_row or 0) > _MF_MAX_ROWS
            max_row = min(ws.max_row or 0, _MF_MAX_ROWS)
            max_col = min(ws.max_column or 0, _MF_MAX_COLS)
            if max_row and max_col:
                for r_i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col)):
                    vals = []
                    for c_i, c in enumerate(row):
                        vals.append(_mf_cell_text(c.value))
                        if c.value is not None and c_i < 20 and getattr(getattr(c, "font", None), "strike", False):
                            struck.add((r_i, c_i))
                    rows.append(vals)
            hidden_rows = {i - 1 for i, d in ws.row_dimensions.items() if d.hidden}
            tables.append({"name": ws.title, "rows": rows, "hidden_rows": hidden_rows,
                           "hidden_sheet": ws.sheet_state != "visible", "struck": struck, "truncated": truncated})
        return tables
    try:
        book = xlrd.open_workbook(file_contents=raw, formatting_info=True)
        fmt = True
    except Exception:
        book = xlrd.open_workbook(file_contents=raw)
        fmt = False
    for sh in book.sheets():
        nrows, ncols = min(sh.nrows, _MF_MAX_ROWS), min(sh.ncols, _MF_MAX_COLS)
        rows = [[_mf_cell_text(v) for v in sh.row_values(r, 0, ncols)] for r in range(nrows)]
        hidden_rows, struck = set(), set()
        if fmt:
            hidden_rows = {r for r, info in sh.rowinfo_map.items() if info.hidden}
            for r in range(nrows):
                for c in range(min(len(rows[r]), 20)):
                    if rows[r][c].strip():
                        try:
                            if book.font_list[book.xf_list[sh.cell_xf_index(r, c)].font_index].struck_out:
                                struck.add((r, c))
                        except Exception:
                            pass
        tables.append({"name": sh.name, "rows": rows, "hidden_rows": hidden_rows,
                       "hidden_sheet": getattr(sh, "visibility", 0) != 0, "struck": struck,
                       "truncated": sh.nrows > _MF_MAX_ROWS})
    return tables


def _mf_plain_table(name, rows):
    return {"name": name, "rows": [[_mf_cell_text(v) for v in r][:_MF_MAX_COLS] for r in rows[:_MF_MAX_ROWS]],
            "hidden_rows": set(), "hidden_sheet": False, "struck": set(), "truncated": len(rows) > _MF_MAX_ROWS}


class _MfHtmlTables(html.parser.HTMLParser):
    """Collects every <table> in an HTML document as rows of cell text -
    some systems export an ".xls" that is really an HTML page."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self._stack = [], []

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._stack.append({"rows": [], "row": None, "cell": None})
        elif not self._stack:
            return
        t = self._stack[-1]
        if tag == "tr":
            t["row"] = []
            t["rows"].append(t["row"])
        elif tag in ("td", "th"):
            if t["row"] is None:
                t["row"] = []
                t["rows"].append(t["row"])
            t["cell"] = []
            colspan = str(dict(attrs).get("colspan") or "1").strip()
            t["span"] = min(50, int(colspan)) if colspan.isdigit() and int(colspan) > 0 else 1
        elif tag == "br" and t.get("cell") is not None:
            t["cell"].append("\n")

    def handle_endtag(self, tag):
        if not self._stack:
            return
        t = self._stack[-1]
        if tag in ("td", "th") and t.get("cell") is not None:
            t["row"].append("".join(t["cell"]).strip())
            t["row"].extend([""] * (t.get("span", 1) - 1))
            t["cell"] = None
        elif tag == "table":
            self.tables.append(self._stack.pop()["rows"])

    def handle_data(self, data):
        if self._stack and self._stack[-1].get("cell") is not None:
            self._stack[-1]["cell"].append(data)


def _mf_tables_from_markup_or_text(raw):
    """An ".xls" that isn't a real Excel file: an HTML table page, an
    Excel 2003 XML spreadsheet, or plain tab/comma-separated text."""
    text = _mf_decode_csv(raw)
    head = text.lstrip("\ufeff \r\n\t")[:4000].lower()
    if "urn:schemas-microsoft-com:office:spreadsheet" in head:
        import xml.etree.ElementTree as ET
        ss = "{urn:schemas-microsoft-com:office:spreadsheet}"
        root = ET.fromstring(raw)
        tables = []
        for ws in root.iter(ss + "Worksheet"):
            rows = []
            for row in ws.iter(ss + "Row"):
                if row.get(ss + "Index", "").isdigit():
                    rows.extend([] for _ in range(int(row.get(ss + "Index")) - 1 - len(rows)))
                vals = []
                for c in row.findall(ss + "Cell"):
                    if c.get(ss + "Index", "").isdigit():
                        vals.extend([""] * (int(c.get(ss + "Index")) - 1 - len(vals)))
                    d = c.find(ss + "Data")
                    vals.append("".join(d.itertext()) if d is not None else "")
                rows.append(vals)
            tables.append(_mf_plain_table(ws.get(ss + "Name") or "Sheet", rows))
        return tables
    if head.startswith("<") or "<table" in head:
        parser = _MfHtmlTables()
        parser.feed(text)
        return [_mf_plain_table(f"Table {n + 1}", rows) for n, rows in enumerate(parser.tables)]
    # Tab/comma-separated text saved with an .xls name. Anything else (a
    # damaged file, binary junk) is reported as unreadable, not guessed at.
    sample = text[:4000]
    printable = sum(ch.isprintable() or ch in "\r\n\t" for ch in sample)
    lines = [ln for ln in sample.splitlines() if ln.strip()]
    delim = _mf_sniff_delimiter(text)
    if not lines or printable < 0.95 * len(sample) or sum(1 for ln in lines if delim in ln) < min(2, len(lines)):
        raise ValueError("not a spreadsheet")
    return [_mf_plain_table("Sheet", list(csv.reader(io.StringIO(text), delimiter=delim)))]


def _mf_sniff_delimiter(text):
    """The separator a CSV/text export actually uses - comma, semicolon
    (European Excel), tab or pipe - judged from the first lines."""
    sample = [ln for ln in text[:6000].splitlines() if ln.strip()][:20]
    best, best_score = ",", 0
    for d in (",", ";", "\t", "|"):
        counts = [ln.count(d) for ln in sample]
        score = sum(1 for c in counts if c) * 10 + sum(counts)
        if score > best_score:
            best, best_score = d, score
    return best


def _mf_decode_csv(raw):
    for enc in ("utf-8-sig", "gb18030", "cp1256"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="ignore")


# ----- finding the B/Ls -----

def _mf_free_column(rows):
    """For a table with no B/L header: the column (if any) holding nothing
    but distinct, strictly B/L-shaped values of one series. Returns
    (col, family) or (None, None)."""
    ncols = max((len(r) for r in rows), default=0)
    best = (0, None, None)
    for j in range(ncols):
        vals, nonempty = [], 0
        for r in rows:
            lines = _mf_cell_lines(r[j] if j < len(r) else "")
            if not lines:
                continue
            nonempty += 1
            if len(lines) > 2:
                continue
            got = _mf_bls_from_line(lines[0], strict=True)
            if got:
                vals.append(got[0])
        if not vals or len(vals) < 0.8 * nonempty or len(set(vals)) < 0.8 * len(vals):
            continue
        fams = {}
        for v in vals:
            fams[_mf_family(v)] = fams.get(_mf_family(v), 0) + 1
        fam, fam_n = max(fams.items(), key=lambda kv: kv[1])
        if fam is None or fam_n < 0.8 * len(vals):
            continue
        if len(vals) > best[0]:
            best = (len(vals), j, fam)
    return best[1], best[2]


def _mf_text_bls(text):
    """B/Ls from a document with no table at all (a PDF printed without grid
    lines, a Word file typed as plain lines). Only the first token of a line
    (or the value after a "B/L NO.:" label) is considered, it must be
    strictly B/L-shaped, and - unless it was labelled - its series has to
    occur at least twice, so stray words like a model or HS code can't slip in."""
    cands = []
    for line in str(text or "").splitlines():
        n = _mf_norm(line).strip()
        if not n:
            continue
        lab = re.search(r"(?:B\s*/\s*L|\bBL|BILL\s+OF\s+LADING)\s*(?:NO\b|NUMBER|NR\b)?\.?\s*[:#]\s*(\S+)", n)
        if lab:
            for b in _mf_bls_from_line(lab.group(1), strict=True):
                cands.append((b, True))
            continue
        toks = n.split()
        if len(toks) > 1 and toks[0].isdigit() and len(toks[0]) <= 4:
            toks = toks[1:]                 # "1  QCLYGJD02  N/M ..." (serial no. first)
        for b in _mf_bls_from_line(toks[0], strict=True)[:1]:
            cands.append((b, False))
    counts = {}
    for b, _ in cands:
        f = _mf_family(b)
        counts[f] = counts.get(f, 0) + 1
    out = []
    for b, labelled in cands:
        f = _mf_family(b)
        if f and (labelled or (counts[f] >= 2 and len(f) >= 3)) and b not in out:
            out.append(b)
    return out


# A second table further down the manifest sheet listing vehicles/serials
# ("NO. | CHASSIS NO. | ENGINE NO."): reading stops there until another
# B/L header row appears.
_MF_SUBTABLE_HEADER_RE = re.compile(r"\b(?:VIN|CHASSIS|ENGINE|SERIAL|FRAME)\b|车架|发动机|底盘")
_MF_REFERENCE_SHEET_RE = re.compile(r"CONTACT|REMARK|NOTES?\b|ADDRESS|联系|备注", re.IGNORECASE)
_MF_ATTACHMENT_SHEET_RE = re.compile(r"ATTACH|附件|VIN|CHASSIS|ENGINE", re.IGNORECASE)


def _mf_note(t, i, col):
    if t["hidden_sheet"]:
        return "hidden_sheet"
    if i in t["hidden_rows"]:
        return "hidden_row"
    if (i, col) in t["struck"]:
        return "struck"
    return None


def _mf_analyse(tables):
    """Decide where the B/Ls are. Returns an ordered list of
    {"bl", "note"} (note None = add it; 'hidden_row' / 'hidden_sheet' /
    'struck' = left unticked in the preview). Marks each table with its role
    ('source' = B/Ls read from it, 'reference' = contact details only) and
    the rows/column read, for the contact reader."""
    headed = []
    for t in tables:
        t["role"], t["data_rows"] = None, []
        hi, hc = _mf_find_header(t["rows"])
        t["hdr"], t["bl_col"] = hi, hc
        if hi is None:
            continue
        block = t["rows"][max(0, hi - 1):hi + 2]
        t["cargo"] = any(_mf_row_has_cargo_words(r, skip=hc) for r in block)
        t["ref_name"] = bool(_MF_REFERENCE_SHEET_RE.search(t["name"] or ""))
        headed.append(t)

    sources = [t for t in headed if t["cargo"] and not t["ref_name"]]
    if not sources:
        sources = [t for t in headed if not t["ref_name"]] or headed[:1]
    for t in headed:
        t["role"] = "source" if t in sources else "reference"

    entries, seen, families = [], {}, set()

    def add(bl, note):
        k = bl_match_key(bl)
        if k in seen:
            # A later visible copy upgrades an earlier hidden/crossed-out one.
            if note is None and entries[seen[k]]["note"] is not None:
                entries[seen[k]]["note"] = None
            return
        seen[k] = len(entries)
        entries.append({"bl": bl, "note": note})
        if note is None and _mf_family(bl):
            families.add(_mf_family(bl))

    for t in sources:
        rows, col = t["rows"], t["bl_col"]
        data = range(t["hdr"] + 1, len(rows))

        def count(j):
            return sum(1 for i in data if j < len(rows[i]) and _mf_bls_from_cell(rows[i][j], allow_numeric=True))
        # Merged header cells occasionally put the label one column off from
        # the values - if the header's column holds no B/Ls at all but a
        # neighbour does, the neighbour is the real B/L column.
        if count(col) == 0:
            alt = max((j for j in (col - 1, col + 1) if j >= 0), key=count, default=col)
            if count(alt) > 0:
                col = alt
        t["bl_col"] = col
        reading = True
        for i in data:
            row = rows[i]
            bls = _mf_bls_from_cell(row[col], allow_numeric=True) if (reading and col < len(row)) else []
            if not bls:
                # A second section further down with its own header row
                # (another page / discharge port): follow its B/L column.
                hc = _mf_header_col(row)
                if hc is not None:
                    col, reading = hc, True
                elif any(_MF_SUBTABLE_HEADER_RE.search(_mf_norm(c)) for c in row):
                    reading = False
                continue
            t["data_rows"].append((i, col, bls))
        # A vehicle VIN that still ended up in the B/L column (a chassis list
        # pasted under the manifest without its own header) is dropped -
        # unless it belongs to the same series as the sheet's real B/Ls.
        fam_count = {}
        for _, _, bls in t["data_rows"]:
            for b in bls:
                fam_count[_mf_family(b)] = fam_count.get(_mf_family(b), 0) + 1
        main_fam = max(fam_count, key=fam_count.get) if fam_count else None
        kept = []
        for i, c, bls in t["data_rows"]:
            bls = [b for b in bls if not (_mf_is_vin(b) and _mf_family(b) != main_fam)]
            if bls:
                kept.append((i, c, bls))
                for b in bls:
                    add(b, _mf_note(t, i, c))
        t["data_rows"] = kept

    # Tables with no B/L header at all. With a headed manifest in the file,
    # only an obvious continuation (same B/L series) counts. With none, the
    # first non-empty table may hold the manifest (a Word table without a
    # header row), plus its continuations.
    seed_allowed = not any(t["role"] == "source" for t in tables)
    for t in tables:
        if t["hdr"] is not None:
            continue
        rows = t["rows"]
        if not any(str(c).strip() for r in rows for c in r):
            continue
        may_seed = seed_allowed and not families and not _MF_ATTACHMENT_SHEET_RE.search(t["name"] or "")
        seed_allowed = False
        col, fam = _mf_free_column(rows)
        if col is None or not (fam in families or may_seed):
            continue
        t["role"], t["hdr"], t["bl_col"] = "source", -1, col
        for i, r in enumerate(rows):
            lines = _mf_cell_lines(r[col] if col < len(r) else "")
            if not lines or len(lines) > 2:
                continue
            bls = _mf_bls_from_line(lines[0], strict=True)
            if not bls or _mf_family(bls[0]) != fam:
                continue
            t["data_rows"].append((i, col, bls[:1]))
            add(bls[0], _mf_note(t, i, col))
    return entries


# ----- consignee contact details -----

_MF_PARTY_MARK_RE = re.compile(
    r"(?im)^[ \t]*(?:<[ \t]*(SHIPPER|SHPR|SH|CONSIGNEE|CNEE|CO|CN|ATTN|NOTIFY[ \t]*PARTY|NOTIFY|NF|NT)[ \t]*\d?[ \t]*:?[ \t]*>"
    r"|(SHIPPER|SHPR|SH|CONSIGNEE|CNEE|CO|CN|ATTN|NOTIFY[ \t]*PARTY|NOTIFY|NF|NT)[ \t]*\d?[ \t]*:)[ \t]*"
)
_MF_PARTY_KIND = {"SHIPPER": "sh", "SHPR": "sh", "SH": "sh", "CONSIGNEE": "co", "CNEE": "co", "CO": "co", "CN": "co",
                  "ATTN": "co", "NOTIFY PARTY": "nf", "NOTIFY": "nf", "NF": "nf", "NT": "nf"}


def _mf_split_parties(text):
    """A combined 'SHIPPER & CONSIGNEE & NOTIFY' cell ("<SH:>... <CO:>...
    <NF:>...", or "SH:... ATTN:... NF:...") -> {"sh": ..., "co": ..., "nf": ...}.
    Returns {} when the cell has no such markers."""
    text = unicodedata.normalize("NFKC", str(text or ""))
    marks = []
    for m in _MF_PARTY_MARK_RE.finditer(text):
        word = re.sub(r"\s+", " ", (m.group(1) or m.group(2)).upper())
        marks.append((m.start(), m.end(), _MF_PARTY_KIND[word]))
    kinds = {k for _, _, k in marks}
    if "sh" not in kinds or not ({"co", "nf"} & kinds):
        return {}
    out = {}
    for n, (_s, e, kind) in enumerate(marks):
        end = marks[n + 1][0] if n + 1 < len(marks) else len(text)
        out.setdefault(kind, text[e:end].strip())
    return out


_MF_TO_ORDER_RE = re.compile(r"^\W*(?:TO\s+(?:THE\s+)?ORDER|TO\s+SHIPPER'?S?\s+ORDER|ORDER\s+OF)\b", re.IGNORECASE)
_MF_SAME_AS_RE = re.compile(r"^\W*SAME\s+AS\s+(?:THE\s+)?CONSIGNEE", re.IGNORECASE)
_MF_NAME_CONT_RE = re.compile(r"^(?:AND\b|&|CO\b|CO\.|COMPANY\b|LTD|LIMITED\b|LLC\b|L\.L\.C|W\.L\.L|EST\b|EST\.|ESTABLISHMENT\b|FOR\b)", re.IGNORECASE)
_MF_ADDRESS_START_RE = re.compile(r"[\s,]+(?:P\.?\s?O\.?\s*BOX\b|POB\b|BUILDING\s*(?:NO\b|NUMBER\b|#|\d)|BLDG\b|ADD(?:RESS)?\s*[:.]|\d)", re.IGNORECASE)
_MF_NAME_LABEL_RE = re.compile(r"^(?:CONSIGNEE|CNEE|CO|CN|ATTN|NAME|COMPANY\s+NAME)\s*[:：]\s*", re.IGNORECASE)


def _mf_party_name(block):
    """Company name = first line of the party block; a name that runs onto
    a second line ("SALEM BIN AHMED ... / AND SONS CO. LTD.") is joined."""
    lines = [re.sub(r"\s+", " ", ln.replace("\xa0", " ")).strip(" ,;:-") for ln in str(block or "").splitlines()]
    lines = [ln for ln in lines if ln]
    if not lines:
        return ""
    name = _MF_NAME_LABEL_RE.sub("", lines[0]).strip(" ,;:-")
    if len(name) > 40:
        # Name and address typed on one line ("TALAI AL-FOLATH FOR
        # MANUFACTURING CO 2ND INDUSTRIAL AREA P.O BOX 22766, JEDDAH"):
        # keep the part before the address starts.
        cut = _MF_ADDRESS_START_RE.search(name)
        if cut and cut.start() >= 8:
            name = name[:cut.start()].strip(" ,;:-")
    if len(lines) > 1 and _MF_NAME_CONT_RE.match(lines[1]) and not re.search(r"\d", lines[1]):
        name = f"{name} {lines[1]}"
    return name[:120]


_MF_PHONE_GROUP_RE = re.compile(r"\+?\d+")


def _mf_phones(text):
    """Saudi phone numbers in a block of text -> (mobiles, landlines), each
    normalized to +966... . Digit groups are joined only across spaces,
    dots, dashes and brackets on the same line, and a number is only read
    from where a phone number can start (the first group of a run, a group
    starting with 0 / +, or right after another number) - so two numbers
    typed side by side ('0126081236 0505584273') aren't merged, and the tail
    of a foreign number ('+971 50 739 7579') is never mistaken for a Saudi
    mobile. Numbers labelled FAX are skipped."""
    mob, land = [], []
    for line in str(text or "").replace("\xa0", " ").splitlines():
        groups = [(m.start(), m.end(), m.group(0)) for m in _MF_PHONE_GROUP_RE.finditer(line)]

        def joined(a, b):
            return re.fullmatch(r"[ \t().-]{0,3}", line[groups[a][1]:groups[b][0]]) and not groups[b][2].startswith("+")
        k, after_number = 0, False
        while k < len(groups):
            can_start = k == 0 or not joined(k - 1, k) or after_number or groups[k][2].startswith(("0", "+"))
            used = None
            if can_start:
                digits = ""
                for e in range(k, min(k + 6, len(groups))):
                    if e > k and not joined(e - 1, e):
                        break
                    digits += groups[e][2]
                    p = _normalize_phone(digits)
                    if re.fullmatch(r"\+966[15]\d{8}", p):
                        used = (e, p)
                        break
            if used:
                e, p = used
                label = re.findall(r"[A-Z]+", line[:groups[k][0]].upper()[-14:])
                if not (label and label[-1] == "FAX"):
                    bucket = mob if p.startswith("+9665") else land
                    if p not in bucket:
                        bucket.append(p)
                k, after_number = e + 1, True
            else:
                k, after_number = k + 1, False
    return mob, land


def _mf_emails(text):
    out = []
    for e in _EMAIL_RE.findall(unicodedata.normalize("NFKC", str(text or ""))):
        e = _clean_email(e.strip(".;,-"))
        if e and e not in out:
            out.append(e)
    return out


_MF_GENERIC_MAIL_LABELS = {"GMAIL", "HOTMAIL", "YAHOO", "OUTLOOK", "ICLOUD", "LIVE", "MAIL", "EMAIL", "FOXMAIL",
                           "SINA", "SOHU", "ALIYUN", "YMAIL", "YANDEX", "PROTONMAIL", "MSN", "AOL"}


def _mf_compact(text):
    return re.sub(r"[^A-Z0-9]", "", _mf_norm(text))


def _mf_contact_from_parts(consignee, notify, extras, shipper=""):
    """name/email/phone for one BL. The consignee is the customer, unless
    the B/L is consigned 'TO ORDER' (of a bank / the shipper) - then the
    consignee block is the bank's own address, so it is skipped and the
    notify party (the actual buyer) is who gets contacted; the consignee name
    is then left blank rather than filled with someone who isn't the
    consignee. The shipper is never used. extras = other contact text for
    the BL (CONTACT / TEL / E-MAIL / REMARKS columns, a contact sheet).
    An e-mail whose domain is the shipper's own name (TOM@TOBEESTEEL.COM
    with shipper "TOBEE STEEL LIMITED") is the shipper's, not the
    customer's, and is skipped - unless the consignee/notify party carries
    that name too (a local branch of the same group)."""
    consignee = str(consignee or "").strip()
    notify = str(notify or "").strip()
    if _MF_SAME_AS_RE.match(notify):
        notify = ""
    to_order = bool(_MF_TO_ORDER_RE.match(_mf_party_name(consignee)))
    name = "" if to_order else _mf_party_name(consignee)
    sources = ([] if to_order else [consignee]) + [notify] + [x for x in extras if x]
    shipper_key = _mf_compact(_mf_party_name(shipper))
    party_key = _mf_compact(name + " " + _mf_party_name(notify))

    def shippers_own(e):
        labels = [_mf_compact(lb) for lb in e.split("@", 1)[1].upper().split(".")]
        return any(len(lb) >= 4 and lb not in _MF_GENERIC_MAIL_LABELS and lb in shipper_key and lb not in party_key
                   for lb in labels)
    email = next((e for s in sources for e in _mf_emails(s) if not (shipper_key and shippers_own(e))), "")
    phones = [_mf_phones(s) for s in sources]
    phone = next((m[0] for m, _ in phones if m), "") or next((l[0] for _, l in phones if l), "")
    return {"name": name, "email": email, "phone": phone}


_MF_CONTACT_COL_WORDS = ("CONTACT", "TEL", "PHONE", "MOBILE", "EMAIL", "E-MAIL", "REMARK", "TERM", "联系", "备注", "电话", "邮箱")


def _mf_contacts(tables):
    """{bl_match_key(bl): {"name", "email", "phone"}} from the tables the
    B/Ls were read from (and any contact/remarks sheet keyed by B/L)."""
    ref = {}   # individual B/L -> contact text from a contact/remarks sheet
    for t in tables:
        if t.get("role") != "reference":
            continue
        rows, col = t["rows"], t["bl_col"]
        for row in rows[t["hdr"] + 1:]:
            if col >= len(row):
                continue
            text = "\n".join(str(c) for j, c in enumerate(row) if j != col and str(c).strip())
            for b in _mf_bls_from_cell(row[col], allow_numeric=True):
                for mbr in bl_members(b):
                    ref.setdefault(mbr, text)

    out = {}
    for t in tables:
        if t.get("role") != "source" or not t["data_rows"]:
            continue
        rows = t["rows"]
        cols = {"party": None, "sh": None, "co": None, "nf": None, "extra": []}
        if t["hdr"] >= 0:
            hdr = t["hdr"]
            # The header row, plus the row above/below when it is the other
            # half of a bilingual (Chinese + English) header.
            block = [rows[hdr]] + [rows[i] for i in (hdr - 1, hdr + 1)
                                   if 0 <= i < len(rows) and _mf_row_has_cargo_words(rows[i])
                                   and not any(i == d[0] for d in t["data_rows"])]
            width = max(len(r) for r in block)
            labelled_upto = -1
            for j in range(width):
                text = re.sub(r"\s+", "", " ".join(_mf_norm(r[j]) for r in block if j < len(r)))
                if not text:
                    continue
                labelled_upto = j
                if j == t["bl_col"]:
                    continue
                is_sh = "SHIPPER" in text or "货主" in text or "发货人" in text
                is_co = "CONSIGNEE" in text or "收货人" in text
                is_nf = "NOTIFY" in text or "通知人" in text
                if is_sh and (is_co or is_nf):
                    cols["party"] = j if cols["party"] is None else cols["party"]
                elif is_sh:
                    cols["sh"] = j if cols["sh"] is None else cols["sh"]
                elif is_co:
                    cols["co"] = j if cols["co"] is None else cols["co"]
                elif is_nf:
                    cols["nf"] = j if cols["nf"] is None else cols["nf"]
                elif any(w in text for w in _MF_CONTACT_COL_WORDS):
                    cols["extra"].append(j)
            # Unlabelled columns to the right of the header often hold a
            # consignee phone/e-mail typed in afterwards.
            data_width = max(len(rows[i]) for i, _, _ in t["data_rows"])
            cols["extra"].extend(range(labelled_upto + 1, data_width))
        else:
            # No header (a Word table): find the combined party column by its markers.
            for j in range(max(len(r) for r in rows)):
                hits = sum(1 for i, _, _ in t["data_rows"] if j < len(rows[i]) and _mf_split_parties(rows[i][j]))
                if hits >= 0.5 * len(t["data_rows"]):
                    cols["party"] = j
                    break

        def cell(row, j):
            return str(row[j]) if j is not None and j < len(row) else ""

        for i, _, bls in t["data_rows"]:
            row = rows[i]
            consignee, notify, shipper = cell(row, cols["co"]), cell(row, cols["nf"]), cell(row, cols["sh"])
            if cols["party"] is not None:
                parts = _mf_split_parties(cell(row, cols["party"]))
                consignee = consignee or parts.get("co", "")
                notify = notify or parts.get("nf", "")
                shipper = shipper or parts.get("sh", "")
            extras = [cell(row, j) for j in cols["extra"]]
            for b in bls:
                sheet = next((ref[m] for m in bl_members(b) if m in ref), "")
                c = _mf_contact_from_parts(consignee, notify, extras + [sheet], shipper)
                if c["name"] or c["email"] or c["phone"]:
                    out.setdefault(bl_match_key(b), c)
    return out


def read_manifest(raw, filename):
    """Reads one uploaded manifest file (.xlsx/.xlsm/.xls/.csv/.docx/.pdf).
    Returns {"entries": [{"bl", "note", "contact"}...]}; entries is empty
    when no B/L column could be found. Raises for a file that can't be
    opened at all (corrupt / password-protected / unsupported)."""
    name = (filename or "").lower()
    if raw[:2] == b"PK":   # xlsx/docx are zip files: refuse ones that swell to an absurd size when opened
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                if sum(i.file_size for i in z.infolist()) > 300 * 1024 * 1024 or len(z.infolist()) > 5000:
                    raise ValueError("archive too large when unpacked")
        except zipfile.BadZipFile:
            pass
    tables, text = [], ""
    if name.endswith(".csv"):
        text_csv = _mf_decode_csv(raw)
        tables = [_mf_plain_table("CSV", list(csv.reader(io.StringIO(text_csv), delimiter=_mf_sniff_delimiter(text_csv))))]
    elif name.endswith(".docx"):
        import docx
        document = docx.Document(io.BytesIO(raw))
        tables = [_mf_plain_table(f"Table {n + 1}", [[c.text for c in r.cells] for r in tb.rows])
                  for n, tb in enumerate(document.tables)]
        text = "\n".join(p.text for p in document.paragraphs)
    elif name.endswith(".pdf"):
        import pdfplumber
        pages = []
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for page in pdf.pages:
                for tb in (page.extract_tables() or []):
                    if tb:
                        tables.append(_mf_plain_table(f"Page {page.page_number}", [[c or "" for c in r] for r in tb]))
                pages.append(page.extract_text() or "")
        text = "\n".join(pages)
    elif name.endswith((".xlsx", ".xlsm", ".xls")):
        if raw[:4] == _ZIP_MAGIC or raw[:8] == _OLE_MAGIC:
            tables = _mf_tables_from_excel(raw, name)
        else:
            tables = _mf_tables_from_markup_or_text(raw)
    else:
        raise ValueError("unsupported file type")

    entries = _mf_analyse(tables)
    if not any(e["note"] is None for e in entries) and text.strip():
        known = {bl_match_key(e["bl"]) for e in entries}
        for b in _mf_text_bls(text):
            if bl_match_key(b) not in known:
                entries.append({"bl": b, "note": None})
    contacts = _mf_contacts(tables) if entries else {}
    for e in entries:
        e["contact"] = contacts.get(bl_match_key(e["bl"]), {"name": "", "email": "", "phone": ""})
    return {"entries": entries, "truncated": any(t.get("truncated") for t in tables)}


# ----- adding manifest B/Ls to the board -----

_MANIFEST_EXTS = (".xlsx", ".xlsm", ".xls", ".csv", ".docx", ".pdf")
_MANIFEST_MAX_BYTES = 25 * 1024 * 1024


def _manifest_board_index():
    """bl_match_key -> the existing record, for every B/L on the board (any
    owner, archived included) - one query instead of one per B/L."""
    rows = get_db().execute(
        "SELECT bl_number, vessel, port, created_by, consignee, consignee_email, consignee_phone FROM records"
    ).fetchall()
    return {bl_match_key(r["bl_number"]): dict(r) for r in rows}


def _manifest_add(port, vessel, items):
    """Adds manifest B/Ls to the board under one Port + Vessel.
    items: [{"bl", "name", "email", "phone"}]. A B/L already on the board
    (compared by bl_match_key, so separator/space differences don't create a
    second copy) is skipped - re-uploading only fills in consignee contacts
    that are still blank, never overwrites one (it may have been corrected
    by hand), and only on the uploader's own B/Ls (or any, for an admin).
    A B/L already on the board under a DIFFERENT port/vessel is reported
    back, since that is either a data error or a reused number."""
    db = get_db()
    index = _manifest_board_index()
    me = session.get("username")
    is_admin = session.get("role") == "admin"
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M")
    added = skipped = contacts_filled = 0
    duplicate_elsewhere = []
    for item in items:
        if not isinstance(item, dict):
            continue
        # Same rule as reading a manifest: one clean B/L number, nothing else
        # (so a hand-crafted request can't put junk or a spreadsheet formula
        # on the board either).
        cleaned = _mf_bls_from_line(str(item.get("bl") or "")[:80], allow_numeric=True)
        if len(cleaned) != 1:
            continue
        bl = bl_match_key(cleaned[0])
        found = {"name": re.sub(r"\s+", " ", str(item.get("name") or "")).strip()[:120],
                 "email": _clean_email(item.get("email")),
                 "phone": _normalize_phone(item.get("phone"))}
        existing = index.get(bl)
        if existing is None:
            inserted = db.execute(
                "INSERT INTO records (bl_number, port, vessel, created_at, created_by, consignee, consignee_email, consignee_phone) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT (bl_number) DO NOTHING RETURNING bl_number",
                (bl, port, vessel, now, me, found["name"], found["email"], found["phone"]),
            ).fetchone()
            if inserted:
                index[bl] = {"bl_number": bl, "vessel": vessel, "port": port, "created_by": me,
                             "consignee": found["name"], "consignee_email": found["email"], "consignee_phone": found["phone"]}
                if found["email"] or found["phone"]:
                    contacts_filled += 1
                _log_audit(bl, "added", "", "", f"{port} / {vessel}")
                added += 1
                continue
            # Someone else added it a moment ago - treat like any existing B/L.
            existing = dict(db.execute(
                "SELECT bl_number, vessel, port, created_by, consignee, consignee_email, consignee_phone FROM records WHERE bl_number = ?",
                (bl,),
            ).fetchone())
            index[bl] = existing
        skipped += 1
        mine = is_admin or existing["created_by"] == me
        if mine and ((existing["vessel"] or "") != vessel or (existing["port"] or "") != port):
            duplicate_elsewhere.append({"bl_number": existing["bl_number"],
                                        "existing_port": existing["port"], "existing_vessel": existing["vessel"]})
        if mine:
            fills = {k: found[src] for k, src in (("consignee", "name"), ("consignee_email", "email"), ("consignee_phone", "phone"))
                     if found[src] and not (existing[k] or "").strip()}
            if fills:
                db.execute("UPDATE records SET " + ", ".join(f"{k} = ?" for k in fills) + " WHERE bl_number = ?",
                           (*fills.values(), existing["bl_number"]))
                existing.update(fills)
                contacts_filled += 1
    db.commit()
    return {"added": added, "skipped": skipped, "duplicate_elsewhere": duplicate_elsewhere,
            "contacts_found": contacts_filled}


def _manifest_read_upload(file, with_truncation=False):
    """(entries, error_code) for one uploaded manifest file - plus whether
    a huge sheet had to be cut short, when with_truncation is set."""
    def done(entries, err, truncated=False):
        return (entries, err, truncated) if with_truncation else (entries, err)
    name = file.filename or ""
    if not name.lower().endswith(_MANIFEST_EXTS):
        return done([], "unsupported")
    raw = file.read(_MANIFEST_MAX_BYTES + 1)
    if len(raw) > _MANIFEST_MAX_BYTES:
        return done([], "too_large")
    if not raw:
        return done([], "empty")
    try:
        result = read_manifest(raw, name)
    except Exception:
        return done([], "unreadable")
    return done(result["entries"], None if result["entries"] else "no_bls", result.get("truncated", False))


_MANIFEST_ERROR_TEXT = {
    "unsupported": "Unsupported file type. Please upload .xlsx, .xls, .csv, .docx or .pdf.",
    "too_large": "That file is larger than 25MB.",
    "empty": "That file is empty.",
    "unreadable": "Couldn't read that file. Make sure it isn't corrupted or password-protected.",
    "no_bls": "No B/L numbers were found in that file - it needs a B/L NO. column.",
}


@app.route("/api/manifest/preview", methods=["POST"])
@login_required
def manifest_preview():
    """Reads one or more manifests WITHOUT adding anything, and reports
    every B/L found with its status, so the user checks the list before it
    goes on the board ("Add a manifest" -> preview -> confirm)."""
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "No file received", "error_code": "no_file"}), 400
    port = request.form.get("port", "").strip().upper()
    vessel = request.form.get("vessel", "").strip().upper()
    index = _manifest_board_index()
    me = session.get("username")
    is_admin = session.get("role") == "admin"
    in_batch = set()
    out = []
    for f in files:
        entries, err, truncated = _manifest_read_upload(f, with_truncation=True)
        rows = []
        for e in entries:
            k = bl_match_key(e["bl"])
            ex = index.get(k)
            fills = False
            if ex is not None and not (is_admin or ex["created_by"] == me):
                # Another user's BL: say it's taken, but not where it is or
                # whose it is (staff only ever see their own BLs).
                status, ex = "taken", None
            elif ex is not None:
                same = (ex["vessel"] or "") == vessel and (ex["port"] or "") == port
                status = "on_board" if same else "elsewhere"
                c = e["contact"]
                fills = any(c.get(src) and not (ex[k2] or "").strip()
                            for k2, src in (("consignee", "name"), ("consignee_email", "email"), ("consignee_phone", "phone")))
            elif k in in_batch:
                status = "repeat"
            else:
                status = "new"
            if status == "new" and e["note"] is None:
                in_batch.add(k)
            rows.append({
                "bl": k, "note": e["note"], "status": status, "contact": e["contact"], "fills_contacts": fills,
                "existing_bl": ex["bl_number"] if ex else None,
                "existing_port": ex["port"] if ex else None, "existing_vessel": ex["vessel"] if ex else None,
            })
        out.append({"name": f.filename, "entries": rows, "error_code": err, "truncated": truncated})
    return jsonify({"files": out, "port": port, "vessel": vessel})


@app.route("/api/manifest/commit", methods=["POST"])
@login_required
def manifest_commit():
    """Adds the B/Ls the user ticked in the manifest preview."""
    data = request.get_json(force=True, silent=True) or {}
    items = data.get("items")
    if not isinstance(items, list) or not items:
        return jsonify({"error": "Nothing selected to add.", "error_code": "nothing_selected"}), 400
    if len(items) > 5000:
        return jsonify({"error": "Too many B/Ls in one go.", "error_code": "too_many"}), 400
    port = str(data.get("port") or "").strip().upper()
    vessel = str(data.get("vessel") or "").strip().upper()
    return jsonify(_manifest_add(port, vessel, items))


@app.route("/api/manifest/upload", methods=["POST"])
@login_required
def upload_manifest_excel():
    """One-step upload (read + add everything that should be added). The
    board itself now goes through preview + commit; this stays for anything
    still calling the old endpoint."""
    if "file" not in request.files:
        return jsonify({"error": "No file received"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400
    entries, err = _manifest_read_upload(file)
    if err and err != "no_bls":
        return jsonify({"error": _MANIFEST_ERROR_TEXT[err], "error_code": err}), 400
    port = request.form.get("port", "").strip().upper()
    vessel = request.form.get("vessel", "").strip().upper()
    items = [{"bl": e["bl"], **e["contact"]} for e in entries if e["note"] is None]
    result = _manifest_add(port, vessel, items) if items else {"added": 0, "skipped": 0, "duplicate_elsewhere": [], "contacts_found": 0}
    # Hidden rows / hidden sheets / crossed-out B/Ls aren't added; say how many.
    result["left_out"] = sum(1 for e in entries if e["note"] is not None)
    result["no_bls"] = not entries
    return jsonify(result)


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


ATTACHMENT_TRASH_DAYS = 7


def _stash_attachments(db, bl_number):
    """Copies a BL's attached files into record_attachments_trash right
    before the BL itself is deleted (which cascade-deletes the originals),
    so an Undo can put them back. Also purges stashed files older than
    ATTACHMENT_TRASH_DAYS. Shares the caller's transaction."""
    now_dt = datetime.utcnow()
    cutoff = (now_dt - timedelta(days=ATTACHMENT_TRASH_DAYS)).strftime("%Y-%m-%d %H:%M")
    db.execute("DELETE FROM record_attachments_trash WHERE deleted_at < ?", (cutoff,))
    db.execute("DELETE FROM record_attachments_trash WHERE bl_number = ?", (bl_number,))
    db.execute(
        """INSERT INTO record_attachments_trash
             (bl_number, kind, filename, content_type, data, file_size, uploaded_by, uploaded_at, deleted_by, deleted_at)
           SELECT bl_number, kind, filename, content_type, data, file_size, uploaded_by, uploaded_at, ?, ?
           FROM record_attachments WHERE bl_number = ?""",
        (session.get("username", ""), now_dt.strftime("%Y-%m-%d %H:%M"), bl_number),
    )


def _unstash_attachments(db, bl_number):
    """Moves a BL's stashed files back into record_attachments on Undo.
    Only for an admin or whoever removed the BL - otherwise anyone who
    knew a removed BL's number could "restore" it under their own name
    and pick up someone else's files along with it."""
    params = [bl_number]
    owner_sql = ""
    if session.get("role") != "admin":
        owner_sql = " AND deleted_by = ?"
        params.append(session.get("username", ""))
    db.execute(
        f"""INSERT INTO record_attachments
              (bl_number, kind, filename, content_type, data, file_size, uploaded_by, uploaded_at)
            SELECT bl_number, kind, filename, content_type, data, file_size, uploaded_by, uploaded_at
            FROM record_attachments_trash WHERE bl_number = ?{owner_sql}
            ON CONFLICT (bl_number, kind) DO NOTHING""",
        tuple(params),
    )
    db.execute(f"DELETE FROM record_attachments_trash WHERE bl_number = ?{owner_sql}", tuple(params))


@app.route("/api/records/<path:bl_number>", methods=["DELETE"])
@login_required
def delete_record(bl_number):
    if not _owns_record(bl_number.upper()):
        return "Not your record.", 403
    db = get_db()
    _stash_attachments(db, bl_number.upper())
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

    # Staff can only ever remove their own BLs, so an Undo by staff always
    # restores to themselves - the client-supplied created_by is only
    # trusted from an admin (who may be restoring someone else's BL).
    if session.get("role") == "admin":
        owner = data.get("created_by") or session.get("username")
    else:
        owner = session.get("username")

    # Keep the BL's tracking link working after an Undo (the customer may
    # already have it) - unless that code is somehow in use elsewhere, in
    # which case the BL just gets a fresh one when next shared.
    token = str(data.get("track_token") or "").strip()
    if token and (not re.fullmatch(r"[A-Za-z0-9_-]{16,64}", token)
                  or db.execute("SELECT 1 FROM records WHERE track_token = ?", (token,)).fetchone()):
        token = ""

    db.execute(
        """INSERT INTO records
           (bl_number, consignee, port, vessel, invoice_issued, invoice_by, invoice_at,
            approval_received, approval_by, approval_at, do_issued, do_by, do_at,
            remarks, created_at, created_by, eta, archived,
            consignee_email, consignee_phone, broker_email, broker_phone, track_token)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
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
            owner,
            str(data.get("eta") or ""),
            1 if data.get("archived") else 0,
            _clean_email(data.get("consignee_email")),
            _normalize_phone(data.get("consignee_phone")),
            _clean_email(data.get("broker_email")),
            _normalize_phone(data.get("broker_phone")),
            token,
        ),
    )
    _unstash_attachments(db, bl_number)
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
        _stash_attachments(db, bl_number)
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
    # Stored upper-case, like everything a manifest upload files BLs under
    # (otherwise "Tai Knight" and "TAI KNIGHT" become two different groups).
    new_value = re.sub(r"\s+", " ", str(data.get("new_value") or "")).strip().upper()
    db = get_db()
    is_admin = session.get("role") == "admin"
    if group_type not in ("port", "vessel"):
        return jsonify({"error": "invalid type"}), 400
    bl_numbers = data.get("bl_numbers")
    if isinstance(bl_numbers, list):
        # The board sends the group's own BLs - this also works for an
        # "Unassigned" group, whose port/vessel is stored blank.
        bls = [str(b).strip().upper() for b in bl_numbers if str(b).strip()]
        if bls:
            sql = f"UPDATE records SET {group_type} = ? WHERE bl_number = ANY(?)"
            params = [new_value, bls]
            if not is_admin:
                sql += " AND created_by = ?"
                params.append(session.get("username"))
            db.execute(sql, tuple(params))
            db.commit()
        return jsonify({"ok": True})
    old_port = data.get("old_port", "")
    old_port = "" if old_port == "Unassigned" else old_port
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
        old_vessel = "" if old_vessel == "Unassigned" else old_vessel
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

    # One UPDATE for the whole selection (it used to be ~3 database round
    # trips per BL, which on Render made a big vessel's Mark/Unmark slow
    # enough for the board's background refresh to flip sliders back).
    # Only rows whose status actually changes are touched: marking a BL
    # that's already marked must NOT overwrite who marked it and when.
    db = get_db()
    sql = (
        f"UPDATE records SET {field} = ?, {by_field} = ?, {at_field} = ? "
        f"WHERE bl_number = ANY(?) AND COALESCE({field}, 0) <> ?"
    )
    params = [value, by_val, now, bl_numbers, value]
    if session.get("role") != "admin":
        sql += " AND created_by = ?"  # staff can only touch their own BLs, same as _owns_record
        params.append(session.get("username"))
    sql += " RETURNING bl_number"
    updated = [r["bl_number"] for r in db.execute(sql, tuple(params)).fetchall()]
    for bl_number in updated:
        _log_audit(bl_number, "toggle", field, "" if value else "1", "1" if value else "")
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
    "BO1234") doesn't produce a false match. A combined board entry
    ("BO26215XJED001-003", "QCLYGJD26/27/28") also matches a document that
    names just one of the B/Ls it covers (see bl_members). Returns
    [(bl, exact)] - exact=False when only such a member matched; the caller
    treats exactly one as a confident auto-match and zero-or-many as
    needing a human to pick."""
    upper = _mf_norm(text)
    tokens = set(re.findall(r"[A-Z0-9]+", upper))

    def present(s):
        if s.isalnum():
            return s in tokens
        return re.search(r"(?<![A-Z0-9])" + re.escape(s) + r"(?![A-Z0-9])", upper) is not None

    found = []
    for bl in known_bls:
        bl_u = (bl or "").strip().upper()
        if not bl_u:
            continue
        members = bl_members(bl_u)
        if present(bl_u) or present(members[0]):
            found.append((bl, True))
        elif any(present(m) for m in members[1:]):
            found.append((bl, False))
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
    found = _find_bl_in_text(text, known_bls)
    matches = [bl for bl, _ in found]
    reason = None
    if len(found) == 1 and not found[0][1]:
        # Matched only through a combined entry (the document names ONE of
        # the B/Ls it covers, e.g. 001 of "001-003"). Attaching it marks the
        # whole entry issued, so a person confirms instead of it happening
        # automatically - and if the entry already has this kind of file,
        # confirming would replace it, which the message says.
        reason = "member_match"
        if kind and db.execute(
            "SELECT 1 FROM record_attachments WHERE bl_number = ? AND kind = ?", (matches[0], kind)
        ).fetchone():
            reason = "already_attached"

    return jsonify({
        "kind": kind,
        "matched_bl": matches[0] if len(matches) == 1 and not reason else None,
        "candidates": matches if len(matches) > 1 or reason else [],
        "reason": reason,
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
        headers={"Content-Disposition": _content_disposition(filename, f"{kind}_{bl_number}.pdf")},
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
    """Looks up one BL by its exact number for the Documents popup (the doc
    chips next to a BL's own row). Admins can open any BL; staff only
    their own."""
    bl_number = bl_number.strip().upper()
    if not bl_number:
        return jsonify({"error": "Enter a BL number."}), 400
    # Staff only ever see their own BLs on the board (cross-staff "Find a
    # BL" was removed on purpose), so this must not hand out another
    # user's consignee contacts or tracking link either.
    if not _owns_record(bl_number):
        return jsonify({"error": f'No BL found matching "{bl_number}".'}), 404
    db = get_db()
    row = db.execute(
        f"""SELECT bl_number, port, vessel, created_by, invoice_issued, invoice_by, invoice_at,
                   do_issued, do_by, do_at, consignee, consignee_email, consignee_phone,
                   broker_email, broker_phone, track_token{ATTACHMENT_FLAGS_SQL}
            FROM records r WHERE bl_number = ?""",
        (bl_number,),
    ).fetchone()
    if not row:
        return jsonify({"error": f'No BL found matching "{bl_number}".'}), 404
    result = dict(row)
    token = result.pop("track_token", "") or ""
    result["track_url"] = f"{_public_base_url()}/t/{token}" if token else ""
    result["email_configured"] = email_configured()
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


# ---------- Customer sharing: contacts, tracking link, QR, notifications ----------

# Arabic port names for customer-facing text (the board stores the English,
# upper-case name; matches the DO Tracker's own Arabic labels).
PORT_NAMES_AR = {
    "DAMMAM PORT": "ميناء الدمام",
    "JUBAIL COMMERCIAL PORT": "ميناء الجبيل التجاري",
    "JEDDAH PORT": "ميناء جدة",
    "YANBU COMMERCIAL PORT": "ميناء ينبع التجاري",
    "YANBU INDUSTRIAL PORT": "ميناء ينبع الصناعي",
}
COMPANY_NAME_EN = "Sea Power Marine Services Co. Ltd"
COMPANY_NAME_AR = "شركة سي باور للخدمات البحرية المحدودة"


def _public_base_url():
    """Base URL for links sent to customers. PUBLIC_BASE_URL (env) wins if
    set - e.g. once Compass has its own domain; otherwise the address this
    request came in on. Render terminates HTTPS at its proxy, so the real
    scheme comes from X-Forwarded-Proto, not request.scheme ("http")."""
    env = os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/")
    if env:
        return env
    proto = (request.headers.get("X-Forwarded-Proto") or request.scheme or "https").split(",")[0].strip()
    return f"{proto}://{request.host}"


def _ensure_track_token(db, bl_number, force_new=False):
    """The BL's tracking code, creating one on first use. 16 random bytes
    (~128 bits) - impossible to guess, so the link itself is the key."""
    row = db.execute("SELECT track_token FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
    if not row:
        return None
    if row["track_token"] and not force_new:
        return row["track_token"]
    token = secrets.token_urlsafe(16)
    db.execute("UPDATE records SET track_token = ? WHERE bl_number = ?", (token, bl_number))
    db.commit()
    return token


@app.route("/api/records/<path:bl_number>/contacts", methods=["POST"])
@login_required
def update_contacts(bl_number):
    bl_number = bl_number.upper()
    if not _owns_record(bl_number):
        return jsonify({"error": "Not your record.", "error_code": "not_owner"}), 403
    data = request.get_json(force=True) or {}
    vals = {}
    for key in ("consignee_email", "broker_email"):
        raw = str(data.get(key) or "").strip()
        if raw and not _clean_email(raw):
            return jsonify({"error": f"That doesn't look like an email address: {raw}", "error_code": "bad_email", "field": key}), 400
        vals[key] = _clean_email(raw)
    for key in ("consignee_phone", "broker_phone"):
        raw = str(data.get(key) or "").strip()
        if raw and not _normalize_phone(raw):
            return jsonify({"error": f"That doesn't look like a phone number: {raw}", "error_code": "bad_phone", "field": key}), 400
        vals[key] = _normalize_phone(raw)
    if "consignee" in data:
        vals["consignee"] = str(data.get("consignee") or "").strip()[:120]
    db = get_db()
    db.execute(
        "UPDATE records SET " + ", ".join(f"{k} = ?" for k in vals) + " WHERE bl_number = ?",
        (*vals.values(), bl_number),
    )
    _log_audit(bl_number, "contacts")
    db.commit()
    return jsonify({"ok": True, **vals})


@app.route("/api/records/<path:bl_number>/share-link", methods=["POST"])
@login_required
def share_link(bl_number):
    """Returns the BL's public tracking link, creating it on first use.
    ?reset=1 replaces it with a new one - the old link stops working
    (use when a link has gone to the wrong person)."""
    bl_number = bl_number.upper()
    if not _owns_record(bl_number):
        return jsonify({"error": "Not your record.", "error_code": "not_owner"}), 403
    db = get_db()
    reset = request.args.get("reset") == "1"
    token = _ensure_track_token(db, bl_number, force_new=reset)
    if not token:
        return jsonify({"error": "BL not found."}), 404
    if reset:
        _log_audit(bl_number, "link_reset")
        db.commit()
    return jsonify({"ok": True, "url": f"{_public_base_url()}/t/{token}"})


@app.route("/api/records/<path:bl_number>/qr", methods=["GET"])
@login_required
def share_qr(bl_number):
    """QR code of the tracking link - SVG to show on screen, ?format=png
    (and &download=1) for pasting into a message or printing."""
    import segno
    bl_number = bl_number.upper()
    if not _owns_record(bl_number):
        return "Not your record.", 403
    db = get_db()
    token = _ensure_track_token(db, bl_number)
    if not token:
        return "BL not found.", 404
    qr = segno.make(f"{_public_base_url()}/t/{token}", error="m")
    buf = io.BytesIO()
    as_png = request.args.get("format") == "png"
    if as_png:
        qr.save(buf, kind="png", scale=10, border=3)
    else:
        qr.save(buf, kind="svg", scale=6, border=2, dark="#0b2740", xmldecl=False)
    headers = {"Cache-Control": "no-store"}
    if request.args.get("download") == "1":
        headers["Content-Disposition"] = _content_disposition(f"QR_{bl_number}.{'png' if as_png else 'svg'}", "qr.png")
    return Response(buf.getvalue(), mimetype="image/png" if as_png else "image/svg+xml", headers=headers)


def _notify_texts(rec, url):
    """Bilingual (English, then Arabic) customer messages. Formal, plain
    wording - reviewed Arabic should replace this draft if needed."""
    bl, vessel = rec["bl_number"], (rec["vessel"] or "-")
    port_en = (rec["port"] or "-").title()
    port_ar = PORT_NAMES_AR.get((rec["port"] or "").upper(), rec["port"] or "-")
    name = (rec["consignee"] or "").strip()
    subject = f"Delivery Order Issued - B/L {bl} - {vessel} | تم إصدار أمر التسليم - بوليصة {bl}"
    body = (
        f"Dear {name or 'Valued Customer'},\n\n"
        f"The Delivery Order for the shipment below has been issued and is attached to this email.\n\n"
        f"B/L No.: {bl}\nVessel: {vessel}\nPort of Discharge: {port_en}\n\n"
        f"Track the status of this shipment at any time:\n{url}\n\n"
        f"Regards,\n{COMPANY_NAME_EN}\n\n"
        f"------------------------------------------------------------\n\n"
        f"{('السادة ' + name) if name else 'عميلنا العزيز'}،\n\n"
        f"نفيدكم بأنه تم إصدار أمر التسليم للشحنة الموضحة أدناه، وتجدونه مرفقًا بهذه الرسالة.\n\n"
        f"رقم البوليصة: {bl}\nالسفينة: {vessel}\nميناء التفريغ: {port_ar}\n\n"
        f"يمكنكم متابعة حالة الشحنة في أي وقت عبر الرابط التالي:\n{url}\n\n"
        f"مع أطيب التحيات،\n{COMPANY_NAME_AR}\n"
    )
    if rec["do_issued"]:
        wa = (f"{COMPANY_NAME_EN}: The Delivery Order for B/L {bl} ({vessel}) has been issued. Track the shipment: {url}\n\n"
              f"{COMPANY_NAME_AR}: تم إصدار أمر التسليم للبوليصة {bl} ({vessel}). لمتابعة الشحنة: {url}")
    else:
        wa = (f"{COMPANY_NAME_EN}: Track the status of B/L {bl} ({vessel}) here: {url}\n\n"
              f"{COMPANY_NAME_AR}: لمتابعة حالة البوليصة {bl} ({vessel}): {url}")
    return subject, body, wa


@app.route("/api/records/<path:bl_number>/notify/email", methods=["POST"])
@login_required
def notify_email(bl_number):
    """Emails the attached Delivery Order (plus the tracking link) to the
    BL's consignee and/or broker. Only ever runs on a staff click - nothing
    is sent automatically."""
    bl_number = bl_number.upper()
    if not _owns_record(bl_number):
        return jsonify({"error": "Not your record.", "error_code": "not_owner"}), 403
    data = request.get_json(force=True) or {}
    who = [w for w in (data.get("to") or []) if w in ("consignee", "broker")]
    db = get_db()
    rec = db.execute("SELECT * FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
    if not rec:
        return jsonify({"error": "BL not found."}), 404
    addrs = []
    for w in who:
        e = _clean_email(rec[f"{w}_email"])
        if e and e.lower() not in (a.lower() for a in addrs):
            addrs.append(e)
    if not addrs:
        return jsonify({"error": "Choose at least one recipient with an email address.", "error_code": "no_recipient"}), 400
    att = db.execute(
        "SELECT filename, data FROM record_attachments WHERE bl_number = ? AND kind = 'do'", (bl_number,)
    ).fetchone()
    if not att:
        return jsonify({"error": "Attach the Delivery Order PDF first.", "error_code": "no_do"}), 400
    if not email_configured():
        return jsonify({"error": "Email isn't set up yet.", "error_code": "email_not_configured"}), 400
    token = _ensure_track_token(db, bl_number)
    url = f"{_public_base_url()}/t/{token}"
    subject, body, _ = _notify_texts(rec, url)
    fname = att["filename"] or f"DO_{bl_number}.pdf"
    ok, err = send_email(addrs, subject, body, attachments=[(fname, bytes(att["data"]), "application/pdf")])
    if not ok:
        code = {"unreachable": "email_unreachable", "auth": "email_auth", "recipient": "email_recipient"}.get(err, "send_failed")
        return jsonify({"error": f"The email could not be sent: {err}", "error_code": code}), 502
    _log_audit(bl_number, "notified", "email", "", ", ".join(addrs))
    db.commit()
    return jsonify({"ok": True, "sent_to": addrs})


@app.route("/api/records/<path:bl_number>/notify/whatsapp", methods=["POST"])
@login_required
def notify_whatsapp(bl_number):
    """Builds the wa.me link (WhatsApp opens with the message ready - the
    staff member presses Send from their own WhatsApp) and records it in
    the BL's history. No WhatsApp Business account needed for this."""
    bl_number = bl_number.upper()
    if not _owns_record(bl_number):
        return jsonify({"error": "Not your record.", "error_code": "not_owner"}), 403
    who = (request.get_json(force=True) or {}).get("to")
    if who not in ("consignee", "broker"):
        return jsonify({"error": "Invalid recipient."}), 400
    db = get_db()
    rec = db.execute("SELECT * FROM records WHERE bl_number = ?", (bl_number,)).fetchone()
    if not rec:
        return jsonify({"error": "BL not found."}), 404
    phone = _normalize_phone(rec[f"{who}_phone"])
    if not phone:
        return jsonify({"error": "No phone number saved for that contact.", "error_code": "no_phone"}), 400
    token = _ensure_track_token(db, bl_number)
    _, _, text = _notify_texts(rec, f"{_public_base_url()}/t/{token}")
    _log_audit(bl_number, "notified", "whatsapp", "", phone)
    db.commit()
    return jsonify({"ok": True, "url": f"https://wa.me/{phone.lstrip('+')}?text={url_quote(text, safe='')}"})


# ---------- Public tracking page (no login) ----------

TRACK_TEXT = {
    "en": {
        "title": "Shipment Status", "bl": "B/L No.", "vessel": "Vessel", "port": "Port of Discharge", "eta": "ETA",
        "steps": ["Invoice issued", "Approval received", "Delivery Order issued"],
        "status": ["Your shipment is registered and being processed.",
                   "Invoice issued - awaiting approval.",
                   "Approval received - Delivery Order in preparation.",
                   "Delivery Order issued."],
        "pending": "Pending", "updated": "Last updated", "not_found_title": "Link not active",
        "download_invoice": "Download invoice (PDF)",
        "not_found": "This tracking link is no longer active. Please contact Sea Power for an updated link.",
        "contact": "Questions about this shipment?", "lang_switch": "العربية",
        "company": COMPANY_NAME_EN,
        "months": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    },
    "ar": {
        "title": "حالة الشحنة", "bl": "رقم البوليصة", "vessel": "السفينة", "port": "ميناء التفريغ", "eta": "الوصول المتوقع",
        "steps": ["صدور الفاتورة", "استلام الموافقة", "صدور أمر التسليم"],
        "status": ["تم تسجيل شحنتكم وجارٍ العمل عليها.",
                   "تم إصدار الفاتورة - بانتظار الموافقة.",
                   "تم استلام الموافقة - جارٍ إعداد أمر التسليم.",
                   "تم إصدار أمر التسليم."],
        "pending": "قيد الانتظار", "updated": "آخر تحديث", "not_found_title": "الرابط غير فعال",
        "download_invoice": "تحميل الفاتورة (PDF)",
        "not_found": "رابط التتبع هذا لم يعد فعالًا. يرجى التواصل مع سي باور للحصول على رابط محدّث.",
        "contact": "هل لديكم استفسار حول هذه الشحنة؟", "lang_switch": "English",
        "company": COMPANY_NAME_AR,
        "months": ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"],
    },
}


def _track_date(raw, lang):
    """Stored 'YYYY-MM-DD HH:MM' (UTC) or 'YYYY-MM-DD' -> '5 Oct 2026' in
    Saudi time (UTC+3), with Arabic month names on the Arabic page."""
    if not raw:
        return ""
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            d = datetime.strptime(raw, fmt)
            if fmt == "%Y-%m-%d %H:%M":
                d = d + timedelta(hours=3)
            return f"{d.day} {TRACK_TEXT[lang]['months'][d.month - 1]} {d.year}"
        except ValueError:
            continue
    return ""


@app.route("/t/<token>")
def public_tracking(token):
    """The page a consignee/broker opens from a shared link or QR code.
    No login. Shows status - BL, vessel, port, ETA and the three steps with
    their dates - plus a download for the invoice PDF once one is attached.
    Never remarks, staff names, contacts or the Delivery Order."""
    accept = request.headers.get("Accept-Language", "")
    lang = request.args.get("lang") or ("ar" if accept.lower().startswith("ar") else "en")
    lang = "ar" if lang == "ar" else "en"
    tx = TRACK_TEXT[lang]
    rec = None
    if re.fullmatch(r"[A-Za-z0-9_-]{16,64}", token or ""):
        rec = get_db().execute(
            """SELECT bl_number, vessel, port, eta, invoice_issued, invoice_at, approval_received, approval_at,
                      do_issued, do_at, created_at FROM records WHERE track_token = ?""",
            (token,),
        ).fetchone()
    ctx = {"lang": lang, "tx": tx, "logo": LOGO_B64, "found": bool(rec),
           "contact_phone": os.environ.get("COMPANY_PHONE", ""), "contact_email": os.environ.get("COMPANY_EMAIL", "")}
    if rec:
        done = [bool(rec["invoice_issued"]), bool(rec["approval_received"]), bool(rec["do_issued"])]
        dates = [_track_date(rec["invoice_at"], lang), _track_date(rec["approval_at"], lang), _track_date(rec["do_at"], lang)]
        stage = 3 if done[2] else 2 if done[1] else 1 if done[0] else 0
        stamps = [s for s in (rec["invoice_at"], rec["approval_at"], rec["do_at"], rec["created_at"]) if s]
        port = rec["port"] or ""
        ctx.update({
            "bl": rec["bl_number"], "vessel": rec["vessel"] or "-",
            "port": (PORT_NAMES_AR.get(port.upper(), port) if lang == "ar" else port.title()) or "-",
            "eta": _track_date(rec["eta"], lang),
            "steps": [{"label": tx["steps"][i], "done": done[i], "date": dates[i]} for i in range(3)],
            "stage": stage, "status": tx["status"][stage],
            "updated": _track_date(max(stamps), lang) if stamps else "",
            "invoice_url": f"/t/{token}/invoice" if _track_invoice(token) else "",
        })
    html = render_template_string(TRACK_HTML, **ctx)
    resp = Response(html, status=200 if rec else 404, mimetype="text/html")
    # Not for search engines, never cached by shared proxies, and the
    # secret link isn't leaked to other sites via the Referer header.
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Referrer-Policy"] = "no-referrer"
    return resp


def _track_invoice(token):
    """The invoice attachment (filename + bytes) for a tracking code, or None
    if the code isn't valid or no invoice has been attached yet."""
    if not re.fullmatch(r"[A-Za-z0-9_-]{16,64}", token or ""):
        return None
    return get_db().execute(
        """SELECT a.bl_number, a.filename, a.data FROM record_attachments a
           JOIN records r ON r.bl_number = a.bl_number
           WHERE r.track_token = ? AND a.kind = 'invoice'""",
        (token,),
    ).fetchone()


@app.route("/t/<token>/invoice")
def public_tracking_invoice(token):
    """Invoice PDF download for whoever holds the tracking link / QR code -
    the same unguessable code is the key, so "New link" in the Documents
    popup cuts off downloads too. Only the invoice: the Delivery Order
    stays internal. Each download is noted in the BL's History."""
    row = _track_invoice(token)
    if not row:
        return Response("This link is no longer active.", status=404, mimetype="text/plain")
    db = get_db()
    db.execute(
        "INSERT INTO audit_log (bl_number, action, field, old_value, new_value, by_user, at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (row["bl_number"], "customer_download", "invoice", "", "", "tracking link", datetime.utcnow().strftime("%Y-%m-%d %H:%M")),
    )
    db.commit()
    resp = Response(bytes(row["data"]), mimetype="application/pdf")
    resp.headers["Content-Disposition"] = _content_disposition(row["filename"] or f"invoice_{row['bl_number']}.pdf", f"invoice_{row['bl_number']}.pdf")
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Referrer-Policy"] = "no-referrer"
    return resp


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
        ws.cell(row=ws.max_row, column=1).data_type = "s"   # text, never a formula
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
        headers={"Content-Disposition": _content_disposition(fname, "export.xlsx")},
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
  form > label, form > input, form > .pw-wrap, form > .remember-row, form > button, form > .error {
    animation: field-in .5s ease both;
  }
  form > label:nth-of-type(1) { animation-delay: .18s; }
  form > input:nth-of-type(1), form > .pw-wrap:nth-of-type(1) { animation-delay: .24s; }
  form > label:nth-of-type(2) { animation-delay: .30s; }
  form > input:nth-of-type(2), form > .pw-wrap:nth-of-type(2) { animation-delay: .36s; }
  form > .remember-row { animation-delay: .40s; }
  form > button { animation-delay: .46s; }
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

  .notice {
    background: color-mix(in srgb, var(--gold) 16%, transparent); color: var(--text); padding: 10px 12px; border-radius: 10px;
    font-size: 13px; margin-top: 16px; text-align: center; font-weight: 600; border: 1px solid color-mix(in srgb, var(--gold) 40%, transparent);
  }
  input.code { text-align: center; font-size: 24px; font-weight: 700; letter-spacing: .32em; padding: 13px 10px; font-variant-numeric: tabular-nums; }
  input.code.recovery { font-size: 18px; letter-spacing: .12em; }
  .auth-link { display: block; text-align: center; margin-top: 16px; font-size: 13px; font-weight: 600; color: var(--muted); text-decoration: none; background: none; border: 0; width: 100%; padding: 6px; cursor: pointer; }
  .auth-link:hover { color: var(--text); }
  .step { display: flex; gap: 12px; align-items: flex-start; margin-top: 18px; font-size: 13.5px; line-height: 1.5; }
  .step b.n { flex: none; width: 24px; height: 24px; border-radius: 50%; background: var(--navy); color: #fff; font-size: 12px; display: inline-flex; align-items: center; justify-content: center; margin-top: 1px; }
  .qr-wrap { margin: 14px auto 0; width: 190px; height: 190px; padding: 8px; background: #fff; border-radius: 14px; border: 1px solid var(--border); }
  .qr-wrap img { width: 100%; height: 100%; display: block; }
  .manual { margin-top: 10px; text-align: center; font-size: 12px; color: var(--muted); }
  .manual code { display: block; margin-top: 4px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 13.5px; letter-spacing: .06em; color: var(--text); user-select: all; word-break: break-all; }
  .codes-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-top: 16px; }
  .codes-grid span { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 14.5px; font-weight: 600; text-align: center; padding: 9px 4px; border-radius: 9px; background: var(--bg); border: 1px solid var(--border); letter-spacing: .05em; user-select: all; }
  .btn-row { display: flex; gap: 10px; margin-top: 14px; }
  .btn-row button { margin-top: 0; background: none; color: var(--text); border: 1px solid var(--border); font-size: 13px; padding: 10px; }
  .btn-row button:hover { background: var(--border); }
  a.btn-link { display: block; text-align: center; text-decoration: none; width: 100%; background: var(--navy); color: #fff; border-radius: 999px; padding: 13px; font-size: 14px; font-weight: 700; margin-top: 22px; }
  a.btn-link:hover { background: var(--navy-light); }
  .ok-badge { width: 54px; height: 54px; border-radius: 50%; margin: 4px auto 14px; display: flex; align-items: center; justify-content: center; background: color-mix(in srgb, #2e9e6b 16%, transparent); color: #2e9e6b; }
  .ok-badge svg { width: 28px; height: 28px; }

  /* live password hint */
  .pw-meter { margin-top: 8px; }
  .pw-meter .bar { height: 4px; border-radius: 99px; background: var(--border); overflow: hidden; }
  .pw-meter .bar i { display: block; height: 100%; width: 0; border-radius: 99px; background: var(--danger); transition: width .25s ease, background .25s ease; }
  .pw-meter .hint { font-size: 12px; margin-top: 5px; min-height: 16px; color: var(--muted); font-weight: 500; }
  .pw-meter.bad .hint { color: var(--danger); }
  .pw-meter.ok .bar i { background: var(--gold); }
  .pw-meter.ok .hint { color: var(--text); }
  .pw-meter.strong .bar i { background: #2e9e6b; }
  .pw-meter.strong .hint { color: #2e9e6b; }

  /* Keep me signed in */
  .remember-row { margin-top: 16px; }
  .remember-row label.remember {
    display: flex; align-items: center; gap: 11px; margin: 0; cursor: pointer;
    text-transform: none; letter-spacing: 0; font-size: 13.5px; font-weight: 600; color: var(--text);
    user-select: none; -webkit-user-select: none;
  }
  .remember input { position: absolute; opacity: 0; width: 0; height: 0; }
  .remember-box {
    flex: none; width: 22px; height: 22px; border-radius: 7px; display: inline-flex; align-items: center; justify-content: center;
    border: 1.5px solid var(--border); background: var(--bg); color: transparent;
    transition: background .18s ease, border-color .18s ease, color .18s ease, box-shadow .2s ease, transform .12s ease;
  }
  .remember-box svg { width: 14px; height: 14px; stroke-dasharray: 24; stroke-dashoffset: 24; transition: stroke-dashoffset .22s ease .04s; }
  .remember:hover .remember-box { border-color: var(--gold); }
  .remember input:focus-visible + .remember-box { box-shadow: 0 0 0 4px color-mix(in srgb, var(--gold) 28%, transparent); border-color: var(--gold); }
  .remember input:checked + .remember-box { background: var(--navy); border-color: var(--navy); color: var(--gold-light); }
  .remember input:checked + .remember-box svg { stroke-dashoffset: 0; }
  .remember:active .remember-box { transform: scale(.9); }
  .remember-text small { display: block; font-size: 11.5px; font-weight: 500; color: var(--muted); margin-top: 1px; }

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


PW_METER_JS = """<script>
(function () {
  var CFG = {{ pw_config|tojson }};
  function personal(ctx) {
    var out = [];
    (ctx || []).forEach(function (c) {
      c = (c || '').toLowerCase();
      if (c.indexOf('@') > -1) c = c.split('@')[0];
      var whole = c.replace(/[^a-z0-9]/g, '');
      if (whole.length >= 3) out.push(whole);
      c.split(/[^a-z0-9]+/).forEach(function (t) { if (t.length >= 4) out.push(t); });
    });
    return out;
  }
  window.pwCheck = function (pw, ctx) {
    if (!CFG.on) return {state: pw ? 'ok' : '', pct: pw ? 100 : 0, msg: ''};
    if (!pw) return {state: '', pct: 0, msg: 'At least ' + CFG.min + ' characters. Longer is stronger.'};
    var low = pw.toLowerCase(), pct = Math.min(100, Math.round(pw.length / 14 * 100));
    if (pw.length < CFG.min) return {state: 'bad', pct: Math.min(pct, 55), msg: pw.length + ' of ' + CFG.min + ' characters - keep going.'};
    if (/^[0-9]+$/.test(low)) return {state: 'bad', pct: 35, msg: 'Use letters as well as numbers.'};
    var uniq = {}; low.split('').forEach(function (ch) { uniq[ch] = 1; });
    if (Object.keys(uniq).length <= 3) return {state: 'bad', pct: 30, msg: 'Avoid repeating the same few characters.'};
    var letters = low.replace(/[^a-z]/g, ''), left = low.length - letters.length;
    if (CFG.common.indexOf(letters) > -1 && left <= 6) return {state: 'bad', pct: 40, msg: 'That password is too common. Pick something harder to guess.'};
    var flat = low.replace(/[^a-z0-9]/g, ''), hit = personal(ctx).some(function (p) { return flat.indexOf(p) > -1; });
    if (hit) return {state: 'bad', pct: 45, msg: "Don't use your name, email or username inside the password."};
    var classes = [/[a-z]/, /[A-Z]/, /[0-9]/, /[^A-Za-z0-9]/].filter(function (r) { return r.test(pw); }).length;
    if (pw.length >= 14 || (pw.length >= 12 && classes >= 3)) return {state: 'strong', pct: 100, msg: 'Strong password.'};
    return {state: 'ok', pct: 75, msg: 'Good enough. A few more characters would make it stronger.'};
  };
  window.attachPwMeter = function (input, box, getCtx) {
    function paint() {
      var r = window.pwCheck(input.value, getCtx ? getCtx() : []);
      box.className = 'pw-meter ' + r.state;
      box.querySelector('.bar i').style.width = r.pct + '%';
      box.querySelector('.hint').textContent = r.msg;
    }
    input.addEventListener('input', paint);
    if (getCtx) { (box.dataset.watch || '').split(',').forEach(function (id) { var el = id && document.getElementById(id); if (el) el.addEventListener('input', paint); }); }
    paint();
  };
})();
</script>"""

PW_METER_HTML = """<div class="pw-meter"><div class="bar"><i></i></div><div class="hint"></div></div>"""

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
    <label>Your full name</label>
    <input type="text" name="full_name" id="setupName" required autofocus autocomplete="name" placeholder="e.g. Ahmed Al-Harbi">
    <label>Company email</label>
    <input type="email" name="email" id="setupEmail" required autocomplete="email" placeholder="name@seapower.com.sa">
    <label>Choose a password</label>
    <div class="pw-wrap">
      <input type="password" name="password" id="setupPw" required autocomplete="new-password">
      """ + PW_TOGGLE_BTN + """
    </div>
    <div id="setupMeter" data-watch="setupName,setupEmail">""" + PW_METER_HTML + """</div>
    <button type="submit">Create Admin Account</button>
  </form>
</div>
""" + PW_METER_JS + """
<script>
(function () {
  var box = document.querySelector('#setupMeter .pw-meter');
  box.dataset.watch = 'setupName,setupEmail';
  attachPwMeter(document.getElementById('setupPw'), box, function () {
    return [document.getElementById('setupName').value, document.getElementById('setupEmail').value];
  });
})();
</script>
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
  {% if notice %}<div class="notice">{{ notice }}</div>{% endif %}
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post">
    <label>Email or username</label>
    <input type="text" name="username" value="{{ typed or '' }}" required {% if not typed %}autofocus{% endif %}
           placeholder="name@seapower.com.sa" autocomplete="username" autocapitalize="none" autocorrect="off" spellcheck="false">
    <label>Password</label>
    <div class="pw-wrap">
      <input type="password" name="password" required {% if typed %}autofocus{% endif %} autocomplete="current-password">
      """ + PW_TOGGLE_BTN + """
    </div>
    <div class="remember-row">
      <label class="remember">
        <input type="checkbox" name="remember" value="1">
        <span class="remember-box" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></span>
        <span class="remember-text">Keep me signed in</span>
      </label>
    </div>
    <button type="submit">Sign In</button>
  </form>
</div>
</body></html>
"""

VERIFY_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Verify</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">""" + AUTH_STYLE + """</head><body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
""" + THEME_TOGGLE_SNIPPET + """
<div class="box">
  <div class="brand-mark">
    <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
    <span class="co">Compass</span>
    <span class="tag">Sea Power Marine Services Co. Ltd</span>
  </div>
  <h1>Hello, {{ name }}</h1>
  <div class="sub" id="vsub">Enter the 6-digit code from your authenticator app.</div>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="post" id="vform">
    <label id="vlabel">Verification code</label>
    <input class="code" type="text" name="code" id="vcode" required autofocus inputmode="numeric" autocomplete="one-time-code"
           autocapitalize="none" autocorrect="off" spellcheck="false" maxlength="12" placeholder="000000">
    <button type="submit">Verify</button>
  </form>
  <button type="button" class="auth-link" id="vtoggle" style="margin-top:12px;">Lost your phone? Use a recovery code</button>
  <a class="auth-link" href="/login" style="margin-top:0;">&larr; Back to sign in</a>
</div>
<script>
(function () {
  var inp = document.getElementById('vcode'), form = document.getElementById('vform'), tog = document.getElementById('vtoggle'), recovery = false;
  inp.addEventListener('input', function () {
    if (!recovery && /^\\d{6}$/.test(inp.value.replace(/\\s/g, ''))) { form.requestSubmit(); }
  });
  tog.addEventListener('click', function () {
    recovery = !recovery;
    inp.value = ''; inp.classList.toggle('recovery', recovery);
    inp.placeholder = recovery ? 'xxxx-xxxx' : '000000';
    inp.inputMode = recovery ? 'text' : 'numeric'; inp.maxLength = recovery ? 12 : 12;
    document.getElementById('vlabel').textContent = recovery ? 'Recovery code' : 'Verification code';
    document.getElementById('vsub').textContent = recovery ? 'Enter one of the recovery codes you saved. Each one works once.' : 'Enter the 6-digit code from your authenticator app.';
    tog.textContent = recovery ? 'Use my authenticator app instead' : 'Lost your phone? Use a recovery code';
    inp.focus();
  });
})();
</script>
</body></html>
"""

SECURITY_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Compass - Two-step verification</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">""" + AUTH_STYLE + """
<style>body { overflow-y: auto; align-items: flex-start; } .box { max-width: 420px; margin: 24px auto; }</style></head><body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
""" + THEME_TOGGLE_SNIPPET + """
<div class="box">
  <div class="brand-mark">
    <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
    <span class="co">Compass</span>
    <span class="tag">Sea Power Marine Services Co. Ltd</span>
  </div>
  {% if mode == 'setup' %}
    <h1>Secure your account</h1>
    <div class="sub">Hello {{ name }} - one quick step. Compass now asks for a code from your phone each time you sign in.</div>
    {% if error %}<div class="error">{{ error }}</div>{% endif %}
    <div class="step"><b class="n">1</b><div>Open <b>Google Authenticator</b> or <b>Microsoft Authenticator</b> on your phone, tap <b>+</b>, and scan this code.</div></div>
    <div class="qr-wrap"><img src="{{ qr }}" alt="QR code"></div>
    <div class="manual">Can't scan? Enter this key by hand:<code>{{ secret }}</code></div>
    <div class="step"><b class="n">2</b><div>Type the 6-digit code the app shows to finish.</div></div>
    <form method="post" id="sform">
      <input class="code" type="text" name="code" id="scode" required autofocus inputmode="numeric" autocomplete="one-time-code" maxlength="7" placeholder="000000" style="margin-top:12px;">
      <button type="submit">Turn on two-step verification</button>
    </form>
    <a class="auth-link" href="/logout">Sign out</a>
    <script>
      var sc = document.getElementById('scode');
      sc.addEventListener('input', function () { if (/^\\d{6}$/.test(sc.value.replace(/\\s/g, ''))) document.getElementById('sform').requestSubmit(); });
    </script>
  {% elif mode == 'codes' %}
    <h1>Save your recovery codes</h1>
    <div class="sub">Two-step verification is on. If you ever lose your phone, each of these codes lets you in once. Keep them somewhere safe - they are shown only now.</div>
    <div class="codes-grid" id="codesGrid">{% for c in codes %}<span>{{ c }}</span>{% endfor %}</div>
    <div class="btn-row">
      <button type="button" onclick="copyCodes(this)">Copy</button>
      <button type="button" onclick="downloadCodes()">Download</button>
    </div>
    <a class="btn-link" href="/">I've saved them - continue</a>
    <script>
      function codesText() { return 'Compass recovery codes\\n' + [].map.call(document.querySelectorAll('#codesGrid span'), function (x) { return x.textContent; }).join('\\n') + '\\n'; }
      function copyCodes(b) { try { navigator.clipboard.writeText(codesText()); } catch (e) {} b.textContent = 'Copied'; setTimeout(function () { b.textContent = 'Copy'; }, 1600); }
      function downloadCodes() { var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([codesText()], {type: 'text/plain'})); a.download = 'compass-recovery-codes.txt'; a.click(); }
    </script>
  {% else %}
    <div class="ok-badge"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>
    <h1>Two-step verification is on</h1>
    <div class="sub">{{ name }}{% if email %} &middot; {{ email }}{% endif %}<br>Your account asks for an authenticator code at every sign-in. Lost your phone? Ask an admin to reset it.</div>
    <a class="btn-link" href="/">Back to Compass</a>
    <a class="auth-link" href="/logout">Sign out</a>
  {% endif %}
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
  .pw-meter { margin-top: 8px; }
  .pw-meter .bar { height: 4px; border-radius: 99px; background: var(--border); overflow: hidden; }
  .pw-meter .bar i { display: block; height: 100%; width: 0; border-radius: 99px; background: var(--danger); transition: width .25s ease, background .25s ease; }
  .pw-meter .hint { font-size: 12px; margin-top: 5px; min-height: 16px; color: var(--muted); font-weight: 500; }
  .pw-meter.bad .hint { color: var(--danger); }
  .pw-meter.ok .bar i { background: var(--gold); }
  .pw-meter.ok .hint { color: var(--text); }
  .pw-meter.strong .bar i { background: #2e9e6b; }
  .pw-meter.strong .hint { color: #2e9e6b; }
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="card">
    <div class="card-label">Add a user</div>
    <div class="row">
      <input type="text" id="newName" placeholder="Full name" autocomplete="off" style="flex:1 1 170px;">
      <input type="text" id="newEmail" placeholder="Company email (name@seapower.com.sa)" autocomplete="off" autocapitalize="none" style="flex:1.4 1 230px;">
      <input type="password" id="newPassword" placeholder="Temporary password" autocomplete="new-password" style="flex:1 1 150px;">
      <select id="newRole"><option value="staff">Staff</option><option value="admin">Admin</option></select>
      <button onclick="addUser()">Add User</button>
    </div>
    <div id="newPwMeter" data-watch="newName,newEmail">""" + PW_METER_HTML + """</div>
    <div style="font-size:12px;color:var(--muted);margin-top:10px;">They sign in with this email and set up their authenticator app the first time.</div>
  </div>
  <div class="card">
    <table>
      <thead><tr><th>Name</th><th>Email / username</th><th>Role</th><th>2-step</th><th>Created</th><th></th></tr></thead>
      <tbody>
        {% for u in users %}
        <tr id="u{{ u['id'] }}">
          <td><b>{{ u['full_name'] or u['username'] }}</b></td>
          <td>{% if u['email'] %}{{ u['email'] }}<div style="font-size:11.5px;color:var(--muted);">{{ u['username'] }}</div>{% else %}{{ u['username'] }}<div style="font-size:11.5px;color:var(--muted);">no email yet</div>{% endif %}</td>
          <td><span class="role-pill {{ u['role'] }}">{{ u['role'] }}</span></td>
          <td>{% if u['totp_enabled'] %}<span class="role-pill staff" style="background:color-mix(in srgb,#2e9e6b 16%,transparent);color:#2e9e6b;">On</span>{% else %}<span class="role-pill" style="background:var(--border);color:var(--muted);">Not yet</span>{% endif %}</td>
          <td class="local-time" data-utc="{{ u['created_at'] }}">{{ u['created_at'] }}</td>
          <td style="white-space:nowrap;"><button class="del" style="color:var(--navy-light);" onclick="toggleEdit({{ u['id'] }})">Edit</button><button class="del" onclick='delUser({{ u["id"] }}, {{ (u["full_name"] or u["username"])|tojson }})'>Remove</button></td>
        </tr>
        <tr class="edit-row" id="e{{ u['id'] }}" data-username="{{ u['username'] }}" style="display:none;">
          <td colspan="6" style="background:color-mix(in srgb,var(--border) 30%,transparent);">
            <div class="row">
              <input type="text" id="en{{ u['id'] }}" value="{{ u['full_name'] }}" placeholder="Full name" style="flex:1 1 160px;">
              <input type="text" id="ee{{ u['id'] }}" value="{{ u['email'] }}" placeholder="name@seapower.com.sa" autocapitalize="none" style="flex:1.3 1 220px;">
              <input type="password" id="ep{{ u['id'] }}" placeholder="New password (optional)" autocomplete="new-password" style="flex:1 1 150px;">
              <button onclick="saveUser({{ u['id'] }})">Save</button>
              {% if u['totp_enabled'] %}<button class="del" style="border:1px solid var(--border);" onclick="reset2fa({{ u['id'] }})">Reset 2-step</button>{% endif %}
            </div>
            <div class="pw-meter" id="epm{{ u['id'] }}" style="display:none;max-width:420px;"><div class="bar"><i></i></div><div class="hint"></div></div>
          </td>
        </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
  <div id="toastHost"></div>
""" + PW_METER_JS + """
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
  const full_name = document.getElementById('newName').value.trim();
  const email = document.getElementById('newEmail').value.trim();
  const password = document.getElementById('newPassword').value;
  const role = document.getElementById('newRole').value;
  if (!full_name || !email || !password) { showToast('Fill in the name, company email and a temporary password.', {error:true}); return; }
  const res = await fetch('/api/users', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({full_name, email, password, role})});
  const data = await res.json();
  if (!res.ok || data.error) { showToast(data.error || 'Could not add that user.', {error:true}); return; }
  showToast(full_name + ' added.');
  setTimeout(() => location.reload(), 500);
}
function toggleEdit(id) {
  const r = document.getElementById('e' + id);
  r.style.display = r.style.display === 'none' ? '' : 'none';
}
// live password hints
(function () {
  var nb = document.querySelector('#newPwMeter .pw-meter');
  nb.dataset.watch = 'newName,newEmail';
  attachPwMeter(document.getElementById('newPassword'), nb, function () {
    return [document.getElementById('newName').value, document.getElementById('newEmail').value];
  });
  document.querySelectorAll('.edit-row').forEach(function (row) {
    var id = row.id.slice(1), pw = document.getElementById('ep' + id), box = document.getElementById('epm' + id);
    if (!pw || !box) return;
    box.dataset.watch = 'en' + id + ',ee' + id;
    attachPwMeter(pw, box, function () { return [document.getElementById('en' + id).value, document.getElementById('ee' + id).value, row.dataset.username || '']; });
    box.style.display = 'none';
    pw.addEventListener('input', function () { box.style.display = pw.value ? '' : 'none'; });
  });
})();
async function saveUser(id) {
  const body = {full_name: document.getElementById('en' + id).value.trim(), email: document.getElementById('ee' + id).value.trim()};
  const pw = document.getElementById('ep' + id).value;
  if (pw) body.password = pw;
  const res = await fetch('/api/users/' + id, {method:'PATCH', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body)});
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.error) { showToast(data.error || 'Could not save.', {error:true}); return; }
  showToast('Saved.');
  setTimeout(() => location.reload(), 500);
}
async function reset2fa(id) {
  const res = await fetch('/api/users/' + id, {method:'PATCH', headers:{'Content-Type':'application/json'}, body: JSON.stringify({reset_2fa: true})});
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.error) { showToast(data.error || 'Could not reset.', {error:true}); return; }
  showToast('Two-step reset. They will set it up again at their next sign-in.');
  setTimeout(() => location.reload(), 900);
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
      <a href="/account/security">Security</a>
      <a href="/logout">Log out</a>
    </div>
  </div>

  <div class="hero">
    <div class="eyebrow">Welcome, {{ first_name }}</div>
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
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

// BL numbers, port/vessel names and usernames are user/file data - escape
// before putting them into HTML so they can't inject markup or script.
function esc(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

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
        <div class="b-bl">${esc(e.bl_number)}</div>
        <div class="b-meta">${[e.port, e.vessel].filter(Boolean).map(esc).join(' &middot; ') || 'No port/vessel set'}${IS_ADMIN ? ' &middot; ' + esc(e.created_by || 'unknown') : ''}</div>
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
              <td><b>${esc(w.agent)}</b></td>
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
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
// BL numbers, reasons and notes come from uploaded packing lists - escape
// them before putting them into HTML so they can't inject markup/script.
function esc(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function jsq(s) { return esc(JSON.stringify(String(s === undefined || s === null ? '' : s))); }

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
            <td><b>${esc(r.bl_number)}</b></td>
            <td style="color:var(--muted);">${esc(r.reason || '')}</td>
            <td style="color:var(--muted);">${esc(r.classified_by || '')}${r.classified_at ? ' - ' + esc(r.classified_at) : ''}</td>
            <td><button class="btn ghost danger" style="padding:4px 10px;font-size:11.5px;" onclick="removeDirectDelivery(${jsq(r.bl_number)})">Remove</button></td>
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
            <td><b>${esc(r.bl_number)}</b></td>
            <td>${ddBadgeHtml(!!r.is_direct)}</td>
            <td style="color:var(--muted);">${esc(r.review_note || '')}</td>
            <td><button class="btn ghost danger" style="padding:4px 10px;font-size:11.5px;" onclick="removeDirectDelivery(${jsq(r.bl_number)}, true)">Remove</button></td>
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
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
      <span style="padding:6px 4px;">Signed in as <b>{{ display_name }}</b></span>
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

# Public tracking page (/t/<token>) - what a consignee or broker sees.
# Self-contained, mobile-first (most will open it from WhatsApp), Arabic
# right-to-left when lang=ar. Jinja autoescapes every value.
TRACK_HTML = """<!DOCTYPE html>
<html lang="{{ lang }}" dir="{{ 'rtl' if lang == 'ar' else 'ltr' }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<meta name="referrer" content="no-referrer">
<title>{{ tx.title }}{% if found %} - {{ bl }}{% endif %} | Sea Power</title>
{% if lang == 'ar' %}
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
{% endif %}
<style>
  :root {
    --bg: #f3f5f8; --card: #ffffff; --text: #14212e; --muted: #64748b; --border: #e2e8f0;
    --navy: #123a56; --navy-deep: #0b2740; --gold: #c9a227; --ok: #1f9d55; --ok-bg: #e7f6ee;
  }
  @media (prefers-color-scheme: dark) {
    :root { --bg: #0d1520; --card: #152131; --text: #e6edf5; --muted: #93a3b8; --border: #26364a;
            --navy: #5aa2d1; --navy-deep: #274a67; --ok: #4cc283; --ok-bg: #12342a; }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; background: var(--bg); color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
    -webkit-text-size-adjust: 100%;
  }
  [dir="rtl"] body { font-family: 'Cairo', 'Segoe UI', Tahoma, Arial, sans-serif; }
  .wrap { max-width: 520px; margin: 0 auto; padding: 18px 16px 40px; }
  .brand { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 18px; }
  .brand-id { display: flex; align-items: center; gap: 10px; min-width: 0; }
  .brand img { height: 38px; width: auto; }
  .brand b { font-size: 14px; line-height: 1.25; }
  .lang { font-size: 13px; font-weight: 700; color: var(--navy); text-decoration: none; padding: 7px 12px;
          border: 1px solid var(--border); border-radius: 999px; background: var(--card); white-space: nowrap; }
  .card { background: var(--card); border: 1px solid var(--border); border-radius: 18px; padding: 20px;
          box-shadow: 0 1px 2px rgba(15, 30, 50, .05), 0 8px 24px rgba(15, 30, 50, .06); }
  .eyebrow { font-size: 12px; font-weight: 700; color: var(--muted); letter-spacing: .04em; text-transform: uppercase; }
  [dir="rtl"] .eyebrow { letter-spacing: 0; text-transform: none; font-size: 13px; }
  .bl { font-size: 24px; font-weight: 800; margin: 4px 0 14px; word-break: break-all; }
  .status { display: flex; gap: 10px; align-items: center; padding: 12px 14px; border-radius: 12px;
            background: color-mix(in srgb, var(--gold) 14%, transparent); font-weight: 700; font-size: 15px; }
  .status.done { background: var(--ok-bg); color: var(--ok); }
  .status .dot { width: 10px; height: 10px; border-radius: 50%; background: var(--gold); flex-shrink: 0; }
  .status.done .dot { background: var(--ok); }
  .details { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 16px; margin: 18px 0 4px; }
  .details div span { display: block; font-size: 12px; color: var(--muted); margin-bottom: 2px; }
  .details div b { font-size: 15px; }
  .details .full { grid-column: 1 / -1; }
  .steps { list-style: none; margin: 22px 0 0; padding: 0; }
  .step { position: relative; display: flex; gap: 14px; padding-bottom: 22px; }
  .step:last-child { padding-bottom: 0; }
  .step::before { content: ""; position: absolute; inset-inline-start: 13px; top: 28px; bottom: 0; width: 2px; background: var(--border); }
  .step:last-child::before { display: none; }
  .step.done::before { background: var(--ok); }
  .mark { width: 28px; height: 28px; border-radius: 50%; flex-shrink: 0; display: flex; align-items: center; justify-content: center;
          border: 2px solid var(--border); background: var(--card); color: transparent; font-size: 15px; font-weight: 800; position: relative; z-index: 1; }
  .step.done .mark { background: var(--ok); border-color: var(--ok); color: #fff; }
  .step.current .mark { border-color: var(--gold); box-shadow: 0 0 0 4px color-mix(in srgb, var(--gold) 22%, transparent); }
  .step-text b { display: block; font-size: 15px; padding-top: 3px; }
  .step-text span { font-size: 13px; color: var(--muted); }
  .updated { margin-top: 18px; font-size: 12px; color: var(--muted); }
  .dl { display: flex; align-items: center; justify-content: center; gap: 8px; margin-top: 22px;
        padding: 13px 16px; border-radius: 12px; background: var(--navy); color: #fff;
        font-weight: 700; font-size: 15px; text-decoration: none; }
  .dl svg { width: 18px; height: 18px; flex-shrink: 0; }
  .dl:active { opacity: .85; }
  .foot { margin-top: 18px; text-align: center; font-size: 13px; color: var(--muted); line-height: 1.7; }
  .foot a { color: var(--navy); font-weight: 700; text-decoration: none; }
  .nf h1 { font-size: 20px; margin: 0 0 8px; }
  .nf p { margin: 0; color: var(--muted); line-height: 1.6; }
</style>
</head>
<body>
<div class="wrap">
  <div class="brand">
    <div class="brand-id">
      <img src="data:image/png;base64,{{ logo }}" alt="Sea Power">
      <b>{{ tx.company }}</b>
    </div>
    <a class="lang" href="?lang={{ 'en' if lang == 'ar' else 'ar' }}">{{ tx.lang_switch }}</a>
  </div>

  {% if found %}
  <div class="card">
    <div class="eyebrow">{{ tx.title }}</div>
    <div class="bl"><span dir="ltr">{{ bl }}</span></div>
    <div class="status{% if stage == 3 %} done{% endif %}"><span class="dot"></span>{{ status }}</div>

    <div class="details">
      <div><span>{{ tx.vessel }}</span><b dir="auto">{{ vessel }}</b></div>
      <div><span>{{ tx.eta }}</span><b>{{ eta or '-' }}</b></div>
      <div class="full"><span>{{ tx.port }}</span><b>{{ port }}</b></div>
    </div>

    <ol class="steps">
      {% for s in steps %}
      <li class="step{% if s.done %} done{% elif loop.index0 == stage %} current{% endif %}">
        <span class="mark">&#10003;</span>
        <div class="step-text"><b>{{ s.label }}</b><span>{{ s.date if s.done else tx.pending }}</span></div>
      </li>
      {% endfor %}
    </ol>
    {% if invoice_url %}
    <a class="dl" href="{{ invoice_url }}" rel="nofollow">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 4v12M12 16l-5-5M12 16l5-5"/><path d="M4 20h16"/></svg>
      {{ tx.download_invoice }}
    </a>
    {% endif %}
    {% if updated %}<div class="updated">{{ tx.updated }}: {{ updated }}</div>{% endif %}
  </div>
  {% else %}
  <div class="card nf">
    <h1>{{ tx.not_found_title }}</h1>
    <p>{{ tx.not_found }}</p>
  </div>
  {% endif %}

  <div class="foot">
    {{ tx.contact }}<br>
    {% if contact_phone %}<a href="tel:{{ contact_phone }}" dir="ltr">{{ contact_phone }}</a>{% endif %}
    {% if contact_phone and contact_email %} &middot; {% endif %}
    {% if contact_email %}<a href="mailto:{{ contact_email }}">{{ contact_email }}</a>{% endif %}
    {% if not contact_phone and not contact_email %}{{ tx.company }}{% endif %}
  </div>
</div>
</body>
</html>
"""


PAGE_HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Compass - DO Tracker</title><link rel="icon" type="image/png" href="data:image/png;base64,""" + LOGO_B64 + """">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap" rel="stylesheet">
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
  /* Wraps as a whole group (never word-by-word inside a link) - on a phone,
     "Signed in as admin" / "Log out" used to break into stacks of 1-2
     words, and in Arabic (longer labels) Log out was pushed off-screen. */
  .topbar-right {
    display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted);
    flex-wrap: wrap; justify-content: flex-end; row-gap: 6px; min-width: 0;
  }
  .topbar-right a, .who { white-space: nowrap; }
  .topbar-right a {
    color: var(--navy); text-decoration: none; font-weight: 600; font-size: 13px;
    padding: 6px 12px; border-radius: 20px; transition: background .15s ease;
  }
  :root[data-theme="dark"] .topbar-right a { color: var(--navy-light); }
  .topbar-right a:hover { background: var(--border); }
  .who { padding: 6px 4px; }
  .who b { color: var(--text); }

  /* Sun/moon theme switch */
  /* direction:ltr - a self-contained control with no reading direction.
     Left to mirror in Arabic, its sun/moon icons swapped sides while the
     knob didn't, so it showed the wrong icon in both themes. */
  .theme-switch { position: relative; display: inline-flex; width: 54px; height: 29px; cursor: pointer; flex-shrink: 0; direction: ltr; }
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
    font-family: inherit; /* browsers default buttons to Arial - every button
                             was in a different font from the page (and in
                             Arabic, not Cairo) */
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
    text-align: start; font-family: inherit; cursor: pointer;
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
    font-size: 11px; font-weight: 700; padding: 5px 11px;
    border-radius: 999px; flex-shrink: 0; text-decoration: none;
  }
  /* Export / Archive / Remove all always travel together as one block at
     the far end of the header (inline-start auto margin = right in
     English, left in Arabic). Previously each button wrapped on its own,
     so on a phone they landed in different spots depending on how long
     the (translated) text before them was. */
  .group-actions { display: flex; align-items: center; gap: 8px; margin-inline-start: auto; flex-shrink: 0; }
  /* Port/vessel name boxes size to their content - the browser default
     (~20 characters) clipped longer names like "YANBU COMMERCIAL PORT". */
  .group-name { field-sizing: content; min-width: 8ch; max-width: 100%; }
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
  th, td { text-align: start; padding: 12px 10px; border-bottom: 1px solid var(--border); overflow: hidden; }
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
    padding: 3px 10px; border-radius: 999px; margin-inline-end: 2px;
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
    background: none; color: var(--danger); margin-inline-start: auto;
    border: 1px solid color-mix(in srgb, var(--danger) 45%, var(--border));
  }
  .bulk-bar button.bulk-remove-btn:hover { background: var(--danger-bg); }
  /* Phones: the selection bar is pinned to the bottom of the screen (like
     Gmail/WhatsApp selection mode) instead of sitting above the BL list.
     Above the list it appeared off-screen when you ticked a BL far down a
     long vessel - you couldn't see it, and had to scroll all the way up to
     use it and then find your place again. Only one bar is pinned at a
     time: the vessel you're working in. */
  @media (max-width: 600px) {
    .bulk-bar.active.docked {
      position: fixed; inset-inline: 10px; bottom: 10px; z-index: 950; margin: 0;
      background: color-mix(in srgb, var(--gold) 12%, var(--card));
      box-shadow: var(--shadow-md); max-height: 45vh; overflow-y: auto;
    }
    body:has(.bulk-bar.active.docked) .scroll-top-btn { display: none; }
    body:has(.bulk-bar.active.docked) #toastHost { bottom: calc(var(--dock-h, 0px) + 20px); }
    body:has(.bulk-bar.active.docked)::after { height: calc(var(--dock-h, 0px) + 70px); }
  }

  /* Above the floating back-to-top button (900) so the dimmed backdrop
     really covers everything behind the popup; toasts (1200) stay above it
     so "Invoice uploaded" etc. are still visible while the popup is open. */
  .history-overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 1100;
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
  /* Title takes the free space so the × always sits in the far corner
     (right in English, left in Arabic) instead of hugging the title. */
  .history-modal-head b { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
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

  /* Customer sharing sections in the Documents popup */
  #docsOverlay .history-modal { max-width: 560px; max-height: 86vh; }
  .docs-section .docs-actions button { padding: 7px 14px; font-size: 12.5px; }
  .docs-section .docs-actions button:disabled { opacity: .45; cursor: not-allowed; }
  .share-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px 12px; margin-bottom: 10px; }
  .share-full { grid-column: 1 / -1; }
  .share-field { display: flex; flex-direction: column; gap: 4px; min-width: 0; }
  .share-field span { font-size: 11.5px; font-weight: 600; color: var(--muted); }
  .share-field b { font-size: 13px; font-weight: 600; overflow-wrap: anywhere; }
  /* All contact boxes styled alike - type=email/tel otherwise fall back to
     the browser's default look while type=text gets the app's. */
  .share-field input {
    width: 100%; padding: 8px 10px; font-size: 13px; font-family: inherit; color: var(--text);
    background: var(--bg); border: 1px solid var(--border); border-radius: 10px;
  }
  .share-field input:focus { outline: none; border-color: var(--navy-light); }
  .share-link-row { margin-bottom: 8px; }
  .share-url { font-size: 12.5px; padding: 8px 10px; color: var(--muted); }
  .qr-panel { margin-top: 10px; display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
  .qr-panel img { width: 148px; height: 148px; background: #fff; border-radius: 10px; border: 1px solid var(--border); padding: 4px; }
  .qr-panel a { font-size: 12.5px; font-weight: 600; color: var(--navy-light); }
  .share-checks { display: flex; gap: 8px 16px; flex-wrap: wrap; margin-bottom: 10px; }
  .share-check { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; cursor: pointer; }
  .share-check small { color: var(--muted); font-size: 11.5px; }
  .share-check.disabled { opacity: .45; cursor: not-allowed; }
  .share-hint { margin-top: 6px; }
  .wa-row { margin-top: 12px; }
  .docs-section .docs-actions button.btn-whatsapp { background: #1f9d55; }
  .docs-section .docs-actions button.btn-whatsapp:hover { background: #17834a; }
  @media (max-width: 600px) { .share-grid { grid-template-columns: 1fr; } }

  /* "Attach documents" auto-match batch rows - one per dropped file, while
     it's being read/matched, once it's auto-attached, or (when the BL
     couldn't be pinned down automatically) while it waits for the user to
     pick the right one by hand. */
  .automatch-modal { max-width: 640px; }
  .automatch-summary { font-size: 12.5px; color: var(--muted); padding: 4px 0 6px; }
  .match-row {
    display: flex; align-items: center; gap: 10px; padding: 9px 0;
    border-bottom: 1px solid var(--border); font-size: 12.5px;
  }
  .match-row:last-child { border-bottom: none; }
  .match-row .match-file { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .match-row .match-status { font-size: 11.5px; color: var(--muted); }
  .match-row.ok .match-status { color: var(--success, #1f9d55); }
  .match-row.review { flex-wrap: wrap; }
  /* Controls always get their own line under the file name + status (the
     longer Arabic status text used to push them to a different spot than
     in English), aligned to the start of the line. */
  .match-row.review .match-review-controls { display: flex; gap: 8px; align-items: center; flex: 0 0 100%; min-width: 0; margin-top: 6px; }
  /* Were plain unstyled browser dropdowns - now match the app's inputs. */
  .match-row.review .match-review-controls select,
  .match-row.review .match-review-controls input {
    flex: 0 1 260px; min-width: 0; font-family: inherit; font-size: 13px; color: var(--text);
    background: var(--bg); border: 1px solid var(--border); border-radius: 10px; padding: 7px 10px;
  }
  .match-row.review .match-review-controls select:focus { outline: none; border-color: var(--navy-light); }
  .match-row.review .match-review-controls button { padding: 7px 16px; flex-shrink: 0; }

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
    /* Actions get their own full row: Export + Archive at the start,
       Remove all alone at the far end - identical in both languages. */
    .group-actions { flex-basis: 100%; margin-inline-start: 0; }
    .group-actions .group-remove { margin-inline-start: auto; }
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
  #toastHost { position: fixed; bottom: 20px; inset-inline-end: 20px; display: flex; flex-direction: column; gap: 8px; z-index: 1200; pointer-events: none; max-width: min(320px, calc(100vw - 40px)); }
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
    position: fixed; bottom: 86px; inset-inline-end: 24px; z-index: 900;
    width: 46px; height: 46px; border-radius: 50%; padding: 0;
    background: var(--navy-deep); color: #fff; border: none;
    display: flex; align-items: center; justify-content: center;
    box-shadow: var(--shadow-md); cursor: pointer;
    /* Hidden until the manifest form has scrolled out of view - it was
       showing at the very top of the page too, sitting on top of
       "Collapse all" / the archived count with nothing to scroll back to. */
    opacity: 0; visibility: hidden; transform: translateY(8px);
    transition: opacity .2s ease, transform .2s ease, visibility .2s;
  }
  .scroll-top-btn.show { opacity: 1; visibility: visible; transform: none; }
  .scroll-top-btn:hover { background: var(--navy-light); }
  .scroll-top-btn svg { width: 20px; height: 20px; }
  /* Room under the last card so the floating button never covers it. */
  body::after { content: ""; display: block; height: 70px; }

  /* Search box + "/" shortcut hint. The hint only shows while the box is
     empty and not focused, and never on touch devices (no keyboard). */
  .search-wrap { position: relative; flex: 1; min-width: 180px; }
  .search-wrap #searchBox { width: 100%; padding-inline-end: 34px; }
  .kbd-hint {
    position: absolute; inset-inline-end: 9px; top: 50%; transform: translateY(-50%);
    font: 600 11px/1 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; color: var(--muted);
    border: 1px solid var(--border); border-bottom-width: 2px; border-radius: 5px;
    padding: 3px 6px; background: var(--card); pointer-events: none;
  }
  #searchBox:focus + .kbd-hint, #searchBox:not(:placeholder-shown) + .kbd-hint { display: none; }
  @media (hover: none) { .kbd-hint { display: none; } }
  /* Brief highlight on the row the search jumped to. */
  @keyframes row-flash { 0%, 40% { background: color-mix(in srgb, var(--gold) 30%, transparent); } 100% { background: transparent; } }
  tr.row-flash > td { animation: row-flash 1.8s ease-out; }

  /* Phone layout */
  @media (max-width: 600px) {
    /* Stats: three compact tiles in one row instead of three half-width
       cards stacked ragged down the page. */
    .summary { gap: 8px; flex-wrap: nowrap; }
    .stat { flex: 1 1 0; min-width: 0; max-width: none; padding: 10px 12px; display: block; }
    .stat-icon { display: none; }
    .stat b { font-size: 19px; }
    /* Toolbar: search + Collapse all share row 1, then each dropdown gets a
       full-width row (they were three different widths across three rows;
       side by side they're too narrow and cut off "Select a vessel to view"). */
    .toolbar-row .search-wrap { flex: 1 1 calc(100% - 130px); min-width: 0; }
    .toolbar-row .btn-neutral { flex: 0 0 auto; }
    .toolbar-row .glass-select-wrap { flex: 1 1 100%; min-width: 0 !important; order: 2; }
    /* Top bar: drop the "Signed in as" words (the username stays), and let
       it scroll away - pinned, its 2-3 rows took ~15% of a phone screen
       permanently, away from the BL list. */
    .who-label { display: none; }
    .topbar { position: static; }
    .topbar-right { gap: 6px; }
    .topbar-right a { padding: 6px 8px; }
  }

  /* English/Arabic switch - plain two-option pill, same spot as the
     theme toggle. The active language is bold/highlighted; clicking the
     other one switches (see setLang() in the shared I18N script). */
  .lang-switch {
    display: flex; align-items: center; gap: 2px; padding: 2px;
    border: 1px solid var(--border); border-radius: 999px; background: var(--card);
  }
  .lang-opt {
    background: none; color: var(--muted); font-size: 11.5px; font-weight: 700;
    padding: 5px 10px; border-radius: 999px; line-height: 1;
  }
  .lang-opt.active { background: var(--navy); color: #fff; }
  .lang-opt:hover:not(.active) { background: var(--border); color: var(--text); }

  /* ---------- RTL (Arabic) overrides ----------
     Setting dir="rtl" on <html> flips text direction and the order of
     flex-row children. Spacing/positions in this page use logical
     properties (margin-inline-start, inset-inline-end, text-align:start)
     that flip with the language on their own, so English and Arabic can't
     drift apart - don't add plain margin-left/right etc. for layout. Only
     things with no logical equivalent are overridden here. */
  [dir="rtl"] body { font-family: 'Cairo', 'Segoe UI', Tahoma, Arial, sans-serif; }
  [dir="rtl"] .cs-option { text-align: start; }
  [dir="rtl"] .port-header.collapsed .chev,
  [dir="rtl"] .vessel-header.collapsed .chev { transform: rotate(90deg); }
  /* Arabic is a joined-up script: letter-spacing pulls the letters apart
     and UPPERCASE means nothing - and the tiny 10-11px label sizes chosen
     for English capitals are too small to read in Arabic. */
  [dir="rtl"] .card-label, [dir="rtl"] .tag-fields label, [dir="rtl"] th,
  [dir="rtl"] .eta-label, [dir="rtl"] .doc-chip, [dir="rtl"] .badge-complete,
  [dir="rtl"] tbody tr td[data-label]::before { letter-spacing: 0; text-transform: none; }
  [dir="rtl"] .card-label { font-size: 14px; }
  [dir="rtl"] .tag-fields label { font-size: 12.5px; }
  [dir="rtl"] th { font-size: 12px; }
  [dir="rtl"] .eta-label, [dir="rtl"] .eta-unset-hint, [dir="rtl"] .meta { font-size: 11.5px; }
  [dir="rtl"] tbody tr td[data-label]::before { font-size: 12px; }
</style>
</head>
<body>
  <div class="topbar">
    <a href="/" class="brand" style="text-decoration:none;">
      <img src="data:image/png;base64,""" + LOGO_B64 + """" alt="Sea Power">
      <div class="brand-text">
        <span class="app-name">Compass</span>
        <span class="app-tag" data-i18n="app_tagline">DO Tracker</span>
      </div>
    </a>
    <div class="topbar-right">
      <div class="lang-switch" title="EN / &#1593;&#1585;&#1576;&#1610;">
        <button type="button" class="lang-opt active" data-lang-btn="en" onclick="setLang('en')">EN</button>
        <button type="button" class="lang-opt" data-lang-btn="ar" onclick="setLang('ar')">&#1593;&#1585;&#1576;&#1610;</button>
      </div>
      <label class="theme-switch" title="Toggle dark mode" data-i18n-title="toggle_dark_mode">
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
      {% if role == 'admin' %}<a href="/users" data-i18n="manage_users">Manage Users</a>{% endif %}
      <span class="who"><span class="who-label" data-i18n="signed_in_as">Signed in as</span> <b>{{ display_name }}</b></span>
      <a href="/logout" data-i18n="log_out">Log out</a>
    </div>
  </div>

  <div class="card" id="manifestCard">
    <div class="card-label" data-i18n="add_a_manifest">Add a manifest</div>
    <div class="tag-fields">
      <div>
        <label for="portField" data-i18n="discharge_port">Discharge Port</label>
        <div class="glass-select-wrap">
          <select id="portField" class="nice-select">
            <option value="" data-i18n="select_a_port">Select a port...</option>
            <option value="DAMMAM PORT" data-i18n="port_dammam">Dammam Port</option>
            <option value="JUBAIL COMMERCIAL PORT" data-i18n="port_jubail">Jubail Commercial Port</option>
            <option value="JEDDAH PORT" data-i18n="port_jeddah">Jeddah Port</option>
            <option value="YANBU COMMERCIAL PORT" data-i18n="port_yanbu_commercial">Yanbu Commercial Port</option>
            <option value="YANBU INDUSTRIAL PORT" data-i18n="port_yanbu_industrial">Yanbu Industrial Port</option>
            <option value="KAP" data-i18n="port_kap">KAP</option>
          </select>
        </div>
      </div>
      <div>
        <label for="vesselField" data-i18n="vessel_label">Vessel</label>
        <input type="text" id="vesselField" placeholder="e.g. TAI KNIGHT" data-i18n-ph="vessel_placeholder" style="text-transform:uppercase;" oninput="this.value = this.value.toUpperCase();"
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
        <div class="dropzone-text"><b data-i18n="click_to_upload">Click to upload</b> <span data-i18n="or_drag_drop_manifest">or drag &amp; drop your manifest</span></div>
        <div class="dropzone-sub" data-i18n="manifest_dropzone_sub">.xlsx, .xls, .csv, .docx or .pdf - the BL numbers are read automatically</div>
        <div class="dropzone-filename" id="dropzoneFilename"></div>
      </div>
      <input type="file" id="manifestFile" accept=".xlsx,.xlsm,.xls,.csv,.docx,.pdf" multiple style="display:none" onchange="stageManifestFile()">
    </label>
    <div class="row" style="margin-top:14px;">
      <button type="button" id="addManifestBtn" onclick="uploadExcel()" disabled data-i18n="add_to_board">Add to board</button>
    </div>
  </div>

  <div class="card">
    <div class="card-label" data-i18n="attach_documents">Attach documents</div>
    <div style="font-size:12.5px; color:var(--muted); margin-bottom:12px;">
      <span data-i18n="attach_docs_help">Drop Invoice / Delivery Order PDFs here - each one is read and matched to its BL automatically, same as the manifest upload above.</span>
      <span class="doc-chip has-file" style="cursor:default;" data-i18n="doc_chip_inv">INV</span> / <span class="doc-chip has-file" style="cursor:default;" data-i18n="doc_chip_do">DO</span>
      <span data-i18n="attach_docs_help2">chips next to a BL number below show what's already attached.</span>
    </div>
    <label class="dropzone" id="autoMatchDropzone" for="autoMatchFile">
      <div class="dropzone-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="1.8">
          <path d="M12 16V4M12 4l-4 4M12 4l4 4"/><path d="M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3"/>
        </svg>
      </div>
      <div>
        <div class="dropzone-text"><b data-i18n="click_to_upload">Click to upload</b> <span data-i18n="or_drag_drop_docs">or drag &amp; drop Invoice/DO PDFs</span></div>
        <div class="dropzone-sub" data-i18n="auto_match_dropzone_sub">Drop as many at once as you like - each is matched to its BL automatically</div>
      </div>
      <input type="file" id="autoMatchFile" accept=".pdf" multiple style="display:none" onchange="handleAutoMatchFiles(this.files)">
    </label>
  </div>

  <div class="summary" id="summary"></div>

  <div class="card">
    <div class="row toolbar-row" style="margin-bottom:14px; flex-wrap:wrap;">
      <div class="search-wrap" title="Shortcut: press / from anywhere to search. Enter jumps to the BL, Esc clears." data-i18n-title="search_shortcut_title">
        <input type="text" id="searchBox" placeholder="Search BL number..." data-i18n-ph="search_bl_placeholder" oninput="render()" onkeydown="onSearchKeydown(event)" autocomplete="off">
        <kbd class="kbd-hint" aria-hidden="true">/</kbd>
      </div>
      <div class="glass-select-wrap" style="width:auto; min-width:200px;">
        <select id="jumpSelect" class="nice-select" onchange="jumpToVessel(this.value)"><option value="" data-i18n="select_vessel_to_view">Select a vessel to view</option></select>
      </div>
      {% if role == 'admin' %}
      <div class="glass-select-wrap" style="width:auto; min-width:140px;">
        <select id="operatorFilter" class="nice-select" onchange="render()"><option value="" data-i18n="all_operators">All operators</option></select>
      </div>
      {% endif %}
      <button type="button" class="btn-neutral" data-i18n="collapse_all" onclick="setAllGroupsCollapsed(true)">Collapse all</button>
    </div>
    <div id="portTabs"></div>
    <div id="groups"></div>
  </div>

  <div class="card" id="archivedCard" style="display:none;">
    <div class="row" style="margin-bottom:14px; cursor:pointer;" onclick="archivedSectionOpen = !archivedSectionOpen; render();">
      <b style="flex:1;" data-i18n="archived_vessels">Archived vessels</b>
      <span class="group-count" id="archivedCount"></span>
    </div>
    <div id="archivedGroups"></div>
  </div>

  <div id="toastHost"></div>
  <button type="button" id="scrollTopBtn" class="scroll-top-btn" title="Back to Discharge Port / Vessel" data-i18n-title="scroll_top_title" onclick="scrollToManifestForm()">
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

  <!-- Results of dropping Invoice/DO PDFs. Shown in a popup (closed with x,
       a click outside, or Esc) instead of stacking up on the front page. -->
  <div id="autoMatchOverlay" class="history-overlay" style="display:none;" onclick="if(event.target===this) closeAutoMatch()">
    <div class="history-modal automatch-modal">
      <div class="history-modal-head">
        <b data-i18n="attach_results_title">Attaching documents</b>
        <button type="button" onclick="closeAutoMatch()" style="background:none; color:var(--text); padding:4px 10px;" aria-label="Close">&times;</button>
      </div>
      <div class="history-modal-body">
        <div id="autoMatchSummary" class="automatch-summary"></div>
        <div id="autoMatchList"></div>
      </div>
    </div>
  </div>

  <div id="docsOverlay" class="history-overlay" style="display:none;" onclick="if(event.target===this) closeDocs()">
    <div class="history-modal">
      <div class="history-modal-head">
        <b id="docsTitle" data-i18n="documents_title">Documents</b>
        <button type="button" onclick="closeDocs()" style="background:none; color:var(--text); padding:4px 10px;">&times;</button>
      </div>
      <div id="docsBody" class="history-modal-body"></div>
    </div>
  </div>

<script>
""" + I18N_JS + """
/* ---------- HTML escaping ----------
   Everything on this board that came from a person or a file (BL numbers
   from a carrier's manifest, remarks, vessel/port names, uploaded file
   names, usernames) MUST go through esc() before being put into HTML, or
   jsq() when it's an argument inside an inline onclick="...". Without
   this, text like <img onerror=...> typed into a remark ran as code in
   whoever viewed it - including an admin opening History. */
function esc(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
// A JS string literal, HTML-escaped for use inside an inline handler
// attribute: onclick="showDocs(${jsq(bl)})". Safe for both single- and
// double-quoted attributes and for any characters in the value.
function jsq(s) {
  return esc(JSON.stringify(String(s === undefined || s === null ? '' : s)));
}
// Text direction for a free-text box: follow its own content once it has
// some (an English remark on the Arabic board reads left-to-right instead
// of being cut off at the start), but stay with the page's direction while
// empty so the placeholder sits on the correct side.
function autoDir(v) {
  return v ? 'auto' : (currentLang === 'ar' ? 'rtl' : 'ltr');
}

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

/* ---------- Saves vs. the 4-second background refresh ----------
   The board re-reads /api/records every 4s. A refresh that runs while a
   save is still on its way to the server (or that STARTED before the save
   landed and comes back after) carries the OLD values - and used to
   overwrite what was just clicked: bulk "Unmark" turned the sliders off,
   a refresh flipped them back on, then the save finished and they flipped
   off again. The 2-second pause after a click only hid this when the
   server answered quickly. Now every save is tracked, and a refresh result
   is thrown away if any save was in flight at any point while it ran. */
let pendingWrites = 0;  // saves currently in flight
let writeVersion = 0;   // bumped whenever a save starts or finishes
let fetchSeq = 0;       // only the newest refresh may apply its result
function beginWrite() { pendingWrites++; writeVersion++; }
function endWrite() { pendingWrites = Math.max(0, pendingWrites - 1); writeVersion++; }
// fetch() for anything that changes data on the server.
async function apiWrite(url, opts) {
  beginWrite();
  try { return await fetch(url, opts); } finally { endWrite(); }
}
let collapsedGroups = {};
let archivedSectionOpen = false;
let selectedPortTab = '';

// Set when a background refresh brought new data while someone was typing
// in a Remarks box (the board isn't redrawn mid-typing, or it would wipe
// their cursor). Leaving the box used to redraw the whole board instantly,
// every time - and since leaving the box happens on mouse-DOWN of whatever
// they clicked next, the button under the mouse was replaced before the
// click landed and that click was silently lost (type a remark, click a
// slider or Remove -> nothing happened). Now it only redraws if there's
// actually something new, and waits for the click to go through first.
let renderDeferred = false;
function markEditing(delta) {
  editingCount = Math.max(0, editingCount + delta);
  if (editingCount === 0 && renderDeferred) {
    renderDeferred = false;
    setTimeout(() => { if (editingCount === 0) render(); }, 300);
  }
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
  const re = /(\\d+)|(\\D+)/g;
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

// force: the refresh a save does right after it lands - skips the short
// post-click pause (which only exists for the background timer).
async function fetchRecords(force) {
  if (!force && Date.now() < suppressPollUntil) return;
  if (pendingWrites > 0) return;  // a save is in flight - its own follow-up refresh will sync
  const mySeq = ++fetchSeq;
  const versionAtStart = writeVersion;
  const res = await fetch('/api/records');
  if (res.status === 401 || res.redirected) { location.reload(); return; }
  const fresh = await res.json();
  // Stale: a save started/finished while this was in flight, or a newer
  // refresh has already been issued. Drop it rather than roll the board back.
  if (mySeq !== fetchSeq || writeVersion !== versionAtStart || pendingWrites > 0) return;
  fresh.forEach(nr => {
    if (remarksTimers[nr.bl_number]) {
      const old = records.find(r => r.bl_number === nr.bl_number);
      if (old) nr.remarks = old.remarks;
    }
  });
  const changed = JSON.stringify(fresh) !== JSON.stringify(records);
  records = fresh;
  if (changed) {
    if (editingCount === 0) render();
    else renderDeferred = true;
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
  if (!e.dataTransfer.files.length) return;
  document.getElementById('manifestFile').files = e.dataTransfer.files;
  stageManifestFile();
});

// A file dropped a few pixels outside a dropzone would otherwise make the
// browser open it in place of the board (losing an open preview and any
// unsaved remark). Dropzones handle their own drops before this runs.
['dragover', 'drop'].forEach(evt => window.addEventListener(evt, e => {
  if (e.dataTransfer && [...(e.dataTransfer.types || [])].includes('Files')) e.preventDefault();
}));

// Several manifest files can be added at once (e.g. one vessel's cargo split
// across a few files) - all go to the same Port + Vessel typed above.
function stageManifestFile() {
  const files = Array.from(document.getElementById('manifestFile').files || []);
  const btn = document.getElementById('addManifestBtn');
  const label = document.getElementById('dropzoneFilename');
  if (!files.length) { label.textContent = ''; btn.disabled = true; return; }
  label.textContent = files.length === 1 ? files[0].name
    : t('n_files_selected', {n: files.length, names: files.map(f => f.name).join(', ')});
  btn.disabled = false;
}

// One file at a time, so a BL listed in two of the files is simply
// "already on the board" by the second one instead of racing itself.
async function uploadExcel() {
  const fileInput = document.getElementById('manifestFile');
  const files = Array.from(fileInput.files || []);
  if (!files.length) { showToast(t('choose_manifest_first')); return; }

  const btn = document.getElementById('addManifestBtn');
  btn.disabled = true;
  const originalLabel = btn.textContent;
  const port = document.getElementById('portField').value.trim().toUpperCase();
  const vessel = document.getElementById('vesselField').value.trim().toUpperCase();

  let added = 0, skipped = 0, contacts = 0, leftOut = 0;
  const dupElsewhere = [], failed = [];
  for (let i = 0; i < files.length; i++) {
    btn.textContent = files.length > 1 ? t('adding_progress', {i: i + 1, n: files.length}) : t('adding_ellipsis');
    const formData = new FormData();
    formData.append('file', files[i]);
    formData.append('port', port);
    formData.append('vessel', vessel);
    let data;
    try {
      const res = await apiWrite('/api/manifest/upload', { method: 'POST', body: formData });
      data = await res.json();
    } catch (e) {
      data = {error: t('could_not_save_retry')};
    }
    if (data.error) {
      const msg = data.error_code && t('mp_err_' + data.error_code) !== 'mp_err_' + data.error_code ? t('mp_err_' + data.error_code) : data.error;
      failed.push(t('file_failed', {name: files[i].name, error: msg}));
      continue;
    }
    if (data.no_bls) { failed.push(t('file_failed', {name: files[i].name, error: t('mp_err_no_bls')})); continue; }
    added += data.added || 0;
    skipped += data.skipped || 0;
    contacts += data.contacts_found || 0;
    leftOut += data.left_out || 0;
    (data.duplicate_elsewhere || []).forEach(d => dupElsewhere.push(d));
  }

  // Reset the dropzone so the same "Add to board" flow can be repeated.
  fileInput.value = '';
  document.getElementById('dropzoneFilename').textContent = '';
  btn.textContent = originalLabel;
  btn.disabled = true;

  if (failed.length < files.length) {
    showToast(t('bl_records_added', {added}) + (skipped ? t('already_on_board_skipped', {skipped}) : '') +
      (contacts ? t('contacts_found', {n: contacts}) : '') + '.' + (leftOut ? ' ' + t('left_out_hidden', {n: leftOut}) : ''),
      {duration: leftOut ? 8000 : 3500});
  }
  failed.forEach(msg => showToast(msg, {duration: 9000}));
  try { await fetchRecords(true); } catch (e) {}

  // A BL that's already on the board under a DIFFERENT vessel than the one
  // just uploaded is worth a second look - either this file re-lists a BL
  // that's really a different shipment (a shipper reusing a number), or a
  // genuine mistake.
  if (dupElsewhere.length) {
    const lines = dupElsewhere.slice(0, 5).map(d =>
      t('already_under', {bl: d.bl_number, vessel: d.existing_vessel || t('unassigned_vessel_ph'), port: d.existing_port || t('unassigned_port_ph')})
    ).join('; ');
    const more = dupElsewhere.length > 5 ? t('and_n_more', {n: dupElsewhere.length - 5}) : '';
    showToast(t('heads_up_duplicate', {n: dupElsewhere.length, lines, more}), {duration: 9000});
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
  // In Arabic mode, Arabic month names - but Western digits (nu-latn) and
  // the Gregorian calendar (ca-gregory), matching how dates are written in
  // the office; a bare 'ar-SA' would switch to Hijri dates and ٠-٩ digits.
  const locale = currentLang === 'ar' ? 'ar-u-nu-latn-ca-gregory' : undefined;
  return d.toLocaleString(locale, {
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
    badge.innerHTML = '&check; ' + t('complete_badge');
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
  // The save counts as "in flight" from the click itself, not just once
  // the debounced request goes out - otherwise a refresh landing inside
  // the 350ms debounce window could still flip the slider back.
  if (!toggleSendTimers[key]) beginWrite();
  clearTimeout(toggleSendTimers[key]);
  toggleSendTimers[key] = setTimeout(async () => {
    delete toggleSendTimers[key];
    let ok = false;
    try {
      const res = await fetch(`/api/records/${encodeURIComponent(bl)}/toggle`, {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({field, value})
      });
      ok = res.ok;
    } catch (e) { ok = false; }
    endWrite();
    if (!ok) showToast(t('could_not_save_retry'));
    fetchRecords(true);  // sync with the server either way (shows the real state if it failed)
  }, 350);
}

let remarksTimers = {};
function onRemarksInput(bl, value) {
  const rec = records.find(r => r.bl_number === bl);
  if (rec) rec.remarks = value;
  clearTimeout(remarksTimers[bl]);
  remarksTimers[bl] = setTimeout(async () => {
    delete remarksTimers[bl];
    await apiWrite(`/api/records/${encodeURIComponent(bl)}/remarks`, {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({remarks: value})
    });
  }, 500);
}

/* ---------- Per-BL history ---------- */
function auditActionLabel(action) {
  return {
    added: t('action_added'), deleted: t('action_deleted'), restored: t('action_restored'),
    toggle: t('action_toggle'), remarks: t('action_remarks'),
    attachment: t('action_attachment'), attachment_removed: t('action_attachment_removed'),
    contacts: t('action_contacts'), link_reset: t('action_link_reset'),
  }[action] || action;
}

function historyFieldLabel(field) {
  return {invoice_issued: t('th_invoice_issued'), approval_received: t('th_approval_received'), do_issued: t('th_do_issued'), remarks: t('th_remarks')}[field] || field;
}

async function showHistory(bl) {
  const overlay = document.getElementById('historyOverlay');
  const body = document.getElementById('historyBody');
  document.getElementById('historyTitle').textContent = t('history_for', {bl});
  body.innerHTML = `<div style="color:var(--muted); padding:10px 0;">${t('loading')}</div>`;
  overlay.style.display = 'flex';

  const res = await fetch(`/api/records/${encodeURIComponent(bl)}/history`);
  if (!res.ok) { body.innerHTML = `<div style="color:var(--muted); padding:10px 0;">${t('could_not_load_history')}</div>`; return; }
  const entries = await res.json();
  if (!entries.length) {
    body.innerHTML = `<div style="color:var(--muted); padding:10px 0;">${t('no_history_yet')}</div>`;
    return;
  }
  body.innerHTML = entries.map(e => {
    let line;
    const byUser = esc(e.by_user || t('unknown_user'));
    if (e.action === 'customer_download') {
      line = esc(t('history_customer_download'));
    } else if (e.action === 'toggle') {
      line = `<b>${byUser}</b> ${t('history_set_field', {field: esc(historyFieldLabel(e.field)), value: e.new_value ? t('yes') : t('no')})}`;
    } else if (e.action === 'remarks') {
      line = `<b>${byUser}</b> ${e.new_value ? t('history_edited_remarks', {value: esc(e.new_value)}) : t('history_cleared_remarks')}`;
    } else if (e.action === 'notified') {
      line = `<b>${byUser}</b> ${t(e.field === 'whatsapp' ? 'history_notified_whatsapp' : 'history_notified_email', {value: esc(e.new_value)})}`;
    } else if (e.action === 'added') {
      line = `<b>${byUser}</b> ${e.new_value ? t('history_added_this_bl', {value: esc(e.new_value)}) : t('history_added_bl_plain')}`;
    } else {
      line = `<b>${byUser}</b> ${esc(auditActionLabel(e.action))}`;
    }
    return `<div class="history-row"><div>${line}</div><div class="when">${esc(formatLocalTime(e.at))}</div></div>`;
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
const DOC_KINDS = ['invoice', 'do'];

function docChip(bl, kind, hasFile) {
  const label = kind === 'invoice' ? t('doc_chip_inv') : t('doc_chip_do');
  const full = kind === 'invoice' ? t('doc_invoice') : t('doc_delivery_order');
  const title = hasFile ? t('attached_tooltip', {full}) : t('not_attached_tooltip', {full});
  return `<span class="doc-chip ${hasFile ? 'has-file' : 'no-file'}" title="${esc(title)}"
            onclick="event.stopPropagation(); showDocs(${jsq(bl)})">${esc(label)}</span>`;
}

function renderDocsSections(data, canManage) {
  const atts = data.attachments || {};
  return DOC_KINDS.map(kind => {
    const label = kind === 'invoice' ? t('doc_invoice') : t('doc_delivery_order');
    const att = atts[kind];
    let status, actions;
    if (att) {
      status = `${t('attached_label', {filename: ''})}<b>${esc(att.filename || (kind + '.pdf'))}</b><br>${t('by_at', {user: esc(att.uploaded_by || t('unknown_user')), time: esc(formatLocalTime(att.uploaded_at))})}`;
      actions = `
        <a href="/api/records/${encodeURIComponent(data.bl_number)}/attachment/${kind}" target="_blank" rel="noopener">
          <button type="button">${t('download')}</button>
        </a>
        ${canManage ? `
          <label class="btn-upload">${t('replace')}<input type="file" accept=".pdf,application/pdf" onchange="uploadAttachment(${jsq(data.bl_number)}, '${kind}', this)"></label>
          <button type="button" class="btn-danger" onclick="removeAttachment(${jsq(data.bl_number)}, '${kind}')">${t('remove')}</button>
        ` : ''}`;
    } else {
      status = canManage ? t('not_attached_yet') : t('not_attached_waiting');
      actions = canManage ? `
        <label class="btn-upload">${t('upload_pdf')}<input type="file" accept=".pdf,application/pdf" onchange="uploadAttachment(${jsq(data.bl_number)}, '${kind}', this)"></label>
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
  document.getElementById('docsTitle').textContent = t('documents_for', {bl});
  body.innerHTML = `<div style="color:var(--muted); padding:10px 0;">${t('loading')}</div>`;
  overlay.style.display = 'flex';
  overlay.dataset.bl = bl;

  const res = await fetch(`/api/records/${encodeURIComponent(bl)}/lookup`);
  const data = await res.json();
  if (!res.ok) { body.innerHTML = `<div style="color:var(--muted); padding:10px 0;">${esc(data.error || t('could_not_load_bl'))}</div>`; return; }
  const canManage = IS_ADMIN || data.created_by === CURRENT_USER;
  body.innerHTML = renderDocsSections(data, canManage) + renderShareSections(data, canManage);
}

/* ---------- Customer sharing: contacts, tracking link, notify ----------
   Lives in the same Documents popup. Nothing is ever sent automatically:
   "Email the DO" and "WhatsApp" only act on a click. */
function apiErrorText(data) {
  const code = data && data.error_code;
  const translated = code ? t('err_' + code) : '';
  if (translated && translated !== 'err_' + code) {
    return code === 'send_failed' && data.error ? `${translated} (${data.error.replace(/^The email could not be sent: /, '')})` : translated;
  }
  return (data && data.error) || t('could_not_save_retry');
}

function renderShareSections(data, canManage) {
  const bl = data.bl_number;
  const field = (key, label, type) => canManage
    ? `<label class="share-field"><span>${t(label)}</span>
         <input type="${type}" id="ct_${key}" value="${esc(data[key] || '')}" dir="${autoDir(data[key])}"
                oninput="this.dir = autoDir(this.value)" autocomplete="off"></label>`
    : `<div class="share-field"><span>${t(label)}</span><b dir="auto">${esc(data[key] || '-')}</b></div>`;
  const contacts = `
    <div class="docs-section">
      <div class="docs-section-title">${t('customer_contacts')}</div>
      <div class="share-grid">
        <div class="share-full">${field('consignee', 'consignee_name', 'text')}</div>
        ${field('consignee_email', 'consignee_email', 'email')}
        ${field('consignee_phone', 'consignee_phone', 'tel')}
        ${field('broker_email', 'broker_email', 'email')}
        ${field('broker_phone', 'broker_phone', 'tel')}
      </div>
      ${canManage ? `<div class="docs-actions"><button type="button" onclick="saveContacts(${jsq(bl)})">${t('save_contacts')}</button></div>` : ''}
    </div>`;

  const link = data.track_url
    ? `<div class="share-link-row">
         <input type="text" class="share-url" id="shareUrl" value="${esc(data.track_url)}" readonly dir="ltr" onclick="this.select()">
       </div>
       <div class="docs-actions">
         <button type="button" onclick="copyShareLink()">${t('copy_link')}</button>
         ${canManage ? `<button type="button" class="btn-neutral" onclick="toggleQr(${jsq(bl)})">${t('show_qr')}</button>
         <button type="button" class="btn-neutral" title="${esc(t('new_link_title'))}" onclick="confirmNewLink(${jsq(bl)})">${t('new_link')}</button>` : ''}
       </div>
       <div id="qrPanel" class="qr-panel" style="display:none;"></div>`
    : (canManage ? `<div class="docs-actions"><button type="button" onclick="createShareLink(${jsq(bl)}, false)">${t('create_link')}</button></div>` : `<div class="docs-status">-</div>`);
  const share = `
    <div class="docs-section">
      <div class="docs-section-title">${t('share_with_customer')}</div>
      <div class="docs-status">${t('share_help')}</div>
      ${link}
    </div>`;

  if (!canManage) return contacts + share;
  const hasDo = !!(data.attachments && data.attachments.do);
  const recip = (who, label) => {
    const email = data[who + '_email'];
    return `<label class="share-check${email ? '' : ' disabled'}">
      <input type="checkbox" class="notify-to" value="${who}" ${email ? 'checked' : 'disabled'}>
      <span>${t(label)}</span>${email ? `<small dir="ltr">${esc(email)}</small>` : ''}</label>`;
  };
  const emailBlocked = !hasDo ? t('attach_do_first') : (!data.email_configured ? t('email_not_setup') : '');
  const waBtn = (who, label) => data[who + '_phone']
    ? `<button type="button" class="btn-whatsapp" onclick="openWhatsApp(${jsq(bl)}, '${who}')">${t(label)}</button>` : '';
  const anyPhone = data.consignee_phone || data.broker_phone;
  const notify = `
    <div class="docs-section">
      <div class="docs-section-title">${t('notify_customer')}</div>
      <div class="docs-status">${t('email_do_to')}</div>
      <div class="share-checks">${recip('consignee', 'recipient_consignee')}${recip('broker', 'recipient_broker')}</div>
      <div class="docs-actions">
        <button type="button" id="emailDoBtn" onclick="emailDo(${jsq(bl)})" ${emailBlocked ? 'disabled' : ''}>${t('send_do_email')}</button>
      </div>
      ${emailBlocked ? `<div class="docs-status share-hint">${emailBlocked}</div>` : ''}
      ${anyPhone ? `<div class="docs-actions wa-row">${waBtn('consignee', 'whatsapp_consignee')}${waBtn('broker', 'whatsapp_broker')}</div>
                    <div class="docs-status share-hint">${t('whatsapp_help')}</div>` : ''}
    </div>`;
  return contacts + share + notify;
}

async function saveContacts(bl) {
  const val = id => (document.getElementById(id) || {}).value || '';
  const payload = {
    consignee: val('ct_consignee'), consignee_email: val('ct_consignee_email'), consignee_phone: val('ct_consignee_phone'),
    broker_email: val('ct_broker_email'), broker_phone: val('ct_broker_phone'),
  };
  const res = await apiWrite(`/api/records/${encodeURIComponent(bl)}/contacts`, {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    showToast(apiErrorText(data));
    const bad = data.field && document.getElementById('ct_' + data.field);
    if (bad) bad.focus();
    return;
  }
  showToast(t('contacts_saved'));
  await showDocs(bl);
  fetchRecords(true);
}

async function createShareLink(bl, reset) {
  const res = await apiWrite(`/api/records/${encodeURIComponent(bl)}/share-link${reset ? '?reset=1' : ''}`, {method: 'POST'});
  const data = await res.json().catch(() => ({}));
  if (!res.ok) { showToast(apiErrorText(data)); return; }
  await showDocs(bl);
  if (reset) showToast(t('link_replaced'));
  else copyShareLink();
}

function confirmNewLink(bl) {
  showToast(t('confirm_new_link'), {actionLabel: t('confirm'), duration: 6000, onAction: () => createShareLink(bl, true)});
}

async function copyShareLink() {
  const box = document.getElementById('shareUrl');
  if (!box) return;
  try {
    await navigator.clipboard.writeText(box.value);
  } catch (e) {
    box.select(); document.execCommand('copy');  // older browsers / non-HTTPS
  }
  showToast(t('link_copied'));
}

function toggleQr(bl) {
  const panel = document.getElementById('qrPanel');
  if (!panel) return;
  if (panel.style.display !== 'none') { panel.style.display = 'none'; return; }
  const base = `/api/records/${encodeURIComponent(bl)}/qr`;
  panel.innerHTML = `<img src="${base}?v=${Date.now()}" alt="QR">
    <a href="${base}?format=png&download=1">${t('download_qr')}</a>`;
  panel.style.display = '';
}

async function emailDo(bl) {
  const to = [...document.querySelectorAll('.notify-to:checked')].map(c => c.value);
  if (!to.length) { showToast(t('err_no_recipient')); return; }
  const btn = document.getElementById('emailDoBtn');
  const label = btn.textContent;
  btn.disabled = true; btn.textContent = t('sending_ellipsis');
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 35000);
  try {
    const res = await apiWrite(`/api/records/${encodeURIComponent(bl)}/notify/email`, {
      method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({to}), signal: ctl.signal
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) { showToast(apiErrorText(data), {duration: 9000}); return; }
    showToast(t('do_emailed', {to: data.sent_to.join(', ')}));
  } catch (e) {
    showToast(t('err_email_timeout'), {duration: 9000});
  } finally {
    clearTimeout(timer);
    btn.disabled = false; btn.textContent = label;
  }
}

async function openWhatsApp(bl, who) {
  // Open the tab right away (inside the click), then point it at WhatsApp
  // once the server returns the link - browsers block a window opened
  // later, after a network round trip, as an unwanted popup.
  const win = window.open('about:blank', '_blank');
  const res = await apiWrite(`/api/records/${encodeURIComponent(bl)}/notify/whatsapp`, {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({to: who})
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) { if (win) win.close(); showToast(apiErrorText(data)); return; }
  if (win) { win.opener = null; win.location.href = data.url; } else { location.href = data.url; }
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
  try {
    const res = await apiWrite(`/api/records/${encodeURIComponent(bl)}/attachment/${kind}`, {method: 'POST', body: form});
    let data = {};
    try { data = await res.json(); } catch (e) {}
    return {ok: res.ok, data};
  } catch (e) {
    return {ok: false, data: {error: t('could_not_save_retry')}};
  }
}

async function uploadAttachment(bl, kind, input) {
  const file = input.files && input.files[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.pdf')) { showToast(t('only_pdf_accepted')); input.value = ''; return; }
  if (file.size > 10 * 1024 * 1024) { showToast(t('file_too_large')); input.value = ''; return; }

  const {ok, data} = await submitAttachmentFile(bl, kind, file);
  if (!ok) { showToast(data.error || t('upload_failed')); return; }
  const label = kind === 'invoice' ? t('doc_invoice') : t('doc_delivery_order');
  // Attaching the file auto-flips the matching Issued slider server-side
  // (see upload_attachment) - say so, so it's obvious the status change
  // wasn't a separate click someone forgot to make.
  showToast(data.auto_issued_field ? t('uploaded_marked_issued', {label}) : t('uploaded', {label}));
  await fetchRecords(true);
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
  return kind === 'invoice' ? t('doc_invoice') : kind === 'do' ? t('doc_delivery_order') : t('unrecognized_document');
}

/* The results popup. Opening it when it's closed starts a fresh list;
   dropping more files while it's already open adds to the same list. */
let autoMatchRunning = 0;  // batches still being processed
function autoMatchSummaryText() {
  const list = document.getElementById('autoMatchList');
  const attached = list.querySelectorAll('.match-row.ok').length;
  const review = list.querySelectorAll('.match-row.review').length;
  const pending = list.querySelectorAll('.match-row').length - attached - review;
  return t('attached_automatically', {n: attached}) + (review ? t('need_your_input', {n: review}) : '') +
    (pending ? t('processing_more', {n: pending}) : '.');
}
function updateAutoMatchSummary() {
  document.getElementById('autoMatchSummary').textContent = autoMatchSummaryText();
}
function openAutoMatch() {
  const overlay = document.getElementById('autoMatchOverlay');
  if (overlay.style.display === 'none') {
    document.getElementById('autoMatchList').innerHTML = '';
    document.getElementById('autoMatchSummary').textContent = '';
  }
  overlay.style.display = 'flex';
}
function closeAutoMatch() {
  document.getElementById('autoMatchOverlay').style.display = 'none';
}
function autoMatchIsOpen() {
  return document.getElementById('autoMatchOverlay').style.display !== 'none';
}

async function handleAutoMatchFiles(fileList) {
  const files = Array.from(fileList || []).filter(f => f.name.toLowerCase().endsWith('.pdf'));
  const input = document.getElementById('autoMatchFile');
  if (input) input.value = '';  // so the same file can be picked again later
  if (!files.length) { showToast(t('drop_pdf_only')); return; }

  openAutoMatch();
  autoMatchRunning++;
  const listEl = document.getElementById('autoMatchList');
  const updateSummary = updateAutoMatchSummary;

  // One row per file straight away, so the popup shows the whole batch
  // (and "processing N more") from the start.
  const rows = files.map((file, i) => {
    const row = document.createElement('div');
    row.className = 'match-row';
    row.id = `matchrow_${Date.now()}_${i}`;
    row.innerHTML = `<div class="match-file" title="${esc(file.name)}">${esc(file.name)}</div><div class="match-status">${t('reading_ellipsis')}</div>`;
    listEl.appendChild(row);
    return row;
  });
  updateSummary();

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const row = rows[i];

    if (file.size > 10 * 1024 * 1024) {
      row.querySelector('.match-status').textContent = t('too_large_skipped');
      updateSummary();
      continue;
    }

    let detect;
    try {
      const form = new FormData();
      form.append('file', file);
      const res = await fetch('/api/attachments/detect', {method: 'POST', body: form});
      detect = await res.json();
      if (!res.ok) throw new Error(detect.error || t('could_not_read_file'));
    } catch (err) {
      row.className = 'match-row review';
      row.innerHTML = `<div class="match-file" title="${esc(file.name)}">${esc(file.name)}</div><div class="match-status">${esc(err.message)}</div>`;
      updateSummary();
      continue;
    }

    if (detect.kind && detect.matched_bl) {
      const {ok, data} = await submitAttachmentFile(detect.matched_bl, detect.kind, file);
      if (ok) {
        row.className = 'match-row ok';
        const issuedNote = data.auto_issued_field ? t('marked_issued_suffix') : '';
        row.innerHTML = `<div class="match-file" title="${esc(file.name)}">${esc(file.name)}</div><div class="match-status">&check; ${esc(detect.matched_bl)} - ${esc(autoMatchKindLabel(detect.kind))}${esc(issuedNote)}</div>`;
        updateSummary();
        continue;
      }
      // Fall through to manual review if the upload itself was rejected
      // (e.g. not the owner of that BL after all) - rare, since the
      // candidate list was already scoped server-side, but don't just
      // drop the file silently if it happens.
      row.className = 'match-row review';
      renderAutoMatchReviewRow(row, file, detect, data.error);
      updateSummary();
      continue;
    }

    row.className = 'match-row review';
    renderAutoMatchReviewRow(row, file, detect, null);
    updateSummary();
  }

  autoMatchRunning--;
  await fetchRecords(true);
  // Closed the popup before it finished? Say how it went, briefly.
  if (!autoMatchIsOpen() && !autoMatchRunning) showToast(t('attach_batch_done', {summary: autoMatchSummaryText()}), {duration: 7000});
}

function renderAutoMatchReviewRow(row, file, detect, errorMsg) {
  const blOptions = records.map(r => r.bl_number).sort();
  const candidates = (detect.candidates && detect.candidates.length) ? detect.candidates : blOptions;
  const statusText = errorMsg ? errorMsg
    : detect.reason === 'already_attached' && detect.candidates && detect.candidates.length === 1
      ? t('already_has_file_confirm', {bl: detect.candidates[0], kind: autoMatchKindLabel(detect.kind)})
    : detect.reason === 'member_match' && detect.candidates && detect.candidates.length === 1
      ? t('member_match_confirm', {bl: detect.candidates[0]})
    : detect.candidates && detect.candidates.length ? t('matches_multiple')
    : !detect.kind ? t('could_not_tell_kind')
    : t('no_bl_matched');

  row.innerHTML = `
    <div class="match-file" title="${esc(file.name)}">${esc(file.name)}</div>
    <div class="match-status">${esc(statusText)}</div>
    <div class="match-review-controls">
      <select class="review-bl">
        <option value="">${t('select_bl_ellipsis')}</option>
        ${candidates.map(bl => `<option value="${esc(bl)}" ${(bl === detect.matched_bl || (detect.reason === 'member_match' && candidates.length === 1)) && detect.reason !== 'already_attached' ? 'selected' : ''}>${esc(bl)}</option>`).join('')}
      </select>
      <select class="review-kind">
        <option value="">${t('kind_ellipsis')}</option>
        <option value="invoice" ${detect.kind === 'invoice' ? 'selected' : ''}>${t('doc_invoice')}</option>
        <option value="do" ${detect.kind === 'do' ? 'selected' : ''}>${t('doc_delivery_order')}</option>
      </select>
      <button type="button" class="review-attach-btn">${t('attach_btn')}</button>
    </div>`;

  row.querySelector('.review-attach-btn').onclick = async () => {
    const bl = row.querySelector('.review-bl').value;
    const kind = row.querySelector('.review-kind').value;
    if (!bl || !kind) { showToast(t('pick_bl_and_kind')); return; }
    const btn = row.querySelector('.review-attach-btn');
    btn.disabled = true;
    btn.textContent = t('attaching_ellipsis');
    const {ok, data} = await submitAttachmentFile(bl, kind, file);
    if (!ok) { showToast(data.error || t('attach_failed')); btn.disabled = false; btn.textContent = t('attach_btn'); return; }
    row.className = 'match-row ok';
    const issuedNote = data.auto_issued_field ? t('marked_issued_suffix') : '';
    row.innerHTML = `<div class="match-file" title="${esc(file.name)}">${esc(file.name)}</div><div class="match-status">&check; ${esc(bl)} - ${esc(autoMatchKindLabel(kind))}${esc(issuedNote)}</div>`;
    updateAutoMatchSummary();
    await fetchRecords(true);
  };
}

async function removeAttachment(bl, kind) {
  const res = await apiWrite(`/api/records/${encodeURIComponent(bl)}/attachment/${kind}`, {method: 'DELETE'});
  if (!res.ok) { showToast(t('could_not_remove_file')); return; }
  showToast(t('doc_removed', {kind: kind === 'invoice' ? t('doc_invoice') : t('doc_delivery_order')}));
  await fetchRecords(true);
  if (document.getElementById('docsOverlay').style.display !== 'none') await showDocs(bl);
}

function deleteRecord(bl) {
  const idx = records.findIndex(r => r.bl_number === bl);
  if (idx === -1) return;
  const removed = records[idx];
  records.splice(idx, 1);
  render();
  suppressPollUntil = Date.now() + 4000;
  const deleting = apiWrite(`/api/records/${encodeURIComponent(bl)}`, {method: 'DELETE'}).catch(() => { fetchRecords(true); });

  showToast(t('removed_bl', {bl}), {
    actionLabel: t('undo'),
    duration: 3000,
    onAction: async () => {
      // Undo pressed while the delete is still on its way: let it finish
      // first, or the restore finds the BL still there, does nothing, and
      // the delete then removes it for good.
      await deleting;
      await apiWrite('/api/records/restore', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(removed)
      });
      await fetchRecords(true);
      showToast(t('restored_bl', {bl}));
    }
  });
}

/* ---------- Bulk remove (vessel group / port group / whole board) ----------
   Same staged-confirm + Undo pattern as the single-row delete above, just
   operating on a whole list of records at once via the bulk API so a
   500-BL manifest doesn't fire 500 individual requests. */
function confirmBulkRemove(label, list) {
  if (!list.length) { showToast(t('nothing_to_remove')); return; }
  const n = list.length;
  showToast(t('confirm_remove_all', {n, p: n === 1 ? '' : 's', inLabel: label ? t('in_label', {label}) : ''}), {
    actionLabel: t('confirm'),
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

  const res = await apiWrite('/api/records/bulk-delete', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers})
  });
  const data = await res.json();
  const deleted = (data.deleted && data.deleted.length) ? data.deleted : snapshot;

  showToast(t('bls_removed', {n: deleted.length, p: deleted.length === 1 ? '' : 's'}), {
    actionLabel: t('undo'),
    duration: 5000,
    onAction: async () => {
      await apiWrite('/api/records/bulk-restore', {
        method: 'POST', headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({records: deleted})
      });
      await fetchRecords(true);
      showToast(t('restored'));
    }
  });
  await fetchRecords(true);
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
  // The group's own BLs (archived ones included), so an "Unassigned" group -
  // stored with a blank port/vessel - can be named too.
  const blNumbers = records.filter(r => (r.port || 'Unassigned') === oldPort &&
    (type === 'port' || (r.vessel || 'Unassigned') === oldVessel)).map(r => r.bl_number);
  await apiWrite('/api/groups/rename', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({type, old_port: oldPort, old_vessel: oldVessel, bl_numbers: blNumbers,
                          new_value: val === fallbackLabel ? '' : val.toUpperCase()})
  });
  await fetchRecords(true);
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
        <input type="checkbox" id="${esc(id)}" ${checked ? 'checked' : ''}
          onchange="toggle(${jsq(bl)}, '${field}', this.checked)">
        <span class="slider"></span>
      </label>
      ${checked ? `<span class="meta" title="${esc((by || '') + ' - ' + formatLocalTime(at))}">${esc((by || '') + ' - ' + formatLocalTime(at))}</span>` : ''}
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
    <div class="stat"><div class="stat-icon">${icons.total}</div><div><b>${total}</b>${t('total_bls')}</div></div>
    <div class="stat ${remaining ? 'gold' : 'done'}"><div class="stat-icon">${icons.check}</div><div><b>${remaining}</b>${t('remaining')}</div></div>
    <div class="stat done"><div class="stat-icon">${icons.check}</div><div><b>${complete}</b>${t('fully_complete')}</div></div>
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
            onchange="toggleRowSelect(${jsq(r.bl_number)}, this.checked)"></td>
      <td>
        <div class="bl-cell">
          <b>${esc(r.bl_number)}</b>
          <div class="bl-cell-chips">
            ${docChip(r.bl_number, 'invoice', r.has_invoice_file)}
            ${docChip(r.bl_number, 'do', r.has_do_file)}
            ${complete ? `<span class="badge-complete">&check; ${t('complete_badge')}</span>` : ''}
          </div>
        </div>
      </td>
      <td data-label="${t('th_invoice_issued')}">${checkbox(r.bl_number, 'invoice_issued', !!r.invoice_issued, r.invoice_by, r.invoice_at)}</td>
      <td data-label="${t('th_approval_received')}">${checkbox(r.bl_number, 'approval_received', !!r.approval_received, r.approval_by, r.approval_at)}</td>
      <td data-label="${t('th_do_issued')}">${checkbox(r.bl_number, 'do_issued', !!r.do_issued, r.do_by, r.do_at)}</td>
      <td data-label="${t('th_remarks')}"><input class="remarks-input" type="text" dir="${autoDir(r.remarks)}" value="${esc(r.remarks || '')}"
            oninput="this.dir = autoDir(this.value); onRemarksInput(${jsq(r.bl_number)}, this.value)"
            onfocus="markEditing(1)" onblur="markEditing(-1)" placeholder="${t('notes_placeholder')}"></td>
      <td>{% if role == 'admin' %}<button type="button" class="hist-btn" title="${t('history')}" onclick="showHistory(${jsq(r.bl_number)})">${t('history')}</button>{% endif %}</td>
      <td><button class="del" onclick="deleteRecord(${jsq(r.bl_number)})">${t('remove')}</button></td>
    </tr>`;
  }).join('');
}

function toggleRowSelect(bl, checked) {
  if (checked) selectedBLs.add(bl); else selectedBLs.delete(bl);
  updateBulkBars();
}

let dockedBarKey = null;  // vessel whose selection bar is pinned (phones)
function updateBulkBars() {
  // The selection bar appears above the BL list when the first row is
  // ticked (and disappears with the last), which pushed the whole list -
  // including the row just clicked - down/up by the bar's height (57px on
  // a desktop, ~200px on a phone). Keep the row that was clicked fixed on
  // screen by scrolling the page by exactly that amount.
  const active = document.activeElement;
  const anchor = active && active.closest ? active.closest('tr') : null;
  const anchorTop = anchor ? anchor.getBoundingClientRect().top : null;
  document.querySelectorAll('.bulk-bar').forEach(bar => {
    const scope = (bar.dataset.bls || '').split('|').filter(Boolean);
    const count = scope.filter(bl => selectedBLs.has(bl)).length;
    const label = bar.querySelector('.bulk-count');
    if (label) label.textContent = count ? t('selected_count', {n: count}) : '';
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
  // Which selection bar gets pinned to the bottom on a phone: the one for
  // the vessel just clicked in, else the one already pinned, else any.
  let dock = null;
  if (anchor && anchor.isConnected) {
    const g = anchor.closest('.vessel-group');
    dock = g ? g.querySelector('.bulk-bar.active') : null;
  }
  if (!dock && dockedBarKey) {
    dock = [...document.querySelectorAll('.bulk-bar.active')].find(b => b.dataset.barKey === dockedBarKey) || null;
  }
  if (!dock) dock = document.querySelector('.bulk-bar.active');
  dockedBarKey = dock ? dock.dataset.barKey : null;
  document.querySelectorAll('.bulk-bar').forEach(b => b.classList.toggle('docked', b === dock));
  const pinned = dock && getComputedStyle(dock).position === 'fixed';
  document.documentElement.style.setProperty('--dock-h', pinned ? dock.offsetHeight + 'px' : '0px');

  if (anchor && anchor.isConnected && anchorTop !== null) {
    const shift = anchor.getBoundingClientRect().top - anchorTop;
    if (Math.abs(shift) > 1) window.scrollBy(0, shift);
  }
}

async function bulkSetField(vesselKey, field, value) {
  const listEl = document.querySelector(`[data-bar-key="${CSS.escape(vesselKey)}"]`);
  const scope = listEl ? (listEl.dataset.bls || '').split('|').filter(Boolean) : [];
  const blNumbers = scope.filter(bl => selectedBLs.has(bl));
  if (!blNumbers.length) { showToast(t('select_at_least_one')); return; }

  // Only BLs whose state actually changes are touched - bulk "Mark" on a
  // selection where some were already marked used to overwrite who/when
  // on those too (making it look like you'd issued Ahmed's invoice today,
  // and skewing KPI turnaround times). Same rule on the server.
  blNumbers.forEach(bl => {
    const rec = records.find(r => r.bl_number === bl);
    if (rec && !!rec[field] !== !!value) {
      rec[field] = value ? 1 : 0;
      const byField = field.replace('_issued', '_by').replace('_received', '_by');
      const atField = field.replace('_issued', '_at').replace('_received', '_at');
      rec[byField] = value ? CURRENT_USER : '';
      rec[atField] = value ? nowLabel() : '';
    }
  });
  suppressPollUntil = Date.now() + 2000;
  render();

  let data = null;
  try {
    const res = await apiWrite('/api/records/bulk-toggle', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({bl_numbers: blNumbers, field, value})
    });
    if (res.ok) data = await res.json();
  } catch (e) { data = null; }
  if (data && Array.isArray(data.updated)) showToast(t('bls_updated', {n: data.updated.length}));
  else showToast(t('could_not_save_retry'));
  await fetchRecords(true);
}

function bulkRemoveSelected(vesselKey) {
  // Scoped to only the checked BLs in this vessel's bar - NOT the whole
  // vessel group (that's what the header's "Remove all" button is for) -
  // and passes an explicit bl_numbers list to the backend, same as every
  // other per-group bulk action here.
  const listEl = document.querySelector(`[data-bar-key="${CSS.escape(vesselKey)}"]`);
  const scope = listEl ? (listEl.dataset.bls || '').split('|').filter(Boolean) : [];
  const blNumbers = scope.filter(bl => selectedBLs.has(bl));
  if (!blNumbers.length) { showToast(t('select_at_least_one')); return; }
  const list = blNumbers.map(bl => records.find(r => r.bl_number === bl)).filter(Boolean);
  confirmBulkRemove(null, list);
}

function bulkBarHtml(vesselKey, list) {
  const key = esc(vesselKey);
  const blsAttr = list.map(r => r.bl_number).join('|');
  return `
    <div class="bulk-bar" data-bar-key="${key}" data-bls="${esc(blsAttr)}">
      <span class="bulk-count"></span>
      <span class="bulk-group">
        <button type="button" onclick="bulkSetField(${jsq(vesselKey)}, 'invoice_issued', true)">${t('mark_invoice_issued')}</button>
        <button type="button" class="bulk-unmark-btn" onclick="bulkSetField(${jsq(vesselKey)}, 'invoice_issued', false)">${t('unmark')}</button>
      </span>
      <span class="bulk-group">
        <button type="button" onclick="bulkSetField(${jsq(vesselKey)}, 'approval_received', true)">${t('mark_approval_received')}</button>
        <button type="button" class="bulk-unmark-btn" onclick="bulkSetField(${jsq(vesselKey)}, 'approval_received', false)">${t('unmark')}</button>
      </span>
      <span class="bulk-group">
        <button type="button" onclick="bulkSetField(${jsq(vesselKey)}, 'do_issued', true)">${t('mark_do_issued')}</button>
        <button type="button" class="bulk-unmark-btn" onclick="bulkSetField(${jsq(vesselKey)}, 'do_issued', false)">${t('unmark')}</button>
      </span>
      <button type="button" class="bulk-remove-btn" onclick="bulkRemoveSelected(${jsq(vesselKey)})">${t('remove')}</button>
    </div>`;
}

function tableHtml(list, vesselKey) {
  const key = esc(vesselKey);
  return `
    ${bulkBarHtml(vesselKey, list)}
    <div class="overflow">
      <table>
        <thead>
          <tr>
            <th class="select-col"><input type="checkbox" class="select-all-vessel" data-bar-key="${key}" title="${t('select_deselect_vessel_title')}" onchange="list_selectAllVessel(${jsq(vesselKey)}, this.checked)"></th>
            <th>${t('th_bl_number')}</th>
            <th>${t('th_invoice_issued')}</th>
            <th>${t('th_approval_received')}</th>
            <th>${t('th_do_issued')}</th>
            <th>${t('th_remarks')}</th>
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
  // Same as updateBulkBars(): keep the list still while the selection bar
  // appears/disappears above it.
  const box = bar ? bar.parentElement.querySelector('.overflow') : null;
  const groupId = box ? box.closest('.vessel-group').id : null;
  const before = box ? box.getBoundingClientRect().top : null;
  render();
  const after = groupId ? document.querySelector('#' + CSS.escape(groupId) + ' .overflow') : null;
  if (after && before !== null) {
    const shift = after.getBoundingClientRect().top - before;
    if (Math.abs(shift) > 1) window.scrollBy(0, shift);
  }
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
  // HTML-escaped JSON array, safe inside a double-quoted attribute.
  const blsJson = esc(JSON.stringify(blList));
  return `
    <div class="vessel-group" id="group_${cssEscape(vesselKey)}">
      <div class="vessel-header ${vesselCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT' && event.target.tagName!=='BUTTON' && event.target.tagName!=='A') toggleGroup(${jsq(collapseKey)})">
        ${CHEVRON}
        <input class="group-name" dir="${autoDir(vesselName === 'Unassigned' ? '' : vesselName)}" oninput="this.dir = autoDir(this.value)" value="${vesselName === 'Unassigned' ? '' : esc(vesselName)}" placeholder="${t('unassigned_vessel_ph')}"
          onclick="event.stopPropagation()"
          onchange="renameGroup('vessel', ${jsq(portName)}, ${jsq(vesselName)}, this.value, 'Unassigned')">
        ${archivedView ? '' : `<span class="eta-wrap" onclick="event.stopPropagation()">
          <span class="eta-label">${t('eta_label')}</span>
          <input type="date" class="eta-input${eta ? '' : ' eta-unset'}" value="${esc(eta)}" title="${t('expected_arrival_title')}"
            onchange="this.classList.toggle('eta-unset', !this.value); setVesselEta(JSON.parse(this.dataset.bls), this.value)" data-bls="${blsJson}">
          ${eta ? '' : `<span class="eta-unset-hint">${t('not_set')}</span>`}
        </span>`}
        <span class="group-count">${t('bl_count', {n: sortedList.length, p: sortedList.length === 1 ? '' : 's'})}${left ? t('left_suffix', {n: left}) : t('done_suffix')}</span>
        <span class="vessel-progress ${pct >= 100 ? 'done' : ''}" title="${t('pct_complete', {pct})}"><span class="vessel-progress-fill" style="width:${pct}%"></span></span>
        <span class="group-actions">
          <a onclick="event.stopPropagation()" href="/api/export?port=${encodeURIComponent(rawPort)}&vessel=${encodeURIComponent(rawVessel)}" class="group-export" title="${t('export_vessel_title')}">${t('export')}</a>
          <button type="button" class="group-neutral" onclick="event.stopPropagation(); setVesselArchived(${blsJson}, ${archivedView ? 'false' : 'true'})">${archivedView ? t('unarchive') : t('archive')}</button>
          <button type="button" class="group-remove" onclick="event.stopPropagation(); removeVesselGroup(${jsq(portName)}, ${jsq(vesselName)})">${t('remove_all')}</button>
        </span>
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
  return `
    <div class="port-group">
      <div class="port-header ${portCollapsed ? 'collapsed' : ''}" onclick="if(event.target.tagName!=='INPUT' && event.target.tagName!=='A') toggleGroup(${jsq(portKey)})">
        ${CHEVRON}
        <input class="group-name" dir="${autoDir(portName === 'Unassigned' ? '' : portName)}" oninput="this.dir = autoDir(this.value)" value="${portName === 'Unassigned' ? '' : esc(portName)}" placeholder="${t('unassigned_port_ph')}"
          onclick="event.stopPropagation()"
          onchange="renameGroup('port', ${jsq(portName)}, '', this.value, 'Unassigned')">
      </div>
      <div class="port-body ${portCollapsed ? 'collapsed' : ''}">${vesselsHtml}</div>
    </div>`;
}

// Each vessel's BL list scrolls inside its own box. render() rebuilds the
// board from scratch, and a rebuilt box always starts at the top - so ticking
// a slider halfway down a 70-BL vessel threw you back to the first BL once
// the save finished and the board redrew. Remember every box's scroll
// position before the rebuild and put it back afterwards.
function saveListScroll() {
  const saved = {};
  document.querySelectorAll('.vessel-group').forEach(g => {
    const box = g.querySelector('.overflow');
    if (box && (box.scrollTop || box.scrollLeft)) saved[g.id] = [box.scrollTop, box.scrollLeft];
  });
  return saved;
}
function restoreListScroll(saved) {
  Object.keys(saved).forEach(id => {
    const g = document.getElementById(id);
    const box = g && g.querySelector('.overflow');
    if (box) { box.scrollTop = saved[id][0]; box.scrollLeft = saved[id][1]; }
  });
}

// Search match: part of the BL number as shown, or one of the individual
// B/Ls a combined entry stands for ("...002" finds "...001-003").
function blMatches(r, q) {
  if (!q) return true;
  if (r.bl_number.toLowerCase().includes(q)) return true;
  return (r.bl_members || []).some(m => m.toLowerCase().includes(q));
}
function blExact(r, q) {
  return r.bl_number.toLowerCase() === q || (r.bl_members || []).some(m => m.toLowerCase() === q);
}

function render() {
  const savedListScroll = saveListScroll();
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  const operatorFilterEl = document.getElementById('operatorFilter');
  if (operatorFilterEl) {
    const operators = [...new Set(records.map(r => r.created_by).filter(Boolean))].sort();
    const current = operatorFilterEl.value;
    operatorFilterEl.innerHTML = `<option value="">${t('all_operators')}</option>` + operators.map(a => `<option value="${esc(a)}">${esc(a)}</option>`).join('');
    if (operators.includes(current)) operatorFilterEl.value = current;
    syncGlassSelectLabel('operatorFilter');
  }
  const operatorFilter = operatorFilterEl ? operatorFilterEl.value : '';

  const base = records.filter(r => blMatches(r, q) && (!operatorFilter || r.created_by === operatorFilter));
  const activeList = base.filter(r => !r.archived);
  const archivedList = base.filter(r => !!r.archived);

  const ports = groupRecordsByPortVessel(activeList);
  const portNames = sortedPortNames(ports);

  // Port landing nav - rebuilt fresh every render from whatever ports are
  // currently on the board (post search/operator filter), same pattern
  // as jumpSelect/operatorFilter just above. If the previously-selected
  // port disappeared (renamed away, last BL removed, etc.) fall back to
  // "All ports" rather than showing an empty board.
  // A search/filter with no hits in the open port shows all ports for the
  // moment, but clearing it goes back into that port. Only a port that
  // really disappeared (renamed away, last BL removed) is forgotten.
  if (selectedPortTab && !portNames.includes(selectedPortTab) && !q && !operatorFilter) selectedPortTab = '';
  const activeTab = selectedPortTab && portNames.includes(selectedPortTab) ? selectedPortTab : '';
  const portTabsEl = document.getElementById('portTabs');
  if (portTabsEl) {
    if (portNames.length === 0) {
      portTabsEl.innerHTML = '';
    } else if (activeTab === '') {
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
        const label = esc(portName === 'Unassigned' ? t('unassigned_port_ph') : portName);
        return `<button type="button" class="port-card" onclick="selectPortTab(${jsq(portName)})">
          <div class="port-card-icon">${anchorIcon}</div>
          <div class="port-card-name">${label}</div>
          <div class="port-card-meta">${t('port_card_meta', {vesselCount, vp: vesselCount === 1 ? '' : 's', blCount, bp: blCount === 1 ? '' : 's'})}</div>
          <span class="vessel-progress port-card-progress ${done ? 'done' : ''}"><span class="vessel-progress-fill" style="width:${pct}%"></span></span>
          <div class="port-card-pct ${done ? 'done' : ''}">${done ? t('all_done') : t('pct_complete', {pct})}</div>
        </button>`;
      }).join('');
      portTabsEl.innerHTML = `<div class="port-card-grid">${cards}</div>`;
    } else {
      // Drilled-in state: a port is selected, so the card grid is hidden
      // and replaced by a small breadcrumb/back control + heading above
      // that one port's (unchanged) groups.
      const label = esc(activeTab === 'Unassigned' ? t('unassigned_port_ph') : activeTab);
      const backArrow = currentLang === 'ar' ? '&rarr;' : '&larr;';
      portTabsEl.innerHTML = `<div class="port-breadcrumb">
        <button type="button" class="port-back-btn" onclick="selectPortTab('')">${backArrow} ${t('all_ports_back')}</button>
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
  const displayPortNames = activeTab ? portNames.filter(p => p === activeTab) : (q ? portNames : []);

  const groupsEl = document.getElementById('groups');
  if (portNames.length === 0) {
    groupsEl.innerHTML = `<div style="color:var(--muted); padding:24px 4px;">${t('no_bls_yet')}</div>`;
  } else {
    // Vessel/port jump menu - with 20-25 manifests a month, scrolling down
    // the whole board to find one vessel doesn't scale. Built fresh every
    // render so it always reflects what's actually on the board right now.
    const jumpEl = document.getElementById('jumpSelect');
    if (jumpEl) {
      const current = jumpEl.value;
      let options = `<option value="">${t('select_vessel_to_view')}</option>`;
      portNames.forEach(portName => {
        const vessels = ports[portName];
        sortedVesselNames(vessels).forEach(vesselName => {
          const key = 'vessel:' + portName + ':' + vesselName;
          options += `<option value="${esc(key)}">${esc(vesselName === 'Unassigned' ? t('unassigned_vessel_ph') : vesselName)}</option>`;
        });
      });
      jumpEl.innerHTML = options;
      if ([...jumpEl.options].some(o => o.value === current)) jumpEl.value = current;
      syncGlassSelectLabel('jumpSelect');
    }

    groupsEl.innerHTML = displayPortNames.map(portName => {
      const vessels = ports[portName];
      return portGroupHtml(portName, sortedVesselNames(vessels), vessels, false, !!activeTab);
    }).join('');
  }

  const archivedCard = document.getElementById('archivedCard');
  if (archivedCard) {
    if (archivedList.length === 0) {
      archivedCard.style.display = 'none';
    } else {
      archivedCard.style.display = '';
      document.getElementById('archivedCount').textContent = t('bl_count', {n: archivedList.length, p: archivedList.length === 1 ? '' : 's'});
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

  restoreListScroll(savedListScroll);
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
  await apiWrite('/api/vessel/eta', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers, eta})
  });
  await fetchRecords(true);
}

async function setVesselArchived(blNumbers, archived) {
  blNumbers.forEach(bl => {
    const rec = records.find(r => r.bl_number === bl);
    if (rec) rec.archived = archived ? 1 : 0;
  });
  suppressPollUntil = Date.now() + 1500;
  render();
  await apiWrite('/api/vessel/archive', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({bl_numbers: blNumbers, archived})
  });
  showToast(archived ? t('vessel_archived') : t('vessel_restored'));
  await fetchRecords(true);
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

// Only show the back-to-top button once the manifest form is out of view -
// there's nothing to scroll back to before that, and it was covering
// controls ("Collapse all", the archived count) near the top of the page.
(function initScrollTopVisibility() {
  const btn = document.getElementById('scrollTopBtn');
  const card = document.getElementById('manifestCard');
  if (!btn || !card) return;
  const update = () => btn.classList.toggle('show', card.getBoundingClientRect().bottom < 0);
  window.addEventListener('scroll', update, { passive: true });
  window.addEventListener('resize', update);
  update();
})();

function setAllGroupsCollapsed(collapsed) {
  const q = document.getElementById('searchBox').value.trim().toLowerCase();
  records.filter(r => blMatches(r, q)).forEach(r => {
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

/* ---------- Keyboard shortcuts ----------
     /      jump to the BL search box from anywhere (selects what's there,
            so you can just type the next BL)
     Enter  (in search) jump straight to the matching BL - see onSearchKeydown
     Esc    (in search) clear it; otherwise close whatever popup is open,
            or collapse all groups if nothing is
   All are no-ops while typing in another field. "/" is matched by the
   physical key (e.code 'Slash') as well as the character: on an Arabic
   keyboard that same key types "ظ", so e.key alone never fired for
   Arabic-keyboard users. */
function isTypingTarget(el) {
  const tag = el ? el.tagName : '';
  return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || (el && el.isContentEditable);
}
function focusSearch() {
  const box = document.getElementById('searchBox');
  if (!box) return;
  box.focus();
  box.select();
}
document.addEventListener('keydown', (e) => {
  const isTyping = isTypingTarget(document.activeElement);

  const popupOpen = ['autoMatchOverlay', 'docsOverlay', 'historyOverlay']
    .some(id => { const el = document.getElementById(id); return el && el.style.display !== 'none'; });
  if ((e.key === '/' || e.code === 'Slash') && !isTyping && !popupOpen && !e.ctrlKey && !e.metaKey && !e.altKey) {
    e.preventDefault();
    focusSearch();
    return;
  }

  if (e.key === 'Escape') {
    if (isTyping) return;
    // A glass-select panel being open means its own trigger keydown
    // handler (onGlassTriggerKeydown) already closed it and stopped this
    // keystroke from bubbling here - this is just a safety net.
    if (Object.values(customSelects).some(s => s.open)) return;
    const docsOverlay = document.getElementById('docsOverlay');
    if (docsOverlay && docsOverlay.style.display !== 'none') { closeDocs(); return; }
    if (autoMatchIsOpen()) { closeAutoMatch(); return; }
    const historyOverlay = document.getElementById('historyOverlay');
    if (historyOverlay && historyOverlay.style.display !== 'none') { closeHistory(); return; }
    if (document.querySelector('.bulk-bar.active')) return;
    setAllGroupsCollapsed(true);
  }
});

// Enter in the search box: if the query identifies exactly one BL (an
// exact number, or the only partial match), open its port and vessel,
// scroll it into view and flash it. The text stays selected so the next
// BL can be typed straight over it: type - Enter - type - Enter.
function onSearchKeydown(e) {
  const box = e.target;
  if (e.key === 'Escape') {
    e.preventDefault();
    e.stopPropagation();
    if (box.value) { box.value = ''; render(); } else { box.blur(); }
    return;
  }
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const q = box.value.trim().toLowerCase();
  if (!q) return;
  const opEl = document.getElementById('operatorFilter');
  const op = opEl ? opEl.value : '';
  const matches = records.filter(r => blMatches(r, q) && (!op || r.created_by === op));
  const exactFull = matches.filter(r => r.bl_number.toLowerCase() === q);
  const exact = matches.filter(r => blExact(r, q));
  const target = (exactFull.length === 1 ? exactFull[0] : null) || (exact.length === 1 ? exact[0] : null) || (matches.length === 1 ? matches[0] : null);
  if (!target) {
    showToast(matches.length ? t('many_bls_match', {n: matches.length}) : t('no_bl_match', {q: box.value.trim()}));
    return;
  }
  jumpToBl(target.bl_number);
  box.select();
}

function jumpToBl(bl) {
  const rec = records.find(r => r.bl_number === bl);
  if (!rec) return;
  const port = rec.port || 'Unassigned';
  const vessel = rec.vessel || 'Unassigned';
  if (rec.archived) {
    archivedSectionOpen = true;
    collapsedGroups['port:' + port + ':archived'] = false;
    collapsedGroups['vessel:' + port + ':' + vessel + ':archived'] = false;
  } else {
    selectedPortTab = port;
    collapsedGroups['port:' + port] = false;
    collapsedGroups['vessel:' + port + ':' + vessel] = false;
  }
  render();
  requestAnimationFrame(() => {
    const row = document.getElementById('row_' + cssEscape(bl));
    if (!row) return;
    row.scrollIntoView({ behavior: 'smooth', block: 'center' });
    row.classList.remove('row-flash');
    void row.offsetWidth;  // restart the animation if it's the same row again
    row.classList.add('row-flash');
  });
}

applyI18n();
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
