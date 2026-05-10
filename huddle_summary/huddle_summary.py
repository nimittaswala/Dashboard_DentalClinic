"""
Open Dental - Daily Huddle Summary
------------------------------------
Sends office manager an 8AM morning summary with:
- Today's schedule overview per provider
- New patients today
- Unconfirmed appointments needing follow-up calls
- Yesterday's no-shows to reschedule
- Insurance verification status for today's appointments

Runs daily at 8AM via Windows Task Scheduler
Reads all config from office_config.py
"""

import mysql.connector
import smtplib
import logging
import os
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta

# =============================================================================
# CONFIG
# =============================================================================

sys.path.insert(0, r"C:\dental_automation\config")

try:
    from office_config import (
        OFFICE, DB_CONFIG, GMAIL_CONFIG,
        MANAGER_EMAIL, FRONT_DESK_EMAIL, FOLDERS, INS_VERIFY_SETTINGS
    )
except ImportError as e:
    print(f"ERROR: Could not load office_config.py: {e}")
    sys.exit(1)

PRACTICE_NAME = OFFICE["name"]
LOG_FILE      = os.path.join(FOLDERS["huddle"], "huddle_summary.log")

CONFIRMED_STATUSES  = ("Confirmed", "Ready", "In Room", "Arrived")
NEEDS_CALL_STATUSES = ("Unconfrm", "LeftMsg", "MsgFam")
GREEN_DAYS          = INS_VERIFY_SETTINGS["green_days"]
YELLOW_DAYS         = INS_VERIFY_SETTINGS["yellow_days"]

# =============================================================================
# LOGGING
# =============================================================================

os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log(msg):
    """Log to file with full unicode, print to console with ascii fallback."""
    logging.info(msg)
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode("ascii"))


# =============================================================================
# QUERIES
# =============================================================================

QUERY_TODAY_SCHEDULE = """
SELECT
    a.AptNum,
    a.AptDateTime,
    a.Confirmed,
    a.ProvNum,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName,
    COALESCE(NULLIF(pat.WirelessPhone,''), NULLIF(pat.HmPhone,''), NULLIF(pat.WkPhone,'')) AS Phone,
    pat.DateFirstVisit,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName,
    op.OpName
FROM appointment a
INNER JOIN patient pat ON pat.PatNum = a.PatNum
INNER JOIN provider prov ON prov.ProvNum = a.ProvNum
LEFT JOIN operatory op ON op.OperatoryNum = a.Op
WHERE a.AptStatus IN (1, 2)
AND DATE(a.AptDateTime) = %s
ORDER BY a.AptDateTime ASC;
"""

QUERY_NOSHOWS = """
SELECT
    a.AptNum,
    a.AptDateTime,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName,
    COALESCE(NULLIF(pat.WirelessPhone,''), NULLIF(pat.HmPhone,''), NULLIF(pat.WkPhone,'')) AS Phone,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName
FROM appointment a
INNER JOIN patient pat ON pat.PatNum = a.PatNum
INNER JOIN provider prov ON prov.ProvNum = a.ProvNum
WHERE a.AptStatus IN (3, 5)
AND DATE(a.AptDateTime) = %s
ORDER BY a.AptDateTime ASC;
"""

QUERY_INS_VERIFICATION = """
SELECT
    a.AptNum,
    a.AptDateTime,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName,
    COALESCE(NULLIF(pat.WirelessPhone,''), NULLIF(pat.HmPhone,''), NULLIF(pat.WkPhone,'')) AS Phone,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName,
    c.CarrierName,
    is2.SubscriberID,
    iv.DateLastVerified
FROM appointment a
INNER JOIN patient pat ON pat.PatNum = a.PatNum
INNER JOIN provider prov ON prov.ProvNum = a.ProvNum
INNER JOIN patplan pp ON pp.PatNum = a.PatNum AND pp.IsPending = 0 AND pp.Ordinal = 1
INNER JOIN inssub is2 ON is2.InsSubNum = pp.InsSubNum
INNER JOIN insplan ip ON ip.PlanNum = is2.PlanNum
INNER JOIN carrier c ON c.CarrierNum = ip.CarrierNum
LEFT JOIN (
            SELECT pp_inner.InsSubNum,
                MAX(iv_all.DateLastVerified) AS DateLastVerified
            FROM patplan pp_inner
            INNER JOIN (
                SELECT FKey, VerifyType, DateLastVerified FROM insverify
                WHERE DateLastVerified > '0001-01-01'
                UNION ALL
                SELECT FKey, VerifyType, DateLastVerified FROM insverifyhist
                WHERE DateLastVerified > '0001-01-01'
            ) iv_all ON (
                (iv_all.VerifyType = 1 AND iv_all.FKey = pp_inner.InsSubNum)
                OR
                (iv_all.VerifyType = 2 AND iv_all.FKey = pp_inner.PatPlanNum)
            )
            GROUP BY pp_inner.InsSubNum
        ) iv ON iv.InsSubNum = pp.InsSubNum
WHERE a.AptStatus = 1
AND DATE(a.AptDateTime) = %s
GROUP BY a.AptNum, a.AptDateTime, PatientName, Phone,
         ProviderName, c.CarrierName, is2.SubscriberID, iv.DateLastVerified
ORDER BY a.AptDateTime ASC;
"""

# Routing slip category — dynamic lookup so it works on any server
QUERY_ROUTING_SLIP_DEFNUM = """
SELECT DefNum FROM definition
WHERE LOWER(ItemName) LIKE '%routing%'
AND Category = 18
LIMIT 1;
"""

# Yesterday's completed appointments for routing slip check
QUERY_COMPLETED_YESTERDAY = """
SELECT DISTINCT
    a.PatNum,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName,
    DATE_FORMAT(a.AptDateTime, '%h:%i %p') AS AptTime,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName
FROM appointment a
INNER JOIN patient pat ON pat.PatNum = a.PatNum
INNER JOIN provider prov ON prov.ProvNum = a.ProvNum
WHERE a.AptStatus = 2
AND DATE(a.AptDateTime) = %s
ORDER BY a.AptDateTime;
"""

# =============================================================================
# DATABASE
# =============================================================================

def get_db_connection():
    return mysql.connector.connect(**DB_CONFIG)

# =============================================================================
# HELPERS
# =============================================================================

def get_verify_status(date_last_verified, today):
    """Return (status_label, color, days_since) for a verification date."""
    if not date_last_verified or str(date_last_verified).startswith("0001-01-01"):
        return "Never Verified", "#e74c3c", None
    # Normalise to date — MySQL can return date, datetime, or str
    try:
        if hasattr(date_last_verified, "date"):
            date_last_verified = date_last_verified.date()
        elif isinstance(date_last_verified, str):
            from datetime import datetime as _dt2
            date_last_verified = _dt2.strptime(date_last_verified[:10], "%Y-%m-%d").date()
    except Exception:
        return "Never Verified", "#e74c3c", None
    days = (today - date_last_verified).days
    if days <= GREEN_DAYS:
        return f"Verified ({days}d ago)", "#27ae60", days
    elif days <= YELLOW_DAYS:
        return f"Aging ({days}d ago)", "#e67e22", days
    else:
        return f"Overdue ({days}d ago)", "#e74c3c", days

# =============================================================================
# EMAIL BUILDER
# =============================================================================

def build_huddle_html(today_str, schedule, noshows, ins_verifications, today, missing_slips=None):

    from collections import defaultdict

    total_patients = len(schedule)
    today_str_short = today.strftime("%Y-%m-%d")
    new_patients   = [a for a in schedule
                      if a["DateFirstVisit"] and
                      str(a["DateFirstVisit"]) == today_str_short]
    unconfirmed    = [a for a in schedule if a["Confirmed"] in NEEDS_CALL_STATUSES]

    # Provider counts
    prov_counts = defaultdict(int)
    for a in schedule:
        prov_counts[a["ProviderName"]] += 1

    # Filter insurance — only yellow and red
    ins_needs_action = []
    for row in ins_verifications:
        label, color, days = get_verify_status(row["DateLastVerified"], today)
        if color != "#27ae60":  # exclude green
            ins_needs_action.append({**row, "verify_label": label, "verify_color": color})

    # ---- Schedule Section ----
    prov_rows = ""
    for prov, count in sorted(prov_counts.items(), key=lambda x: -x[1]):
        prov_rows += f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{prov}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee; text-align:center;">
                <strong>{count}</strong>
            </td>
        </tr>"""

    schedule_section = f"""
    <div style="padding:24px 32px 16px;">
        <h3 style="color:#2c3e50; margin-top:0; border-bottom:2px solid #eee; padding-bottom:8px;">
            [RECALL] Today's Schedule
        </h3>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
            <thead>
                <tr style="background:#f0f4f8;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Provider</th>
                    <th style="padding:8px 14px; text-align:center; font-size:13px; color:#666;">Patients</th>
                </tr>
            </thead>
            <tbody>{prov_rows}</tbody>
        </table>
        <div style="margin-top:12px; padding:10px 14px; background:#f8f9fa;
                    border-radius:4px; font-size:14px;">
            <strong>Total: {total_patients} patients</strong>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            [NEW] New: <strong>{len(new_patients)}</strong>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            [WARN] Unconfirmed: <strong>{len(unconfirmed)}</strong>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            [FAIL] No-Shows Yesterday: <strong>{len(noshows)}</strong>
        </div>
    </div>"""

    # ---- New Patients ----
    new_pat_section = ""
    if new_patients:
        rows = "".join([f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">
                {a['AptDateTime'].strftime('%I:%M %p')}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['PatientName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['ProviderName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['Phone'] or '—'}</td>
        </tr>""" for a in new_patients])
        new_pat_section = f"""
        <div style="padding:0 32px 16px;">
            <h3 style="color:#27ae60; border-bottom:2px solid #eee; padding-bottom:8px;">
                [NEW] New Patients Today ({len(new_patients)})
            </h3>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#f0fdf4;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Provider</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Phone</th>
                </tr></thead>
                <tbody>{rows}</tbody>
            </table>
        </div>"""

    # ---- Unconfirmed ----
    unconf_section = ""
    if unconfirmed:
        rows = "".join([f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">
                {a['AptDateTime'].strftime('%I:%M %p')}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['PatientName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['ProviderName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">
                {a['Phone'] or '<span style="color:#e74c3c;">[WARN] No phone</span>'}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee; color:#e67e22; font-size:13px;">
                {a['Confirmed']}</td>
        </tr>""" for a in unconfirmed])
        unconf_section = f"""
        <div style="padding:0 32px 16px;">
            <h3 style="color:#e67e22; border-bottom:2px solid #eee; padding-bottom:8px;">
                [WARN] Unconfirmed Appointments ({len(unconfirmed)}) — Call Now
            </h3>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#fffbe6;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Provider</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Phone</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Status</th>
                </tr></thead>
                <tbody>{rows}</tbody>
            </table>
        </div>"""

    # ---- No-Shows ----
    noshow_section = ""
    if noshows:
        rows = "".join([f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">
                {a['AptDateTime'].strftime('%I:%M %p')}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['PatientName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['ProviderName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['Phone'] or '—'}</td>
        </tr>""" for a in noshows])
        noshow_section = f"""
        <div style="padding:0 32px 16px;">
            <h3 style="color:#c0392b; border-bottom:2px solid #eee; padding-bottom:8px;">
                [FAIL] Yesterday's No-Shows / Broken ({len(noshows)}) — Reschedule
            </h3>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#fdf2f2;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Provider</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Phone</th>
                </tr></thead>
                <tbody>{rows}</tbody>
            </table>
        </div>"""

    # ---- Insurance Verification ----
    ins_section = ""
    if ins_needs_action:
        rows = "".join([f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">
                {a['AptDateTime'].strftime('%I:%M %p')}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['PatientName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{a['ProviderName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee; font-size:13px;">
                {a['CarrierName']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee; font-size:13px;
                       color:{a['verify_color']}; font-weight:bold;">
                {a['verify_label']}</td>
        </tr>""" for a in ins_needs_action])
        ins_section = f"""
        <div style="padding:0 32px 16px;">
            <h3 style="color:#8e44ad; border-bottom:2px solid #eee; padding-bottom:8px;">
                [VERIFY] Insurance Verification — Today's Appointments ({len(ins_needs_action)} need action)
            </h3>
            <p style="color:#666; font-size:13px; margin-top:-8px;">
                🟡 Yellow = verified 8-15 days ago &nbsp;|&nbsp;
                🔴 Red = overdue 15+ days or never verified
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#f5f0ff;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Provider</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Insurance</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Status</th>
                </tr></thead>
                <tbody>{rows}</tbody>
            </table>
        </div>"""
    else:
        ins_section = """
        <div style="padding:0 32px 16px;">
            <h3 style="color:#27ae60; border-bottom:2px solid #eee; padding-bottom:8px;">
                [VERIFY] Insurance Verification — [OK] All patients verified
            </h3>
        </div>"""

    # ---- Missing Routing Slips (Yesterday) ----
    missing_slips = missing_slips or []
    slip_section = ""
    if missing_slips:
        rows = "".join([f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{s['time']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{s['patient']}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #eee;">{s['provider']}</td>
        </tr>""" for s in missing_slips])
        slip_section = f"""
        <div style="padding:0 32px 16px;">
            <h3 style="color:#d97706; border-bottom:2px solid #eee; padding-bottom:8px;">
                &#128203; Yesterday's Missing Routing Slips ({len(missing_slips)}) — Scan Before EOD
            </h3>
            <p style="color:#666; font-size:13px; margin:-6px 0 10px;">
                These patients were seen yesterday but no routing slip was found in Open Dental.
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#fffbeb;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Time</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Patient</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#666;">Provider</th>
                </tr></thead>
                <tbody>{rows}</tbody>
            </table>
        </div>"""

    # ---- Assemble ----
    html = f"""
    <!DOCTYPE html>
    <html>
    <body style="font-family:Arial,sans-serif; margin:0; padding:0; background:#f5f5f5;">
    <div style="max-width:750px; margin:20px auto; background:white; border-radius:8px;
                overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.1);">

        <!-- Header -->
        <div style="background:#1a252f; padding:24px 32px;">
            <div style="color:#bdc3c7; font-size:13px;">{PRACTICE_NAME}</div>
            <div style="color:white; font-size:24px; font-weight:bold; margin-top:4px;">
                [HUDDLE] Daily Huddle Summary
            </div>
            <div style="color:#95a5a6; font-size:13px; margin-top:4px;">{today_str}</div>
        </div>

        <!-- Quick Banner -->
        <div style="background:#2c3e50; padding:12px 32px; font-size:14px;">
            <span style="color:white;">[PATIENTS] <strong>{total_patients}</strong> Patients</span>
            <span style="color:#2ecc71; margin-left:20px;">[NEW] <strong>{len(new_patients)}</strong> New</span>
            <span style="color:#f39c12; margin-left:20px;">[WARN] <strong>{len(unconfirmed)}</strong> Unconfirmed</span>
            <span style="color:#e74c3c; margin-left:20px;">[FAIL] <strong>{len(noshows)}</strong> No-Shows</span>
            <span style="color:#bb8fce; margin-left:20px;">[VERIFY] <strong>{len(ins_needs_action)}</strong> Ins. Verify</span>
            {f'<span style="color:#fde68a; margin-left:20px;">&#128203; <strong>{len(missing_slips)}</strong> Missing Slips</span>' if missing_slips else ''}
        </div>

        {schedule_section}
        {new_pat_section}
        {unconf_section}
        {noshow_section}
        {ins_section}
        {slip_section}

        <!-- Footer -->
        <div style="background:#f8f9fa; padding:16px 32px; border-top:1px solid #eee;">
            <p style="color:#aaa; font-size:12px; margin:0;">
                Automated daily huddle report from {PRACTICE_NAME}. Do not reply to this email.
            </p>
        </div>
    </div>
    </body>
    </html>"""

    return html

# =============================================================================
# EMAIL SENDER
# =============================================================================

def send_email(to_emails, subject, html_body):
    if isinstance(to_emails, str):
        to_emails = [to_emails]
    to_emails = [e for e in to_emails if e]  # drop blanks
    if not to_emails:
        return
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_CONFIG["sender_email"]
    msg["To"]      = ", ".join(to_emails)
    msg.attach(MIMEText(html_body, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_CONFIG["sender_email"], GMAIL_CONFIG["sender_password"])
        server.sendmail(GMAIL_CONFIG["sender_email"], to_emails, msg.as_string())

# =============================================================================
# MAIN
# =============================================================================

def main():
    today     = datetime.now().date()
    yesterday = today - timedelta(days=1)
    today_str = datetime.now().strftime("%A, %B %d, %Y")

    log(f"=== Daily Huddle Started — {today_str} ===")

    if not MANAGER_EMAIL:
        log("ERROR: MANAGER_EMAIL not set in office_config.py")
        return

    try:
        conn   = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        log("Database connected.")
    except Exception as e:
        log(f"ERROR: DB connection failed: {e}")
        return

    try:
        cursor.execute(QUERY_TODAY_SCHEDULE, (today,))
        schedule = cursor.fetchall()
        log(f"Schedule: {len(schedule)} appointments")

        cursor.execute(QUERY_NOSHOWS, (yesterday,))
        noshows = cursor.fetchall()
        log(f"No-shows yesterday: {len(noshows)}")

        cursor.execute(QUERY_INS_VERIFICATION, (today,))
        ins_verifications = cursor.fetchall()
        log(f"Insurance records fetched: {len(ins_verifications)}")

        # Routing slip check for yesterday
        missing_slips = []
        try:
            cursor.execute(QUERY_ROUTING_SLIP_DEFNUM)
            cat_row     = cursor.fetchone()
            routing_cat = cat_row["DefNum"] if cat_row else None

            cursor.execute(QUERY_COMPLETED_YESTERDAY, (yesterday,))
            completed_yesterday = cursor.fetchall()

            if routing_cat and completed_yesterday:
                pat_nums = list(set(p["PatNum"] for p in completed_yesterday))
                fmt      = ",".join(["%s"] * len(pat_nums))
                cursor.execute(
                    f"SELECT DISTINCT PatNum FROM document "
                    f"WHERE DocCategory = %s AND DATE(DateCreated) = %s "
                    f"AND PatNum IN ({fmt})",
                    [routing_cat, yesterday] + pat_nums
                )
                slipped = set(r["PatNum"] for r in cursor.fetchall())
                missing_slips = [
                    {"time": p["AptTime"].lstrip("0"), "patient": p["PatientName"], "provider": p["ProviderName"]}
                    for p in completed_yesterday if p["PatNum"] not in slipped
                ]
            log(f"Missing routing slips yesterday: {len(missing_slips)}")
        except Exception as e:
            log(f"WARNING: Routing slip check failed: {e}")

    except Exception as e:
        log(f"ERROR: Query failed: {e}")
        import traceback
        log(traceback.format_exc())
        cursor.close()
        conn.close()
        return

    html_body = build_huddle_html(today_str, schedule, noshows, ins_verifications, today, missing_slips=missing_slips)

    # Build subject
    unconfirmed_count = len([a for a in schedule if a["Confirmed"] in NEEDS_CALL_STATUSES])
    def _dlv_days(dlv):
        if not dlv or str(dlv).startswith("0001-01-01"):
            return 9999
        try:
            if hasattr(dlv, "date"):
                d = dlv.date()
            elif isinstance(dlv, str):
                from datetime import datetime as _dt2
                d = _dt2.strptime(dlv[:10], "%Y-%m-%d").date()
            else:
                d = dlv
            return (today - d).days
        except Exception:
            return 9999
    ins_needs_action  = sum(1 for r in ins_verifications
                            if _dlv_days(r["DateLastVerified"]) > GREEN_DAYS)
    alerts = []
    if unconfirmed_count:        alerts.append(f"[WARN] {unconfirmed_count} Unconfirmed")
    if len(noshows):             alerts.append(f"[FAIL] {len(noshows)} No-Shows")
    if ins_needs_action:         alerts.append(f"[VERIFY] {ins_needs_action} Ins. Verify")
    if len(missing_slips):       alerts.append(f"📋 {len(missing_slips)} Missing Slips")
    alert_str = " | ".join(alerts) if alerts else "[OK] All Clear"
    subject   = f"🌅 Huddle {today_str} — {len(schedule)} Patients | {alert_str}"

    try:
        recipients = list(dict.fromkeys([MANAGER_EMAIL, FRONT_DESK_EMAIL]))  # dedup
        send_email(recipients, subject, html_body)
        log(f"✅ Huddle email sent to {', '.join(r for r in recipients if r)}")
    except Exception as e:
        log(f"❌ Failed to send: {e}")
        import traceback
        log(traceback.format_exc())

    cursor.close()
    conn.close()
    log("=== Done ===\n")

if __name__ == "__main__":
    main()
