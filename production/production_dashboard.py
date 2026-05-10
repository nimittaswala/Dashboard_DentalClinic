"""
Open Dental - Daily Production Dashboard
-----------------------------------------
Sends office manager a 4:30PM end-of-day production report with:
- Total patients seen today vs same day last week
- Procedures completed today per provider
- Patient payments collected today by method
- Insurance payments received today
- Claims submitted today
- Print-friendly layout

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
from collections import defaultdict

# =============================================================================
# CONFIG
# =============================================================================

sys.path.insert(0, r"C:\dental_automation\config")

try:
    from office_config import (
        OFFICE, DB_CONFIG, GMAIL_CONFIG,
        MANAGER_EMAIL, FOLDERS,
        PAYMENT_TYPES, PAYMENT_GROUPS,
        PRODUCTION_SETTINGS
    )
except ImportError as e:
    print(f"ERROR: Could not load office_config.py: {e}")
    sys.exit(1)

PRACTICE_NAME = OFFICE["name"]
LOG_FILE      = os.path.join(FOLDERS["production"], "production.log")

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

# Appointments seen today (Scheduled or Complete, past only)
QUERY_PATIENTS_SEEN = """
SELECT
    a.AptNum,
    a.ProvNum,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName
FROM appointment a
INNER JOIN provider prov ON prov.ProvNum = a.ProvNum
INNER JOIN patient pat ON pat.PatNum = a.PatNum
WHERE a.AptStatus IN (1, 2)
AND DATE(a.AptDateTime) = %s;
"""

# Week-over-week patient counts by date range
QUERY_PATIENTS_DATE_RANGE = """
SELECT COUNT(DISTINCT a.AptNum) AS PatientCount
FROM appointment a
WHERE a.AptStatus IN (1, 2)
AND DATE(a.AptDateTime) >= %s
AND DATE(a.AptDateTime) <= %s;
"""

# Week-over-week collections by date range
QUERY_COLLECTIONS_DATE_RANGE = """
SELECT
    COALESCE(SUM(ps.SplitAmt), 0) AS TotalCollected,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName
FROM payment p
INNER JOIN paysplit ps ON ps.PayNum = p.PayNum
INNER JOIN provider prov ON prov.ProvNum = ps.ProvNum
WHERE DATE(p.PayDate) >= %s
AND DATE(p.PayDate) <= %s
AND p.PayAmt > 0
AND ps.SplitAmt > 0
AND ps.IsDiscount = 0
GROUP BY ProviderName
ORDER BY TotalCollected DESC;
"""

# Procedures completed today per provider
QUERY_PROCEDURES_TODAY = """
SELECT
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName,
    pc.ProcCode,
    pc.Descript,
    COUNT(*) AS ProcCount
FROM procedurelog pl
INNER JOIN provider prov ON prov.ProvNum = pl.ProvNum
INNER JOIN procedurecode pc ON pc.CodeNum = pl.CodeNum
WHERE pl.ProcStatus = 2
AND DATE(pl.ProcDate) = %s
GROUP BY ProviderName, pc.ProcCode, pc.Descript
ORDER BY ProviderName, ProcCount DESC;
"""

# Patient payments today with provider split breakdown
QUERY_PATIENT_PAYMENTS = """
SELECT
    p.PayNum,
    p.PayAmt,
    p.PayType,
    p.PatNum,
    ps.SplitAmt,
    ps.ProvNum,
    CONCAT(prov.FName, ' ', prov.LName) AS ProviderName,
    CONCAT(pat.LName, ', ', pat.FName) AS PatientName
FROM payment p
INNER JOIN paysplit ps ON ps.PayNum = p.PayNum
INNER JOIN patient pat ON pat.PatNum = p.PatNum
LEFT JOIN provider prov ON prov.ProvNum = ps.ProvNum
WHERE DATE(p.PayDate) = %s
AND p.PayAmt > 0
AND ps.SplitAmt > 0
AND ps.IsDiscount = 0;
"""

# Insurance payments today
QUERY_INSURANCE_PAYMENTS = """
SELECT
    cp.ClaimPaymentNum,
    cp.CheckAmt,
    cp.CarrierName,
    cp.CheckNum
FROM claimpayment cp
WHERE DATE(cp.CheckDate) = %s
AND cp.CheckAmt > 0;
"""

# Claims submitted today
QUERY_CLAIMS_TODAY = """
SELECT
    COUNT(*) AS ClaimCount,
    SUM(cl.ClaimFee) AS TotalBilled
FROM claim cl
WHERE DATE(cl.DateSent) = %s
AND cl.ClaimStatus IN ('S', 'W');
"""

# Same day last week comparisons
QUERY_PATIENTS_LAST_WEEK = """
SELECT COUNT(DISTINCT AptNum) AS PatientCount
FROM appointment
WHERE AptStatus = 2
AND DATE(AptDateTime) = %s;
"""

QUERY_PAYMENTS_LAST_WEEK = """
SELECT COALESCE(SUM(PayAmt), 0) AS TotalPaid
FROM payment
WHERE DATE(PayDate) = %s
AND PayAmt > 0;
"""

# Routing slip category — looked up dynamically so it works on any server
QUERY_ROUTING_SLIP_DEFNUM = """
SELECT DefNum FROM definition
WHERE LOWER(ItemName) LIKE '%routing%'
AND Category = 18
LIMIT 1;
"""

# Completed appointments today with patient/provider info for slip check
QUERY_COMPLETED_TODAY = """
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
# DATA PROCESSING
# =============================================================================

def get_payment_group(pay_type):
    """Return payment group name for a given PayType ID."""
    for group_name, type_ids in PAYMENT_GROUPS.items():
        if pay_type in type_ids:
            return group_name
    return PAYMENT_TYPES.get(pay_type, f"Other (Type {pay_type})")

def group_payments_by_method(payments):
    """Group payments by method — office total using SplitAmt to avoid double counting."""
    # Use unique PayNums to avoid counting same payment multiple times
    seen_pays = {}
    for p in payments:
        pay_num = p["PayNum"]
        if pay_num not in seen_pays:
            seen_pays[pay_num] = {"PayType": p["PayType"], "PayAmt": float(p["PayAmt"])}

    grouped = defaultdict(float)
    for pay in seen_pays.values():
        group = get_payment_group(pay["PayType"])
        if group != "Insurance Check":
            grouped[group] += pay["PayAmt"]

    refunds = grouped.pop("Patient Refund", 0)
    return dict(grouped), refunds

def group_payments_by_provider(payments):
    """Group split amounts by provider and payment method."""
    # prov_name -> { method -> amount }
    prov_pay = defaultdict(lambda: defaultdict(float))
    prov_total = defaultdict(float)

    for p in payments:
        prov = p["ProviderName"] or "Unassigned"
        group = get_payment_group(p["PayType"])
        if group == "Insurance Check":
            continue
        if group == "Patient Refund":
            continue
        amt = float(p["SplitAmt"])
        prov_pay[prov][group] += amt
        prov_total[prov] += amt

    return dict(prov_pay), dict(prov_total)

def group_procs_by_provider(procedures):
    """Group procedures by provider."""
    prov_procs = defaultdict(list)
    for p in procedures:
        prov_procs[p["ProviderName"]].append(p)
    return dict(prov_procs)

def group_patients_by_provider(patients):
    """Count patients per provider."""
    prov_counts = defaultdict(int)
    for p in patients:
        prov_counts[p["ProviderName"]] += 1
    return dict(prov_counts)

# =============================================================================
# EMAIL BUILDER
# =============================================================================

def build_production_html(today_str, today,
                           patients, procedures,
                           patient_payments, ins_payments,
                           claims, last_week_patients, last_week_payment_total,
                           this_week_patients=0, last_week_patients_wow=0,
                           this_week_collections=None, last_week_collections=None,
                           this_week_total=0, last_week_total=0,
                           week_start=None, lw_start=None, lw_end=None,
                           missing_slips=None, total_completed=0):

    # --- Summaries ---
    total_patients     = len(patients)
    last_week_pts      = last_week_patients
    patient_diff       = total_patients - last_week_pts
    patient_diff_str   = (f'<span style="color:#2ecc71;">▲ {patient_diff}</span>'
                          if patient_diff > 0 else
                          f'<span style="color:#e74c3c;">▼ {abs(patient_diff)}</span>'
                          if patient_diff < 0 else
                          '<span style="color:#95a5a6;">—</span>')

    total_procs        = sum(p["ProcCount"] for p in procedures)
    prov_patient_map   = group_patients_by_provider(patients)
    prov_proc_map      = group_procs_by_provider(procedures)

    pay_grouped, refunds   = group_payments_by_method(patient_payments)
    prov_pay, prov_total   = group_payments_by_provider(patient_payments)
    total_patient_pay      = sum(pay_grouped.values())
    total_ins_pay          = sum(float(p["CheckAmt"]) for p in ins_payments)
    total_collected        = total_patient_pay + total_ins_pay

    last_week_pay_diff = total_patient_pay - float(last_week_payment_total)
    pay_diff_str       = (f'<span style="color:#2ecc71;">+${last_week_pay_diff:,.2f}</span>'
                          if last_week_pay_diff > 0 else
                          f'<span style="color:#e74c3c;">-${abs(last_week_pay_diff):,.2f}</span>'
                          if last_week_pay_diff < 0 else
                          '<span style="color:#95a5a6;">same</span>')

    claim_count   = claims["ClaimCount"] if claims else 0
    claim_billed  = float(claims["TotalBilled"]) if claims and claims["TotalBilled"] else 0

    # --- Provider Summary Table ---
    all_provs = sorted(set(
        list(prov_patient_map.keys()) +
        list(prov_proc_map.keys()) +
        list(prov_total.keys())
    ))
    prov_rows = ""
    for prov in all_provs:
        pts        = prov_patient_map.get(prov, 0)
        procs      = sum(p["ProcCount"] for p in prov_proc_map.get(prov, []))
        collected  = prov_total.get(prov, 0)
        pay_detail = prov_pay.get(prov, {})
        # Build payment type mini-breakdown
        pay_mini   = " &nbsp;|&nbsp; ".join(
            [f"{k}: <strong>${v:,.2f}</strong>"
             for k, v in sorted(pay_detail.items(), key=lambda x: -x[1])]
        ) if pay_detail else "—"
        prov_rows += f"""
        <tr>
            <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a; color:#eee;
                       font-weight:600;">{prov}</td>
            <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a; text-align:center;
                       color:#eee;">{pts}</td>
            <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a; text-align:center;
                       color:#eee;">{procs}</td>
            <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a; text-align:right;
                       color:#2ecc71; font-weight:700;">${collected:,.2f}</td>
            <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                       color:#8b949e; font-size:11px;">{pay_mini}</td>
        </tr>"""

    # --- Payment Method Breakdown ---
    pay_rows = ""
    for method, amount in sorted(pay_grouped.items(), key=lambda x: -x[1]):
        pay_rows += f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; color:#eee;">{method}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; text-align:right;
                       color:#2ecc71; font-weight:bold;">${amount:,.2f}</td>
        </tr>"""
    if refunds > 0:
        pay_rows += f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; color:#e74c3c;">
                Patient Refunds</td>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; text-align:right;
                       color:#e74c3c;">-${refunds:,.2f}</td>
        </tr>"""

    # --- Insurance Payments ---
    ins_rows = ""
    for p in ins_payments:
        ins_rows += f"""
        <tr>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; color:#eee;">
                {p['CarrierName'] or 'Unknown Carrier'}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; color:#aaa;
                       font-size:12px;">{p['CheckNum'] or '—'}</td>
            <td style="padding:8px 14px; border-bottom:1px solid #2a2a2a; text-align:right;
                       color:#2ecc71; font-weight:bold;">${float(p['CheckAmt']):,.2f}</td>
        </tr>"""

    ins_note = ""
    if PRODUCTION_SETTINGS.get("insurance_payment_note"):
        ins_note = """
        <p style="color:#718096; font-size:12px; font-style:italic; margin-top:8px;">
            [WARN] Insurance payments are entered as received and may not reflect today's actual
            insurance receipts. ERA postings may be from earlier dates.
        </p>"""


    # --- Week Over Week calculations ---
    this_week_collections = this_week_collections or []
    last_week_collections  = last_week_collections or []
    tw_prov = {r["ProviderName"]: float(r["TotalCollected"]) for r in this_week_collections}
    lw_prov = {r["ProviderName"]: float(r["TotalCollected"]) for r in last_week_collections}
    all_wow_provs = sorted(set(list(tw_prov.keys()) + list(lw_prov.keys())))

    def wow_change(this_val, last_val):
        diff = this_val - last_val
        if diff > 0:   return f'<span style="color:#2ecc71">+${diff:,.2f}</span>'
        elif diff < 0: return f'<span style="color:#e74c3c">-${abs(diff):,.2f}</span>'
        return '<span style="color:#8b949e">same</span>'

    pts_diff    = this_week_patients - last_week_patients_wow
    pts_wow_str = (f'<span style="color:#2ecc71">+{pts_diff}</span>' if pts_diff > 0
                   else f'<span style="color:#e74c3c">{pts_diff}</span>' if pts_diff < 0
                   else '<span style="color:#8b949e">same</span>')
    col_wow_str = wow_change(this_week_total, last_week_total)

    wow_prov_rows = ""
    for prov in all_wow_provs:
        tw_amt = tw_prov.get(prov, 0)
        lw_amt = lw_prov.get(prov, 0)
        wow_prov_rows += (
            '<tr>' +
            f'<td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;color:#a0aec0;font-size:12px;">&nbsp;&nbsp;' + prov + '</td>' +
            f'<td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;text-align:center;color:#2ecc71;">${tw_amt:,.2f}</td>' +
            f'<td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;text-align:center;color:#8b949e;">${lw_amt:,.2f}</td>' +
            f'<td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;text-align:center;">' + wow_change(tw_amt, lw_amt) + '</td>' +
            '</tr>'
        )

    week_start_str = week_start.strftime("%b %d") if week_start else "Mon"
    lw_start_str   = lw_start.strftime("%b %d") if lw_start else "--"
    lw_end_str     = lw_end.strftime("%b %d, %Y") if lw_end else "--"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        @media print {{
            body {{ background: white !important; margin: 0; padding: 0; }}
            .no-print {{ display: none !important; }}
            .main-card {{
                box-shadow: none !important;
                border: none !important;
                max-width: 100% !important;
                margin: 0 !important;
                border-radius: 0 !important;
            }}
            * {{ color: black !important; background: white !important;
                 -webkit-print-color-adjust: exact; }}
            .header-bar {{ background: #f0f0f0 !important; padding: 12px 20px !important; }}
            .stats-bar {{ background: #e8e8e8 !important; padding: 10px 20px !important; }}
            table {{ width: 100% !important; border-collapse: collapse !important; }}
            th {{ background: #e0e0e0 !important; padding: 8px !important;
                  border: 1px solid #ccc !important; text-align: left; }}
            td {{ padding: 8px !important; border: 1px solid #ddd !important; }}
            tfoot td {{ background: #f0f0f0 !important; font-weight: bold; }}
            h3 {{ color: #333 !important; border-bottom: 2px solid #333 !important; }}
            .summary-box {{ background: #f5f5f5 !important;
                            border: 1px solid #ccc !important; }}
            @page {{ margin: 1.5cm; size: A4; }}
        }}
    </style>
    </head>
    <body style="font-family:Arial,sans-serif; margin:0; padding:0; background:#1a1a2e;">
    <div class="main-card" style="max-width:800px; margin:20px auto; background:#16213e;
         border-radius:8px; overflow:hidden; box-shadow:0 4px 16px rgba(0,0,0,0.4);">

        <!-- Header -->
        <div style="background:#0f3460; padding:24px 32px;">
            <div style="color:#a0aec0; font-size:13px;">{PRACTICE_NAME}</div>
            <div style="color:white; font-size:24px; font-weight:bold; margin-top:4px;">
                [PROD] Daily Production Report
            </div>
            <div style="color:#718096; font-size:13px; margin-top:4px;">{today_str}</div>
            <div class="no-print" style="margin-top:10px;">
                <span style="color:#a0aec0; font-size:12px; font-style:italic;">
                    To print or save as PDF &mdash; press <strong>Ctrl+P</strong> in your browser
                </span>
            </div>
        </div>

        <!-- Quick Stats -->
        <div style="background:#1a365d; padding:16px 32px; display:flex; flex-wrap:wrap; gap:24px;">
            <div style="color:white; font-size:15px;">
                [PATIENTS] Patients Seen: <strong>{total_patients}</strong>
                <span style="font-size:12px; margin-left:8px;">vs last week: {patient_diff_str}</span>
            </div>
            <div style="color:white; font-size:15px;">
                [PERIO] Procedures: <strong>{total_procs}</strong>
            </div>
            <div style="color:#2ecc71; font-size:15px;">
                [MONEY] Total Collected: <strong>${total_collected:,.2f}</strong>
            </div>
        </div>

        <!-- Provider Summary -->
        <div style="padding:24px 32px 16px;">
            <h3 style="color:#63b3ed; margin-top:0; border-bottom:2px solid #2a2a2a;
                       padding-bottom:8px;">[PROV] Provider Summary</h3>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#2a2a2a;">
                    <th style="padding:10px 14px; text-align:left; font-size:13px; color:#a0aec0;">
                        Provider</th>
                    <th style="padding:10px 14px; text-align:center; font-size:13px; color:#a0aec0;">
                        Patients</th>
                    <th style="padding:10px 14px; text-align:center; font-size:13px; color:#a0aec0;">
                        Procedures</th>
                    <th style="padding:10px 14px; text-align:right; font-size:13px; color:#a0aec0;">
                        Collected</th>
                    <th style="padding:10px 14px; text-align:left; font-size:13px; color:#a0aec0;">
                        Payment Types</th>
                </tr></thead>
                <tbody>{prov_rows}</tbody>
                <tfoot>
                    <tr style="background:#1a365d;">
                        <td style="padding:10px 14px; color:white; font-weight:bold;">TOTAL</td>
                        <td style="padding:10px 14px; text-align:center; color:white;
                                   font-weight:bold;">{total_patients}</td>
                        <td style="padding:10px 14px; text-align:center; color:white;
                                   font-weight:bold;">{total_procs}</td>
                        <td style="padding:10px 14px; text-align:right; color:#2ecc71;
                                   font-weight:bold;">${total_patient_pay:,.2f}</td>
                        <td style="padding:10px 14px; color:#a0aec0; font-size:12px;">
                            Patient payments total</td>
                    </tr>
                </tfoot>
            </table>
        </div>

        <!-- Patient Payments -->
        <div style="padding:0 32px 16px;">
            <h3 style="color:#63b3ed; border-bottom:2px solid #2a2a2a; padding-bottom:8px;">
                [PAY] Patient Payments — ${total_patient_pay:,.2f}
                <span style="font-size:13px; font-weight:normal; color:#718096; margin-left:8px;">
                    vs last week: {pay_diff_str}
                </span>
            </h3>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#2a2a2a;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#a0aec0;">
                        Payment Method</th>
                    <th style="padding:8px 14px; text-align:right; font-size:13px; color:#a0aec0;">
                        Amount</th>
                </tr></thead>
                <tbody>{pay_rows}</tbody>
                <tfoot>
                    <tr style="background:#1a365d;">
                        <td style="padding:10px 14px; color:white; font-weight:bold;">TOTAL</td>
                        <td style="padding:10px 14px; text-align:right; color:#2ecc71;
                                   font-weight:bold;">${total_patient_pay:,.2f}</td>
                    </tr>
                </tfoot>
            </table>
        </div>

        <!-- Insurance Payments -->
        <div style="padding:0 32px 16px;">
            <h3 style="color:#63b3ed; border-bottom:2px solid #2a2a2a; padding-bottom:8px;">
                [INS] Insurance Payments — ${total_ins_pay:,.2f}
            </h3>
            {'<p style="color:#718096; font-style:italic; font-size:13px;">No insurance payments entered today.</p>' if not ins_payments else f"""
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#2a2a2a;">
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#a0aec0;">Carrier</th>
                    <th style="padding:8px 14px; text-align:left; font-size:13px; color:#a0aec0;">Check #</th>
                    <th style="padding:8px 14px; text-align:right; font-size:13px; color:#a0aec0;">Amount</th>
                </tr></thead>
                <tbody>{ins_rows}</tbody>
                <tfoot>
                    <tr style="background:#1a365d;">
                        <td style="padding:10px 14px; color:white; font-weight:bold;" colspan="2">TOTAL</td>
                        <td style="padding:10px 14px; text-align:right; color:#2ecc71; font-weight:bold;">${total_ins_pay:,.2f}</td>
                    </tr>
                </tfoot>
            </table>"""}
            {ins_note}
        </div>

        <!-- Claims Submitted -->
        <div style="padding:0 32px 16px;">
            <h3 style="color:#63b3ed; border-bottom:2px solid #2a2a2a; padding-bottom:8px;">
                [NOTES] Claims Submitted Today
            </h3>
            <div style="background:#2a2a2a; padding:16px; border-radius:6px; color:#eee;">
                Claims sent: <strong>{claim_count}</strong>
                &nbsp;&nbsp;|&nbsp;&nbsp;
                Total billed: <strong>${claim_billed:,.2f}</strong>
            </div>
        </div>

        <!-- Routing Slips -->
        <div style="padding:0 32px 16px;">
            <h3 style="color:#63b3ed; border-bottom:2px solid #2a2a2a; padding-bottom:8px;">
                &#128203; Routing Slip Compliance
            </h3>
            __ROUTING_SLIP_BLOCK__
        </div>

        <!-- Total Summary Box -->
        <div style="padding:0 32px 24px;">
            <div style="background:#0f3460; padding:20px; border-radius:8px;
                        border:1px solid #2a6496;">
                <div style="color:#a0aec0; font-size:13px; margin-bottom:8px;">
                    END OF DAY SUMMARY</div>
                <div style="display:flex; flex-wrap:wrap; gap:24px;">
                    <div style="color:white;">
                        Patient Collections:
                        <strong style="color:#2ecc71;">${total_patient_pay:,.2f}</strong>
                    </div>
                    <div style="color:white;">
                        Insurance Collections:
                        <strong style="color:#2ecc71;">${total_ins_pay:,.2f}</strong>
                    </div>
                    <div style="color:white; font-size:16px;">
                        Total:
                        <strong style="color:#f1c40f;">${total_collected:,.2f}</strong>
                    </div>
                </div>
            </div>
        </div>

        <!-- Week Over Week Section -->
        <div style="padding:0 32px 24px;">
            <h3 style="color:#63b3ed; border-bottom:2px solid #2a2a2a; padding-bottom:8px;">
                📈 Week-Over-Week Comparison
            </h3>
            <p style="color:#718096; font-size:12px; margin-bottom:12px;">
                This week: {week_start_str} — Today &nbsp;|&nbsp;
                Last week: {lw_start_str} — {lw_end_str}
            </p>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#2a2a2a;">
                    <th style="padding:10px 14px; text-align:left; font-size:13px; color:#a0aec0;">Metric</th>
                    <th style="padding:10px 14px; text-align:center; font-size:13px; color:#a0aec0;">This Week</th>
                    <th style="padding:10px 14px; text-align:center; font-size:13px; color:#a0aec0;">Last Week</th>
                    <th style="padding:10px 14px; text-align:center; font-size:13px; color:#a0aec0;">Change</th>
                </tr></thead>
                <tbody>
                    <tr>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a; color:#eee;">
                            Patients Seen</td>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                                   text-align:center; color:#eee; font-weight:700;">
                            {this_week_patients}</td>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                                   text-align:center; color:#8b949e;">{last_week_patients_wow}</td>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                                   text-align:center; font-weight:700;">
                            {pts_wow_str}</td>
                    </tr>
                    <tr>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a; color:#eee;">
                            Patient Collections</td>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                                   text-align:center; color:#2ecc71; font-weight:700;">
                            ${this_week_total:,.2f}</td>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                                   text-align:center; color:#8b949e;">${last_week_total:,.2f}</td>
                        <td style="padding:10px 14px; border-bottom:1px solid #2a2a2a;
                                   text-align:center; font-weight:700;">
                            {col_wow_str}</td>
                    </tr>
                    {wow_prov_rows}
                </tbody>
            </table>
        </div>

        <!-- Footer -->
        <div style="background:#0f1923; padding:16px 32px; border-top:1px solid #2a2a2a;">
            <p style="color:#4a5568; font-size:12px; margin:0;">
                Automated daily production report from {PRACTICE_NAME}.
                Do not reply to this email.
            </p>
        </div>
    </div>
    </body>
    </html>"""

    # Build routing slip block and inject into HTML
    missing_slips  = missing_slips or []
    slips_scanned  = total_completed - len(missing_slips)
    slip_color     = "#2ecc71" if len(missing_slips) == 0 else "#f1c40f" if len(missing_slips) <= 3 else "#e74c3c"

    slip_block = f"""
            <div style="background:#2a2a2a;padding:14px 16px;border-radius:6px;margin-bottom:10px;
                        display:flex;gap:32px;flex-wrap:wrap;align-items:center;">
                <div style="color:white;font-size:14px;">
                    Scanned: <strong style="color:{slip_color};">{slips_scanned} / {total_completed}</strong>
                </div>"""

    if len(missing_slips) == 0:
        slip_block += """
                <div style="color:#2ecc71;font-size:13px;">&#10003; All routing slips accounted for</div>"""
    else:
        slip_block += f"""
                <div style="color:#e74c3c;font-size:13px;font-weight:600;">
                    &#9888; {len(missing_slips)} missing slip{'s' if len(missing_slips) != 1 else ''}
                </div>"""

    slip_block += "</div>"

    if missing_slips:
        slip_block += """
            <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
                <thead><tr style="background:#2a2a2a;">
                    <th style="padding:8px 14px;font-size:12px;color:#a0aec0;text-align:left;">Time</th>
                    <th style="padding:8px 14px;font-size:12px;color:#a0aec0;text-align:left;">Patient</th>
                    <th style="padding:8px 14px;font-size:12px;color:#a0aec0;text-align:left;">Provider</th>
                    <th style="padding:8px 14px;font-size:12px;color:#a0aec0;text-align:left;">Status</th>
                </tr></thead><tbody>"""
        for s in missing_slips:
            slip_block += f"""
                <tr>
                    <td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;color:#a0aec0;font-size:12px;">{s['time']}</td>
                    <td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;color:#eee;">{s['patient']}</td>
                    <td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;color:#a0aec0;font-size:12px;">{s['provider']}</td>
                    <td style="padding:8px 14px;border-bottom:1px solid #2a2a2a;">
                        <span style="padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;
                                     background:rgba(231,76,60,.15);color:#e74c3c;
                                     border:1px solid rgba(231,76,60,.3);">&#9888; Slip Missing</span>
                    </td>
                </tr>"""
        slip_block += "</tbody></table>"

    html = html.replace("__ROUTING_SLIP_BLOCK__", slip_block)
    return html

# =============================================================================
# EMAIL SENDER
# =============================================================================

def send_email(to_email, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_CONFIG["sender_email"]
    msg["To"]      = to_email
    msg.attach(MIMEText(html_body, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_CONFIG["sender_email"], GMAIL_CONFIG["sender_password"])
        server.sendmail(GMAIL_CONFIG["sender_email"], to_email, msg.as_string())

# =============================================================================
# MAIN
# =============================================================================

def main():
    today         = datetime.now().date()
    last_week     = today - timedelta(days=7)
    today_str     = datetime.now().strftime("%A, %B %d, %Y")

    log(f"=== Production Dashboard Started — {today_str} ===")

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
        cursor.execute(QUERY_PATIENTS_SEEN, (today,))
        patients = cursor.fetchall()

        cursor.execute(QUERY_PROCEDURES_TODAY, (today,))
        procedures = cursor.fetchall()

        cursor.execute(QUERY_PATIENT_PAYMENTS, (today,))
        patient_payments = cursor.fetchall()

        cursor.execute(QUERY_INSURANCE_PAYMENTS, (today,))
        ins_payments = cursor.fetchall()

        cursor.execute(QUERY_CLAIMS_TODAY, (today,))
        claims = cursor.fetchone()

        cursor.execute(QUERY_PATIENTS_LAST_WEEK, (last_week,))
        lw_pts = cursor.fetchone()
        last_week_patients = lw_pts["PatientCount"] if lw_pts else 0

        cursor.execute(QUERY_PAYMENTS_LAST_WEEK, (last_week,))
        lw_pay = cursor.fetchone()
        last_week_payment_total = lw_pay["TotalPaid"] if lw_pay else 0

        # Week-over-week data
        # This week: Monday to today
        today_dow   = today.weekday()  # 0=Monday
        week_start  = today - timedelta(days=today_dow)
        lw_start    = week_start - timedelta(days=7)
        lw_end      = lw_start + timedelta(days=today_dow)

        cursor.execute(QUERY_PATIENTS_DATE_RANGE, (week_start, today))
        tw_pts = cursor.fetchone()
        this_week_patients = tw_pts["PatientCount"] if tw_pts else 0

        cursor.execute(QUERY_PATIENTS_DATE_RANGE, (lw_start, lw_end))
        lw_pts2 = cursor.fetchone()
        last_week_patients_wow = lw_pts2["PatientCount"] if lw_pts2 else 0

        cursor.execute(QUERY_COLLECTIONS_DATE_RANGE, (week_start, today))
        this_week_collections = cursor.fetchall()

        cursor.execute(QUERY_COLLECTIONS_DATE_RANGE, (lw_start, lw_end))
        last_week_collections = cursor.fetchall()

        this_week_total = sum(float(r["TotalCollected"]) for r in this_week_collections)
        last_week_total = sum(float(r["TotalCollected"]) for r in last_week_collections)

        # Routing slip compliance
        missing_slips   = []
        total_completed = 0
        try:
            cursor.execute(QUERY_ROUTING_SLIP_DEFNUM)
            cat_row      = cursor.fetchone()
            routing_cat  = cat_row["DefNum"] if cat_row else None

            cursor.execute(QUERY_COMPLETED_TODAY, (today,))
            completed_pts = cursor.fetchall()
            total_completed = len(completed_pts)

            if routing_cat and completed_pts:
                pat_nums = list(set(p["PatNum"] for p in completed_pts))
                fmt      = ",".join(["%s"] * len(pat_nums))
                cursor.execute(
                    f"SELECT DISTINCT PatNum FROM document "
                    f"WHERE DocCategory = %s AND DATE(DateCreated) = %s "
                    f"AND PatNum IN ({fmt})",
                    [routing_cat, today] + pat_nums
                )
                slipped = set(r["PatNum"] for r in cursor.fetchall())
                missing_slips = [
                    {"time": p["AptTime"].lstrip("0"), "patient": p["PatientName"], "provider": p["ProviderName"]}
                    for p in completed_pts if p["PatNum"] not in slipped
                ]
            log(f"Routing slips: {total_completed - len(missing_slips)}/{total_completed} scanned, {len(missing_slips)} missing")
        except Exception as e:
            log(f"WARNING: Routing slip check failed: {e}")
        log(f"Procedures: {sum(p['ProcCount'] for p in procedures)}")
        log(f"Patient payments: {len(patient_payments)}")
        log(f"Insurance payments: {len(ins_payments)}")
        log(f"This week collections: ${this_week_total:,.2f} vs last week: ${last_week_total:,.2f}")

    except Exception as e:
        log(f"ERROR: Query failed: {e}")
        import traceback
        log(traceback.format_exc())
        cursor.close()
        conn.close()
        return

    # ── Skip email on closed-office days (Saturdays at offices that don't work) ──
    # Robust signal: no procedures with ProcStatus=2 today AND no appointments
    # marked AptStatus=2 today. Both being zero means no work happened.
    # We check both because some offices update AptStatus, some don't, but
    # ProcStatus gets updated whenever procedures are billed (universal).
    if len(procedures) == 0 and total_completed == 0:
        log(f"No completed procedures or appointments today ({today_str}) — "
            f"skipping email (office likely closed).")
        cursor.close()
        conn.close()
        return {
            "patients": 0,
            "total_collected": 0,
            "patient_pay": 0,
            "ins_pay": 0,
            "skipped": "no_production",
        }

    html_body = build_production_html(
        today_str, today,
        patients, procedures,
        patient_payments, ins_payments,
        claims, last_week_patients, last_week_payment_total,
        this_week_patients, last_week_patients_wow,
        this_week_collections, last_week_collections,
        this_week_total, last_week_total,
        week_start, lw_start, lw_end,
        missing_slips=missing_slips,
        total_completed=total_completed
    )

    pay_grouped, _ = group_payments_by_method(patient_payments)
    total_patient  = sum(pay_grouped.values())
    total_ins      = sum(float(p["CheckAmt"]) for p in ins_payments)
    total          = total_patient + total_ins
    subject        = (f"📊 Production {today_str} — "
                     f"{len(patients)} Patients | "
                     f"${total:,.2f} Collected")

    try:
        send_email(MANAGER_EMAIL, subject, html_body)
        log(f"✅ Production report sent to {MANAGER_EMAIL}")
        log(f"   Total collected: ${total:,.2f}")
    except Exception as e:
        log(f"❌ Failed to send: {e}")
        import traceback
        log(traceback.format_exc())

    cursor.close()
    conn.close()
    log("=== Done ===\n")
    return {
        "patients": len(patients),
        "total_collected": round(total, 2),
        "patient_pay": round(total_patient, 2),
        "ins_pay": round(total_ins, 2),
    }

# group_payments_by_method defined above

if __name__ == "__main__":
    main()
