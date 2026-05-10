"""
Dental AR Push — Pushes outstanding insurance AR data to central Google Sheet.
Runs every 2 hours 8am-6pm. One row per outstanding claim.
Overwrites today's data on each run to keep sheet clean.
"""

import sys, os, logging
from datetime import datetime, date, timedelta

sys.path.insert(0, r"C:\dental_automation\config")

try:
    from office_config import (
        OFFICE, DB_CONFIG, FOLDERS, AR_SETTINGS
    )
except ImportError as e:
    print(f"ERROR: Could not load office_config.py: {e}")
    sys.exit(1)

import mysql.connector
import gspread
from google.oauth2.service_account import Credentials

PRACTICE_NAME = OFFICE["name"]
CREDS_PATH    = os.path.join(FOLDERS["base"], "ar", "credentials.json")
SHEET_ID      = AR_SETTINGS["sheet_id"]
LOG_FILE      = os.path.join(FOLDERS["base"], "ar", "ar_push.log")

os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE, level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def log(msg):
    logging.info(msg)
    print(msg)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ── Queries ───────────────────────────────────────────────────────────────────

QUERY_OUTSTANDING_CLAIMS = """
SELECT
    c.ClaimNum,
    CONCAT(pat.LName, ', ', pat.FName)   AS PatientName,
    DATE_FORMAT(pat.Birthdate, '%m/%d/%Y') AS PatientDOB,
    (
        SELECT MIN(pl2.ProcDate)
        FROM claimproc cp2
        INNER JOIN procedurelog pl2 ON pl2.ProcNum = cp2.ProcNum
        WHERE cp2.ClaimNum = c.ClaimNum
    )                                    AS DateOfServiceMin,
    (
        SELECT MAX(pl2.ProcDate)
        FROM claimproc cp2
        INNER JOIN procedurelog pl2 ON pl2.ProcNum = cp2.ProcNum
        WHERE cp2.ClaimNum = c.ClaimNum
    )                                    AS DateOfServiceMax,
    DATE(c.DateSent)                     AS DateSent,
    DATE(c.DateSentOrig)                 AS DateSentOrig,
    ca.CarrierName,
    ca.ElectID                           AS PayorID,
    ip.PlanType,
    CONCAT(prov.FName, ' ', prov.LName)  AS ProviderName,
    COALESCE(cp_billed.FeeBilled, 0)     AS Billed,
    COALESCE(cp_paid.InsPayAmt, 0)       AS InsPaid,
    COALESCE(cp_est.InsPayEst, 0)
        - COALESCE(cp_paid.InsPayAmt, 0) AS InsEstimate,
    COALESCE(cp_wr.WriteOff, 0)          AS WriteOff,
    c.ClaimStatus,
    c.ClaimType,
    DATEDIFF(CURDATE(),
        CASE
            WHEN DATE(c.DateSent)     > '2000-01-01' THEN DATE(c.DateSent)
            WHEN DATE(c.DateSentOrig) > '2000-01-01' THEN DATE(c.DateSentOrig)
            ELSE NULL
        END
    )                                    AS AgeDays,
    (
        SELECT COUNT(*) FROM claimtracking ct
        WHERE ct.ClaimNum = c.ClaimNum
    )                                    AS ResubmitCount,
    (
        SELECT GROUP_CONCAT(DISTINCT pc2.ProcCode ORDER BY pc2.ProcCode SEPARATOR ', ')
        FROM claimproc cp3
        INNER JOIN procedurelog pl3 ON pl3.ProcNum = cp3.ProcNum
        INNER JOIN procedurecode pc2 ON pc2.CodeNum = pl3.CodeNum
        WHERE cp3.ClaimNum = c.ClaimNum
    )                                    AS ProcCodes
FROM claim c
INNER JOIN patient pat    ON pat.PatNum    = c.PatNum
INNER JOIN insplan ip     ON ip.PlanNum    = c.PlanNum
INNER JOIN carrier ca     ON ca.CarrierNum = ip.CarrierNum
INNER JOIN provider prov  ON prov.ProvNum  = c.ProvTreat
LEFT JOIN (
    SELECT ClaimNum, SUM(FeeBilled) AS FeeBilled
    FROM claimproc WHERE Status NOT IN (2,7)
    GROUP BY ClaimNum
) cp_billed ON cp_billed.ClaimNum = c.ClaimNum
LEFT JOIN (
    SELECT ClaimNum, SUM(InsPayAmt) AS InsPayAmt
    FROM claimproc WHERE Status = 1
    GROUP BY ClaimNum
) cp_paid ON cp_paid.ClaimNum = c.ClaimNum
LEFT JOIN (
    SELECT ClaimNum, SUM(InsPayEst) AS InsPayEst
    FROM claimproc WHERE Status NOT IN (2,7)
    GROUP BY ClaimNum
) cp_est ON cp_est.ClaimNum = c.ClaimNum
LEFT JOIN (
    SELECT ClaimNum, SUM(WriteOff) AS WriteOff
    FROM claimproc WHERE Status NOT IN (2,7)
    GROUP BY ClaimNum
) cp_wr ON cp_wr.ClaimNum = c.ClaimNum
WHERE c.ClaimStatus IN ('S', 'W', 'R', 'H')
AND c.ClaimType <> 'PreAuth'
AND COALESCE(cp_est.InsPayEst, 0)
    - COALESCE(cp_paid.InsPayAmt, 0) > 0
ORDER BY AgeDays DESC;
"""

QUERY_MEDICAID_CARRIERS = """
SELECT LOWER(ca.CarrierName) AS CarrierName
FROM insplan ip
INNER JOIN carrier ca ON ca.CarrierNum = ip.CarrierNum
WHERE ip.PlanNum IN (
    SELECT DISTINCT PlanNum FROM claim
    WHERE ClaimStatus IN ('S','W','R','H')
);
"""

def is_medicaid(carrier_name, medicaid_list):
    cn = carrier_name.lower()
    return any(m in cn for m in medicaid_list)

def connect_sheets():
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID)

def main():
    log(f"=== AR Push Started — {PRACTICE_NAME} — {datetime.now().strftime('%Y-%m-%d %H:%M')} ===")
    today     = date.today()
    today_str = today.isoformat()
    now_str   = datetime.now().strftime("%H:%M")

    try:
        # ── Connect to Open Dental ────────────────────────────────────────────
        conn   = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)
        log("Database connected.")

        cursor.execute(QUERY_OUTSTANDING_CLAIMS)
        claims = cursor.fetchall()
        log(f"Found {len(claims)} outstanding claims (PreAuths excluded, InsEstimate > 0).")

        # ── Get Medicaid carriers from config ─────────────────────────────────
        try:
            from office_config import MEDICAID_CARRIERS
            medicaid_list = MEDICAID_CARRIERS
        except Exception:
            medicaid_list = [
                "keystone first", "upmc", "united healthcare community plan",
                "health partners", "aetna better health", "dentaquest"
            ]

        # ── Date helper ───────────────────────────────────────────────────────
        def to_date(value):
            """Coerce date | datetime | str | None into a date or None."""
            if value is None or value == "":
                return None
            if isinstance(value, datetime):
                return value.date()
            if isinstance(value, date):
                return value
            if isinstance(value, str):
                try:
                    return datetime.strptime(value[:10], "%Y-%m-%d").date()
                except ValueError:
                    return None
            return None

        # ── Calculate summary stats (all based on InsEstimate) ────────────────
        bucket_0_30  = sum(float(c["InsEstimate"] or 0) for c in claims if (c["AgeDays"] or 0) <= 30)
        bucket_31_60 = sum(float(c["InsEstimate"] or 0) for c in claims if 31 <= (c["AgeDays"] or 0) <= 60)
        bucket_61_90 = sum(float(c["InsEstimate"] or 0) for c in claims if 61 <= (c["AgeDays"] or 0) <= 90)
        bucket_90p   = sum(float(c["InsEstimate"] or 0) for c in claims if (c["AgeDays"] or 0) > 90)
        total_ar     = sum(float(c["InsEstimate"] or 0) for c in claims)
        medicaid_ar  = sum(float(c["InsEstimate"] or 0) for c in claims if is_medicaid(c["CarrierName"] or "", medicaid_list))
        ppo_ar       = total_ar - medicaid_ar
        oldest_days  = max((c["AgeDays"] or 0) for c in claims) if claims else 0
        oldest_date  = min((c["DateSent"] for c in claims if c["DateSent"]), default="")

        cursor.close()
        conn.close()
        log("Database disconnected.")

        # ── Connect to Google Sheets ──────────────────────────────────────────
        log("Connecting to Google Sheets...")
        sh = connect_sheets()

        # ── Update Summary tab ────────────────────────────────────────────────
        summary_ws = sh.worksheet("Summary")
        all_rows   = summary_ws.get_all_values()

        # Find and remove today's existing rows for this office
        rows_to_keep = [all_rows[0]]  # keep header
        for row in all_rows[1:]:
            if not (len(row) >= 2 and row[0] == today_str and row[1] == PRACTICE_NAME):
                rows_to_keep.append(row)

        # Add today's summary row
        summary_row = [
            today_str, PRACTICE_NAME,
            round(bucket_0_30, 2), round(bucket_31_60, 2),
            round(bucket_61_90, 2), round(bucket_90p, 2),
            round(total_ar, 2), len(claims),
            oldest_days, str(oldest_date),
            round(medicaid_ar, 2), round(ppo_ar, 2),
            "", ""  # claims paid/new — calculated on next push
        ]
        rows_to_keep.append(summary_row)
        summary_ws.clear()
        summary_ws.update("A1", rows_to_keep)
        log(f"Summary updated: ${total_ar:,.2f} total AR")

        # ── Update Claims tab ─────────────────────────────────────────────────
        claims_ws   = sh.worksheet("Claims")
        all_claims  = claims_ws.get_all_values()

        # Remove today's existing claims for this office
        claims_keep = [all_claims[0]]  # keep header
        for row in all_claims[1:]:
            if not (len(row) >= 3 and row[0] == today_str and row[2] == PRACTICE_NAME):
                claims_keep.append(row)

        # Add new claim rows
        def dos_from(d_min):
            d_min = to_date(d_min)
            return d_min.isoformat() if d_min else ""

        def dos_to(d_min, d_max):
            """Only populated when the claim spans multiple service dates."""
            d_min = to_date(d_min)
            d_max = to_date(d_max)
            if d_min and d_max and d_min != d_max:
                return d_max.isoformat()
            return ""

        for c in claims:
            claims_keep.append([
                today_str,
                now_str,
                PRACTICE_NAME,
                str(c["ClaimNum"] or ""),
                c["PatientName"] or "",
                str(c["PatientDOB"] or ""),
                dos_from(c["DateOfServiceMin"]),
                dos_to(c["DateOfServiceMin"], c["DateOfServiceMax"]),
                str(c["DateSent"] or ""),
                c["CarrierName"] or "",
                str(c["PayorID"] or ""),
                c["PlanType"] or "",
                c["ProviderName"] or "",
                c["ProcCodes"] or "",
                round(float(c["Billed"] or 0), 2),
                round(float(c["InsPaid"] or 0), 2),
                round(float(c["InsEstimate"] or 0), 2),
                round(float(c["WriteOff"] or 0), 2),
                c["ClaimStatus"] or "",
                int(c["AgeDays"] or 0),
                int(c["ResubmitCount"] or 0),
                str(c["ClaimNum"] or ""),  # Claim ID
            ])

        claims_ws.clear()
        claims_ws.update("A1", claims_keep)
        log(f"Claims pushed: {len(claims)} claims")

        # ── Update Overview tab ───────────────────────────────────────────────
        overview_ws  = sh.worksheet("Overview")
        overview_all = overview_ws.get_all_values()

        # Find this office's row or add it
        office_row_idx = None
        for i, row in enumerate(overview_all):
            if len(row) >= 1 and row[0] == PRACTICE_NAME:
                office_row_idx = i + 1  # 1-indexed
                break

        overview_row = [
            PRACTICE_NAME,
            f"{today_str} {now_str}",
            f"${bucket_0_30:,.0f}",
            f"${bucket_31_60:,.0f}",
            f"${bucket_61_90:,.0f}",
            f"${bucket_90p:,.0f}",
            f"${total_ar:,.0f}",
            len(claims),
            f"{oldest_days}d"
        ]

        if office_row_idx and office_row_idx > 3:
            overview_ws.update(f"A{office_row_idx}", [overview_row])
        else:
            # Find next empty row after row 3
            next_row = len(overview_all) + 1
            if next_row < 4:
                next_row = 4
            overview_ws.update(f"A{next_row}", [overview_row])

        log(f"Overview updated for {PRACTICE_NAME}")
        log(f"=== AR Push Complete — {len(claims)} claims, ${total_ar:,.2f} outstanding ===")
        print(f"AR Push successful | {PRACTICE_NAME} | {len(claims)} claims | ${total_ar:,.2f} outstanding")

    except Exception as e:
        log(f"ERROR: {e}")
        import traceback
        log(traceback.format_exc())
        print(f"AR Push failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
