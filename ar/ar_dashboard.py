"""
Dental AR Dashboard — Central AR management tool.
Runs on port 5002. Reads from Google Sheet.
All offices connect to the same Sheet and see all outstanding claims.
"""

import sys, os, json, logging, smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from logging.handlers import RotatingFileHandler

sys.path.insert(0, r"C:\dental_automation\config")

try:
    from office_config import OFFICE, GMAIL_CONFIG, MANAGER_EMAIL, FOLDERS, AR_SETTINGS
except ImportError as e:
    print(f"ERROR: Could not load office_config.py: {e}")
    sys.exit(1)

from flask import Flask, request, jsonify, render_template_string
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

PRACTICE_NAME = OFFICE["name"]
SHEET_ID      = AR_SETTINGS["sheet_id"]
CREDS_PATH    = AR_SETTINGS["credentials"]
PORT          = AR_SETTINGS.get("dashboard_port", 5002)

# ── Logging setup ─────────────────────────────────────────────────────────────
LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ar_dashboard.log")

logger = logging.getLogger("ar_dashboard")
logger.setLevel(logging.INFO)
logger.propagate = False

if not logger.handlers:
    _fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    # Rotating file: 5 MB per file, keep 5 backups (~25 MB max on disk)
    _fh = RotatingFileHandler(LOG_PATH, maxBytes=5*1024*1024, backupCount=5, encoding="utf-8")
    _fh.setFormatter(_fmt)
    logger.addHandler(_fh)

    _ch = logging.StreamHandler()
    _ch.setFormatter(_fmt)
    logger.addHandler(_ch)

# Quiet werkzeug's per-request chatter (flask logs every request at INFO)
logging.getLogger("werkzeug").setLevel(logging.WARNING)

logger.info(f"━━━ AR Dashboard initialized for {PRACTICE_NAME} on port {PORT} ━━━")
logger.info(f"Sheet ID: {SHEET_ID}")
logger.info(f"Credentials path: {CREDS_PATH}")
logger.info(f"Log file: {LOG_PATH}")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ── Google Sheets connection ──────────────────────────────────────────────────
def get_sheet():
    creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID)

# ── HTML Template ─────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>AR Dashboard — __PRACTICE__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',sans-serif;background:#f4f6fb;color:#0f172a;min-height:100vh;font-size:14px}

/* Header */
.hdr{background:linear-gradient(135deg,#7f1d1d,#991b1b,#b91c1c);
     padding:14px 32px;display:flex;align-items:center;justify-content:space-between;
     position:sticky;top:0;z-index:100;box-shadow:0 4px 20px rgba(127,29,29,.35)}
.hdr h1{font-size:20px;font-weight:800;color:#fff;letter-spacing:-.4px}
.hdr p{font-size:12px;color:rgba(254,202,202,.85);font-weight:500;margin-top:1px}
.hdr-right{display:flex;align-items:center;gap:12px}
.back-btn{padding:7px 16px;background:rgba(255,255,255,.12);color:#fff;
          border:1px solid rgba(255,255,255,.2);border-radius:8px;
          font-size:13px;font-weight:600;text-decoration:none}
.refresh-btn{padding:7px 16px;background:rgba(255,255,255,.12);color:#fecaca;
             border:1px solid rgba(254,202,202,.3);border-radius:8px;
             font-size:13px;font-weight:600;cursor:pointer;border:none}

/* Summary stats */
.stats-row{display:flex;background:#fff;border-bottom:2px solid #e8edf5}
.stat-box{flex:1;padding:18px 24px;border-right:1px solid #f1f5f9}
.stat-box:last-child{border-right:none}
.stat-num{font-size:32px;font-weight:800;letter-spacing:-1px;color:#1e3a8a}
.stat-num.red{color:#dc2626}
.stat-num.amber{color:#d97706}
.stat-num.green{color:#16a34a}
.stat-lbl{font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:.8px;margin-top:5px;font-weight:600}

/* Filters */
.filters{padding:14px 32px;background:#fff;border-bottom:1px solid #e8edf5;
         display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.filter-select, .filter-input{
  padding:8px 12px;border:1.5px solid #e2e8f0;border-radius:8px;
  font-size:13px;font-family:inherit;color:#0f172a;background:#fff;min-width:140px}
.filter-select:focus,.filter-input:focus{outline:none;border-color:#dc2626}
.filter-clear{padding:8px 14px;background:#f1f5f9;color:#64748b;border:1.5px solid #e2e8f0;
              border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.filter-clear:hover{background:#e2e8f0}
.results-count{margin-left:auto;font-size:13px;color:#64748b;font-weight:500}

/* Claims table */
.table-wrap{padding:20px 32px;overflow-x:auto}
.claims-table{width:100%;border-collapse:collapse;background:#fff;
              border-radius:12px;overflow:hidden;
              box-shadow:0 1px 4px rgba(0,0,0,.05)}
.claims-table th{padding:11px 14px;text-align:left;font-size:11px;font-weight:700;
                 color:#94a3b8;text-transform:uppercase;letter-spacing:.6px;
                 background:#f8fafc;border-bottom:2px solid #e8edf5;white-space:nowrap}
.claims-table td{padding:11px 14px;border-bottom:1px solid #f8fafc;font-size:13px;
                 white-space:nowrap}
.claims-table tr:hover td{background:#f8faff;cursor:pointer}
.claims-table tr.selected td{background:#fff1f2}
.age-badge{padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}
.age-0{background:#dcfce7;color:#15803d}
.age-30{background:#fef3c7;color:#b45309}
.age-60{background:#ffedd5;color:#c2410c}
.age-90{background:#fee2e2;color:#dc2626}
.status-badge{padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;
              background:#dbeafe;color:#1d4ed8}
.note-preview{color:#64748b;font-size:11px;max-width:180px;
              overflow:hidden;text-overflow:ellipsis;white-space:nowrap}

/* Detail panel */
.detail-overlay{position:fixed;inset:0;background:rgba(15,23,42,.4);z-index:200;display:none}
.detail-panel{position:fixed;right:0;top:0;bottom:0;width:520px;background:#fff;
              z-index:201;display:flex;flex-direction:column;
              box-shadow:-4px 0 24px rgba(0,0,0,.15);transform:translateX(100%);
              transition:transform .25s ease}
.detail-panel.open{transform:translateX(0)}
.detail-hdr{padding:20px 24px;background:linear-gradient(135deg,#7f1d1d,#991b1b);
            color:#fff;flex-shrink:0}
.detail-hdr h3{font-size:17px;font-weight:800;margin-bottom:4px}
.detail-hdr p{font-size:12px;color:rgba(254,202,202,.85)}
.detail-close{position:absolute;top:16px;right:20px;background:rgba(255,255,255,.15);
              border:none;color:#fff;width:30px;height:30px;border-radius:50%;
              font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center}
.detail-body{flex:1;overflow-y:auto;padding:20px 24px}
.detail-section{margin-bottom:20px}
.detail-section-title{font-size:11px;font-weight:700;color:#94a3b8;
                      text-transform:uppercase;letter-spacing:.8px;
                      margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #f1f5f9}
.detail-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.detail-field label{font-size:11px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.detail-field p{font-size:14px;font-weight:600;color:#0f172a;margin-top:2px}

/* Notes */
.note-item{padding:12px 14px;background:#f8fafc;border-radius:8px;
           margin-bottom:8px;border-left:3px solid #e2e8f0}
.note-item.called{border-left-color:#3b82f6}
.note-item.narrative{border-left-color:#f59e0b}
.note-item.resubmit{border-left-color:#8b5cf6}
.note-item.waiting{border-left-color:#64748b}
.note-item.resolved{border-left-color:#16a34a}
.note-item.escalate{border-left-color:#dc2626}
.note-meta{font-size:11px;color:#94a3b8;margin-bottom:4px;display:flex;gap:8px;align-items:center}
.note-cat-badge{padding:1px 7px;border-radius:20px;font-size:10px;font-weight:700;
                background:#e2e8f0;color:#64748b}
.note-text{font-size:13px;color:#0f172a;line-height:1.5}
.no-notes{text-align:center;color:#cbd5e1;font-size:13px;padding:24px 0}

/* Add note form */
.note-form{background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:10px;padding:14px}
.note-form textarea{width:100%;padding:10px 12px;border:1.5px solid #e2e8f0;
                    border-radius:8px;font-size:13px;font-family:inherit;resize:vertical;
                    min-height:80px;margin-bottom:10px}
.note-form textarea:focus{outline:none;border-color:#dc2626}
.note-form-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.cat-select{flex:1;min-width:140px;padding:8px 12px;border:1.5px solid #e2e8f0;
            border-radius:8px;font-size:13px;font-family:inherit}
.btn-save-note{padding:8px 18px;background:#dc2626;color:#fff;border:none;
               border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.btn-save-note:hover{background:#b91c1c}
.btn-send-mgr{padding:8px 18px;background:#1e3a8a;color:#fff;border:none;
              border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
.btn-send-mgr:hover{background:#1e40af}

/* Action buttons */
.detail-actions{padding:16px 24px;border-top:1px solid #e8edf5;display:flex;gap:8px;flex-shrink:0}

/* Loading */
.loading{text-align:center;padding:60px;color:#94a3b8;font-size:15px}

/* Toast */
.toast{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);
       background:#0f172a;color:#fff;padding:12px 24px;border-radius:10px;
       font-size:14px;font-weight:500;z-index:500;opacity:0;transition:opacity .3s;
       pointer-events:none}
.toast.show{opacity:1}
</style>
</head>
<body>

<div class="hdr">
  <div>
    <h1>📊 AR Dashboard</h1>
    <p>Central Insurance AR Management — All Offices</p>
  </div>
  <div class="hdr-right">
    <a href="http://localhost:5000" class="back-btn">← Office Dashboard</a>
    <button class="refresh-btn" onclick="loadClaims()">↻ Refresh</button>
  </div>
</div>

<!-- Stats row -->
<div class="stats-row" id="stats-row">
  <div class="stat-box"><div class="stat-num" id="stat-total">--</div><div class="stat-lbl">Total AR</div></div>
  <div class="stat-box"><div class="stat-num" id="stat-claims">--</div><div class="stat-lbl">Open Claims</div></div>
  <div class="stat-box"><div class="stat-num red" id="stat-90">--</div><div class="stat-lbl">90+ Days</div></div>
  <div class="stat-box"><div class="stat-num amber" id="stat-offices">--</div><div class="stat-lbl">Offices</div></div>
  <div class="stat-box"><div class="stat-num" id="stat-updated">--</div><div class="stat-lbl">Last Updated</div></div>
</div>

<!-- Filters -->
<div class="filters">
  <select class="filter-select" id="f-office" onchange="applyFilters()">
    <option value="">All Offices</option>
  </select>
  <select class="filter-select" id="f-carrier" onchange="applyFilters()">
    <option value="">All Carriers</option>
  </select>
  <select class="filter-select" id="f-age" onchange="applyFilters()">
    <option value="">All Ages</option>
    <option value="0">0-30 Days</option>
    <option value="30">31-60 Days</option>
    <option value="60">61-90 Days</option>
    <option value="90">90+ Days</option>
  </select>
  <select class="filter-select" id="f-status" onchange="applyFilters()">
    <option value="">All Statuses</option>
    <option value="S">Sent to Insurance</option>
    <option value="W">Waiting on Response</option>
    <option value="R">Received</option>
    <option value="H">On Hold</option>
  </select>
  <select class="filter-select" id="f-date-type" onchange="applyFilters()" style="min-width:130px">
    <option value="sent">Date Sent</option>
    <option value="dos">Date of Service</option>
  </select>
  <input type="date" class="filter-input" id="f-date-from" onchange="applyFilters()" style="min-width:130px">
  <span style="color:#94a3b8;font-size:13px">to</span>
  <input type="date" class="filter-input" id="f-date-to" onchange="applyFilters()" style="min-width:130px">
  <input class="filter-input" id="f-search" placeholder="Search patient or carrier..." oninput="applyFilters()">
  <button class="filter-clear" onclick="clearFilters()">Clear</button>
  <div class="results-count" id="results-count"></div>
</div>

<!-- Claims table -->
<div class="table-wrap">
  <div id="table-container"><div class="loading">Loading claims...</div></div>
</div>

<!-- Detail panel overlay -->
<div class="detail-overlay" id="overlay" onclick="closeDetail()"></div>

<!-- Detail panel -->
<div class="detail-panel" id="detail-panel">
  <div class="detail-hdr">
    <button class="detail-close" onclick="closeDetail()">✕</button>
    <h3 id="dp-patient">--</h3>
    <p id="dp-carrier">--</p>
  </div>
  <div class="detail-body">
    <div class="detail-section">
      <div class="detail-section-title">Claim Details</div>
      <div class="detail-grid">
        <div class="detail-field"><label>Office</label><p id="dp-office">--</p></div>
        <div class="detail-field"><label>Date of Birth</label><p id="dp-dob">--</p></div>
        <div class="detail-field"><label>Claim #</label><p id="dp-claim-num">--</p></div>
        <div class="detail-field"><label>Payor ID</label><p id="dp-payor-id" style="color:#64748b;font-size:12px">--</p></div>
        <div class="detail-field"><label>Date of Service</label><p id="dp-dos">--</p></div>
        <div class="detail-field"><label>Date Sent</label><p id="dp-sent">--</p></div>
        <div class="detail-field"><label>Provider</label><p id="dp-provider">--</p></div>
        <div class="detail-field"><label>Status</label><p id="dp-status">--</p></div>
        <div class="detail-field"><label>Age</label><p id="dp-age">--</p></div>
        <div class="detail-field"><label>Resubmits</label><p id="dp-resub">--</p></div>
        <div class="detail-field"><label>Billed</label><p id="dp-billed">--</p></div>
        <div class="detail-field"><label>Ins Paid</label><p id="dp-paid">--</p></div>
        <div class="detail-field"><label>Write Off</label><p id="dp-writeoff">--</p></div>
        <div class="detail-field"><label>Ins Estimate</label><p id="dp-ins-estimate" style="color:#dc2626;font-size:16px">--</p></div>
      </div>
      <div style="margin-top:10px">
        <div class="detail-field"><label>Procedure Codes</label><p id="dp-codes" style="font-size:13px;color:#64748b">--</p></div>
      </div>
    </div>

    <div class="detail-section">
      <div class="detail-section-title">Notes</div>
      <div id="dp-notes-list"></div>
      <div class="note-form" style="margin-top:12px">
        <textarea id="note-text" placeholder="Add a note about this claim..."></textarea>
        <div class="note-form-row">
          <select class="cat-select" id="note-cat">
            <option value="general">📝 General</option>
            <option value="called">📞 Called Insurance</option>
            <option value="narrative">📄 Narrative Needed</option>
            <option value="resubmit">🔄 Resubmitting</option>
            <option value="waiting">⏳ Waiting on Response</option>
            <option value="resolved">✅ Resolved</option>
            <option value="escalate">🚩 Escalate to Manager</option>
          </select>
          <button class="btn-save-note" onclick="saveNote()">Save Note</button>
        </div>
      </div>
    </div>
  </div>
  <div class="detail-actions">
    <button class="btn-send-mgr" onclick="sendToManager()" style="flex:1">
      📧 Send to Manager
    </button>
  </div>
</div>

<!-- Toast -->
<div class="toast" id="toast"></div>

<script>
var ALL_CLAIMS   = [];
var ALL_NOTES    = {};
var CURRENT_CLAIM = null;

// ── Load claims from API ──────────────────────────────────────────────────────
function loadClaims() {
  document.getElementById('table-container').innerHTML = '<div class="loading">Loading claims...</div>';
  fetch('/api/claims')
    .then(function(r) { return r.json(); })
    .then(function(d) {
      ALL_CLAIMS = d.claims || [];
      ALL_NOTES  = d.notes  || {};
      populateFilters();
      applyFilters();
      updateStats();
      document.getElementById('stat-updated').textContent = d.last_push || '--';
    })
    .catch(function(e) {
      document.getElementById('table-container').innerHTML =
        '<div class="loading" style="color:#dc2626">Failed to load: ' + e.message + '</div>';
    });
}

// ── Normalize carrier name using PayorID grouping ──────────────────────────────
function normalizeCarrier(claim) {
  // Group by PayorID if available, fallback to carrier name
  if (claim.payor_id && claim.payor_id.trim() !== '') {
    return claim.carrier; // Use carrier name but grouped by payor_id
  }
  return claim.carrier;
}

function getCarrierKey(claim) {
  return (claim.payor_id && claim.payor_id.trim() !== '') ? claim.payor_id : claim.carrier;
}

function getCarrierLabel(claim) {
  return claim.carrier || claim.payor_id || '';
}

// ── Populate filter dropdowns ─────────────────────────────────────────────────
function populateFilters() {
  var offices = [...new Set(ALL_CLAIMS.map(function(c){return c.office}))].sort();
  // Build carrier map: PayorID -> best display name
  var carrierMap = {};
  ALL_CLAIMS.forEach(function(c) {
    var key = getCarrierKey(c);
    if (!carrierMap[key]) carrierMap[key] = c.carrier;
  });
  var carriers = Object.keys(carrierMap).sort(function(a,b){
    return carrierMap[a].localeCompare(carrierMap[b]);
  });

  var fo = document.getElementById('f-office');
  var cur_o = fo.value;
  fo.innerHTML = '<option value="">All Offices</option>';
  offices.forEach(function(o) {
    var opt = document.createElement('option');
    opt.value = o; opt.textContent = o;
    if (o === cur_o) opt.selected = true;
    fo.appendChild(opt);
  });

  var fc = document.getElementById('f-carrier');
  var cur_c = fc.value;
  fc.innerHTML = '<option value="">All Carriers</option>';
  carriers.forEach(function(key) {
    var opt = document.createElement('option');
    opt.value = key; opt.textContent = carrierMap[key];
    if (key === cur_c) opt.selected = true;
    fc.appendChild(opt);
  });
}

// ── Filter claims by a subset of filters (exclude one filter for cascading) ───
function filterClaims(excludeFilter) {
  var fOffice  = document.getElementById('f-office').value;
  var fCarrier = document.getElementById('f-carrier').value;
  var fAge     = document.getElementById('f-age').value;
  var fStatus  = document.getElementById('f-status').value;
  var fSearch  = document.getElementById('f-search').value.toLowerCase();

  return ALL_CLAIMS.filter(function(c) {
    if (excludeFilter !== 'office'  && fOffice  && c.office !== fOffice) return false;
    if (excludeFilter !== 'carrier' && fCarrier && getCarrierKey(c) !== fCarrier) return false;
    if (excludeFilter !== 'status'  && fStatus  && c.status !== fStatus) return false;
    if (excludeFilter !== 'age' && fAge !== '') {
      var age = parseInt(c.age || 0);
      if (fAge === '0'  && !(age <= 30))              return false;
      if (fAge === '30' && !(age >= 31 && age <= 60)) return false;
      if (fAge === '60' && !(age >= 61 && age <= 90)) return false;
      if (fAge === '90' && !(age > 90))               return false;
    }
    // Date range filter
    var fDateFrom = document.getElementById('f-date-from').value;
    var fDateTo   = document.getElementById('f-date-to').value;
    var fDateType = document.getElementById('f-date-type').value;
    if (fDateFrom || fDateTo) {
      var dateStr = fDateType === 'dos' ? c.dos : c.date_sent;
      // Convert MM/DD/YYYY or YYYY-MM-DD to comparable string
      var d = dateStr ? new Date(dateStr) : null;
      if (!d || isNaN(d)) return false;
      var iso = d.toISOString().slice(0,10);
      if (fDateFrom && iso < fDateFrom) return false;
      if (fDateTo   && iso > fDateTo)   return false;
    }
    if (fSearch) {
      var hay = (c.patient + ' ' + c.carrier + ' ' + c.claim_num).toLowerCase();
      // Tokenize: every whitespace-separated word must appear somewhere in hay.
      // Handles "LastName, FirstName" patient format — typing "Smith John" or
      // "John Smith" both match because the comma between them is ignored.
      var tokens = fSearch.split(/\s+/).filter(function(t) { return t.length > 0; });
      for (var i = 0; i < tokens.length; i++) {
        if (!hay.includes(tokens[i])) return false;
      }
    }
    return true;
  });
}

// ── Apply filters and render table ────────────────────────────────────────────
function applyFilters() {
  // Full filtered set for table
  var filtered = filterClaims(null);

  // Update each dropdown to only show options in current context
  updateDropdown('f-office',  filterClaims('office'),
    function(c){return {key:c.office, label:c.office}});
  updateDropdown('f-carrier', filterClaims('carrier'),
    function(c){return {key:getCarrierKey(c), label:c.carrier}});
  updateDropdown('f-status',  filterClaims('status'),
    function(c){return {key:c.status, label:statusLabel(c.status)}});

  document.getElementById('results-count').textContent =
    filtered.length + ' of ' + ALL_CLAIMS.length + ' claims';

  renderTable(filtered);
}

// ── Update a dropdown with only relevant options ──────────────────────────────
function updateDropdown(id, pool, keyLabelFn) {
  var el      = document.getElementById(id);
  var current = el.value;
  var seen    = {};
  var options = [];

  pool.forEach(function(c) {
    var kl = keyLabelFn(c);
    if (kl.key && !seen[kl.key]) {
      seen[kl.key] = true;
      options.push(kl);
    }
  });
  options.sort(function(a,b){return a.label.localeCompare(b.label)});

  var placeholders = {
    'f-office': 'All Offices',
    'f-carrier': 'All Carriers',
    'f-status': 'All Statuses'
  };
  el.innerHTML = '<option value="">' + (placeholders[id]||'All') + '</option>';
  options.forEach(function(o) {
    var opt = document.createElement('option');
    opt.value = o.key; opt.textContent = o.label;
    if (o.key === current) opt.selected = true;
    el.appendChild(opt);
  });

  // If current selection no longer valid, reset it
  if (current && !seen[current]) el.value = '';
}

function clearFilters() {
  document.getElementById('f-office').value    = '';
  document.getElementById('f-carrier').value   = '';
  document.getElementById('f-age').value       = '';
  document.getElementById('f-status').value    = '';
  document.getElementById('f-date-from').value = '';
  document.getElementById('f-date-to').value   = '';
  document.getElementById('f-search').value    = '';
  applyFilters();
}

// ── Render table ──────────────────────────────────────────────────────────────
function renderTable(claims) {
  if (!claims.length) {
    document.getElementById('table-container').innerHTML =
      '<div class="loading">No claims match your filters</div>';
    return;
  }

  var html = '<table class="claims-table">' +
    '<thead><tr>' +
    '<th>Office</th><th>Patient</th><th>Carrier</th>' +
    '<th>DOS</th><th>Age</th><th>Ins Estimate</th>' +
    '<th>Status</th><th>Provider</th><th>Notes</th>' +
    '</tr></thead><tbody>';

  claims.forEach(function(c) {
    var age   = parseInt(c.age || 0);
    var ageCls = age > 90 ? 'age-90' : age > 60 ? 'age-60' : age > 30 ? 'age-30' : 'age-0';
    var insPaid = parseFloat(c.ins_paid || 0);
    var rowBg = '';
    if (c.status === 'R' && insPaid > 0) rowBg = 'background:rgba(22,163,74,.06)';
    else if (c.status === 'R' && insPaid === 0) rowBg = 'background:rgba(217,119,6,.06)';
    var notes = ALL_NOTES[c.claim_id] || [];
    var lastNote = notes.length ? notes[notes.length-1].note : '';
    var isSelected = CURRENT_CLAIM && CURRENT_CLAIM.claim_id === c.claim_id ? ' selected' : '';

    html += '<tr class="claim-row' + isSelected + '" style="' + rowBg + '" onclick="openDetail(\'' + c.claim_id + '\')">' +
      '<td style="color:#64748b;font-size:12px">' + (c.office||'') + '</td>' +
      '<td style="font-weight:600;color:#0f172a">' + (c.patient||'') + '</td>' +
      '<td style="color:#334155">' + (c.carrier||'') + '</td>' +
      '<td style="color:#64748b;font-size:12px">' + (c.dos||'') + (c.dos_to ? ' — ' + c.dos_to : '') + '</td>' +
      '<td><span class="age-badge ' + ageCls + '">' + age + 'd</span></td>' +
      (function(){
        var amt = parseFloat(c.ins_estimate||0);
        var col = amt > 500 ? '#dc2626' : amt > 200 ? '#c2410c' : '#d97706';
        return '<td style="font-weight:700;color:' + col + '">$' + amt.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}) + '</td>';
      })() +
      '<td><span class="status-badge" style="background:' + statusColor(c.status) + '">' + statusLabel(c.status) + '</span></td>' +
      '<td style="color:#64748b;font-size:12px">' + (c.provider||'') + '</td>' +
      '<td><div class="note-preview">' + (lastNote || '<span style="color:#e2e8f0">No notes</span>') + '</div></td>' +
      '</tr>';
  });

  html += '</tbody></table>';
  document.getElementById('table-container').innerHTML = html;
}

// ── Update stats ──────────────────────────────────────────────────────────────
function updateStats() {
  var total    = ALL_CLAIMS.reduce(function(s,c){return s+parseFloat(c.ins_estimate||0)},0);
  var over90   = ALL_CLAIMS.filter(function(c){return parseInt(c.age||0)>90}).length;
  var offices  = new Set(ALL_CLAIMS.map(function(c){return c.office})).size;

  document.getElementById('stat-total').textContent  = '$' + (total/1000).toFixed(1) + 'k';
  document.getElementById('stat-claims').textContent = ALL_CLAIMS.length;
  document.getElementById('stat-90').textContent     = over90;
  document.getElementById('stat-offices').textContent = offices;
}

// ── Open detail panel ─────────────────────────────────────────────────────────
function openDetail(claimId) {
  var claim = ALL_CLAIMS.find(function(c){return c.claim_id === claimId});
  if (!claim) return;
  CURRENT_CLAIM = claim;

  document.getElementById('dp-patient').textContent    = claim.patient || '--';
  document.getElementById('dp-carrier').textContent    = (claim.carrier||'') + ' · ' + (claim.office||'');
  document.getElementById('dp-office').textContent     = claim.office     || '--';
  document.getElementById('dp-dob').textContent        = claim.dob        || '--';
  document.getElementById('dp-payor-id').textContent   = claim.payor_id   || '--';
  document.getElementById('dp-claim-num').textContent  = claim.claim_num  || '--';
  document.getElementById('dp-dos').textContent        = (claim.dos || '--') + (claim.dos_to ? ' — ' + claim.dos_to : '');
  document.getElementById('dp-sent').textContent       = claim.date_sent  || '--';
  document.getElementById('dp-provider').textContent   = claim.provider   || '--';
  document.getElementById('dp-status').textContent     = claim.status     || '--';
  document.getElementById('dp-age').textContent        = (claim.age||0) + ' days';
  document.getElementById('dp-resub').textContent      = claim.resubmit_count || '0';
  document.getElementById('dp-billed').textContent     = '$' + parseFloat(claim.billed||0).toLocaleString();
  document.getElementById('dp-paid').textContent       = '$' + parseFloat(claim.ins_paid||0).toLocaleString();
  document.getElementById('dp-writeoff').textContent   = '$' + parseFloat(claim.write_off||0).toLocaleString();
  document.getElementById('dp-ins-estimate').textContent = '$' + parseFloat(claim.ins_estimate||0).toLocaleString('en-US',{minimumFractionDigits:2});
  document.getElementById('dp-codes').textContent      = claim.proc_codes || '--';

  renderNotes(claimId);

  document.getElementById('overlay').style.display = 'block';
  document.getElementById('detail-panel').classList.add('open');
  document.getElementById('note-text').value = '';
}

function closeDetail() {
  document.getElementById('overlay').style.display = 'none';
  document.getElementById('detail-panel').classList.remove('open');
  CURRENT_CLAIM = null;
}

// ── Render notes ──────────────────────────────────────────────────────────────
function renderNotes(claimId) {
  var notes = ALL_NOTES[claimId] || [];
  var html = '';
  if (!notes.length) {
    html = '<div class="no-notes">No notes yet</div>';
  } else {
    notes.forEach(function(n) {
      html += '<div class="note-item ' + (n.category||'general') + '">' +
        '<div class="note-meta">' +
          '<span>' + (n.timestamp||'') + '</span>' +
          '<span class="note-cat-badge">' + catLabel(n.category) + '</span>' +
          (n.sent_to_manager === 'Yes' ? '<span style="color:#16a34a;font-size:10px">✓ Sent to Manager</span>' : '') +
        '</div>' +
        '<div class="note-text">' + (n.note||'') + '</div>' +
      '</div>';
    });
  }
  document.getElementById('dp-notes-list').innerHTML = html;
}

function statusLabel(s) {
  var labels = {S:'Sent',W:'Waiting',R:'Received',H:'On Hold'};
  return labels[s] || s;
}

function statusColor(s) {
  var colors = {S:'#dbeafe;color:#1d4ed8',W:'#fef3c7;color:#b45309',
                R:'#dcfce7;color:#15803d',H:'#f1f5f9;color:#64748b'};
  return colors[s] || '#e2e8f0;color:#334155';
}

function catLabel(cat) {
  var labels = {general:'📝 General',called:'📞 Called',narrative:'📄 Narrative',
                resubmit:'🔄 Resubmit',waiting:'⏳ Waiting',resolved:'✅ Resolved',escalate:'🚩 Escalate'};
  return labels[cat] || cat;
}

// ── Save note ─────────────────────────────────────────────────────────────────
function saveNote() {
  if (!CURRENT_CLAIM) return;
  var text = document.getElementById('note-text').value.trim();
  if (!text) { showToast('Please enter a note'); return; }
  var cat = document.getElementById('note-cat').value;

  fetch('/api/notes/add', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({
      claim_id: CURRENT_CLAIM.claim_id,
      office:   CURRENT_CLAIM.office,
      patient:  CURRENT_CLAIM.patient,
      carrier:  CURRENT_CLAIM.carrier,
      note:     text,
      category: cat
    })
  }).then(function(r){return r.json()}).then(function(d){
    if (d.success) {
      if (!ALL_NOTES[CURRENT_CLAIM.claim_id]) ALL_NOTES[CURRENT_CLAIM.claim_id] = [];
      ALL_NOTES[CURRENT_CLAIM.claim_id].push({
        timestamp: new Date().toLocaleString(),
        note: text, category: cat, sent_to_manager: 'No'
      });
      renderNotes(CURRENT_CLAIM.claim_id);
      document.getElementById('note-text').value = '';
      renderTable(ALL_CLAIMS);
      showToast('Note saved');
    } else {
      showToast('Error: ' + d.error);
    }
  });
}

// ── Send to manager ───────────────────────────────────────────────────────────
function sendToManager() {
  if (!CURRENT_CLAIM) return;
  var note = document.getElementById('note-text').value.trim();

  fetch('/api/send_to_manager', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({
      claim_id: CURRENT_CLAIM.claim_id,
      claim:    CURRENT_CLAIM,
      note:     note
    })
  }).then(function(r){return r.json()}).then(function(d){
    if (d.success) {
      showToast('Sent to manager successfully');
      if (note) {
        document.getElementById('note-cat').value = 'general';
        saveNote();
      }
    } else {
      showToast('Error: ' + d.error);
    }
  });
}

// ── Toast ─────────────────────────────────────────────────────────────────────
function showToast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); }, 3000);
}

// ── Init ──────────────────────────────────────────────────────────────────────
loadClaims();
setInterval(loadClaims, 300000); // refresh every 5 min
</script>
</body>
</html>"""

# ── API Routes ────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return HTML.replace("__PRACTICE__", PRACTICE_NAME)


@app.route("/api/claims")
def api_claims():
    try:
        sh         = get_sheet()
        claims_ws  = sh.worksheet("Claims")
        notes_ws   = sh.worksheet("Notes")

        all_claims  = claims_ws.get_all_records()
        all_notes_r = notes_ws.get_all_records()

        # Get latest push date per office+claim (overwrite same day)
        latest = {}
        for row in all_claims:
            key = str(row.get("Claim ID", "")) + "_" + str(row.get("Office", ""))
            existing = latest.get(key)
            if not existing or str(row.get("Push Date","")) >= str(existing.get("Push Date","")):
                latest[key] = row

        claims = []
        for row in latest.values():
            claims.append({
                "claim_id":      str(row.get("Claim ID", "")),
                "claim_num":     str(row.get("Claim Num", "")),
                "office":        str(row.get("Office", "")),
                "patient":       str(row.get("Patient Name", "")),
                "dob":           str(row.get("Patient DOB", "")),
                "dos":           str(row.get("Date of Service", "")),
                "dos_to":        str(row.get("DOS To", "")),
                "payor_id":      str(row.get("Payor ID", "")),
                "date_sent":     str(row.get("Date Sent", "")),
                "carrier":       str(row.get("Carrier", "")),
                "payor_id":      str(row.get("Payor ID", "")),
                "plan_type":     str(row.get("Plan Type", "")),
                "provider":      str(row.get("Provider", "")),
                "proc_codes":    str(row.get("Procedure Codes", "")),
                "billed":        str(row.get("Billed", 0)),
                "ins_paid":      str(row.get("Insurance Paid", 0)),
                "ins_estimate":  str(row.get("Ins Estimate", 0)),
                "write_off":     str(row.get("Write Off", 0)),
                "status":        str(row.get("Claim Status", "")),
                "age":           str(row.get("Age (Days)", 0)),
                "resubmit_count": str(row.get("Resubmit Count", 0)),
                "push_date":     str(row.get("Push Date", "")),
                "push_time":     str(row.get("Push Time", "")),
            })

        # Build notes dict keyed by claim_id
        notes = {}
        for row in all_notes_r:
            cid = str(row.get("Claim ID", ""))
            if cid not in notes:
                notes[cid] = []
            notes[cid].append({
                "timestamp":       str(row.get("Timestamp", "")),
                "note":            str(row.get("Note", "")),
                "category":        str(row.get("Category", "general")),
                "added_by":        str(row.get("Added By", "")),
                "sent_to_manager": str(row.get("Sent To Manager", "No")),
            })

        # Get latest push time for header
        push_times = [r.get("Push Date","") + " " + r.get("Push Time","")
                      for r in all_claims if r.get("Push Date")]
        last_push  = max(push_times) if push_times else "--"

        logger.info(f"/api/claims served {len(claims)} claims, last push: {last_push}")
        return jsonify({"claims": claims, "notes": notes, "last_push": last_push})

    except Exception as e:
        import traceback
        logger.exception(f"/api/claims FAILED: {e}")
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/notes/add", methods=["POST"])
def api_add_note():
    try:
        data = request.json
        sh       = get_sheet()
        notes_ws = sh.worksheet("Notes")

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        notes_ws.append_row([
            timestamp,
            data.get("claim_id", ""),
            data.get("office", ""),
            data.get("patient", ""),
            data.get("carrier", ""),
            data.get("note", ""),
            data.get("category", "general"),
            PRACTICE_NAME,
            "No",
            ""
        ])
        logger.info(f"Note added to claim {data.get('claim_id','?')} ({data.get('patient','?')}) by {PRACTICE_NAME}, category={data.get('category','general')}")
        return jsonify({"success": True})
    except Exception as e:
        logger.exception(f"/api/notes/add FAILED for claim {request.json.get('claim_id','?') if request.json else '?'}: {e}")
        return jsonify({"success": False, "error": str(e)})


@app.route("/api/send_to_manager", methods=["POST"])
def api_send_to_manager():
    try:
        data  = request.json
        claim = data.get("claim", {})
        note  = data.get("note", "")

        subject = (f"AR Follow-up: {claim.get('patient','')} — "
                   f"{claim.get('carrier','')} — "
                   f"${float(claim.get('ins_estimate',0)):,.2f} Ins Estimate")

        body = f"""
        <div style="font-family:Arial,sans-serif;max-width:600px">
          <h2 style="color:#991b1b">AR Follow-up Required</h2>
          <table style="width:100%;border-collapse:collapse;margin:16px 0">
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Office</td>
                <td style="padding:8px">{claim.get('office','')}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Patient</td>
                <td style="padding:8px">{claim.get('patient','')}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Carrier</td>
                <td style="padding:8px">{claim.get('carrier','')}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Claim #</td>
                <td style="padding:8px">{claim.get('claim_num','')}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Date of Service</td>
                <td style="padding:8px">{claim.get('dos','')}{(' — ' + claim.get('dos_to','')) if claim.get('dos_to') else ''}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Age</td>
                <td style="padding:8px;color:#dc2626;font-weight:700">{claim.get('age',0)} days</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Ins Estimate</td>
                <td style="padding:8px;color:#dc2626;font-weight:700;font-size:16px">
                  ${float(claim.get('ins_estimate',0)):,.2f}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Provider</td>
                <td style="padding:8px">{claim.get('provider','')}</td></tr>
            <tr><td style="padding:8px;background:#f8fafc;font-weight:600">Procedure Codes</td>
                <td style="padding:8px">{claim.get('proc_codes','')}</td></tr>
          </table>
          {f'<div style="background:#fef3c7;padding:14px;border-radius:8px;border-left:4px solid #d97706"><strong>AR Note:</strong> {note}</div>' if note else ''}
          <p style="color:#64748b;font-size:12px;margin-top:16px">
            Sent from AR Dashboard — {PRACTICE_NAME}
          </p>
        </div>
        """

        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = GMAIL_CONFIG["sender_email"]
        msg["To"]      = MANAGER_EMAIL
        msg.attach(MIMEText(body, "html"))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_CONFIG["sender_email"], GMAIL_CONFIG["sender_password"])
            server.sendmail(GMAIL_CONFIG["sender_email"], MANAGER_EMAIL, msg.as_string())

        # Log as note
        sh       = get_sheet()
        notes_ws = sh.worksheet("Notes")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        notes_ws.append_row([
            timestamp,
            data.get("claim_id", ""),
            claim.get("office", ""),
            claim.get("patient", ""),
            claim.get("carrier", ""),
            note or "Sent to manager",
            "escalate",
            PRACTICE_NAME,
            "Yes",
            timestamp
        ])

        logger.info(f"Sent to manager: claim {claim.get('claim_num','?')} ({claim.get('patient','?')}) "
                    f"${float(claim.get('ins_estimate',0)):,.2f} → {MANAGER_EMAIL}")
        return jsonify({"success": True})
    except Exception as e:
        logger.exception(f"/api/send_to_manager FAILED: {e}")
        return jsonify({"success": False, "error": str(e)})


# ── Global request/error logging ──────────────────────────────────────────────
@app.errorhandler(Exception)
def handle_unexpected(e):
    """Catches anything that escapes route-level try/except."""
    logger.exception(f"Unhandled exception on {request.method} {request.path}: {e}")
    return jsonify({"error": "Internal server error", "detail": str(e)}), 500


if __name__ == "__main__":
    logger.info(f"Starting AR Dashboard server on 0.0.0.0:{PORT}")
    print(f"Starting AR Dashboard on port {PORT}...")
    print(f"Open http://localhost:{PORT} in your browser")
    print(f"Logs: {LOG_PATH}")
    try:
        app.run(host="0.0.0.0", port=PORT, debug=False)
    except Exception as e:
        logger.critical(f"Server failed to start: {e}", exc_info=True)
        raise
