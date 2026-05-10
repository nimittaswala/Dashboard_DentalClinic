"""
Dental Automation Suite — Setup Wizard
Runs a local web server on port 5001, opens in browser.
Writes office_config.py, creates VBS launcher, adds to Windows Startup.
"""

import os, sys, json, threading, time, subprocess
from flask import Flask, request, jsonify, render_template_string

app = Flask(__name__)
INSTALL_PATH = r"C:\dental_automation"
CONFIG_PATH  = os.path.join(INSTALL_PATH, "config", "office_config.py")

# ── Load existing config if present ──────────────────────────────────────────
def load_existing_config():
    """Try to read current office_config.py values for pre-filling the form."""
    defaults = {
        "practice_name": "",
        "timezone": "America/New_York",
        "dashboard_password": "",
        "db_host": "server",
        "db_port": "3306",
        "db_name": "opendental",
        "db_user": "root",
        "db_password": "",
        "gmail_sender": "",
        "gmail_password": "",
        "manager_email": "",
        "front_desk_email": "",
        "providers": [],
        "extra_medicaid": "",
        "install_path": INSTALL_PATH,
        "sheet_id": "1WJecMp3LX9sncgzeJx1Q4z2Ndi4Ts7uHA8448wyngAs",
    }
    try:
        if os.path.exists(CONFIG_PATH):
            ns = {}
            with open(CONFIG_PATH, "r") as f:
                exec(f.read(), ns)
            defaults["practice_name"]  = ns.get("OFFICE", {}).get("name", "")
            defaults["timezone"]       = ns.get("OFFICE", {}).get("timezone", "America/New_York")
            defaults["dashboard_password"] = ns.get("DASHBOARD_PASSWORD", "")
            db = ns.get("DB_CONFIG", {})
            defaults["db_host"]        = db.get("host", "server")
            defaults["db_port"]        = str(db.get("port", 3306))
            defaults["db_name"]        = db.get("database", "opendental")
            defaults["db_user"]        = db.get("user", "root")
            defaults["db_password"]    = db.get("password", "")
            gm = ns.get("GMAIL_CONFIG", {})
            defaults["gmail_sender"]   = gm.get("sender_email", "")
            defaults["gmail_password"] = gm.get("sender_password", "")
            defaults["manager_email"]  = ns.get("MANAGER_EMAIL", "")
            defaults["front_desk_email"] = ns.get("FRONT_DESK_EMAIL", "")
            prov = ns.get("PROVIDER_EMAILS", {})
            defaults["providers"]      = [{"name": k, "email": v} for k, v in prov.items()]
    except Exception as e:
        print(f"Could not load existing config: {e}")
    return defaults


HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Dental Automation — Setup Wizard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',sans-serif;background:#f4f6fb;color:#0f172a;min-height:100vh}
.hdr{background:linear-gradient(135deg,#0f2250,#1a3a7c);padding:20px 40px;
     display:flex;align-items:center;gap:16px;
     box-shadow:0 4px 20px rgba(10,30,90,.35)}
.hdr-icon{font-size:32px}
.hdr h1{color:#fff;font-size:22px;font-weight:800;letter-spacing:-.4px}
.hdr p{color:#93c5fd;font-size:13px;margin-top:2px}
.progress-bar{height:4px;background:#1e3a8a}
.progress-fill{height:100%;background:#fde68a;width:0%;transition:width .4s ease}
.body{max-width:780px;margin:32px auto;padding:0 20px 60px}
.section{background:#fff;border:1px solid #e8edf5;border-radius:16px;
         padding:28px 32px;margin-bottom:20px;
         box-shadow:0 1px 4px rgba(0,0,0,.05)}
.section-hdr{display:flex;align-items:center;gap:12px;margin-bottom:22px;
             padding-bottom:14px;border-bottom:2px solid #f1f5f9}
.section-num{width:32px;height:32px;border-radius:50%;background:#1e3a8a;
             color:#fff;font-size:13px;font-weight:800;
             display:flex;align-items:center;justify-content:center;flex-shrink:0}
.section-title{font-size:16px;font-weight:700;color:#0f172a}
.section-sub{font-size:12px;color:#64748b;margin-top:2px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.grid-3{display:grid;grid-template-columns:2fr 1fr 1fr;gap:16px}
.field{display:flex;flex-direction:column;gap:5px}
.field.full{grid-column:1/-1}
label{font-size:12px;font-weight:600;color:#374151;text-transform:uppercase;letter-spacing:.5px}
input,select{padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;
             font-size:14px;font-family:inherit;color:#0f172a;
             transition:border-color .15s;background:#fff}
input:focus,select:focus{outline:none;border-color:#2563eb;
                          box-shadow:0 0 0 3px rgba(37,99,235,.1)}
input.error{border-color:#dc2626}
.hint{font-size:11px;color:#94a3b8;margin-top:3px}
.hint a{color:#2563eb}
/* Provider table */
.prov-table{width:100%;border-collapse:collapse;margin-top:8px}
.prov-table th{text-align:left;font-size:11px;font-weight:700;color:#94a3b8;
               text-transform:uppercase;letter-spacing:.5px;padding:8px 10px;
               border-bottom:2px solid #f1f5f9}
.prov-table td{padding:6px 6px}
.prov-table td input{padding:8px 10px;font-size:13px}
.prov-row-del{background:none;border:none;color:#dc2626;font-size:16px;
              cursor:pointer;padding:4px 8px;border-radius:6px}
.prov-row-del:hover{background:#fee2e2}
.btn-add-prov{margin-top:10px;padding:8px 16px;background:#eff6ff;color:#1d4ed8;
              border:1.5px solid #bfdbfe;border-radius:8px;font-size:13px;
              font-weight:600;cursor:pointer;font-family:inherit}
.btn-add-prov:hover{background:#dbeafe}
/* Test DB button */
.btn-test{padding:10px 20px;background:#f0f4ff;color:#1e3a8a;
          border:1.5px solid #bfdbfe;border-radius:8px;font-size:13px;
          font-weight:600;cursor:pointer;font-family:inherit;
          display:flex;align-items:center;gap:8px}
.btn-test:hover{background:#dbeafe}
.test-result{margin-top:10px;padding:10px 14px;border-radius:8px;
             font-size:13px;font-weight:500;display:none}
.test-ok{background:#dcfce7;color:#15803d;border:1px solid #86efac}
.test-fail{background:#fee2e2;color:#dc2626;border:1px solid #fca5a5}
/* Medicaid carriers */
.carrier-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:12px}
.carrier-item{display:flex;align-items:center;gap:8px;padding:8px 10px;
              background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;
              font-size:13px}
.carrier-item input[type=checkbox]{width:16px;height:16px;cursor:pointer}
/* Submit button */
.submit-wrap{position:fixed;bottom:0;left:0;right:0;background:#fff;
             border-top:2px solid #e8edf5;padding:16px 40px;
             display:flex;align-items:center;justify-content:space-between;
             box-shadow:0 -4px 20px rgba(0,0,0,.08);z-index:100}
.btn-submit{padding:14px 32px;background:linear-gradient(135deg,#0f2250,#1e4db7);
            color:#fff;border:none;border-radius:10px;font-size:15px;
            font-weight:700;cursor:pointer;font-family:inherit;
            transition:all .2s;letter-spacing:.2px}
.btn-submit:hover{filter:brightness(1.1);transform:translateY(-1px);
                  box-shadow:0 6px 20px rgba(30,53,100,.3)}
.btn-submit:disabled{opacity:.5;cursor:not-allowed;transform:none;box-shadow:none}
.status-msg{font-size:13px;color:#64748b}
/* Success overlay */
.overlay{position:fixed;inset:0;background:rgba(15,34,80,.7);
         display:flex;align-items:center;justify-content:center;
         z-index:200;display:none}
.success-card{background:#fff;border-radius:20px;padding:40px 48px;
              text-align:center;max-width:420px;
              box-shadow:0 20px 60px rgba(0,0,0,.3)}
.success-icon{font-size:56px;margin-bottom:16px}
.success-card h2{font-size:22px;font-weight:800;color:#0f172a;margin-bottom:8px}
.success-card p{color:#64748b;font-size:14px;line-height:1.6}
.success-card .note{margin-top:16px;padding:12px 16px;background:#f0f4ff;
                    border-radius:8px;font-size:13px;color:#1e3a8a;font-weight:500}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-icon">🦷</div>
  <div>
    <h1>Dental Automation Suite — Setup Wizard</h1>
    <p>Complete the form below to configure the automation for this office</p>
  </div>
</div>
<div class="progress-bar"><div class="progress-fill" id="progress"></div></div>

<div class="body">

  <!-- Section 1: Practice Info -->
  <div class="section">
    <div class="section-hdr">
      <div class="section-num">1</div>
      <div>
        <div class="section-title">Practice Information</div>
        <div class="section-sub">Basic details about this office</div>
      </div>
    </div>
    <div class="grid">
      <div class="field full">
        <label>Practice Name *</label>
        <input type="text" id="practice_name" placeholder="e.g. Pristine Dental Center" value="__PRACTICE_NAME__">
      </div>
      <div class="field full">
        <label>Dashboard Password *</label>
        <div style="display:flex;gap:8px;align-items:center">
          <input type="password" id="dashboard_password" placeholder="Min 6 characters" value="__DASHBOARD_PASS__" style="flex:1">
          <button type="button" onclick="toggleVis('dashboard_password',this)"
            style="padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;background:#f8fafc;cursor:pointer;font-size:13px;white-space:nowrap">
            👁 Show
          </button>
        </div>
        <span class="hint">Required to open the dashboard at port 5000. Anyone with this password can sign in. Minimum 6 characters.</span>
      </div>
      <div class="field">
        <label>Install Path</label>
        <input type="text" id="install_path" value="__INSTALL_PATH__">
        <span class="hint">Folder where all scripts will live. Default is fine for most setups.</span>
      </div>
      <div class="field">
        <label>Timezone</label>
        <select id="timezone">
          <option value="America/New_York" __TZ_EAST__>Eastern (ET)</option>
          <option value="America/Chicago" __TZ_CENT__>Central (CT)</option>
          <option value="America/Denver" __TZ_MOUNT__>Mountain (MT)</option>
          <option value="America/Los_Angeles" __TZ_PAC__>Pacific (PT)</option>
        </select>
      </div>
    </div>
  </div>

  <!-- Section 2: Database -->
  <div class="section">
    <div class="section-hdr">
      <div class="section-num">2</div>
      <div>
        <div class="section-title">Database Connection</div>
        <div class="section-sub">Open Dental MySQL database credentials</div>
      </div>
    </div>
    <div class="grid-3">
      <div class="field">
        <label>Host / Server Name *</label>
        <input type="text" id="db_host" placeholder="server or IP address" value="__DB_HOST__">
        <span class="hint">Usually "server" or the server's IP address</span>
      </div>
      <div class="field">
        <label>Port</label>
        <input type="text" id="db_port" value="__DB_PORT__">
      </div>
      <div class="field">
        <label>Database Name</label>
        <input type="text" id="db_name" value="__DB_NAME__">
      </div>
      <div class="field">
        <label>Username</label>
        <input type="text" id="db_user" value="__DB_USER__">
      </div>
      <div class="field">
        <label>Password</label>
        <div style="display:flex;gap:8px;align-items:center">
          <input type="password" id="db_password" placeholder="Leave blank if no password" value="__DB_PASS__" style="flex:1">
          <button type="button" onclick="toggleVis('db_password',this)"
            style="padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;background:#f8fafc;cursor:pointer;font-size:13px;white-space:nowrap">
            👁 Show
          </button>
        </div>
        <span class="hint">Leave empty if your database has no password</span>
      </div>
    </div>
    <div style="margin-top:16px">
      <button class="btn-test" onclick="testDB()">
        <span>🔌</span> Test Connection
      </button>
      <div class="test-result test-ok"  id="test-ok">✓ Connection successful — database is reachable</div>
      <div class="test-result test-fail" id="test-fail"></div>
    </div>
  </div>

  <!-- Section 3: Email -->
  <div class="section">
    <div class="section-hdr">
      <div class="section-num">3</div>
      <div>
        <div class="section-title">Email Configuration</div>
        <div class="section-sub">Gmail account used to send all automated emails</div>
      </div>
    </div>
    <div class="grid">
      <div class="field">
        <label>Gmail Address *</label>
        <input type="email" id="gmail_sender" placeholder="office@gmail.com" value="__GMAIL_SENDER__">
      </div>
      <div class="field">
        <label>Gmail App Password *</label>
        <div style="display:flex;gap:8px;align-items:center">
          <input type="password" id="gmail_password" placeholder="xxxx xxxx xxxx xxxx" value="__GMAIL_PASS__" style="flex:1">
          <button type="button" onclick="toggleVis('gmail_password',this)" 
            style="padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;background:#f8fafc;cursor:pointer;font-size:13px;white-space:nowrap">
            👁 Show
          </button>
        </div>
        <span class="hint">
          Not your regular password. 
          <a href="https://myaccount.google.com/apppasswords" target="_blank">Generate an App Password here</a>.
          Enter with spaces exactly as shown: xxxx xxxx xxxx xxxx
        </span>
      </div>
      <div class="field">
        <label>Manager Email *</label>
        <input type="email" id="manager_email" placeholder="manager@gmail.com" value="__MANAGER_EMAIL__">
        <span class="hint">Receives production reports, recall lists, day summary</span>
      </div>
      <div class="field">
        <label>Front Desk Email *</label>
        <input type="email" id="front_desk_email" placeholder="frontdesk@gmail.com" value="__FRONT_DESK_EMAIL__">
        <span class="hint">Receives morning huddle summary</span>
      </div>
    </div>
  </div>

  <!-- Section 4: Providers -->
  <div class="section">
    <div class="section-hdr">
      <div class="section-num">4</div>
      <div>
        <div class="section-title">Provider Emails</div>
        <div class="section-sub">Unsigned notes are sent to each provider individually</div>
      </div>
    </div>
    <table class="prov-table">
      <thead>
        <tr>
          <th>Provider Full Name (as it appears in Open Dental)</th>
          <th>Email Address</th>
          <th></th>
        </tr>
      </thead>
      <tbody id="prov-tbody">
      </tbody>
    </table>
    <button class="btn-add-prov" onclick="addProvider()">+ Add Provider</button>
  </div>

  <!-- Section 5: Medicaid Carriers -->
  <div class="section">
    <div class="section-hdr">
      <div class="section-num">5</div>
      <div>
        <div class="section-title">AR Dashboard Settings</div>
        <div class="section-sub">Central AR dashboard shared across all offices</div>
      </div>
    </div>
    <div class="grid">
      <div class="field full">
        <label>Google Sheet ID *</label>
        <input type="text" id="sheet_id" placeholder="Paste Sheet ID from Google Sheets URL" value="__SHEET_ID__">
        <span class="hint">From the URL: docs.google.com/spreadsheets/d/<strong>[THIS PART]</strong>/edit</span>
      </div>
    </div>
  </div>

  <!-- Section 6: Medicaid Carriers -->
  <div class="section">
    <div class="section-hdr">
      <div class="section-num">6</div>
      <div>
        <div class="section-title">Medicaid Carriers</div>
        <div class="section-sub">Used to calculate Medicaid collection rates and reports</div>
      </div>
    </div>
    <p style="font-size:13px;color:#64748b;margin-bottom:14px">
      Standard PA carriers are pre-selected. Check any additional carriers used by this office.
    </p>
    <div class="carrier-grid" id="carrier-grid">
      <!-- Populated by JS -->
    </div>
    <div class="field" style="margin-top:12px">
      <label>Additional Carriers (one per line)</label>
      <textarea id="extra_medicaid" rows="3"
        style="padding:10px 14px;border:1.5px solid #e2e8f0;border-radius:8px;
               font-size:13px;font-family:inherit;resize:vertical"
        placeholder="Add any other Medicaid carrier names not listed above">__EXTRA_MEDICAID__</textarea>
      <span class="hint">Enter exact carrier name as it appears in Open Dental</span>
    </div>
  </div>

</div><!-- /body -->

<!-- Submit bar -->
<div class="submit-wrap">
  <div class="status-msg" id="status-msg">Fill in all required fields (*) to continue</div>
  <button class="btn-submit" id="btn-submit" onclick="submitForm()">
    💾 Save and Launch Dashboard
  </button>
</div>

<!-- Success overlay -->
<div class="overlay" id="overlay">
  <div class="success-card">
    <div class="success-icon">🎉</div>
    <h2>Setup Complete!</h2>
    <p>office_config.py has been saved, the startup shortcut has been created, and the dashboard is launching.</p>
    <div class="note">
      Open your browser and go to<br>
      <strong id="dash-url">http://localhost:5000</strong>
      <br>to access the dashboard from any computer on this network.
    </div>
  </div>
</div>

<script>
var DEFAULTS = __DEFAULTS_JSON__;

function toggleVis(fieldId, btn) {
  var input = document.getElementById(fieldId);
  if (input.type === 'password') {
    input.type = 'text';
    btn.textContent = '🙈 Hide';
  } else {
    input.type = 'password';
    btn.textContent = '👁 Show';
  }
}

// ── Medicaid carriers ──────────────────────────────────────────────────────
var STANDARD_CARRIERS = [
  "keystone first", "keystone first dentaquest", "keystone first vip",
  "keystone mercy", "upmc", "united healthcare community plan",
  "health partners", "health partner", "health partner medicare",
  "aetna better health", "dentaquest", "avesis"
];

function buildCarrierGrid() {
  var g = document.getElementById('carrier-grid');
  g.innerHTML = STANDARD_CARRIERS.map(function(c) {
    return '<label class="carrier-item">' +
      '<input type="checkbox" value="' + c + '" checked>' +
      c.replace(/\b\w/g, function(l){return l.toUpperCase()}) +
    '</label>';
  }).join('');
}

// ── Providers ──────────────────────────────────────────────────────────────
function addProvider(name, email) {
  var tbody = document.getElementById('prov-tbody');
  var tr = document.createElement('tr');
  tr.innerHTML =
    '<td><input type="text" placeholder="First Last" value="' + (name||'') + '" class="prov-name" style="width:100%"></td>' +
    '<td><input type="email" placeholder="provider@gmail.com" value="' + (email||'') + '" class="prov-email" style="width:100%"></td>' +
    '<td><button class="prov-row-del" onclick="this.closest(\'tr\').remove()">✕</button></td>';
  tbody.appendChild(tr);
}

// ── Test DB connection ─────────────────────────────────────────────────────
function testDB() {
  var btn  = document.querySelector('.btn-test');
  var ok   = document.getElementById('test-ok');
  var fail = document.getElementById('test-fail');
  ok.style.display = fail.style.display = 'none';
  btn.innerHTML = '<span>⏳</span> Testing...';
  btn.disabled = true;

  fetch('/test_db', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({
      host:     document.getElementById('db_host').value,
      port:     document.getElementById('db_port').value,
      name:     document.getElementById('db_name').value,
      user:     document.getElementById('db_user').value,
      password: document.getElementById('db_password').value,
    })
  })
  .then(function(r) { return r.json(); })
  .then(function(d) {
    btn.innerHTML = '<span>🔌</span> Test Connection';
    btn.disabled = false;
    if (d.success) {
      ok.style.display = 'block';
      fail.style.display = 'none';
    } else {
      fail.textContent = '✕ ' + d.error;
      fail.style.display = 'block';
      ok.style.display = 'none';
    }
  })
  .catch(function(err) {
    btn.innerHTML = '<span>🔌</span> Test Connection';
    btn.disabled = false;
    fail.textContent = '✕ Connection error: ' + err.message;
    fail.style.display = 'block';
  });
}

// ── Submit ─────────────────────────────────────────────────────────────────
function submitForm() {
  var btn = document.getElementById('btn-submit');
  var msg = document.getElementById('status-msg');

  // Gather providers
  var providers = {};
  document.querySelectorAll('#prov-tbody tr').forEach(function(tr) {
    var name  = tr.querySelector('.prov-name').value.trim();
    var email = tr.querySelector('.prov-email').value.trim();
    if (name) providers[name] = email;
  });

  // Gather checked carriers
  var carriers = [];
  document.querySelectorAll('#carrier-grid input:checked').forEach(function(cb) {
    carriers.push(cb.value);
  });
  var extra = document.getElementById('extra_medicaid').value;
  extra.split('\n').forEach(function(line) {
    var l = line.trim().toLowerCase();
    if (l && carriers.indexOf(l) === -1) carriers.push(l);
  });

  var payload = {
    practice_name:    document.getElementById('practice_name').value.trim(),
    timezone:         document.getElementById('timezone').value,
    install_path:     document.getElementById('install_path').value.trim(),
    dashboard_password: document.getElementById('dashboard_password').value,
    db_host:          document.getElementById('db_host').value.trim(),
    db_port:          document.getElementById('db_port').value.trim(),
    db_name:          document.getElementById('db_name').value.trim(),
    db_user:          document.getElementById('db_user').value.trim(),
    db_password:      document.getElementById('db_password').value,
    gmail_sender:     document.getElementById('gmail_sender').value.trim(),
    gmail_password:   document.getElementById('gmail_password').value,
    manager_email:    document.getElementById('manager_email').value.trim(),
    front_desk_email: document.getElementById('front_desk_email').value.trim(),
    providers:        providers,
    medicaid_carriers: carriers,
    sheet_id:         document.getElementById("sheet_id").value.trim(),
  };

  // Basic validation
  var required = ['practice_name','dashboard_password','db_host','gmail_sender','gmail_password','manager_email','front_desk_email'];
  var missing = required.filter(function(k) { return !payload[k]; });
  if (missing.length) {
    msg.textContent = '⚠ Please fill in all required fields';
    msg.style.color = '#dc2626';
    return;
  }
  if (payload.dashboard_password.length < 6) {
    msg.textContent = '⚠ Dashboard password must be at least 6 characters';
    msg.style.color = '#dc2626';
    return;
  }

  btn.disabled = true;
  btn.textContent = '⏳ Saving...';
  msg.textContent = 'Writing configuration...';
  msg.style.color = '#64748b';

  fetch('/save', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify(payload)
  }).then(r => r.json()).then(d => {
    if (d.success) {
      document.getElementById('dash-url').textContent = 'http://' + d.ip + ':5000';
      document.getElementById('overlay').style.display = 'flex';
      setTimeout(function() { window.close(); }, 8000);
    } else {
      btn.disabled = false;
      btn.textContent = '💾 Save and Launch Dashboard';
      msg.textContent = '✕ Error: ' + d.error;
      msg.style.color = '#dc2626';
    }
  });
}

// ── Progress bar (scroll-based) ────────────────────────────────────────────
window.addEventListener('scroll', function() {
  var h = document.body.scrollHeight - window.innerHeight;
  var pct = h > 0 ? (window.scrollY / h * 100) : 100;
  document.getElementById('progress').style.width = pct + '%';
});

// ── Init ───────────────────────────────────────────────────────────────────
buildCarrierGrid();

// Pre-fill providers
if (DEFAULTS.providers && DEFAULTS.providers.length) {
  DEFAULTS.providers.forEach(function(p) { addProvider(p.name, p.email); });
} else {
  addProvider('', '');
}
</script>
</body>
</html>
"""

# ── Flask routes ──────────────────────────────────────────────────────────────
@app.route("/")
def index():
    d = load_existing_config()
    html = HTML
    html = html.replace("__PRACTICE_NAME__",   d["practice_name"])
    html = html.replace("__INSTALL_PATH__",     d["install_path"])
    html = html.replace("__TZ_EAST__",   "selected" if d["timezone"] == "America/New_York"    else "")
    html = html.replace("__TZ_CENT__",   "selected" if d["timezone"] == "America/Chicago"      else "")
    html = html.replace("__TZ_MOUNT__",  "selected" if d["timezone"] == "America/Denver"       else "")
    html = html.replace("__TZ_PAC__",    "selected" if d["timezone"] == "America/Los_Angeles"  else "")
    html = html.replace("__DB_HOST__",   d["db_host"])
    html = html.replace("__DB_PORT__",   d["db_port"])
    html = html.replace("__DB_NAME__",   d["db_name"])
    html = html.replace("__DB_USER__",   d["db_user"])
    html = html.replace("__DB_PASS__",   d["db_password"])
    html = html.replace("__DASHBOARD_PASS__",  d["dashboard_password"])
    html = html.replace("__GMAIL_SENDER__",    d["gmail_sender"])
    html = html.replace("__GMAIL_PASS__",      d["gmail_password"])
    html = html.replace("__MANAGER_EMAIL__",   d["manager_email"])
    html = html.replace("__FRONT_DESK_EMAIL__", d["front_desk_email"])
    html = html.replace("__EXTRA_MEDICAID__",  d["extra_medicaid"])
    html = html.replace("__SHEET_ID__",       d.get("sheet_id", ""))
    html = html.replace("__DEFAULTS_JSON__",   json.dumps(d))
    return html


@app.route("/test_db", methods=["POST"])
def test_db():
    data = request.json
    try:
        import mysql.connector
        conn = mysql.connector.connect(
            host=data["host"], port=int(data["port"]),
            database=data["name"], user=data["user"], password=data["password"],
            connection_timeout=5
        )
        conn.close()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route("/save", methods=["POST"])
def save():
    d = request.json
    try:
        install_path = d["install_path"].rstrip("\\")

        # ── Build office_config.py content ──────────────────────────────
        providers_str = "\n".join(
            f'    "{k}": "{v}",' for k, v in d["providers"].items()
        )
        carriers_str = "\n".join(
            f'        "{c}",' for c in d["medicaid_carriers"]
        )

        gmail_password_safe = d["gmail_password"].replace('"', '\\"').replace('\n', '').replace('\r', '')
        dashboard_password_safe = (d.get("dashboard_password","") or "").replace('"', '\\"').replace('\n', '').replace('\r', '')
        config_content = f'''# ============================================================
# OFFICE CONFIG — Generated by Setup Wizard
# This is the ONLY file that needs to change when deploying
# to a new office. All scripts import from here.
# ============================================================

OFFICE = {{
    "name":     "{d["practice_name"]}",
    "timezone": "{d["timezone"]}",
}}

# Password required to open the office dashboard at port 5000.
# To change it later, run reconfigure.bat.
DASHBOARD_PASSWORD = "{dashboard_password_safe}"

DB_CONFIG = {{
    "host":     "{d["db_host"]}",
    "port":     {d["db_port"]},
    "database": "{d["db_name"]}",
    "user":     "{d["db_user"]}",
    "password": "{d["db_password"]}",
}}

GMAIL_CONFIG = {{
    "sender_email":    "{d["gmail_sender"]}",
    "sender_password": "{gmail_password_safe}",
}}

# =============================================================================
# RECIPIENTS
# =============================================================================

PROVIDER_EMAILS = {{
{providers_str}
}}

MANAGER_EMAIL    = "{d["manager_email"]}"
FRONT_DESK_EMAIL = "{d["front_desk_email"]}"

# =============================================================================
# FOLDERS
# =============================================================================

FOLDERS = {{
    "base":           r"{install_path}",
    "unsigned_notes": r"{install_path}\\unsigned_notes",
    "provider_notes": r"{install_path}\\unsigned_notes\\provider_notes",
    "huddle":         r"{install_path}\\huddle_summary",
    "recall":         r"{install_path}\\recall_list",
    "production":     r"{install_path}\\production",
}}

# =============================================================================
# SCHEDULES
# =============================================================================

SCHEDULES = {{
    "unsigned_notes": "10:00 AM daily",
    "huddle_summary": "08:00 AM daily",
    "recall_list":    "08:00 AM Monday",
    "production":     "07:00 PM daily",
}}

# =============================================================================
# UNSIGNED NOTES SETTINGS
# =============================================================================

UNSIGNED_NOTES_SETTINGS = {{
    "lookback_days":        30,
    "perio_lookahead_days": 30,
    "notes_retention_days": 7,
}}

# =============================================================================
# INSURANCE VERIFICATION SETTINGS
# =============================================================================

INS_VERIFY_SETTINGS = {{
    "green_days":  7,
    "yellow_days": 15,
}}

# =============================================================================
# RECALL SETTINGS
# =============================================================================

RECALL_SETTINGS = {{
    "interval_months":       6,
    "grace_period_months":   1,
    "lookback_appointments": 5,
    "low_priority_carriers": [
{carriers_str}
    ],
}}

# =============================================================================
# PRODUCTION DASHBOARD SETTINGS
# =============================================================================

PAYMENT_TYPES = {{
    69:  "Check",
    70:  "Cash",
    71:  "Credit Card",
    72:  "Patient Refund",
    313: "Cash",
    314: "Check",
    316: "Credit Card",
    317: "Online BillPay",
    318: "EFT",
    364: "Care Credit",
    397: "Modento",
    406: "Proceed Finance",
    412: "Cherry Payments",
    414: "Message-to-Pay",
    415: "FHD",
    315: "Insurance Check",
}}

PAYMENT_GROUPS = {{
    "Cash":          [70, 313],
    "Check":         [69, 314],
    "Credit Card":   [71, 316, 364, 397, 406, 412, 414, 415],
    "EFT / Online":  [317, 318],
    "Insurance Check": [315],
    "Patient Refund": [72],
}}

PRODUCTION_SETTINGS = {{
    "insurance_payment_note": True,
}}

# =============================================================================
# MEDICAID CARRIERS
# =============================================================================

MEDICAID_CARRIERS_BASE = [
{carriers_str}
]

MEDICAID_CARRIERS_EXTRA = []

MEDICAID_CARRIERS = MEDICAID_CARRIERS_BASE + MEDICAID_CARRIERS_EXTRA

# =============================================================================
# AR DASHBOARD SETTINGS
# =============================================================================

AR_SETTINGS = {{
    "sheet_id":       "{d["sheet_id"]}",
    "credentials":    r"{install_path}\\ar\\credentials.json",
    "dashboard_port": 5002,
}}
'''

        # Write office_config.py
        config_dir = os.path.join(install_path, "config")
        os.makedirs(config_dir, exist_ok=True)
        with open(os.path.join(config_dir, "office_config.py"), "w") as f:
            f.write(config_content)

        # ── Create start_dashboard.vbs ──────────────────────────────────
        # The VBS auto-detects Python at runtime so the file works regardless
        # of where Python is installed (per-user, all-users, etc.) and also
        # survives a Python upgrade without needing reconfigure.
        vbs_content = (
            'Set WshShell = CreateObject("WScript.Shell")\r\n'
            'Set fso = CreateObject("Scripting.FileSystemObject")\r\n'
            '\r\n'
            'Dim pyPath, localApp\r\n'
            'pyPath = ""\r\n'
            'localApp = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")\r\n'
            '\r\n'
            'Dim candidates(15)\r\n'
            'candidates(0)  = localApp & "\\Programs\\Python\\Python314\\python.exe"\r\n'
            'candidates(1)  = localApp & "\\Programs\\Python\\Python313\\python.exe"\r\n'
            'candidates(2)  = localApp & "\\Programs\\Python\\Python312\\python.exe"\r\n'
            'candidates(3)  = localApp & "\\Programs\\Python\\Python311\\python.exe"\r\n'
            'candidates(4)  = "C:\\Program Files\\Python314\\python.exe"\r\n'
            'candidates(5)  = "C:\\Program Files\\Python313\\python.exe"\r\n'
            'candidates(6)  = "C:\\Program Files\\Python312\\python.exe"\r\n'
            'candidates(7)  = "C:\\Program Files\\Python311\\python.exe"\r\n'
            'candidates(8)  = "C:\\Program Files (x86)\\Python314\\python.exe"\r\n'
            'candidates(9)  = "C:\\Program Files (x86)\\Python313\\python.exe"\r\n'
            'candidates(10) = "C:\\Program Files (x86)\\Python312\\python.exe"\r\n'
            'candidates(11) = "C:\\Program Files (x86)\\Python311\\python.exe"\r\n'
            'candidates(12) = "C:\\Python314\\python.exe"\r\n'
            'candidates(13) = "C:\\Python313\\python.exe"\r\n'
            'candidates(14) = "C:\\Python312\\python.exe"\r\n'
            'candidates(15) = "C:\\Python311\\python.exe"\r\n'
            '\r\n'
            'Dim i\r\n'
            'For i = 0 To 15\r\n'
            '    If fso.FileExists(candidates(i)) Then\r\n'
            '        pyPath = candidates(i)\r\n'
            '        Exit For\r\n'
            '    End If\r\n'
            'Next\r\n'
            '\r\n'
            'If pyPath = "" Then\r\n'
            '    MsgBox "Python not found. Please run install.bat first.", 16, "Dental Dashboard"\r\n'
            '    WScript.Quit 1\r\n'
            'End If\r\n'
            '\r\n'
            f'WshShell.Run Chr(34) & pyPath & Chr(34) & " " & Chr(34) & "{install_path}\\dashboard\\dashboard.py" & Chr(34), 0, False\r\n'
            f'WshShell.Run Chr(34) & pyPath & Chr(34) & " " & Chr(34) & "{install_path}\\ar\\ar_dashboard.py" & Chr(34), 0, False\r\n'
            'Set WshShell = Nothing\r\n'
        )
        vbs_path = os.path.join(install_path, "start_dashboard.vbs")
        with open(vbs_path, "w") as f:
            f.write(vbs_content)

        # ── Add to Windows Startup folder ──────────────────────────────
        try:
            startup_folder = os.path.join(
                os.environ.get("APPDATA", ""),
                "Microsoft", "Windows", "Start Menu", "Programs", "Startup"
            )
            shortcut_path = os.path.join(startup_folder, "DentalDashboard.vbs")
            import shutil
            shutil.copy2(vbs_path, shortcut_path)
        except Exception as e:
            print(f"Warning: Could not add to startup folder: {e}")

        # ── Get local IP for display ────────────────────────────────────
        import socket
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            local_ip = s.getsockname()[0]
            s.close()
        except Exception:
            local_ip = "localhost"

        # ── Launch the dashboard directly using full Python path ────────
        try:
            subprocess.Popen(
                [python_path, os.path.join(install_path, "dashboard", "dashboard.py")],
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
            )
        except Exception as e:
            print(f"Warning: Could not auto-launch dashboard: {e}")

        # Shut down wizard after a delay
        def shutdown():
            time.sleep(5)
            os._exit(0)
        threading.Thread(target=shutdown, daemon=True).start()

        return jsonify({"success": True, "ip": local_ip})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


if __name__ == "__main__":
    print("=" * 55)
    print("  Dental Automation Suite — Setup Wizard")
    print("  Opening http://localhost:5001 in your browser...")
    print("=" * 55)
    app.run(host="0.0.0.0", port=5001, debug=False)
