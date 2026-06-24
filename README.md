# inventory-dashboard
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Inventory & Cycle Count Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden;font-family:Arial,sans-serif;font-size:13px;color:#222;background:#ffffff;display:flex;flex-direction:column}

/* ── TOP BAR ── */
.topbar{background:#ffffff;border-bottom:2px solid #1a3a5c;padding:0 16px;display:flex;align-items:center;gap:0;flex-shrink:0;height:50px}
.topbar-logo{font-size:15px;font-weight:700;color:#1a3a5c;padding-right:18px;border-right:1px solid #e0e4ea;margin-right:14px;white-space:nowrap}
.topbar-logo span{color:#2980b9}
.client-tabs{display:flex;gap:2px;flex:1;overflow-x:auto;align-items:stretch;height:100%}
.ctab{display:flex;align-items:center;gap:6px;padding:0 14px;font-size:12px;font-weight:600;cursor:pointer;border:none;background:transparent;color:#888;border-bottom:3px solid transparent;white-space:nowrap;transition:all .15s;flex-shrink:0}
.ctab:hover{color:#1a3a5c;background:#f5f8fb}
.ctab.active{color:#1a3a5c;border-bottom-color:#2980b9;background:#f0f6fb}
.ctab .dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.ctab-add{padding:0 12px;font-size:20px;font-weight:300;color:#aaa;cursor:pointer;border:none;background:none;flex-shrink:0}
.ctab-add:hover{color:#1a3a5c}
.topbar-right{display:flex;align-items:center;gap:6px;padding-left:14px;border-left:1px solid #e0e4ea;flex-shrink:0}
.badge{font-size:11px;padding:2px 9px;border-radius:9px;font-weight:600;white-space:nowrap}
.b-blue{background:#e8f2fb;color:#1a3a5c}.b-green{background:#e3f6e8;color:#1a6b2e}.b-amber{background:#fff3cd;color:#8a6500}.b-red{background:#fde8e8;color:#a32020}

/* ── MAIN VIEW TABS (Dashboard / Branch Data / Schedule) ── */
.main-tabs{display:flex;gap:0;background:#f8fafc;border-bottom:1px solid #e0e4ea;flex-shrink:0;padding:0 16px}
.main-tab{padding:11px 20px;font-size:13px;font-weight:700;cursor:pointer;border:none;background:none;color:#888;border-bottom:3px solid transparent;letter-spacing:.02em;display:flex;align-items:center;gap:6px}
.main-tab.active{color:#1a3a5c;border-bottom-color:#1a3a5c}
.main-tab svg{width:15px;height:15px}

/* ── BODY ── */
.body-area{flex:1;overflow:hidden;display:flex;flex-direction:column;background:#fff}
.view{flex:1;overflow:hidden;display:none;flex-direction:column;padding:14px 16px;gap:10px}
.view.active{display:flex}

/* ── CARDS ── */
.cards{display:grid;gap:10px;flex-shrink:0}
.cards.cols-3{grid-template-columns:repeat(3,1fr)}
.cards.cols-4{grid-template-columns:repeat(4,1fr)}
.cards.cols-5{grid-template-columns:repeat(5,1fr)}
.card{background:#fff;border:1px solid #e5e9ef;border-radius:10px;padding:12px 15px;box-shadow:0 1px 2px rgba(0,0,0,.03)}
.card .num{font-size:26px;font-weight:700;line-height:1.1}
.card .lbl{font-size:10px;color:#888;margin-top:3px;text-transform:uppercase;letter-spacing:.05em;font-weight:600}
.card.card-sm{padding:9px 12px}
.card.card-sm .num{font-size:18px}
.card.card-sm .lbl{font-size:9px;margin-top:1px}
.card.card-sm .hint{font-size:9px;margin-top:2px;margin-bottom:0;line-height:1.3}
.card.c-blue{border-left:4px solid #2980b9}.card.c-blue .num{color:#1a3a5c}
.card.c-green{border-left:4px solid #27ae60}.card.c-green .num{color:#1a6b2e}
.card.c-amber{border-left:4px solid #f39c12}.card.c-amber .num{color:#8a6500}
.card.c-red{border-left:4px solid #e74c3c}.card.c-red .num{color:#a32020}
.card.c-purple{border-left:4px solid #8e44ad}.card.c-purple .num{color:#5b2c6f}

/* ── RISK BADGES ── */
.risk-pill{display:inline-block;padding:3px 10px;border-radius:10px;font-size:11px;font-weight:700}
.risk-low{background:#e3f6e8;color:#1a6b2e}
.risk-medium{background:#fff3cd;color:#8a6500}
.risk-high{background:#fde8e8;color:#a32020}
.risk-none{background:#f0f1f3;color:#999}

/* ── FILTER ROW ── */
.filter-row{display:flex;gap:7px;align-items:center;flex-wrap:wrap;flex-shrink:0}
.filter-row select,.filter-row input{font-size:12px;padding:6px 9px;border:1px solid #d8dde5;border-radius:6px;background:#fff;color:#222}
.filter-row select:focus,.filter-row input:focus{outline:none;border-color:#2980b9;box-shadow:0 0 0 2px rgba(41,128,185,.12)}
.f-search{flex:1;min-width:160px}

/* ── BUTTONS ── */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:7px 13px;border-radius:7px;font-size:12px;font-weight:700;cursor:pointer;border:1px solid #d8dde5;background:#fff;color:#333;transition:all .12s;white-space:nowrap}
.btn:hover{background:#f5f8fb;border-color:#2980b9;color:#1a3a5c}
.btn-primary{background:#1a3a5c;color:#fff;border-color:#1a3a5c}.btn-primary:hover{background:#142e4a;color:#fff}
.btn-success{background:#27ae60;color:#fff;border-color:#27ae60}.btn-success:hover{background:#219a52;color:#fff}
.btn-danger{color:#c0392b;border-color:#fadbd8;background:#fff}.btn-danger:hover{background:#fdedec}
.btn-sm{padding:5px 10px;font-size:11px}
.btn-outline{background:transparent}

/* ── TABLE ── */
.table-wrap{flex:1;overflow:auto;background:#fff;border:1px solid #e5e9ef;border-radius:10px;min-height:0}
table{width:100%;border-collapse:collapse;table-layout:auto}
th{background:#1a3a5c;color:#fff;padding:9px 12px;text-align:left;font-size:11px;font-weight:700;position:sticky;top:0;z-index:2;white-space:nowrap;cursor:pointer;user-select:none;letter-spacing:.02em}
th:hover{background:#142e4a}
th.sort-asc::after{content:' \25B2';font-size:9px}
th.sort-desc::after{content:' \25BC';font-size:9px}
td{padding:7px 12px;border-bottom:1px solid #f0f2f5;font-size:12px;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafd}
.td-num{font-weight:700;font-variant-numeric:tabular-nums}
.td-name{max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.num-cell{text-align:right;font-variant-numeric:tabular-nums}

/* ── EMPTY STATE ── */
.empty-state{text-align:center;padding:48px 20px;color:#bbb}
.empty-state svg{display:block;margin:0 auto 12px;opacity:.25}
.empty-state p{font-size:13px;line-height:1.6}

/* ── SIDEBAR-LIKE PANEL FOR BRANCH/UPDATE ── */
.panel{background:#fff;border:1px solid #e5e9ef;border-radius:10px;padding:14px 16px}
.panel h3{font-size:13px;font-weight:700;color:#1a3a5c;margin-bottom:10px;display:flex;align-items:center;gap:6px}
.sec-label{font-size:10px;font-weight:700;color:#999;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px}
label{font-size:11px;font-weight:700;color:#555;display:block;margin-bottom:4px;margin-top:10px}
input[type=text],input[type=number],input[type=date],textarea,select{width:100%;padding:7px 9px;border:1px solid #d8dde5;border-radius:6px;font-size:12px;font-family:Arial,sans-serif;outline:none;background:#fff;color:#222}
input:focus,textarea:focus,select:focus{border-color:#2980b9;box-shadow:0 0 0 2px rgba(41,128,185,.12)}
textarea{height:80px;resize:vertical;font-family:'Courier New',monospace;font-size:11px}
.row2{display:flex;gap:6px}
.row2 > *{flex:1}

/* ── MODAL ── */
.modal-bg{position:fixed;inset:0;background:rgba(20,30,45,.5);z-index:100;display:flex;align-items:center;justify-content:center}
.modal{background:#fff;border-radius:12px;width:440px;max-width:92vw;padding:22px;box-shadow:0 10px 40px rgba(0,0,0,.2);max-height:88vh;overflow-y:auto}
.modal h2{font-size:16px;font-weight:700;color:#1a3a5c;margin-bottom:14px}
.modal-footer{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
.color-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:5px}
.color-swatch{width:26px;height:26px;border-radius:50%;cursor:pointer;border:3px solid transparent;transition:transform .1s}
.color-swatch.selected{border-color:#1a3a5c;transform:scale(1.15)}

/* ── PROGRESS BAR (data completeness) ── */
.progress-wrap{background:#f0f2f5;border-radius:6px;height:8px;overflow:hidden;margin-top:4px}
.progress-fill{height:100%;background:linear-gradient(90deg,#27ae60,#2ecc71);border-radius:6px;transition:width .3s}
.progress-fill.warn{background:linear-gradient(90deg,#f39c12,#f1c40f)}
.progress-fill.danger{background:linear-gradient(90deg,#e74c3c,#ec7063)}

/* ── HINT/TEXT ── */
.hint{font-size:11px;color:#999;line-height:1.5;margin-top:4px}
.section-title{font-size:14px;font-weight:700;color:#1a3a5c;margin-bottom:4px}

/* ── 2-COL LAYOUT (Branch Data view) ── */
.two-col{display:flex;gap:12px;flex:1;overflow:hidden}
.col-main{flex:1;display:flex;flex-direction:column;gap:10px;overflow:hidden}
.col-side{width:280px;flex-shrink:0;display:flex;flex-direction:column;gap:10px;overflow-y:auto}

/* ── CHIP ── */
.chip-list{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;max-height:90px;overflow-y:auto}
.chip{display:inline-flex;align-items:center;gap:3px;border-radius:10px;font-size:11px;font-weight:600;padding:3px 9px}
.chip button{background:none;border:none;cursor:pointer;font-size:13px;line-height:1;padding:0 0 0 3px;opacity:.7}

/* SCROLLBAR */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:#f4f4f4}
::-webkit-scrollbar-thumb{background:#ccc;border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:#aaa}
</style>
</head>
<body>

<!-- TOP BAR: client tabs -->
<div class="topbar">
  <div class="topbar-logo">&#128202; Inventory &amp; <span>Cycle Count</span></div>
  <div class="client-tabs" id="client-tabs"></div>
  <button class="ctab-add" id="btn-add-client" title="Add new client">+</button>
  <div class="topbar-right">
    <span class="badge b-blue" id="badge-branches">0 branches</span>
    <span class="badge b-green" id="badge-complete">0% complete</span>
    <button class="btn btn-sm btn-outline" id="btn-help" title="User manual" style="margin-left:6px">&#10067; Help</button>
    <button class="btn btn-sm btn-outline" id="btn-restore-builtin" title="Restore the original built-in data for all clients (Client A–D)">&#8635; Restore built-in data</button>
  </div>
</div>

<!-- MAIN VIEW TABS -->
<div class="main-tabs">
  <button class="main-tab active" id="mt-home">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12l9-9 9 9"/><path d="M5 10v10h14V10"/></svg>
    Home
  </button>
  <button class="main-tab" id="mt-dashboard">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="9"/><rect x="14" y="3" width="7" height="5"/><rect x="14" y="12" width="7" height="9"/><rect x="3" y="16" width="7" height="5"/></svg>
    Dashboard
  </button>
  <button class="main-tab" id="mt-branches">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3h18v18H3z"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>
    Branch Master Data
  </button>
  <button class="main-tab" id="mt-schedule">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>
    Schedule &amp; Count
  </button>
</div>

<div class="body-area">

<!-- ════════════════════════ VIEW 0: HOME (all-client overview) ════════════════════════ -->
<div class="view active" id="view-home">

  <div class="cards cols-4" style="flex-shrink:0">
    <div class="card c-blue"><div class="num" id="home-total-clients">0</div><div class="lbl">Clients</div></div>
    <div class="card" style="border-left:4px solid #2980b9"><div class="num" id="home-total-branches">0</div><div class="lbl">Total branches (all clients)</div></div>
    <div class="card" style="border-left:4px solid #8e44ad"><div class="num" id="home-total-qty">0</div><div class="lbl">Total stock QTY (all clients)</div></div>
    <div class="card" style="border-left:4px solid #16a085"><div class="num" id="home-total-cost">฿0</div><div class="lbl">Total stock value (all clients)</div></div>
  </div>

  <div class="cards cols-5" style="margin-top:8px;flex-shrink:0">
    <div class="card c-green">
      <div class="num" id="home-counted-branches">0</div>
      <div class="lbl">Counted branches <span id="home-counted-pct" style="font-weight:700">(0%)</span></div>
    </div>
    <div class="card" style="border-left:4px solid #8e44ad">
      <div class="num" id="home-counted-qty">0</div>
      <div class="lbl">Counted QTY <span id="home-counted-qty-pct" style="font-weight:700">(0%)</span></div>
    </div>
    <div class="card" style="border-left:4px solid #16a085">
      <div class="num" id="home-counted-cost">฿0</div>
      <div class="lbl">Counted value <span id="home-counted-cost-pct" style="font-weight:700">(0%)</span></div>
    </div>
    <div class="card c-amber"><div class="num" id="home-pending-branches">0</div><div class="lbl">Pending count (all clients)</div></div>
    <div class="card" style="border-left:4px solid #e67e22"><div class="num" id="home-avg-diff">0%</div><div class="lbl">Avg. stock difference (all clients)</div></div>
  </div>

  <div class="panel" style="flex-shrink:0">
    <h3>&#128202; By client</h3>
    <div class="table-wrap" style="max-height:320px">
      <table>
        <thead><tr>
          <th>Client</th>
          <th>Total branches</th>
          <th>Total QTY</th>
          <th>Total value</th>
          <th>Counted branches</th>
          <th>Counted QTY</th>
          <th>Counted value</th>
          <th>Avg. diff %</th>
          <th>High risk</th>
        </tr></thead>
        <tbody id="home-client-tbl-body"></tbody>
      </table>
    </div>
  </div>

</div>

<!-- ════════════════════════ VIEW 1: DASHBOARD ════════════════════════ -->
<div class="view" id="view-dashboard">

  <div class="cards cols-3">
    <div class="card card-sm c-blue">
      <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px">
        <div>
          <div class="num" id="dash-total-branches">0</div>
          <div class="lbl">Total branches</div>
        </div>
        <div style="text-align:right;border-left:1px solid #c9dcec;padding-left:8px">
          <div class="num" id="dash-total-selected">0</div>
          <div class="lbl">Selected <span id="dash-selected-pct" style="font-weight:700">(0%)</span></div>
        </div>
      </div>
    </div>

    <div class="card card-sm c-green">
      <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px">
        <div>
          <div class="num" id="dash-counted">0</div>
          <div class="lbl">Counted <span id="dash-counted-pct" style="font-weight:700">(0%)</span></div>
        </div>
        <div style="text-align:right;border-left:1px solid #d9eee0;padding-left:8px">
          <div class="num" id="dash-pending" style="color:#8a6500">0</div>
          <div class="lbl" style="color:#8a6500">Pending <span id="dash-pending-pct" style="font-weight:700">(0%)</span></div>
        </div>
      </div>
      <p class="hint">Of threshold-Selected branches</p>
    </div>

    <div class="card card-sm c-purple">
      <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:6px">
        <div>
          <div class="num" id="dash-qty-coverage">0%</div>
          <div class="lbl">Qty counted</div>
        </div>
        <div style="text-align:right;border-left:1px solid #e0d4ea;padding-left:6px">
          <div class="num" id="dash-cost-coverage">0%</div>
          <div class="lbl">Cost counted</div>
        </div>
      </div>
      <p class="hint" id="dash-coverage-detail">0 / 0 qty &middot; &#3647;0 / &#3647;0</p>
    </div>
  </div>

  <div class="cards cols-3" style="margin-top:8px">
    <div class="card card-sm" style="border-left:4px solid #27ae60">
      <div class="num" id="risk-low-count" style="color:#1a6b2e">0</div>
      <div class="lbl">Low risk (&lt;0.5%) <span id="risk-low-pct" style="font-weight:700">(0%)</span></div>
      <p class="hint">of total Counted</p>
    </div>
    <div class="card card-sm" style="border-left:4px solid #f39c12">
      <div class="num" id="risk-medium-count" style="color:#8a6500">0</div>
      <div class="lbl">Medium risk (0.5–1%) <span id="risk-medium-pct" style="font-weight:700">(0%)</span></div>
      <p class="hint">of total Counted</p>
    </div>
    <div class="card card-sm" style="border-left:4px solid #e74c3c">
      <div class="num" id="risk-high-count" style="color:#a32020">0</div>
      <div class="lbl">High risk (&ge;1%) <span id="risk-high-pct" style="font-weight:700">(0%)</span></div>
      <p class="hint">of total Counted</p>
    </div>
  </div>

  <div class="filter-row">
    <div style="display:flex;gap:7px;flex-wrap:nowrap">
      <select id="dash-f-risk">
        <option value="">All risk levels</option>
        <option value="none">Not counted</option>
        <option value="Low">Low risk</option>
        <option value="Medium">Medium risk</option>
        <option value="High">High risk</option>
      </select>
      <select id="dash-f-region"><option value="">All regions</option></select>
    </div>
    <input type="text" id="dash-f-search" class="f-search" placeholder="&#128269; Search branch no., name, province…">
    <button class="btn btn-outline btn-sm" id="btn-dash-reset">&#8635; Reset filters</button>
    <button class="btn btn-outline btn-sm" id="btn-export-dashboard">&#8595; Export</button>
  </div>

  <div class="table-wrap">
    <table>
      <thead><tr>
        <th id="dth0">Branch no.</th>
        <th id="dth1" style="min-width:200px">Branch name</th>
        <th id="dth2">Province</th>
        <th id="dth3">Region</th>
        <th id="dth4">Count date</th>
        <th id="dth5">Diff %</th>
        <th id="dth6">Risk level</th>
      </tr></thead>
      <tbody id="dash-tbl-body"></tbody>
    </table>
  </div>
  <div style="font-size:10px;color:#aaa;text-align:right;flex-shrink:0" id="dash-row-count"></div>
</div>

<!-- ════════════════════════ VIEW 2: BRANCH MASTER DATA ════════════════════════ -->
<div class="view" id="view-branches">

  <div class="two-col">
    <div class="col-main">
      <div class="cards cols-4" style="flex-shrink:0">
        <div class="card card-sm c-blue"><div class="num" id="bm-total">0</div><div class="lbl" id="bm-total-lbl">Total branches</div></div>
        <div class="card card-sm c-green">
          <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px">
            <div>
              <div class="num" id="bm-complete">0</div>
              <div class="lbl">Complete <span id="bm-complete-pct" style="font-weight:700">(0%)</span></div>
            </div>
            <div style="text-align:right;border-left:1px solid #d9eee0;padding-left:8px">
              <div class="num" id="bm-incomplete" style="color:#8a6500">0</div>
              <div class="lbl" style="color:#8a6500">Incomplete <span id="bm-incomplete-pct" style="font-weight:700">(0%)</span></div>
            </div>
          </div>
          <div class="progress-wrap"><div class="progress-fill" id="bm-progress" style="width:0%"></div></div>
        </div>
        <div class="card card-sm c-red">
          <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px">
            <div>
              <div class="num" id="bm-avg-qty">0</div>
              <div class="lbl" id="bm-avg-qty-lbl">Avg. qty / branch</div>
            </div>
            <div style="text-align:right;border-left:1px solid #f5d5d5;padding-left:8px">
              <div class="num" id="bm-total-qty">0</div>
              <div class="lbl" id="bm-total-qty-lbl">Total qty</div>
            </div>
          </div>
        </div>
        <div class="card card-sm c-purple">
          <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:8px">
            <div>
              <div class="num" id="bm-avg-balance">฿0</div>
              <div class="lbl" id="bm-avg-balance-lbl">Avg. balance / branch</div>
            </div>
            <div style="text-align:right;border-left:1px solid #e0d4ea;padding-left:8px">
              <div class="num" id="bm-total-balance">฿0</div>
              <div class="lbl" id="bm-total-balance-lbl">Total balance</div>
            </div>
          </div>
        </div>
      </div>

      <div class="filter-row" style="flex-shrink:0;flex-wrap:wrap">
        <div style="display:flex;gap:7px;flex-wrap:nowrap">
          <select id="bm-f-region"><option value="">All regions</option></select>
          <select id="bm-f-complete">
            <option value="">All branches</option>
            <option value="complete">Complete data</option>
            <option value="incomplete">Incomplete data</option>
          </select>
          <select id="bm-f-status">
            <option value="">All statuses</option>
            <option value="selected">&#10003; Selected</option>
            <option value="manual">&#10003; Manually Match</option>
            <option value="excluded">&#10005; Excluded</option>
            <option value="watchlisted">&#128204; Watchlisted</option>
            <option value="none">Not flagged</option>
          </select>
        </div>
        <input type="text" id="bm-f-search" class="f-search" placeholder="&#128269; Search branch no., name, province…">
        <button class="btn btn-primary btn-sm" id="btn-add-branch-master">+ Add branch</button>
        <button class="btn btn-outline btn-sm" id="btn-import-branches">&#8593; Import CSV/XLSX</button>
        <input type="file" id="import-branch-file" accept=".csv,.xlsx" style="display:none">
        <button class="btn btn-outline btn-sm" id="btn-bm-reset" title="Clear all filters">&#8635;</button>
        <button class="btn btn-outline btn-sm" id="btn-export-branches">&#8595; Export</button>
        <button class="btn btn-outline btn-sm" id="btn-handoff-history-branches" title="View handoff history">&#128203;</button>
      </div>

      <div class="table-wrap">
        <table>
          <thead><tr>
            <th id="bth0">Branch no.</th>
            <th id="bth1" style="min-width:200px">Branch name</th>
            <th id="bth2">Province</th>
            <th id="bth3">Region</th>
            <th id="bth4">Data date</th>
            <th id="bth5">Stock qty</th>
            <th id="bth6">Stock balance</th>
            <th id="bth7">Sales qty</th>
            <th id="bth8">Status</th>
            <th style="width:110px">Action</th>
          </tr></thead>
          <tbody id="bm-tbl-body"></tbody>
        </table>
      </div>
      <div style="font-size:10px;color:#aaa;text-align:right;flex-shrink:0" id="bm-row-count"></div>
    </div>

    <div class="col-side">
      <div class="panel">
        <h3>&#128221; Bulk update</h3>
        <p class="hint">Paste rows or upload a file to update stock/sales data for existing branches (matched by branch no.) or add new ones.</p>
        <label>Format: Branch No., Name, Province, Region, Data Date, Stock Qty, Stock Balance, Sales Qty</label>
        <textarea id="bm-bulk-text" placeholder="624, TESCO BANGPOO, Bangkok, Bangkok, 2026-06-01, 1200, 45000, 300"></textarea>
        <div class="row2" style="margin-top:6px">
          <button class="btn btn-outline btn-sm" id="btn-bulk-csv-pick">&#8593; CSV</button>
          <button class="btn btn-primary btn-sm" id="btn-bulk-apply">&#10003; Apply update</button>
        </div>
        <input type="file" id="bulk-branch-csv" accept=".csv,.txt" style="display:none">
      </div>

      <div class="panel">
        <h3>&#127919; Selection threshold <span id="bm-threshold-stale-badge" class="badge b-amber" style="margin-left:4px;display:none">&#9888; Out of date</span></h3>
        <p class="hint">Minimum stock balance = <strong>Avg. stock balance / branch &times; haircut %</strong>. Branches at or above that amount get auto-flagged for counting.</p>
        <label>Avg. stock balance / branch</label>
        <input type="text" id="bm-threshold-avg-display" readonly style="background:#f5f7fa;color:#555;font-weight:700">
        <label>Haircut % of average <span style="font-weight:400;color:#999">(100% = use average as-is)</span></label>
        <input type="number" id="bm-threshold-haircut" placeholder="e.g. 100" step="5" min="0" value="100">
        <label>Resulting minimum stock balance (THB)</label>
        <input type="text" id="bm-threshold-computed" readonly style="background:#eef6ff;color:#1a3a5c;font-weight:700">
        <div class="row2" style="margin-top:8px">
          <button class="btn btn-primary btn-sm" id="btn-apply-threshold">Apply &amp; select</button>
          <button class="btn btn-outline btn-sm" id="btn-clear-threshold-sel">Clear selection</button>
        </div>
        <div class="hint" id="bm-threshold-result" style="margin-top:6px"></div>
      </div>

      <div class="panel">
        <h3>&#128683; Excluded branches <span id="ex-count-badge" class="badge b-red" style="margin-left:4px">0</span></h3>
        <p class="hint">Branches manually excluded from counting (e.g. WH/DC, online warehouse), with the reason on file.</p>
        <div id="ex-list" style="margin-top:6px;max-height:160px;overflow-y:auto"></div>
      </div>

      <div class="panel">
        <h3>&#9888; Data completeness</h3>
        <p class="hint" id="bm-incomplete-hint">All required fields present.</p>
        <div id="bm-missing-fields" style="margin-top:6px"></div>
      </div>
    </div>
  </div>
</div>

<!-- ════════════════════════ VIEW 3: SCHEDULE & COUNT ════════════════════════ -->
<div class="view" id="view-schedule">

  <div class="cards cols-5" style="flex-shrink:0">
    <div class="card c-blue"><div class="num" id="sch-total">0</div><div class="lbl">Total visits scheduled</div></div>
    <div class="card" style="border-left:4px solid #2980b9"><div class="num" id="sch-sel">0</div><div class="lbl">Threshold-Selected branches</div></div>
    <div class="card c-green"><div class="num" id="sch-match">0</div><div class="lbl">Matched visits</div></div>
    <div class="card" style="border-left:4px solid #8e44ad"><div class="num" id="sch-counted">0</div><div class="lbl">Counted branches</div></div>
    <div class="card c-amber"><div class="num" id="sch-noshow">0</div><div class="lbl">Not in schedule</div></div>
  </div>

  <div class="two-col">
    <div class="col-main">
      <div class="filter-row" style="flex-shrink:0;flex-wrap:wrap">
        <div style="display:flex;gap:7px;flex-wrap:nowrap">
          <select id="sch-f-month"><option value="">All months</option></select>
          <select id="sch-f-week"><option value="">All weeks</option></select>
          <select id="sch-f-region">
            <option value="">All regions</option>
            <option>Bangkok</option><option>Central</option><option>Northeast</option>
            <option>North</option><option>East</option><option>South</option><option>West</option>
          </select>
          <select id="sch-f-status">
            <option value="">All rows</option>
            <option value="match">Matched only</option>
            <option value="nomatch">Not matched</option>
            <option value="watchlisted">&#128204; Watchlisted only</option>
          </select>
        </div>
        <input type="text" id="sch-f-search" class="f-search" placeholder="&#128269; Search…">
        <button class="btn btn-outline btn-sm" id="btn-sch-reset">&#8635;</button>
        <button class="btn btn-success btn-sm" id="btn-sch-export-filtered">&#8595; Export (filtered view)</button>
        <button class="btn btn-outline btn-sm" id="btn-handoff-history-schedule" title="View handoff history">&#128203;</button>
      </div>

      <div class="table-wrap">
        <table>
          <thead><tr>
            <th id="sth0">Month</th><th id="sth1">Week</th><th id="sth2">Date</th><th id="sth3">Day</th>
            <th id="sth4">Branch no.</th><th id="sth5" style="min-width:180px">Branch name</th>
            <th id="sth6">Area</th><th id="sth7">Region</th><th id="sth8">Time</th><th id="sth9">Status</th>
            <th id="sth10" style="min-width:110px">Count date</th><th id="sth11" style="min-width:90px">Diff %</th><th id="sth12">Risk level</th>
            <th style="width:40px">Action</th>
          </tr></thead>
          <tbody id="sch-tbl-body"></tbody>
        </table>
      </div>
      <div style="font-size:10px;color:#aaa;text-align:right;flex-shrink:0" id="sch-row-count"></div>
    </div>

    <div class="col-side">
      <div class="panel">
        <h3>&#128197; Schedule files</h3>
        <div class="hint" id="sch-months-summary" style="margin-bottom:8px">No schedule loaded</div>
        <button class="btn btn-outline btn-sm" id="btn-sch-upload" style="width:100%">&#8593; Upload CSV/XLSX</button>
        <input type="file" id="sch-file-input" accept=".csv,.xlsx" multiple style="display:none">
        <div id="sch-month-selection-bar" style="display:none;align-items:center;gap:6px;margin-top:8px;font-size:10px">
          <button class="btn btn-outline btn-sm" id="btn-sch-select-all-months" style="padding:3px 7px;font-size:10px">Select all</button>
          <button class="btn btn-outline btn-sm" id="btn-sch-deselect-all-months" style="padding:3px 7px;font-size:10px">Deselect all</button>
          <span id="sch-month-sel-count" style="flex:1;color:#888">None selected</span>
          <button class="btn btn-danger btn-sm" id="btn-sch-remove-selected-months" style="padding:3px 7px;font-size:10px" disabled>&#10005; Remove</button>
        </div>
        <div id="sch-month-list" style="margin-top:8px;display:flex;flex-direction:column;gap:4px"></div>
        <button class="btn btn-outline btn-sm" id="btn-sch-load-sample" style="width:100%;margin-top:8px">&#9654; Load Client A sample (May–Jul 2026)</button>
      </div>

      <div class="panel">
        <h3>&#9999;&#65039; Manual schedule entry</h3>
        <p class="hint">For special counts not in any uploaded schedule file. Tagged <strong>Manual</strong> &mdash; kept out of the month list above. Branch name is pulled from Branch Master Data automatically.</p>
        <div class="row2">
          <input type="text" id="man-branchno" placeholder="Branch no.">
          <input type="date" id="man-date">
        </div>
        <div class="hint" id="man-branchname-preview" style="margin-top:4px;margin-bottom:0;min-height:14px"></div>
        <div class="row2" style="margin-top:5px">
          <select id="man-counttime">
            <option value="Morning">Morning</option>
            <option value="Night">Night</option>
          </select>
        </div>
        <div class="row2" style="margin-top:5px">
          <input type="text" id="man-area" placeholder="Area code (optional)">
          <input type="text" id="man-region" placeholder="Region (optional)">
        </div>
        <button class="btn btn-primary btn-sm" id="btn-man-add" style="width:100%;margin-top:6px">+ Add manual entry</button>

        <div style="margin-top:10px;border-top:1px solid #f0f2f5;padding-top:10px">
          <label style="margin-top:0">Bulk paste (one per line)</label>
          <p class="hint" style="margin-top:0">Format: Branch No., Date (YYYY-MM-DD), Count Time, Area, Region. Branch name is looked up automatically.</p>
          <textarea id="man-bulk-text" placeholder="999, 2026-03-15, Morning, BKK-99, Bangkok"></textarea>
          <button class="btn btn-outline btn-sm" id="btn-man-bulk-add" style="width:100%;margin-top:5px">+ Add all manual entries</button>
        </div>
      </div>

      <div class="panel">
        <h3>&#128202; Detailed count entry</h3>
        <p class="hint">Count date and Diff % are entered directly in the table above. Use this form only if you also need to log book/actual quantities for a branch.</p>
        <button class="btn btn-outline btn-sm" id="btn-record-count" style="width:100%">+ Detailed entry form</button>
      </div>

      <div class="panel">
        <h3>&#128221; Watchlist <span id="watchlist-count-badge" class="badge b-blue" style="margin-left:4px">0</span></h3>
        <p class="hint">Flag branches you want to keep an eye on \u2014 e.g. follow up after last count, client asked about it, recurring discrepancy. Watchlisted branches are highlighted with a &#128204; pin in the Schedule and Branch Master tables. This list is separate from Selected/Matched status and doesn\u2019t affect any matching logic.</p>
        <div class="row2">
          <input type="text" id="wl-inp-no" placeholder="Branch no.">
          <input type="text" id="wl-inp-note" placeholder="Note (optional)">
        </div>
        <button class="btn btn-primary btn-sm" id="btn-wl-add" style="width:100%;margin-top:6px">&#128204; Add to watchlist</button>
        <div id="watchlist-list" style="margin-top:8px;max-height:220px;overflow-y:auto"></div>
        <button class="btn btn-danger btn-sm" id="btn-wl-clear" style="width:100%;margin-top:6px">&#10005; Clear watchlist</button>
      </div>
    </div>
  </div>
</div>

</div><!-- /body-area -->

<!-- ═══ MODALS ═══ -->

<!-- Count entry modal -->
<div class="modal-bg" id="count-modal-bg" style="display:none">
  <div class="modal">
    <h2>Record count result</h2>
    <label>Branch number *</label>
    <input type="text" id="cm-branchno" placeholder="e.g. 624">
    <label>Branch name</label>
    <input type="text" id="cm-branchname" placeholder="auto-filled if found">
    <label>Count date *</label>
    <input type="date" id="cm-date">
    <label>Book quantity (system stock)</label>
    <input type="number" id="cm-bookqty" placeholder="optional — system stock qty" step="0.01">
    <label>Actual counted quantity</label>
    <input type="number" id="cm-actualqty" placeholder="optional — physical count qty" step="0.01">
    <label>Difference % *</label>
    <input type="number" id="cm-diff" placeholder="e.g. 0.75" step="0.01" min="0" max="100">
    <div id="cm-diff-preview" style="display:none;margin-top:8px;padding:8px 12px;border-radius:8px;font-size:13px;font-weight:700;text-align:center"></div>
    <div class="modal-footer">
      <button class="btn" id="btn-cm-cancel">Cancel</button>
      <button class="btn btn-primary" id="btn-cm-save">Save</button>
    </div>
  </div>
</div>

<!-- Add/edit branch modal -->
<div class="modal-bg" id="branch-modal-bg" style="display:none">
  <div class="modal">
    <h2 id="bmodal-title">Add branch</h2>
    <label>Branch number *</label>
    <input type="text" id="bm-no" placeholder="e.g. 624">
    <label>Branch name *</label>
    <input type="text" id="bm-name" placeholder="e.g. TESCO BANGPOO">
    <label>Province</label>
    <input type="text" id="bm-province" placeholder="e.g. Bangkok">
    <label>Region</label>
    <select id="bm-region">
      <option value="">— select —</option>
      <option>Bangkok</option><option>Central</option><option>Northeast</option>
      <option>North</option><option>East</option><option>South</option><option>West</option>
      <option>WH/DC</option><option>Online</option><option>Others</option>
    </select>
    <label>Data date</label>
    <input type="date" id="bm-datadate">
    <div class="row2">
      <div><label>Stock quantity</label><input type="number" id="bm-stockqty" step="0.01"></div>
      <div><label>Stock balance (THB)</label><input type="number" id="bm-stockbal" step="0.01"></div>
    </div>
    <div class="row2">
      <div><label>Sales quantity</label><input type="number" id="bm-salesqty" step="0.01"></div>
    </div>
    <div class="modal-footer">
      <button class="btn" id="btn-bm-cancel">Cancel</button>
      <button class="btn btn-danger btn-sm" id="btn-bm-delete" style="display:none;margin-right:auto">Delete</button>
      <button class="btn btn-primary" id="btn-bm-save">Save</button>
    </div>
  </div>
</div>


<!-- Exclude branch modal -->
<div class="modal-bg" id="exclude-modal-bg" style="display:none">
  <div class="modal">
    <h2>Exclude branch from counting</h2>
    <p class="hint" style="margin-bottom:8px">Branches like warehouses, distribution centers, or online-only stock locations often can\u2019t be physically counted the same way. Excluding a branch removes it from threshold-Selected status, with a note kept for reference.</p>
    <label>Branch</label>
    <div id="ex-branchno-display" style="font-weight:700;color:#1a3a5c;padding:6px 0"></div>
    <label>Reason / note *</label>
    <textarea id="ex-note" placeholder="e.g. WH/DC \u2014 distribution center, not a sellable branch&#10;e.g. Online warehouse \u2014 stock held for e-commerce fulfillment only"></textarea>
    <div class="modal-footer">
      <button class="btn" id="btn-ex-cancel">Cancel</button>
      <button class="btn btn-danger btn-sm" id="btn-ex-save">Exclude branch</button>
    </div>
  </div>
</div>

<!-- Add/edit client modal -->
<div class="modal-bg" id="client-modal-bg" style="display:none">
  <div class="modal">
    <h2 id="client-modal-title">Add new client</h2>
    <label>Client name</label>
    <input type="text" id="cl-name" placeholder="e.g. Client B">
    <label>Short label (tab)</label>
    <input type="text" id="cl-label" placeholder="e.g. Client B" maxlength="18">
    <label>Tab colour</label>
    <div class="color-row" id="cl-color-row"></div>
    <div class="modal-footer">
      <button class="btn" id="btn-cl-cancel">Cancel</button>
      <button class="btn btn-primary" id="btn-cl-save">Save</button>
    </div>
  </div>
</div>

<!-- Handoff Note modal -->
<div class="modal-bg" id="handoff-modal-bg" style="display:none">
  <div class="modal">
    <h2>&#128228; Handoff note</h2>
    <p class="hint" style="margin-bottom:8px">Stamp this export with who sent it and what the next person should know. This gets added as a cover sheet in the file, and logged here so you can see the handoff history.</p>
    <label>Your name *</label>
    <input type="text" id="ho-from" placeholder="e.g. Mim">
    <label>For (optional)</label>
    <input type="text" id="ho-to" placeholder="e.g. Praew">
    <label>Note for the next person</label>
    <textarea id="ho-note" placeholder="e.g. Branch 358 still needs sales qty filled in. Threshold applied at 100%. Please count the Selected branches and send back by Friday."></textarea>
    <div class="modal-footer">
      <button class="btn" id="btn-ho-cancel">Cancel</button>
      <button class="btn btn-primary" id="btn-ho-confirm">Export with note</button>
    </div>
  </div>
</div>

<!-- Handoff History modal -->
<div class="modal-bg" id="handoff-history-modal-bg" style="display:none">
  <div class="modal" style="width:560px;max-height:80vh">
    <h2>&#128203; Handoff history <span id="ho-history-client-name" style="font-weight:400;color:#888;font-size:13px"></span></h2>
    <div id="ho-history-list" style="overflow-y:auto;max-height:calc(80vh - 110px);margin-top:8px"></div>
    <div class="modal-footer">
      <button class="btn btn-primary" id="btn-ho-history-close">Close</button>
    </div>
  </div>
</div>

<!-- Help / User Manual modal -->
<div class="modal-bg" id="help-modal-bg" style="display:none">
  <div class="modal" style="width:680px;max-height:85vh">
    <h2>&#128218; User Manual</h2>
    <div style="overflow-y:auto;max-height:calc(85vh - 110px);padding-right:6px;font-size:12px;line-height:1.6;color:#333">

      <h3 style="color:#1a3a5c;font-size:13px;margin-top:4px;margin-bottom:6px">&#127968; Home</h3>
      <p style="margin-bottom:10px">A read-only overview across <strong>every client at once</strong>: total branches, total stock quantity, and total stock value, plus how much of that has been counted (in branches, quantity, and value) with percentages. The "By client" table breaks every number down per client \u2014 click a row to jump straight into that client's Dashboard.</p>

      <h3 style="color:#1a3a5c;font-size:13px;margin-bottom:6px">&#128202; Dashboard</h3>
      <p style="margin-bottom:4px">The monitoring view for the <strong>currently selected client</strong> (use the coloured tabs at the top to switch clients).</p>
      <ul style="margin:0 0 10px 18px;padding:0">
        <li><strong>Total branches | Selected</strong> \u2014 how many branches are currently flagged for counting (Threshold-Selected or Manually Match), with % of total.</li>
        <li><strong>Counted | Pending</strong> \u2014 of the Selected branches, how many already have a recorded result vs. still waiting.</li>
        <li><strong>Qty counted % / Cost counted %</strong> \u2014 coverage of total stock quantity and value that's been physically counted so far, across all branches.</li>
        <li><strong>Low / Medium / High risk cards</strong> \u2014 breakdown of recorded results by your risk thresholds (&lt;0.5% Low, 0.5\u20131% Medium, &ge;1% High), as a % of branches actually counted.</li>
        <li>The table lists every branch with its latest count date, diff %, and risk level. Filter by risk level, region, or search by name/number.</li>
      </ul>

      <h3 style="color:#1a3a5c;font-size:13px;margin-bottom:6px">&#128203; Branch Master Data</h3>
      <p style="margin-bottom:4px">The master list of branches for this client \u2014 stock quantity, stock balance (THB), sales quantity, province, and region.</p>
      <ul style="margin:0 0 10px 18px;padding:0">
        <li><strong>Summary cards</strong> \u2014 Total branches, Complete/Incomplete data (with %), Avg./Total stock balance, Avg./Total stock qty. All four react live to whatever filter is applied.</li>
        <li><strong>+ Add branch</strong> / <strong>Import CSV/XLSX</strong> \u2014 add branches one at a time or in bulk. Bulk Update (right panel) lets you paste or upload rows to update existing branches (matched by branch no.) without overwriting fields you leave blank.</li>
        <li><strong>Selection threshold</strong> \u2014 set a haircut % of the average stock balance (e.g. 100% = use the average as-is). Click <strong>Apply &amp; select</strong> to flag every branch at or above that computed minimum as "Selected." This is a one-time snapshot, not a live formula \u2014 if you edit branch data afterward, a small <strong>"&#9888; Out of date"</strong> badge appears on the panel telling you to re-Apply.</li>
        <li><strong>Status column</strong> in the table shows: <span style="color:#1a6b2e">&#10003; Selected</span> (from the threshold), <span style="color:#4a2db5">&#10003; Manually Match</span> (manual override, independent of the formula), <span style="color:#a32020">&#10005; Excluded</span> (can't be counted, e.g. a warehouse), or a &#128204; Watchlist tag (personal follow-up flag, doesn't affect any matching logic).</li>
        <li><strong>Excluded branches panel</strong> \u2014 every excluded branch with its reason on file; re-include any time.</li>
      </ul>

      <h3 style="color:#1a3a5c;font-size:13px;margin-bottom:6px">&#128197; Schedule &amp; Count</h3>
      <p style="margin-bottom:4px">Upload the client's official stocktake schedule, log count results, and track which scheduled visits matter.</p>
      <ul style="margin:0 0 10px 18px;padding:0">
        <li><strong>Schedule files</strong> \u2014 upload CSV/XLSX (columns: Month, Week, Day, Date, Branch No., Branch Name, Area, Region, Count Time). Each month becomes a pill you can reorder (&#9650;&#9660;), select/deselect (checkbox \u2014 ticking months filters the table to just those), or remove (&#10005;).</li>
        <li><strong>Manual schedule entry</strong> \u2014 for special counts not in any uploaded file. Branch name is pulled automatically from Branch Master Data; if the branch isn't registered there yet, the row is tagged <strong>"Not Registered"</strong> as a reminder to add it. Manual entries are tagged <strong>"Manual"</strong> and excluded from the month pill list.</li>
        <li><strong>Duplicate dates</strong> \u2014 if a branch is scheduled twice (e.g. 19-Feb and 20-Feb), only one date is editable: whichever already has a result, or the latest date if both/neither do. The other date shows a "See [date]" tag \u2014 this is how a branch counts as exactly <strong>one</strong> visit, not two.</li>
        <li><strong>Recording a result</strong> \u2014 type directly into the <strong>Count date</strong> and <strong>Diff %</strong> columns in the table; risk level fills in automatically. Use "+ Detailed entry form" only if you also need to log book/actual quantities.</li>
        <li><strong>Deleting a row</strong> \u2014 every schedule row has a &#10005; delete button if an entry needs removing.</li>
        <li><strong>Export</strong> \u2014 the button in the filter row exports exactly what's currently shown in the table (respecting all active filters), including count results. Clicking it opens a <strong>Handoff note</strong> first \u2014 see below.</li>
      </ul>

      <h3 style="color:#1a3a5c;font-size:13px;margin-bottom:6px">&#128228; Handoff notes (working with a colleague)</h3>
      <p style="margin-bottom:4px">This dashboard saves only to <strong>your own browser</strong> \u2014 it doesn't sync live with anyone else. If two people split the work (e.g. one preps data, the other counts), exporting and re-importing a file is how the work actually moves between you.</p>
      <ul style="margin:0 0 10px 18px;padding:0">
        <li>Every <strong>Export</strong> button (Branch Master Data and Schedule &amp; Count) now opens a small <strong>Handoff note</strong> prompt first \u2014 your name, optionally who it's for, and a note for them. This gets added as a cover sheet in the exported file.</li>
        <li>Send the exported .xlsx file however you normally share files \u2014 email, Slack, a shared drive. You do <strong>not</strong> send the HTML dashboard file itself; each person keeps their own copy of that.</li>
        <li>The other person opens their own copy of the dashboard and imports your file: <strong>Bulk Update</strong> on Branch Master Data, or <strong>Upload CSV/XLSX</strong> on Schedule &amp; Count. Existing data is matched and updated, not wiped.</li>
        <li>Click the <strong>&#128203;</strong> button next to either Export button to see the handoff history for the current client \u2014 who sent what, when, and their note.</li>
        <li>This is a <strong>take-turns</strong> workflow, not real-time \u2014 it works best when one person finishes their step before the other starts theirs on the same data.</li>

      <h3 style="color:#1a3a5c;font-size:13px;margin-bottom:6px">&#128081; Status reference</h3>
      <ul style="margin:0 0 10px 18px;padding:0">
        <li><strong>Selected</strong> \u2014 qualifies under the current threshold formula.</li>
        <li><strong>Manually Match</strong> \u2014 manually flagged as matched, independent of the formula. Useful for one-off overrides.</li>
        <li><strong>Excluded</strong> \u2014 can't be physically counted (e.g. WH/DC, online warehouse). Cannot be Selected or Manually Matched at the same time.</li>
        <li><strong>Watchlist</strong> \u2014 a personal reminder pin (&#128204;), purely informational, never affects matching.</li>
        <li><strong>Not Registered</strong> \u2014 a schedule row references a branch number that isn't yet in Branch Master Data. Register it there and the name will need to be re-synced on that row.</li>
      </ul>

      <h3 style="color:#1a3a5c;font-size:13px;margin-bottom:6px">&#128190; Data &amp; clients</h3>
      <ul style="margin:0 0 4px 18px;padding:0">
        <li>Everything saves automatically to this browser (localStorage) \u2014 no manual save needed, but data won't transfer to a different browser or device on its own.</li>
        <li>Click <strong>+</strong> next to the client tabs to add a new client. Each client has fully independent data.</li>
        <li><strong>&#8635; Restore built-in data</strong> (top right) resets everything back to the original 4-client dataset \u2014 use with caution, this can't be undone.</li>
      </ul>

    </div>
    <div class="modal-footer">
      <button class="btn btn-primary" id="btn-help-close">Close</button>
    </div>
  </div>
</div>

<script>
// ── CONSTANTS ─────────────────────────────────────────────────────────────────
const COLORS=['#2980b9','#e74c3c','#27ae60','#e67e22','#8e44ad','#16a085','#d35400','#2c3e50','#c0392b','#1abc9c','#7f8c8d','#f39c12'];
// Compares two "DD-Mon" date labels (e.g. "20-Feb") chronologically, inferring the year
// from context: dates roll from Dec back to Jan are assumed to cross into the following year
// relative to whichever reference month is "latest" seen so far. For schedule sorting purposes
// within a single client's dataset (which spans at most ~12-18 months), comparing by
// (month-index, day) after normalizing Jan-following-Dec as "next year" is sufficient.
const MONTH_INDEX={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
function parseDateLabel(label){
  if(!label) return null;
  const m=String(label).match(/^(\d{1,2})-([A-Za-z]{3})$/);
  if(!m) return null;
  const day=parseInt(m[1],10);
  const mon=MONTH_INDEX[m[2]];
  if(mon===undefined||isNaN(day)) return null;
  return {day, month:mon};
}
// Returns a sortable numeric key. Since exact years aren't stored in the date label,
// this assumes the dataset's months are added in roughly chronological order and uses
// a rolling 24-month window anchored so that, e.g., Jan after Dec is treated as later.
function dateSortKey(label){
  const p=parseDateLabel(label);
  if(!p) return -1;
  return p.month*31+p.day;
}
function compareDateLabels(a,b){ return dateSortKey(a)-dateSortKey(b); }

// Risk thresholds (editable here if criteria change)
const RISK_THRESHOLDS = { low: 0.5, medium: 1.0 }; // <0.5%=Low, <1.0%=Medium, >=1.0%=High

function calcRiskLevel(diffPct){
  if(diffPct===null||diffPct===undefined||diffPct==='') return null;
  const v=Math.abs(parseFloat(diffPct));
  if(v<RISK_THRESHOLDS.low) return 'Low';
  if(v<RISK_THRESHOLDS.medium) return 'Medium';
  return 'High';
}
function riskPillClass(lvl){
  return {Low:'risk-low',Medium:'risk-medium',High:'risk-high'}[lvl]||'risk-none';
}

// ── CLIENT A MASTER BRANCH DATA (from All_Branch_Data.xlsx) ───────────────────
// Format: [no, name, province, region, dataDate, stockQty, stockBalance, salesQty]
const CLIENT_A_BRANCHES = [
["201","MANEEYA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2486.0,92245.18,670],
["268","PASSIONE SHOPPING DESTINATION","ระยอง","ภาคตะวันออก","2026-05-24",2505.0,87497.0,868],
["517","MUANGTHONG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2523.0,82657.11,1170],
["531","BIG C SAMUTSONGKRAM","สมุทรสงคราม","ภาคตะวันตก","2026-05-24",1707.0,64211.22,683],
["599","SIAM SQUARE 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3395.0,114778.87,2012],
["598","THE MALL KORAT FL:2","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2352.0,86433.86,568],
["760","BIG C BANDUNG","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1773.0,63516.3,666],
["769","TAWANNA BANGKAPI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1600.0,60302.17,469],
["775","Lotus's Thaluang Phichit","พิจิตร","ภาคเหนือ","2026-05-24",1419.0,53751.23,512],
["790","BIG C SATTAHIP","ชลบุรี","ภาคตะวันออก","2026-05-24",1517.0,57536.34,901],
["810","TESCO BANPHUE","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1706.0,61060.61,463],
["890","SIAM CENTER","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2469.0,92808.68,744],
["975","TESCO RAMINDRA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1681.0,61977.49,453],
["991","TESCO CHIANGYUEN","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2036.0,70403.51,372],
["3001","TESCO HATYAI 1","สงขลา","ภาคใต้","2026-05-24",2089.0,74057.09,625],
["3004","TESCO PRAKONCHAI","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1634.0,60277.86,578],
["3020","TESCO BANGKADI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1680.0,59935.01,458],
["216","PACIFIC PARK","ชลบุรี","ภาคตะวันออก","2026-05-24",3053.0,102624.26,1254],
["257","BANG RAK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1919.0,68405.33,470],
["351","GRAND PLAZA NGAMWONGWAN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1975.0,69437.13,346],
["3067","LOTUS UTHUMPHONPHISAI","ศรีสะเกษ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1630.0,59087.74,503],
["3102","MARKET PLACE RANGSIT KLONG 1","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1757.0,60613.2,542],
["3185","The Mall Bangkae MF","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1939.0,69277.75,1014],
["175","คลังคืนสินค้า WH Bangplee","WH/DC","WH/DC","2026-05-24",117.0,3172.56,0],
["155","WH Bangplee","WH/DC","WH/DC","2026-05-24",69898.0,2026363.89,0],
["692","BIG C MAHACHAI","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",1766.0,64279.65,349],
["964","BIG C KHONKAEN 2","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1737.0,62324.33,369],
["3049","MUKDAHAN","มุกดาหาร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2149.0,73170.77,628],
["3043","TESCO SAMCHUK","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",2048.0,72364.18,467],
["742","ROBINSON PETCHBURI","เพชรบุรี","ภาคตะวันตก","2026-05-24",2367.0,87685.74,673],
["3166","Movenpick Chiangmai","เชียงใหม่","ภาคเหนือ","2026-05-24",2386.0,83295.68,378],
["3327","K Avenue Nakornsawan","นครสวรรค์","ภาคเหนือ","2026-05-24",1608.0,55798.8,773],
["3302","Nongbualamphu","หนองบัวลำภู","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1554.0,49654.17,450],
["361","RIEANTHONG MARKET (KLONG 6)","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1911.0,62097.34,723],
["427","BIG C HANG DONG","เชียงใหม่","ภาคเหนือ","2026-05-24",1412.0,54725.97,793],
["436","DIANA HATYAI","สงขลา","ภาคใต้","2026-05-24",1627.0,55947.33,655],
["480","SEACON SQUARE 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1161.0,48006.03,288],
["602","Kad Farang","เชียงใหม่","ภาคเหนือ","2026-05-24",1476.0,56962.95,681],
["663","SOI BEARING","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2181.0,70050.46,347],
["891","SRISAKON MALL","สกลนคร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2028.0,77189.22,459],
["3016","Lotus's Thasala","นครศรีธรรมราช","ภาคใต้","2026-05-24",2155.0,77050.78,542],
["3258","Makro Rangsit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1756.0,63118.21,526],
["3307","Burapha Market","ฉะเชิงเทรา","ภาคตะวันออก","2026-05-24",2062.0,68500.14,500],
["3317","Siri Plaza Market","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2137.0,70254.8,420],
["345","IT SQUARE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1283.0,45119.22,850],
["809","Lotus's Phibunmangsahan","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1788.0,64222.05,494],
["565","TESCO NAKORNNAYOK","นครนายก","ภาคตะวันออก","2026-05-24",2120.0,74809.72,553],
["589","TESCO KLONGLUANG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1569.0,60158.41,657],
["623","TREE ON 3","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2070.0,72657.43,370],
["674","GATEWAY EKAMAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2061.0,72988.17,807],
["767","Lotus's Dankhunthot","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1687.0,61700.76,426],
["904","eStore-Shopee","Online","Online","2026-05-24",1968.0,54375.31,6158],
["872","AMORN CENTER","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1941.0,68276.23,254],
["883","MARKET PLACE DUSIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1644.0,54424.26,442],
["3017","KARON BEACH","ภูเก็ต","ภาคใต้","2026-05-24",1122.0,49875.05,247],
["3126","Homepro Rayong","ระยอง","ภาคตะวันออก","2026-05-24",2740.0,97114.16,214],
["3142","Amnatcharoen","อำนาจเจริญ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2155.0,75503.0,340],
["3195","Mahasarakham University","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2698.0,93771.55,219],
["3228","Market today 64","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1897.0,67135.72,531],
["3270","Det Udom Market","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2109.0,74274.73,217],
["3299","Lotus's Srinakarin","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2130.0,73840.24,808],
["3300","Bangli Suphanburi","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",2871.0,95026.38,338],
["359","RAMKHAMHAENG 24","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1643.0,60071.68,480],
["398","Lotus's Phitsanulok 1","พิษณุโลก","ภาคเหนือ","2026-05-24",1231.0,49363.31,827],
["487","BIG C MUKDAHARN","มุกดาหาร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2187.0,78651.26,397],
["496","DAN NOK","สงขลา","ภาคใต้","2026-05-24",1148.0,49194.68,385],
["554","CP TOWER","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3341.0,116845.14,1324],
["594","BIG C SRISAKET","ศรีสะเกษ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1545.0,53183.77,739],
["765","ROBINSON KAMPHAENGPHET","กำแพงเพชร","ภาคเหนือ","2026-05-24",1882.0,63694.74,788],
["820","Big C Chachoengsao","ฉะเชิงเทรา","ภาคตะวันออก","2026-05-24",1780.0,64741.3,503],
["876","BIG C CHIANGRAI 2","เชียงราย","ภาคเหนือ","2026-05-24",2129.0,77683.52,812],
["878","TESCO BORABUE","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2058.0,72761.8,392],
["893","TESCO DET UDOM","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1903.0,65516.21,477],
["422","CENTRAL CHIANGRAI","เชียงราย","ภาคเหนือ","2026-05-24",2410.0,91054.87,967],
["426","TESCO  SATUN","สตูล","ภาคใต้","2026-05-24",1463.0,53315.51,742],
["495","CURVE CHIANGMAI","เชียงใหม่","ภาคเหนือ","2026-05-24",1877.0,67268.43,298],
["514","ZEER RANGSIT 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1729.0,62704.0,720],
["600","UD TOWN 2","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1832.0,64476.99,487],
["3012","BIG C SOUTH PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",1882.0,68489.06,467],
["3151","Lotus’s Sikhio","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1680.0,60517.26,487],
["3197","True Digital Park","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2517.0,86546.86,768],
["3247","Chumphae","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1732.0,63295.35,553],
["254","TESCO RAMA  III","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1744.0,64668.35,620],
["442","TERMINAL 21","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2361.0,80282.08,1703],
["479","BIG C LUMLUKKA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1787.0,64727.07,426],
["604","ESPLANADE RATCHADA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2119.0,72616.42,583],
["610","Big C Suwinthawong","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1480.0,55710.08,558],
["638","THE BRIGHT RAMA 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1851.0,62313.94,169],
["669","BIG C BURIRUM","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1697.0,62677.14,371],
["654","TAKUAPA PHANGNGA","พังงา","ภาคใต้","2026-05-24",1936.0,67734.06,891],
["899","BIG C NAKHON PHANOM","นครพนม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1888.0,65596.6,608],
["973","METRO MALL CHATUCHAK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1710.0,61988.48,768],
["3019","BIG C NAM DAENG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2436.0,85046.85,186],
["3060","BIG C RAMA 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1958.0,70960.1,605],
["3132","The Seasons Mall","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1708.0,60890.71,617],
["3313","Lotus’s Sriracha","ชลบุรี","ภาคตะวันออก","2026-05-24",3751.0,135808.37,741],
["224","ROBINSON CHANTABURI","จันทบุรี","ภาคตะวันออก","2026-05-24",2748.0,99780.68,746],
["362","UDOMSUK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2062.0,68468.4,607],
["804","Lotus's Nakorn-In","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1723.0,62900.29,705],
["522","MAX VALUE SUKHUMVIT 71","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1558.0,53183.47,307],
["576","ROBINSON DEPARTMENT ROI ET","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2375.0,88332.65,735],
["3061","METRO MALL PHRA RAMA 9","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2008.0,71889.68,503],
["3066","CENTRAL SRIRACHA","ชลบุรี","ภาคตะวันออก","2026-05-24",2451.0,87844.48,772],
["3080","BIG C THATAKO","นครสวรรค์","ภาคเหนือ","2026-05-24",1799.0,64092.71,425],
["3190","Lotus's Mahachai","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",1780.0,64833.61,608],
["905","Online","Online","Online","2026-05-24",3375.0,65507.77,23090],
["488","THE SCENE TOWN IN TOWN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1596.0,58012.2,559],
["894","FUTURE PARK RANGSIT BF","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2300.0,86245.91,1042],
["3014","BIG C THA IT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1869.0,66657.85,503],
["699","THE PLATFORM@WONG WIEN YAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1895.0,68121.34,344],
["572","PINKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1901.0,67670.59,267],
["871","MAX VALU LAKSI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2023.0,71172.8,256],
["3193","PTT Angsila","ชลบุรี","ภาคตะวันออก","2026-05-24",2221.0,76670.09,393],
["3203","DD Mache Market","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1962.0,69412.24,306],
["3269","Wanon Niwat","สกลนคร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2361.0,84244.56,425],
["3301","Muak Lek Saraburi","สระบุรี","ภาคกลาง","2026-05-24",2508.0,83314.44,640],
["3041","WORLD MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1840.0,67591.25,628],
["3056","TRANG HOSPITAL","ตรัง","ภาคใต้","2026-05-24",1950.0,69108.18,389],
["3157","Kimyong Market","สงขลา","ภาคใต้","2026-05-24",1733.0,60865.13,534],
["3207","Metro Mall Queen Sirikit Center","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1911.0,69782.96,588],
["339","Taveekij Burirum","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1778.0,65195.97,785],
["575","PLERNNARY MALL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1865.0,61575.77,101],
["424","BIG C HAT YAI","สงขลา","ภาคใต้","2026-05-24",2760.0,91007.51,785],
["603","FUTURE PARK RANGSIT FL:2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2056.0,73109.59,1424],
["743","JUNGCEYLON PHUKET","ภูเก็ต","ภาคใต้","2026-05-24",2456.0,88280.13,1021],
["889","BIG C SAIMAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2152.0,76921.92,200],
["3021","TESCO SAMUTPRAKARN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1858.0,67335.12,626],
["3174","Suanluang Krathumbaen","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",2485.0,87656.79,277],
["3224","Asiatique","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3217.0,109243.77,673],
["3089","YASOTHON MARKET","ยโสธร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2180.0,72612.68,512],
["3133","Saphan 4 Market","ระยอง","ภาคตะวันออก","2026-05-24",2010.0,66624.89,620],
["403","PURE PLACE SAMMAKORN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1774.0,63612.02,505],
["803","Lotus's Phanatnikhom","ชลบุรี","ภาคตะวันออก","2026-05-24",1560.0,59183.75,774],
["543","ROBINSON SARABURI","สระบุรี","ภาคกลาง","2026-05-24",2184.0,80494.5,874],
["611","CHIANGRAI NIGHT BAZAAR","เชียงราย","ภาคเหนือ","2026-05-24",2299.0,83937.09,401],
["614","UNION MALL 2_2FL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2372.0,82457.45,753],
["794","Vibhavadee Soi 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1944.0,65125.08,609],
["869","TOPS PLAZA PHON","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1674.0,59753.68,827],
["3087","DIANA PATTANI","ปัตตานี","ภาคใต้","2026-05-24",2091.0,74005.77,379],
["3210","Kangsadan Khonkaen","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2186.0,73958.8,417],
["3215","Porto Go Bang Pa-In","อยุธยา","ภาคกลาง","2026-05-24",1919.0,67209.37,317],
["3219","Satuek Market","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1770.0,62613.05,381],
["3231","La-Ngu","สตูล","ภาคใต้","2026-05-24",1820.0,65103.34,672],
["256","PATONG PHUKET","ภูเก็ต","ภาคใต้","2026-05-24",3561.0,124391.91,466],
["293","MAJOR HOLLYWOOD CHANGWATTANA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1499.0,58307.79,681],
["388","AYUTTHAYA PARK  1","อยุธยา","ภาคกลาง","2026-05-24",2056.0,72218.39,826],
["412","BIG C KHAMPHAENGPETCH","กำแพงเพชร","ภาคเหนือ","2026-05-24",1599.0,60959.56,604],
["445","TESCO KORAT 2","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2037.0,73128.85,346],
["484","J MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1680.0,59708.75,865],
["783","Top Plaza Phayao","พะเยา","ภาคเหนือ","2026-05-24",1731.0,63350.67,633],
["778","BIG C SUVARNABHUMI ROI-ET","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1479.0,56794.9,672],
["3097","LOTUS LIAB KLONG SONG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1660.0,61519.07,383],
["720","EXTERNAL PROMOTION","Others","Others","2026-05-24",1716.0,62462.15,338],
["3304","Phu Kiang Khonkaen","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2528.0,84153.08,329],
["368","WONGSAWANG TOWN CENTER","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1887.0,70521.92,931],
["450","CENTRAL RAMA9","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2038.0,78413.37,1187],
["460","TESCO PRACHINBURI","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",1902.0,68132.85,277],
["467","BIG C PATTAYA KLANG","ชลบุรี","ภาคตะวันออก","2026-05-24",1742.0,64732.04,709],
["588","VICTORIA GARDEN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1843.0,68368.18,558],
["618","CENTRAL WEST GATE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",5045.0,164892.11,2991],
["628","BIG C BANGPLEE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2043.0,75892.09,925],
["863","THAMMASAT HOSPITAL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1199.0,46856.93,484],
["969","TESCO NAKHON SRI THAMMARAT","นครศรีธรรมราช","ภาคใต้","2026-05-24",2179.0,71874.91,467],
["3005","CENTRAL PATTAYA BEACH FL.G","ชลบุรี","ภาคตะวันออก","2026-05-24",1657.0,68648.66,439],
["3105","RANONG MARKET","ระนอง","ภาคใต้","2026-05-24",2112.0,74470.4,382],
["3319","Little Walk Phrannok-Phutthamonthon","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3080.0,106087.28,463],
["3054","THE MALL THAPRA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1784.0,69149.91,885],
["3071","BIG C KRABI","กระบี่","ภาคใต้","2026-05-24",1998.0,71288.23,329],
["3107","POSRI MARKET","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2099.0,69905.0,484],
["3128","Sentosa Maliwan Khonkaen","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1629.0,60028.55,334],
["3254","Thungsong Market","นครศรีธรรมราช","ภาคใต้","2026-05-24",2363.0,79185.13,377],
["3292","Green Park Chiangmai","เชียงใหม่","ภาคเหนือ","2026-05-24",3562.0,128471.11,431],
["315","COLISEUM YALA","ยะลา","ภาคใต้","2026-05-24",2064.0,69844.74,321],
["349","ZEER RANGSIT 1","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2017.0,65865.85,406],
["374","BIG C LOPBURI","ลพบุรี","ภาคกลาง","2026-05-24",2174.0,78582.44,1022],
["245","LOTUS SUKUMVIT 50","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1899.0,68148.28,1235],
["253","BIG C SAMUTPRAKARN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1746.0,62670.3,858],
["584","TESCO NAN","น่าน","ภาคเหนือ","2026-05-24",1700.0,62347.27,561],
["641","THE MALL KORAT (B FLOOR)","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2184.0,81508.14,449],
["440","Lotus's Tak","ตาก","ภาคเหนือ","2026-05-24",1561.0,58684.96,643],
["456","HUAY KWANG MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2124.0,72206.34,448],
["579","ROBINSON SAMUTPRAKARN FL:2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1403.0,52246.65,1727],
["346","CHAN ROAD","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1598.0,58646.8,489],
["400","SEACON SQUARE 1","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2924.0,102596.55,1836],
["482","CENTRAL HUAMARK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2186.0,79820.41,952],
["476","CENTRAL BANGNA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3207.0,136855.32,605],
["762","CENTRAL MAHACHAI","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",2483.0,89457.06,1016],
["792","THE ONE PARK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1955.0,69409.98,257],
["3158","Mark Four Plaza","แพร่","ภาคเหนือ","2026-05-24",1870.0,65384.29,660],
["3274","Srifah Market","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1635.0,59416.09,473],
["3287","Central Village","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2237.0,80019.74,513],
["680","TESCO PHAYAKKHAPHUM PHISAI","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1929.0,64710.68,430],
["781","BIG C BANBUENG","ชลบุรี","ภาคตะวันออก","2026-05-24",1771.0,65013.8,387],
["814","TERMINAL 21 PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",2423.0,90299.78,928],
["961","Lotus's Chonburi","ชลบุรี","ภาคตะวันออก","2026-05-24",1676.0,56170.17,340],
["343","SF PETCHAKASEM","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1982.0,66180.12,704],
["407","Lotus's Chumporn","ชุมพร","ภาคใต้","2026-05-24",2431.0,87816.92,580],
["414","TESCO SAMPARN","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1788.0,67827.06,377],
["498","CENTRAL LAMPANG","ลำปาง","ภาคเหนือ","2026-05-24",2152.0,81434.54,1313],
["509","Lotus's Maesod","ตาก","ภาคเหนือ","2026-05-24",1492.0,57976.73,584],
["528","PURE RATCHAPRUEK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2069.0,73686.96,301],
["3042","I'M CHINATOWN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2074.0,74735.71,667],
["3110","THE FOURTH SAI 4","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1744.0,63876.08,689],
["3186","Village Hub Prachauthit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1809.0,64344.47,640],
["3279","Lotus's Kaeng Khoi","สระบุรี","ภาคกลาง","2026-05-24",1754.0,62298.25,395],
["570","Central Salaya","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1980.0,75471.37,814],
["633","TESCO WIENGSA","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1505.0,57250.27,708],
["756","CENTRAL NAKRONRATCHASRIMA 3F","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2357.0,85581.04,680],
["896","Lotus's Banglen","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1451.0,55690.65,497],
["438","TESCO PRACHAUTHIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2020.0,71681.44,237],
["3023","BIG C SAMRONG 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2079.0,74949.41,345],
["3078","JAS GREEN VILLAGE KHUBON","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1861.0,69447.45,411],
["3091","RUNGROJ MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2236.0,79433.36,226],
["3101","SAIMAI MALL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1644.0,61369.62,828],
["3149","Pattaya Avenue","ชลบุรี","ภาคตะวันออก","2026-05-24",2315.0,81670.72,465],
["3239","Kanchanawanit Songkhla","สงขลา","ภาคใต้","2026-05-24",2203.0,78299.74,391],
["3278","Central Park Dusit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3086.0,115111.77,697],
["3295","Lotus's Onnut 80","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2317.0,80743.25,332],
["3296","Loeng Nok Tha","ยโสธร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",4515.0,162552.14,501],
["3312","Central Krabi","กระบี่","ภาคใต้","2026-05-24",2536.0,94680.1,714],
["763","BIG C SURATTHANI","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2066.0,72366.33,388],
["812","CENTRAL PHUKET FLORESTA","ภูเก็ต","ภาคใต้","2026-05-24",2471.0,86019.15,808],
["816","INDIGO PATONG","ภูเก็ต","ภาคใต้","2026-05-24",2530.0,91884.39,355],
["3079","THE SENSE PINKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2048.0,74229.57,381],
["3098","LOTUS TIWANON","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1966.0,70041.06,245],
["3196","Lotus's Ranong","ระนอง","ภาคใต้","2026-05-24",2046.0,73252.08,454],
["976","ROBINSON LADKRABANG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2503.0,85122.2,1225],
["985","ROBINSON TRANG 1F","ตรัง","ภาคใต้","2026-05-24",3468.0,110553.04,1338],
["3009","THE MALL NGAMWONGWAN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2204.0,79537.98,1251],
["3063","MARKET PLACE KRUNGTHEP KREETHA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2010.0,72212.9,262],
["3276","Bang Prong","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1880.0,66141.42,349],
["3281","Government Complex Chaengwattana","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2014.0,68013.98,592],
["3240","Tilokuthit Phuket","ภูเก็ต","ภาคใต้","2026-05-24",2909.0,102907.29,824],
["3284","Udomsuk Market Kabinburi","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",1743.0,63467.69,540],
["881","TESCO  BANSUAN","ชลบุรี","ภาคตะวันออก","2026-05-24",1726.0,57054.63,361],
["485","SEACON BANGKAE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2492.0,81340.73,900],
["3321","Betong Yala","ยะลา","ภาคใต้","2026-05-24",1763.0,59363.73,1091],
["3308","Kanlapapruek Market","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1936.0,64936.94,518],
["630","PHETCHABURI","เพชรบุรี","ภาคตะวันตก","2026-05-24",1149.0,51143.08,201],
["631","SUPREME SAMSEN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2043.0,71898.47,634],
["698","CITY PARK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1867.0,64449.51,516],
["3121","Topland Phitsanulok","พิษณุโลก","ภาคเหนือ","2026-05-24",1627.0,61159.23,546],
["289","SAVELAND BANGSEAN","ชลบุรี","ภาคตะวันออก","2026-05-24",2153.0,73445.81,564],
["532","BANPONG  RATCHABURI","ราชบุรี","ภาคตะวันตก","2026-05-24",1036.0,40684.57,300],
["560","TESCO NAKORNPANOM","นครพนม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1127.0,47536.63,302],
["561","CENTRAL SAMUI","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2152.0,73741.18,262],
["636","FASHION ISLAND 4","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2514.0,92251.82,1243],
["689","Lotus's Phatthalung","พัทลุง","ภาคใต้","2026-05-24",1981.0,72767.43,540],
["995","AO NANG KRABI 2","กระบี่","ภาคใต้","2026-05-24",2483.0,92896.41,803],
["3002","COLISEUM PHATTHALUNG","พัทลุง","ภาคใต้","2026-05-24",1480.0,56356.54,670],
["3075","MARUAY HATHAIRAT MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2321.0,75681.45,494],
["3123","Sahathai Garden Plaza Surat","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2028.0,72671.36,435],
["3188","Phu Khiao Chaiyaphum","ชัยภูมิ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1778.0,60290.55,455],
["465","ISARAPHAP ROAD","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1782.0,65966.2,788],
["993","JAZ URBAN SRINAKARIN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1905.0,66813.47,512],
["418","BIG C RAJDAMRI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1856.0,72498.2,478],
["776","Big C Wichienburi","เพชรบูรณ์","ภาคเหนือ","2026-05-24",1603.0,60041.85,444],
["3027","MARKET VILLAGE RANGSIT KLONG 4","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1699.0,61928.82,598],
["917","Online","Online","Online","2026-05-24",-5.0,57.88,1253],
["3072","BIG C NARATHIWAT","นราธิวาส","ภาคใต้","2026-05-24",2035.0,75267.64,1085],
["3106","LOTUS'S RATTANATHIBET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1822.0,67273.45,765],
["3164","Naresuan University","พิษณุโลก","ภาคเหนือ","2026-05-24",2231.0,71517.11,458],
["3212","Kadluang Chiangrai","เชียงราย","ภาคเหนือ","2026-05-24",2033.0,70909.8,486],
["3315","CDC Ramindra","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1898.0,68160.89,353],
["369","BIG C KORAT","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2589.0,89799.27,183],
["494","TESCO NAVANAKORN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1784.0,68710.68,878],
["547","BIG C TRANG","ตรัง","ภาคใต้","2026-05-24",1965.0,70849.76,529],
["688","TESCO CHO HO","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1992.0,73006.24,312],
["694","THE CHILLED PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",1746.0,61660.73,464],
["813","CENTRAL FESTIVAL PHUKET","ภูเก็ต","ภาคใต้","2026-05-24",2893.0,101483.28,1368],
["3059","LOTUS CHA AM","เพชรบุรี","ภาคตะวันตก","2026-05-24",1673.0,60218.0,478],
["3199","Central Samui","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2839.0,97274.89,1036],
["3290","Somdet Kalasin","กาฬสินธุ์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1826.0,63623.97,551],
["637","CENTRALCHONBURI 2ND FLOOR","ชลบุรี","ภาคตะวันออก","2026-05-24",2227.0,83526.76,703],
["798","Big C Sakaeo","สระแก้ว","ภาคตะวันออก","2026-05-24",2006.0,72274.33,642],
["875","TESCO MAPTAPHUT","ระยอง","ภาคตะวันออก","2026-05-24",1646.0,61798.78,476],
["3074","C-MALL CHAINAT","ชัยนาท","ภาคกลาง","2026-05-24",1916.0,68995.16,527],
["3189","Phan Chiangrai","เชียงราย","ภาคเหนือ","2026-05-24",2110.0,71795.51,568],
["3268","Kosum Phisai","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2188.0,76458.27,298],
["797","ROBINSON CHONBURI","ชลบุรี","ภาคตะวันออก","2026-05-24",2606.0,88008.35,992],
["990","TESCO PAKTHONGCHAI","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1905.0,63247.2,286],
["3073","LOTUS KAENGKHRO","ชัยภูมิ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1842.0,67355.41,713],
["236","THE OLD SIAM PLAZA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2244.0,80967.59,803],
["258","LOTUS BANGKAPI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2041.0,76430.4,841],
["493","UD TOWN","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1981.0,71079.91,487],
["501","TESCO BANG PA-IN","อยุธยา","ภาคกลาง","2026-05-24",2057.0,76850.39,368],
["772","Lotus's Chaiyaphum","ชัยภูมิ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1453.0,54744.97,495],
["3038","TESCO CHOKCHAI","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1887.0,67696.2,471],
["3182","Home Garden Ville","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2255.0,77857.08,252],
["3232","Thai Thanee Navanakorn","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2536.0,82227.96,460],
["3233","The promenade","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1837.0,67799.9,459],
["212","CENTRAL RAMA III","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3211.0,115385.4,921],
["411","NK PANGNGA","พังงา","ภาคใต้","2026-05-24",1626.0,60481.59,737],
["473","TESCO BAN-BUNG","ชลบุรี","ภาคตะวันออก","2026-05-24",1619.0,60030.52,931],
["811","TESCO  SINGBURI","สิงห์บุรี","ภาคกลาง","2026-05-24",1809.0,65234.89,367],
["483","BIG C LOEI","เลย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1346.0,56719.67,437],
["506","BIG C SUANLUANG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2144.0,75890.83,229],
["963","TESCO SADAO","สงขลา","ภาคใต้","2026-05-24",1830.0,64541.74,556],
["3275","Ban Kruat Buriram","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",3472.0,127014.66,385],
["249","LOTUS LAKSI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1735.0,61421.38,580],
["373","NEW SOUTHERN BUS TERMINAL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2016.0,72859.17,461],
["271","MAJOR CINEPLEX CHACHOENGSAO","ฉะเชิงเทรา","ภาคตะวันออก","2026-05-24",2008.0,69144.16,541],
["419","CENTER ONE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1658.0,60002.86,1029],
["805","Lotus's Ratchaburi","ราชบุรี","ภาคตะวันตก","2026-05-24",1866.0,66959.66,508],
["502","SERMTHAI COMPLEX","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2308.0,83024.43,799],
["3112","CENTRAL CHANTHABURI","จันทบุรี","ภาคตะวันออก","2026-05-24",1935.0,72709.5,969],
["3204","PTT Bangyai","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1252.0,52267.38,146],
["404","Lotus's Prachachoen","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1094.0,49746.56,400],
["428","PANTIP NGAMWONGWAN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1741.0,64306.51,727],
["536","TESCO KLAENG RAYONG","ระยอง","ภาคตะวันออก","2026-05-24",1757.0,64843.84,660],
["612","Big C Nan","น่าน","ภาคเหนือ","2026-05-24",918.0,41715.53,441],
["867","CENTRAL CHIANGMAI2  4F","เชียงใหม่","ภาคเหนือ","2026-05-24",2234.0,77988.13,1175],
["887","M  PARK","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1572.0,57611.82,590],
["3050","BIG C SATUN","สตูล","ภาคใต้","2026-05-24",1894.0,69679.3,838],
["3069","LIMELIGHT AVENUE PHUKET","ภูเก็ต","ภาคใต้","2026-05-24",2411.0,86402.25,497],
["3115","JAS RAMINDRA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2005.0,72022.64,280],
["3179","Lotus's Srimahaphot","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",1960.0,69824.16,429],
["807","Big C Onnut","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1629.0,60290.0,739],
["358","UNION MALL 1_GFL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3148.0,95846.79,839],
["640","PRACHUABKHIRIKHAN","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",1835.0,66956.12,499],
["590","ROBINSON DEPARTMENT MUKDAHARN","มุกดาหาร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",3276.0,120196.02,695],
["779","BIG C SUKHOTHAI","สุโขทัย","ภาคเหนือ","2026-05-24",1980.0,67707.32,857],
["3125","Seri Market Phuthamonthon5","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1720.0,62379.16,484],
["3271","PTT U Park Mahasarakham","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2039.0,70610.36,360],
["3323","Asawan Shopping Complex2","หนองคาย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1917.0,67441.98,737],
["3163","Big C Itsaraphap","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1512.0,57116.2,804],
["320","KALASIN PLAZA","กาฬสินธุ์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2359.0,86335.11,912],
["391","SAHATHAI TUNGSONG","นครศรีธรรมราช","ภาคใต้","2026-05-24",1893.0,69478.84,851],
["415","MEECHOK PLAZA","เชียงใหม่","ภาคเหนือ","2026-05-24",1951.0,68578.04,553],
["389","CENTRAL  PLAZA KHONKAEN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2057.0,77648.98,743],
["239","AIRPORT PLAZA CHIANGMAI","เชียงใหม่","ภาคเหนือ","2026-05-24",2170.0,72707.9,402],
["681","ASAWANN PHONPHISAI","หนองคาย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1777.0,65914.61,501],
["683","BIG C BANPHAI","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1641.0,59012.29,456],
["685","JAMPHA SHOPPING MALL","ลำพูน","ภาคเหนือ","2026-05-24",1328.0,51185.69,978],
["606","Lotus's Nongjok","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1559.0,58801.73,774],
["773","Tops Plaza Pichit","พิจิตร","ภาคเหนือ","2026-05-24",1796.0,66899.93,626],
["3146","Int Intersect","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2078.0,73745.07,176],
["3241","Big C Pattani","ปัตตานี","ภาคใต้","2026-05-24",2874.0,105010.38,536],
["3262","Royal Garden Plaza Pattaya","ชลบุรี","ภาคตะวันออก","2026-05-24",1798.0,65233.5,419],
["3282","PTT Kluaynamthai","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2208.0,74946.06,299],
["3036","TESCO KANTHARALAK","ศรีสะเกษ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1530.0,56595.36,762],
["3047","TESCO BUENGSAMPHAN","เพชรบูรณ์","ภาคเหนือ","2026-05-24",1643.0,60646.97,541],
["646","CENTRAL PINKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2682.0,95204.01,1475],
["744","TESCO NANGRONG","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1824.0,67414.84,522],
["3192","Thiandat Market","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",2405.0,80017.84,276],
["3309","Lotus's Saraphi","เชียงใหม่","ภาคเหนือ","2026-05-24",2296.0,74300.15,466],
["586","ROBINSON RANGSIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2214.0,74069.47,841],
["691","KOH PHANGAN","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2369.0,83847.12,799],
["898","BIG C SUKHAPIBAN 5","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1386.0,47405.67,309],
["992","BIG C BANGNA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1719.0,57053.95,330],
["3034","TESCO MAECHAN","เชียงราย","ภาคเหนือ","2026-05-24",1807.0,66951.86,667],
["3114","MAKRO PRADITMANUTHAM","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1743.0,64192.23,377],
["3171","Robinson Chsomalong","ภูเก็ต","ภาคใต้","2026-05-24",2004.0,73206.25,776],
["3218","Maejo University","เชียงใหม่","ภาคเหนือ","2026-05-24",2376.0,78345.63,575],
["3277","City mall Ubon","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1524.0,54252.47,1008],
["3280","Wiang Pa Pao","เชียงราย","ภาคเหนือ","2026-05-24",2749.0,94445.61,303],
["478","TUK COM SRIRACHA","ชลบุรี","ภาคตะวันออก","2026-05-24",1101.0,48324.42,142],
["3245","Banna Nakornnayok","นครนายก","ภาคตะวันออก","2026-05-24",2374.0,84568.71,340],
["577","AYUTTHAYA PARK 2","อยุธยา","ภาคกลาง","2026-05-24",1543.0,58226.69,485],
["505","TESCO HANGDONG","เชียงใหม่","ภาคเหนือ","2026-05-24",1818.0,66342.98,303],
["978","Lotus's Kuchinarai","กาฬสินธุ์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1771.0,59790.03,474],
["980","BIG C PAKCHONG","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2029.0,72446.86,352],
["3082","LOTUS NONGRUEA KHONKAEN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1582.0,57240.5,581],
["3216","Jas Green Village Prewet","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1856.0,68473.61,394],
["278","TESCO RAMA I CHAROENPOL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1601.0,61584.41,575],
["730","AO NANG KRABI","กระบี่","ภาคใต้","2026-05-24",2458.0,87959.81,726],
["750","BIG C ARANYAPRATHET","สระแก้ว","ภาคตะวันออก","2026-05-24",2255.0,81562.81,216],
["3238","Mancha Khiri","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2794.0,95857.55,268],
["3242","PTT Borom 97","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2752.0,96651.69,232],
["3263","Nakornphanom Market","นครพนม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1887.0,67218.22,557],
["806","Lotus's Bangkruay Sainoi","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1846.0,69446.41,1176],
["787","Lotus's Phonthong","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1791.0,65200.95,574],
["3100","RBS BANCHANG","ระยอง","ภาคตะวันออก","2026-05-24",2387.0,87902.07,624],
["3130","Infinite Mall","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2142.0,74348.41,335],
["3134","Bangsaphan","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",1654.0,60266.57,478],
["3222","Bangyai City","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1780.0,62215.54,520],
["3297","Big C Amnatcharoeng","อำนาจเจริญ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",4227.0,154156.26,485],
["615","ROBINSON BURIRUM","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1847.0,67278.0,476],
["3032","METRO MALL PHETCHABURI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1957.0,71060.71,1003],
["3055","PEOPLE PARK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1868.0,67141.19,423],
["3070","LOTUS KABINBURI","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",2030.0,72652.13,603],
["3260","Khukhot Crossing","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1944.0,69728.16,423],
["429","TU DOME","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1947.0,69807.24,378],
["458","TESCO TAYANG","เพชรบุรี","ภาคตะวันตก","2026-05-24",2008.0,67931.36,244],
["564","BIG C SRIMAHAPOH","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",2010.0,73222.53,591],
["3253","Lotus's Nongchang","อุทัยธานี","ภาคเหนือ","2026-05-24",1622.0,61063.62,697],
["555","ROBINSON SURIN","สุรินทร์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",460.0,16589.1,108],
["511","TESCO NAKORNSAWAN","นครสวรรค์","ภาคเหนือ","2026-05-24",1155.0,33870.89,114],
["557","Lotus's Rojana","อยุธยา","ภาคกลาง","2026-05-24",2011.0,76917.85,479],
["596","BIG C ROMKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2331.0,81436.01,208],
["3191","Danchang Suphanburi","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",2069.0,70249.74,373],
["3211","Big C Omyai","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1593.0,57207.22,683],
["472","VOGUE KRABI","กระบี่","ภาคใต้","2026-05-24",1793.0,65702.75,1146],
["879","BIG C TAK","ตาก","ภาคเหนือ","2026-05-24",1550.0,56394.29,645],
["885","TESCO ANGTHONG","อ่างทอง","ภาคกลาง","2026-05-24",1933.0,68470.11,524],
["3214","AC Saimai","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1628.0,59611.5,511],
["539","SAPANKHWAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1184.0,51718.89,221],
["679","TESCO PHIMAI","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1881.0,69019.77,478],
["597","PETCHPAIBOON PLAZA","เพชรบุรี","ภาคตะวันตก","2026-05-24",1802.0,65103.55,601],
["3062","TESCO PHRAE","แพร่","ภาคเหนือ","2026-05-24",1840.0,66913.93,393],
["520","CENTRAL UBONRATCHATHANI","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2426.0,83902.84,1488],
["530","BIG C KALASIN","กาฬสินธุ์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1929.0,68745.27,640],
["684","CENTRAL MARINA","ชลบุรี","ภาคตะวันออก","2026-05-24",2066.0,74538.55,545],
["619","TESCO SAMUTSONGKRAM","สมุทรสงคราม","ภาคตะวันตก","2026-05-24",1338.0,47426.14,582],
["747","YES  BANGPLEE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1849.0,66714.85,916],
["870","ROBINSON CHAIYAPHUM","ชัยภูมิ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2864.0,94358.01,1012],
["3169","The Mall Bangkapi","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2955.0,111110.74,2525],
["3177","Tipgaysorn Market","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1612.0,57404.89,545],
["3206","Mangmee Market","ระยอง","ภาคตะวันออก","2026-05-24",2034.0,72952.55,644],
["448","INDRA PRATUNAM","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2014.0,83746.15,279],
["960","IMPERIAL WORLD SAMRONG 2F","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3121.0,108150.65,1002],
["3010","ROI ET HOSPITAL","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2250.0,76822.75,474],
["3175","Jas Green Village Bangbuathong","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1610.0,59113.54,686],
["3252","Makro Chainat","ชัยนาท","ภาคกลาง","2026-05-24",1677.0,63467.16,520],
["3283","Si Satchanalai","สุโขทัย","ภาคเหนือ","2026-05-24",2075.0,71462.13,378],
["274","Lotus's Khamtieng Chiengmai","เชียงใหม่","ภาคเหนือ","2026-05-24",1855.0,69499.44,616],
["363","MINBURI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2099.0,69923.63,580],
["3234","Lotus's Yala","ยะลา","ภาคใต้","2026-05-24",1700.0,61159.08,493],
["988","ROBINSON SUPHANBURI 1F","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",2860.0,99868.95,845],
["3230","The Sphere Phetkasem","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",2116.0,76241.76,378],
["3235","R.N.Yard","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1607.0,59709.51,491],
["297","SERMTHAI MAHASARAKHAM","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2200.0,77264.74,206],
["617","TESCO KHONKAEN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1573.0,57850.5,357],
["753","Lotus's Pakchong","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2210.0,80488.05,674],
["3076","RBS SRISAMAN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2541.0,82350.26,842],
["3178","Big C Rayong","ระยอง","ภาคตะวันออก","2026-05-24",2737.0,97695.71,222],
["3256","Kad Farang Village Maerim","เชียงใหม่","ภาคเหนือ","2026-05-24",1763.0,63364.99,599],
["3314","Romphobai Market","สระบุรี","ภาคกลาง","2026-05-24",2099.0,66983.55,490],
["287","TAD CHAN HUA HIN","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",2189.0,74154.78,392],
["243","PHAHOLYOTHIN PLACE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1763.0,65687.79,588],
["265","BIG C NAKORN SAWAN","นครสวรรค์","ภาคเหนือ","2026-05-24",3360.0,123715.09,608],
["291","Saneha Hatyai","สงขลา","ภาคใต้","2026-05-24",2548.0,87506.79,953],
["378","SUBSIN PLAZA SONGKLA","สงขลา","ภาคใต้","2026-05-24",1842.0,66993.55,607],
["655","THE CRYSTAL RAMINDRA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1878.0,66938.6,307],
["671","WANGLANG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2093.0,68891.08,961],
["3118","Ubon Square","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2036.0,68053.01,422],
["3251","Suen Heng Srisaket","ศรีสะเกษ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1688.0,61638.75,782],
["3259","Soi Suphaphong","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1898.0,64468.94,495],
["665","MID TOWN PLAZA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1708.0,63547.23,415],
["873","GATEWAY BANGSUE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1677.0,60212.28,674],
["682","TERMINAL 21 KORAT","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1776.0,63559.34,477],
["3040","TESCO KRATHUMBAEN","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",1800.0,63941.54,659],
["981","BIG C SUPHANBURI","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",1632.0,54793.92,351],
["3305","EmQuartier","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2406.0,88070.18,632],
["375","BIG C AYUDTHAYA","อยุธยา","ภาคกลาง","2026-05-24",2149.0,76187.0,335],
["658","TESCO SANGKHA","สุรินทร์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1486.0,57538.21,691],
["759","AMPORN AYUTTHAYA","อยุธยา","ภาคกลาง","2026-05-24",1855.0,68297.27,634],
["3226","Paknam Pranburi","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",2454.0,79591.79,227],
["3237","Phangkhon","สกลนคร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2286.0,82106.54,344],
["3261","Thong Pha Phum","กาญจนบุรี","ภาคตะวันตก","2026-05-24",2072.0,74161.37,436],
["394","TESCO SUPHANBURI","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",1709.0,62076.68,387],
["678","BLU PORT HUA HIN","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",2284.0,79913.85,384],
["771","BIG C WANGNAMYEN","สระแก้ว","ภาคตะวันออก","2026-05-24",2046.0,72952.81,420],
["791","BIG C RAMINDRA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2282.0,79539.23,274],
["3165","Big C Sungaikolok","นราธิวาส","ภาคใต้","2026-05-24",1836.0,66520.68,597],
["3124","Sunpasitthiprasong Hospital","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1751.0,61382.86,479],
["3249","Big C Kanlapapruek","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2835.0,102992.82,355],
["895","CENTRAL CHONBURI  GF","ชลบุรี","ภาคตะวันออก","2026-05-24",2224.0,80707.75,810],
["3058","LOTUS PHANOSARAKARN","ฉะเชิงเทรา","ภาคตะวันออก","2026-05-24",1937.0,69361.32,461],
["3065","LOTUS BANGNA TRAD","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1813.0,64327.82,469],
["644","TESCO SURATTHANEE","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1949.0,71179.01,360],
["3033","TESCO UDONTHANI","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1774.0,63167.9,422],
["3293","Big C Bangpakok","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1800.0,61117.29,633],
["639","ROBINSON MAESOD","ตาก","ภาคเหนือ","2026-05-24",1817.0,69025.47,1494],
["386","ESPLANADE RATTANATIBETH","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2035.0,73887.74,861],
["3152","The Forest Rangsit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2155.0,74964.52,676],
["3173","Big C Nongchok","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1802.0,64841.14,461],
["3162","Makro Theprak","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1836.0,66965.09,532],
["347","BANGPAKOK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1715.0,61295.35,475],
["585","ROBINSON DEPARTMENT PRACHINBURI","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",2414.0,81201.59,1003],
["153","WH","WH/DC","WH/DC","2026-05-24",26047.0,751948.83,0],
["327","BANG KHUNSRI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1811.0,66093.65,551],
["281","BIG C KHON KAEN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1690.0,61029.93,433],
["583","Lotus's Salaya","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1447.0,49872.63,331],
["621","Lotus's Saraburi","สระบุรี","ภาคกลาง","2026-05-24",1694.0,63961.65,821],
["632","CENTRAL EASTVILLE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2505.0,84894.77,678],
["677","TESCO LAMLUKKA KLONG2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1775.0,62857.03,567],
["3093","BIG C NONGKI","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1883.0,69001.07,509],
["3200","Mingle Hill Minburi","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1587.0,58872.19,690],
["3326","Bangchak KKU","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2953.0,91399.39,618],
["263","CENTRAL RAMA II","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2320.0,82492.09,1237],
["283","NATHORN SAMUI","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2471.0,80487.49,433],
["469","CK PLAZA RAYONG","ระยอง","ภาคตะวันออก","2026-05-24",2118.0,79053.65,950],
["544","PASEO LADKRABANG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1804.0,65857.59,571],
["3048","TESCO LOMSAK","เพชรบูรณ์","ภาคเหนือ","2026-05-24",1819.0,66234.19,776],
["634","THE STREET RATCHADA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1914.0,64787.06,495],
["659","BIG C RANONG","ระนอง","ภาคใต้","2026-05-24",1664.0,61313.5,779],
["764","BIG C CHIANGMAI 2","เชียงใหม่","ภาคเหนือ","2026-05-24",1567.0,57801.3,583],
["965","TESCO CHIANGKHONG","เชียงราย","ภาคเหนือ","2026-05-24",1802.0,68762.23,485],
["3208","Central Nakornpathom","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",2216.0,80221.42,1594],
["285","HAPPY PLAZA PICHIT","พิจิตร","ภาคเหนือ","2026-05-24",1599.0,60340.03,499],
["664","BIG C LAMPHUN","ลำพูน","ภาคเหนือ","2026-05-24",1631.0,57997.98,627],
["267","ASAWANN","หนองคาย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2634.0,96720.39,1219],
["356","HUA HIN MARKET VILLAGE","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",1617.0,65233.86,514],
["481","SAHATHAI NAKORNSRITHAMARAT","นครศรีธรรมราช","ภาคใต้","2026-05-24",1132.0,50124.26,209],
["3294","Lotus's Lamnarai","ลพบุรี","ภาคกลาง","2026-05-24",2363.0,79197.31,592],
["292","NANA SQUARE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1912.0,75734.7,161],
["492","BIG C RANGSIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1732.0,64291.04,394],
["608","MARKET VILLAGE SUVARNABHUMI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2281.0,83725.15,1060],
["874","TESCO NORTH PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",1551.0,56740.17,422],
["967","TESCO THATUM","สุรินทร์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1502.0,55021.84,596],
["3145","Nawamin City Avenue","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1659.0,60686.05,302],
["3217","Big C Yala","ยะลา","ภาคใต้","2026-05-24",1831.0,66362.55,412],
["574","ROYAL PARK RATCHABURI","ราชบุรี","ภาคตะวันตก","2026-05-24",1919.0,63371.51,103],
["656","TESCO BUENGKAN","บึงกาฬ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1927.0,69345.82,623],
["736","THE EXPLACE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1902.0,66910.2,232],
["650","CENTRAL HATYAI2","สงขลา","ภาคใต้","2026-05-24",2773.0,95030.81,917],
["379","CENTRAL FESTIVAL PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",1884.0,69437.34,814],
["609","CENTRAL RAYONG","ระยอง","ภาคตะวันออก","2026-05-24",2828.0,101388.76,1034],
["864","THE CIRCLE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2476.0,86160.44,249],
["968","BIG C ANGTHONG","อ่างทอง","ภาคกลาง","2026-05-24",1992.0,68223.06,197],
["660","Green Place","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1358.0,53152.51,563],
["556","CRYSTAL CHAIYAPRUEK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1924.0,68198.4,242],
["784","ZUELLIG HOUSE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1822.0,65635.1,673],
["3011","SALAYA MARKET","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1131.0,44815.3,370],
["3143","Jampha Lamphun","ลำพูน","ภาคเหนือ","2026-05-24",1648.0,59746.84,668],
["3264","Gemmo Market","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1892.0,63765.45,633],
["3318","Chiangkhan","เลย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2413.0,81382.11,331],
["3202","Ramindra 40","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",933.0,42697.33,216],
["217","Mahboonkrong","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2696.0,101575.84,1035],
["261","BIG C TIWANON","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2135.0,80552.78,902],
["390","COLESIUM SURATTHANI","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1816.0,67549.62,303],
["453","Lotus's Klong 7","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1558.0,57649.87,521],
["788","CENTRAL WORLD 4FL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3938.0,140031.64,2471],
["3046","BIG C UDONTHANI 1","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1848.0,65698.08,342],
["296","TUK COM PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",1752.0,63850.49,525],
["731","TESCO PRACHUABKHIRIKHAN","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",1646.0,59733.71,573],
["3244","Rongkluea Market Navananorn","อยุธยา","ภาคกลาง","2026-05-24",1898.0,70105.32,1026],
["3311","Green Park Chiangrai","เชียงราย","ภาคเหนือ","2026-05-24",2115.0,73227.52,691],
["733","ROBINSON BANGRAK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2536.0,92095.72,898],
["755","CENTRAL NAKRONRATCHASRIMA GF","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1600.0,64594.48,590],
["770","TESCO MINBURI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2144.0,77394.26,198],
["3168","Little Walk Krungthep Kreetha","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",971.0,44781.72,175],
["395","HARBOR MALL","ชลบุรี","ภาคตะวันออก","2026-05-24",1386.0,51719.2,1012],
["700","MARKET WALK BANGYAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1424.0,53290.9,926],
["3081","BIG C BORWIN","ระยอง","ภาคตะวันออก","2026-05-24",1660.0,59486.07,377],
["902","eStore-Lazada","Online","Online","2026-05-24",1214.0,15544.69,6399],
["3084","TRAT MARKET","ตราด","ภาคตะวันออก","2026-05-24",2210.0,78519.87,545],
["3267","Dreamland Market","นครสวรรค์","ภาคเหนือ","2026-05-24",1968.0,69550.8,583],
["3030","TESCO MAESAI","เชียงราย","ภาคเหนือ","2026-05-24",2008.0,70692.12,897],
["3116","MAKRO CHARANSANITWONG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1783.0,63297.13,629],
["3243","Dern Plearn Village","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1918.0,69667.7,292],
["567","ROBINSON SUKHUMVIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3261.0,113379.54,558],
["582","TESCO ARANYAPRATHET","สระแก้ว","ภาคตะวันออก","2026-05-24",2141.0,76500.31,656],
["815","MARKET PLACE NANGLINCHEE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1769.0,63765.15,293],
["3096","SENTOSA SRICHAN KHONKAEN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1382.0,52047.82,635],
["279","CENTRAL AIRPORT PLAZA CHIENGMAI 2","เชียงใหม่","ภาคเหนือ","2026-05-24",2089.0,76706.93,946],
["475","PORTO CHINO","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",1755.0,63685.07,396],
["549","Lotus's Chalong Phuket","ภูเก็ต","ภาคใต้","2026-05-24",2019.0,69732.57,411],
["735","TESCO PATHUMTHANI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1577.0,57669.0,943],
["865","TESCO  BANFAH","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1514.0,51033.16,310],
["365","BIG C RATCHBURI","ราชบุรี","ภาคตะวันตก","2026-05-24",1672.0,60852.86,364],
["748","THE GROVE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2139.0,75576.08,149],
["758","TESCO PRASAT SURIN","สุรินทร์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1708.0,62356.45,562],
["884","TESCO JOMTHONG","เชียงใหม่","ภาคเหนือ","2026-05-24",1619.0,58532.56,675],
["3045","BIG C UBONRATCHATHANI","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1900.0,67659.44,387],
["3161","Makro Sanambinnam","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",985.0,43989.57,232],
["286","BIG C NAKORNPATHOM","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1839.0,67465.43,587],
["384","Central Chaeng Wattana","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2278.0,78755.61,1074],
["466","TIME SQUARE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1991.0,65257.66,547],
["552","BIG C PHITSANULOK","พิษณุโลก","ภาคเหนือ","2026-05-24",1458.0,55932.28,557],
["396","SUK ANAN SARABURI","สระบุรี","ภาคกลาง","2026-05-24",2406.0,83835.87,258],
["423","MUANG AKE RANGSIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1741.0,61413.23,397],
["3088","CHAREONSRI WARIN MARKET","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2141.0,74822.28,283],
["3172","Lotus Phuket","ภูเก็ต","ภาคใต้","2026-05-24",2556.0,92148.56,1130],
["3085","BIG C NAKHONSITHAMMARAT","นครศรีธรรมราช","ภาคใต้","2026-05-24",2889.0,97742.32,461],
["3257","Wang Sam Mo","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2092.0,74920.3,366],
["463","ASAWANN 1","หนองคาย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1985.0,71637.05,352],
["3181","Suwan Kleaw Thong Market","อยุธยา","ภาคกลาง","2026-05-24",1840.0,59742.63,325],
["490","CENTRAL SURATTHANI","สุราษฎร์ธานี","ภาคใต้","2026-05-24",2813.0,102136.07,912],
["754","Lotus's Yasothon","ยโสธร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2153.0,74650.74,668],
["259","LOTUS LADPRAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1296.0,53221.84,365],
["413","TESCO PINKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1790.0,61874.65,647],
["534","MAJOR RANGSIT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2096.0,74900.9,279],
["3286","Hatyai Village","สงขลา","ภาคใต้","2026-05-24",1255.0,50848.19,447],
["3213","Century Sukhumvit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2254.0,79507.97,973],
["795","Lotus's U Tong","สุพรรณบุรี","ภาคตะวันตก","2026-05-24",2022.0,72035.21,525],
["330","SURIN PLAZA","สุรินทร์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2211.0,78608.6,479],
["537","NATION TOWER","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2217.0,78316.74,199],
["3006","ROBINSON BORWIN","ชลบุรี","ภาคตะวันออก","2026-05-24",3022.0,95141.89,1073],
["3064","LOTUS CHAIYA","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1893.0,66695.96,511],
["668","CENTRAL NAKORNSRITHAMMARAT","นครศรีธรรมราช","ภาคใต้","2026-05-24",2249.0,86656.64,1258],
["676","THE PLATINUM MALL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2154.0,84550.49,203],
["3090","ATARA MALL","ชลบุรี","ภาคตะวันออก","2026-05-24",1880.0,69314.47,422],
["675","LOTUS THABOR","หนองคาย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",935.0,43568.19,290],
["620","TESCO BANGKAE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",945.0,43453.71,189],
["3221","Big C Rangsit Khlong  3","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1696.0,61532.76,587],
["332","Chaisaeng Singburi","สิงห์บุรี","ภาคกลาง","2026-05-24",1923.0,69749.45,522],
["761","Lotus's Banpong","ราชบุรี","ภาคตะวันตก","2026-05-24",1403.0,53994.98,665],
["3136","Thungsao Market","สงขลา","ภาคใต้","2026-05-24",1578.0,53546.52,509],
["3159","Uthaithani","อุทัยธานี","ภาคเหนือ","2026-05-24",1801.0,63227.55,529],
["3324","Lotus's Amata City","ชลบุรี","ภาคตะวันออก","2026-05-24",2250.0,85702.67,1173],
["3140","Lotus's Nikhompattana","ระยอง","ภาคตะวันออก","2026-05-24",1888.0,67402.97,529],
["3205","Central Nakornsawan","นครสวรรค์","ภาคเหนือ","2026-05-24",2474.0,96221.51,710],
["3220","Lotus's Pattaya 53","ชลบุรี","ภาคตะวันออก","2026-05-24",1745.0,63445.4,610],
["3306","Market Place Prachauthit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3453.0,117590.07,789],
["941","Online","Online","Online","2026-05-24",163578.0,4749872.37,0],
["524","NANA CHAROEN MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1771.0,60320.78,512],
["673","TESCO CHANGWATTANA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1763.0,64875.82,434],
["3068","LOTUS CHATURAT","ชัยภูมิ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1955.0,70038.32,325],
["3273","Seka Buengkan","บึงกาฬ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2003.0,72712.15,369],
["3094","VACHIRA SONGHKLA","สงขลา","ภาคใต้","2026-05-24",1671.0,61759.16,515],
["350","KAD TON PAYAOM","เชียงใหม่","ภาคเหนือ","2026-05-24",2210.0,73986.02,479],
["3225","Nonmuang Khonkaen","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1869.0,64556.26,343],
["3272","Pong Pai Market","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",1949.0,68839.0,409],
["3028","THANOMMIT MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1985.0,68347.91,421],
["3104","EKKAMAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1913.0,70186.68,342],
["3170","Foody Farm","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1809.0,63772.05,335],
["325","WANGHIN SHOPPING CENTRE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1580.0,59885.7,639],
["417","TOPLAND PETCHABOON","เพชรบูรณ์","ภาคเหนือ","2026-05-24",897.0,41560.65,220],
["439","TESCO TALANG","ภูเก็ต","ภาคใต้","2026-05-24",1927.0,71710.28,1012],
["715","EXTERNAL PROMOTION","Others","Others","2026-05-24",1306.0,49630.93,92],
["569","TESCO SAMUTSAKHON","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",918.0,43097.02,232],
["687","ROBINSON LOPBURI","ลพบุรี","ภาคกลาง","2026-05-24",2450.0,92693.09,774],
["989","MARKET VILLAGE PHUKET CHALONG","ภูเก็ต","ภาคใต้","2026-05-24",1596.0,59203.68,528],
["605","Lotus's Petchaboon","เพชรบูรณ์","ภาคเหนือ","2026-05-24",1586.0,59993.46,968],
["766","SAVE ONE KORAT","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1880.0,63387.44,475],
["3141","PHITSANULOK MARKET","พิษณุโลก","ภาคเหนือ","2026-05-24",1784.0,60824.93,432],
["3077","CENTRAL AYUTTHAYA","อยุธยา","ภาคกลาง","2026-05-24",3285.0,112458.97,1029],
["3209","The Mall Bangkae GF","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2677.0,90230.69,867],
["364","PATTAYA KLANG","ชลบุรี","ภาคตะวันออก","2026-05-24",1666.0,58730.16,561],
["3148","Chamchuri Square","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1832.0,66390.4,516],
["504","TESCO PRANBURI","ประจวบคีรีขันธ์","ภาคตะวันตก","2026-05-24",1632.0,60228.2,788],
["3007","MAHBOONKRONG FL.G","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2593.0,92098.45,577],
["777","Kadkam Plaza","แม่ฮ่องสอน","ภาคเหนือ","2026-05-24",1768.0,63792.23,992],
["3013","SAMMAKORN PLACE RANGSIT KLONG2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1746.0,62269.05,298],
["3022","BIG C MAHACHAI 2","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",2131.0,75530.63,231],
["3291","Lasalle's Avenue","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1913.0,69150.4,324],
["3303","SuscoSquarePhutthabucha","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1736.0,59654.98,623],
["446","FORTUNE TOWN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1345.0,55606.17,518],
["974","BIG C SAINOI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1428.0,49331.27,308],
["3039","TESCO LAMPANG","ลำปาง","ภาคเหนือ","2026-05-24",1608.0,59623.6,374],
["3167","Vi Plaza.","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1144.0,52021.32,248],
["464","CENTRAL PLAZA UDONTHANI","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2703.0,96501.56,1614],
["649","V MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1719.0,61597.02,776],
["3201","Central Udon 3F","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1894.0,68581.73,538],
["336","BIG C LAMPANG","ลำปาง","ภาคเหนือ","2026-05-24",872.0,41721.63,343],
["526","BIG C LAMLUKKA 1","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1725.0,62813.96,468],
["3298","BaanFah Le Marche","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1950.0,65508.46,437],
["382","FASHION ISLAND 1","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2444.0,86032.2,1154],
["648","UNION MALL_BFL","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2015.0,70875.07,300],
["734","TESCO BORWIN","ชลบุรี","ภาคตะวันออก","2026-05-24",1646.0,59295.19,592],
["3057","LOTUS BANGBOR","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1664.0,64114.92,848],
["3111","BIG C LATYAO","นครสวรรค์","ภาคเหนือ","2026-05-24",1688.0,61575.3,472],
["3127","Tops Plaza Nonghan","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1640.0,59208.58,617],
["3325","Asoke Tower","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1384.0,53727.22,172],
["3051","MIXT CHATUCHAK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2058.0,75744.69,757],
["3229","Big C Phetkasem 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1774.0,65257.55,838],
["233","FAIRY PLAZA KHON-KAEN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2161.0,80355.75,677],
["420","TESCO CHUMPAE","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1434.0,54463.95,841],
["3003","BIG C KHEHAROMKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2056.0,72901.06,768],
["3288","Santisuk 2 Market","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2002.0,71878.28,255],
["3024","TESCO KHUKHAN","ศรีสะเกษ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1719.0,60248.75,749],
["435","Central Ladprao","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2557.0,90078.24,2372],
["568","TESCO BANGYAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1357.0,50431.05,803],
["3120","Terminal 21 Rama 3","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1818.0,66557.99,487],
["437","TESCO HATYAI 2","สงขลา","ภาคใต้","2026-05-24",1843.0,69189.4,538],
["510","BIG C ROI ET","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1959.0,70788.79,469],
["593","THE PASEO PARK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2069.0,72326.32,405],
["3150","Central Ramindra","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2090.0,71044.14,496],
["749","BIG C PATHUMTHANI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1503.0,56326.39,471],
["799","Lotus's Kamphaengsaen","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1791.0,64212.94,475],
["3053","IMPERIAL SAMRONG GF","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2350.0,81766.13,909],
["491","SILOM COMPLEX","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1898.0,70558.35,862],
["662","BIG C SUKSAWAT","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2018.0,73677.98,1287],
["241","TESCO RAMA IV","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1877.0,70068.98,726],
["645","TESCO RAYONG","ระยอง","ภาคตะวันออก","2026-05-24",2211.0,73412.18,514],
["868","TESCO PHAYAO","พะเยา","ภาคเหนือ","2026-05-24",1768.0,65121.36,538],
["972","Big C Chiangmai 1","เชียงใหม่","ภาคเหนือ","2026-05-24",2369.0,82388.53,741],
["3129","Bangpu Center","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1597.0,58651.86,845],
["540","THE TREE BANGBON","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1234.0,53149.14,138],
["3316","Koh Lanta Krabi","กระบี่","ภาคใต้","2026-05-24",2316.0,80773.3,479],
["247","TAVEEKIT PLAZA SARABURI","สระบุรี","ภาคกลาง","2026-05-24",2080.0,74310.24,223],
["984","SURATTHANI HOSPITAL","สุราษฎร์ธานี","ภาคใต้","2026-05-24",959.0,34842.49,138],
["503","Lotus's Phitsanulok 2","พิษณุโลก","ภาคเหนือ","2026-05-24",1557.0,56377.63,599],
["548","CENTRAL HATYAI","สงขลา","ภาคใต้","2026-05-24",1960.0,71509.87,675],
["3153","Lotus's Songkla","สงขลา","ภาคใต้","2026-05-24",2487.0,90773.47,620],
["3119","Robinson Ratchapruek","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1893.0,66313.27,477],
["3154","Paradise Park","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2019.0,73123.39,722],
["661","CHA-AM","เพชรบุรี","ภาคตะวันตก","2026-05-24",2367.0,70529.03,208],
["3223","The Prom Din Daeng","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1660.0,60981.1,669],
["3044","TESCO CHIANGKAM","พะเยา","ภาคเหนือ","2026-05-24",868.0,41493.94,376],
["443","CENTRAL PHITSANULOK","พิษณุโลก","ภาคเหนือ","2026-05-24",1856.0,67194.23,882],
["774","Lotus's Rangsit","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1506.0,55880.82,657],
["672","TESCO CHANTHABURI","จันทบุรี","ภาคตะวันออก","2026-05-24",1620.0,58576.16,727],
["250","LOTUS KORAT","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2123.0,81018.04,316],
["551","MAYA CHIANG MAI","เชียงใหม่","ภาคเหนือ","2026-05-24",2464.0,77366.29,1029],
["751","Lotus's Fang","เชียงใหม่","ภาคเหนือ","2026-05-24",1729.0,64526.19,738],
["3109","RBS CHACHOENGSAO","ฉะเชิงเทรา","ภาคตะวันออก","2026-05-24",3000.0,107790.14,876],
["3139","Lotus's Bangpakong","ฉะเชิงเทรา","ภาคตะวันออก","2026-05-24",2057.0,72986.77,350],
["635","BURIRAM CASTLE","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1064.0,49029.96,86],
["642","SEACON BANGKAE2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2751.0,92627.16,1754],
["657","BIG C PHUKET","ภูเก็ต","ภาคใต้","2026-05-24",889.0,36037.24,144],
["3099","LOTUS SAWANGDAENDDIN","สกลนคร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1792.0,65292.41,532],
["3255","MuangThaiPhatra Complex","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1775.0,66738.9,476],
["686","Lotus's Samui","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1758.0,62458.79,928],
["3184","Victory Hub","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1802.0,62351.47,1497],
["507","BIG C SUKHAPIBAN 3","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2036.0,73234.8,311],
["342","Century the movie plaza","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1993.0,74357.56,1209],
["745","BIG C PETCHABURI","เพชรบุรี","ภาคตะวันตก","2026-05-24",1868.0,67707.19,507],
["3265","Nong Bua Daeng","ชัยภูมิ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1835.0,66853.39,355],
["489","PLAZA LAGOON","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1150.0,52374.01,139],
["3117","ROBINSON TALANG","ภูเก็ต","ภาคใต้","2026-05-24",1194.0,28616.75,569],
["3310","Tonson Market Mahasarakham","มหาสารคาม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2574.0,92126.61,308],
["3236","Friday Uttaradit","อุตรดิตถ์","ภาคเหนือ","2026-05-24",1499.0,55590.0,630],
["695","Lotus's Sukaphiban 3","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1600.0,59340.59,569],
["880","OZO KATA","ภูเก็ต","ภาคใต้","2026-05-24",1622.0,63155.79,278],
["3246","Dernlen Market","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1801.0,64261.23,501],
["667","HAT YAI HOSPITAL","สงขลา","ภาคใต้","2026-05-24",1580.0,53743.58,598],
["3160","Mae Fah Luang University","เชียงราย","ภาคเหนือ","2026-05-24",2579.0,83722.38,601],
["785","BIG C RAMA 4","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1781.0,64277.81,687],
["3183","Big C Saraburi","สระบุรี","ภาคกลาง","2026-05-24",1735.0,64080.58,364],
["3187","Central Westville","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2587.0,94980.22,725],
["521","THE WALK KASET-NAWAMINTR","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1771.0,65700.21,324],
["367","BIG C PHRAE","แพร่","ภาคเหนือ","2026-05-24",936.0,42923.79,208],
["538","ROBINSON RATCHABURI","ราชบุรี","ภาคตะวันตก","2026-05-24",832.0,38134.93,280],
["3029","TESCO NAKHONCHAISI","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",984.0,44802.69,361],
["3026","TESCO KALASIN","กาฬสินธุ์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1558.0,57515.09,602],
["410","TESCO  LOEI","เลย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1071.0,46642.76,73],
["525","ROBINSON SAKHONNAKORN","สกลนคร","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2334.0,83318.33,620],
["431","GREEN WEALTH","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1839.0,66527.92,555],
["272","MAJOR CINEPLEX PINKLAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1790.0,62429.82,500],
["519","THE TREE PATHUMTHANI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2097.0,74243.71,188],
["3138","Wanasin Market Pattaya","ชลบุรี","ภาคตะวันออก","2026-05-24",1957.0,66966.13,703],
["470","MEGA BANGNA 1","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2450.0,83219.14,1306],
["535","THE WALK RATCHAPRUE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1810.0,65513.36,286],
["255","LOTUS PAKTOR","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1840.0,66756.43,479],
["318","KANOKPETCH KANCHANABURI","กาญจนบุรี","ภาคตะวันตก","2026-05-24",1469.0,50922.5,403],
["996","ROBINSON KANCHANABURI 1F","กาญจนบุรี","ภาคตะวันตก","2026-05-24",2354.0,83047.27,1286],
["3008","SIAM PREMIUM OUTLET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2012.0,72760.51,433],
["624","TESCO BANGPOO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1328.0,43911.8,512],
["782","BIG C CHANTHABURI","จันทบุรี","ภาคตะวันออก","2026-05-24",2223.0,81967.33,502],
["3086","LOTUS LANGSUAN","ชุมพร","ภาคใต้","2026-05-24",920.0,42339.15,361],
["553","PASEO SUKHAPIBAN 3","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1102.0,45832.54,70],
["690","PHUKET GROCERY","ภูเก็ต","ภาคใต้","2026-05-24",937.0,41947.96,380],
["377","MAJOR RATCHAYOTHIN","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2758.0,90270.81,1496],
["541","CENTRAL CHIANGMAI 2","เชียงใหม่","ภาคเหนือ","2026-05-24",1847.0,64534.47,760],
["3176","Ladsawai Market","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2065.0,67851.3,558],
["3092","LOTUS BOROM SAI 2","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2150.0,74511.75,108],
["3095","LOPBURI MARKET","ลพบุรี","ภาคกลาง","2026-05-24",1958.0,67564.83,553],
["3194","Emsphere","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2176.0,80705.97,634],
["983","MAESAI","เชียงราย","ภาคเหนือ","2026-05-24",1508.0,60691.16,655],
["866","TESCO U TAPAO","ชลบุรี","ภาคตะวันออก","2026-05-24",2046.0,67661.39,382],
["3289","Khok Kloi Takuathung","พังงา","ภาคใต้","2026-05-24",2244.0,77665.18,437],
["652","THE CRYSTAL RATCHAPRUEK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",923.0,42667.26,203],
["817","LITTLE WALK PATTAYA","ชลบุรี","ภาคตะวันออก","2026-05-24",1468.0,54271.17,492],
["3052","ROI ET PLAZA","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2002.0,69305.83,517],
["3122","Lotus's Kumphawapi","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1929.0,70634.92,628],
["3198","The Mall Bangkapi 1F","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1884.0,72614.61,550],
["447","Lotus's Khonkaen 2","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1569.0,60173.05,598],
["433","TESCO PAKKRED","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",880.0,41132.6,268],
["300","RAMKHAMHAENG SOI 53","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1556.0,58554.35,901],
["3285","Onnut 17 (Kratiam Market)","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",2108.0,71493.15,421],
["269","LOTUS CHARANSANITWONG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1937.0,68936.61,777],
["962","TESCO NONGBUALAMPHU","หนองบัวลำภู","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1685.0,55755.5,424],
["562","BIG C KANCHANABURI","กาญจนบุรี","ภาคตะวันตก","2026-05-24",1001.0,45814.89,153],
["3083","BIG C NONGBUA","นครสวรรค์","ภาคเหนือ","2026-05-24",1931.0,68909.95,198],
["380","TESCO SRISAKET","ศรีสะเกษ","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1721.0,60933.22,296],
["909","eStore HD-COD","Online","Online","2026-05-24",1909.0,38601.0,14479],
["360","N-MARK","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1663.0,60468.16,747],
["3266","Thai Siri 3 Market","อุดรธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2475.0,85940.9,764],
["903","eStore-Click&Collect","Online","Online","2026-05-24",221.99,4221.88,1084],
["3248","Little walk Rattanathibeth","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1771.0,63722.98,508],
["3250","Big C Surin","สุรินทร์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1903.0,69792.53,591],
["3103","THONBURI MARKET PLACE","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1775.0,64908.74,470],
["918","Lineman","Online","Online","2026-05-24",-12.0,-566.24,555],
["3113","CHACHAWARN MARKET","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1947.0,70035.38,393],
["892","CHIANGMAI 89 PLAZA","เชียงใหม่","ภาคเหนือ","2026-05-24",1567.0,58949.33,564],
["3135","Phahonyothin 54","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1711.0,61509.41,537],
["571","TESCO JANA SONGKHLA","สงขลา","ภาคใต้","2026-05-24",902.0,40660.47,267],
["693","NAMPU PLAZA","สมุทรสาคร","กรุงเทพมหานครและปริมณฑล","2026-05-24",996.0,42144.41,270],
["780","Central Khonkaen 2","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1801.0,63959.47,670],
["486","TESCO WARINCHAMRAP","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1733.0,64023.09,457],
["432","TESCO TRAD","ตราด","ภาคตะวันออก","2026-05-24",1318.0,53267.62,284],
["982","BIG C CHONBURI 1","ชลบุรี","ภาคตะวันออก","2026-05-24",1868.0,65951.14,303],
["786","SRIPONG PARK","อุตรดิตถ์","ภาคเหนือ","2026-05-24",1870.0,68359.94,561],
["977","SPACE@304","ปราจีนบุรี","ภาคตะวันออก","2026-05-24",2047.0,70188.72,545],
["3018","BIG C HATYAI 2","สงขลา","ภาคใต้","2026-05-24",1068.0,47111.47,181],
["3035","TESCO LAMPLAIMAT","บุรีรัมย์","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1457.0,57376.73,688],
["393","TESCO KANCHANABURI","กาญจนบุรี","ภาคตะวันตก","2026-05-24",1721.0,63402.67,578],
["3227","Wangnoi Market","อยุธยา","ภาคกลาง","2026-05-24",1313.0,48535.89,527],
["3025","TESCO LADKRABANG","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",774.0,37726.19,349],
["732","PAI","แม่ฮ่องสอน","ภาคเหนือ","2026-05-24",2046.0,64977.81,351],
["3155","U Ville","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2259.0,76180.85,307],
["670","BIG C BANGYAI","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1475.0,55333.49,793],
["752","Lotus's Ubonratchathani","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1737.0,62675.52,411],
["739","TESCO KRABI","กระบี่","ภาคใต้","2026-05-24",149.0,3527.74,0],
["622","TESCO THATPANOM","นครพนม","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1747.0,62916.17,556],
["3320","Chann 24 Mapyangphon Rayong","ระยอง","ภาคตะวันออก","2026-05-24",3777.0,121010.35,316],
["3144","I'm Park","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1746.0,62514.58,628],
["746","BIG C KORAT 2","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",944.0,44700.89,215],
["966","TESCO RANGSIT KLONG 4","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1819.0,61616.65,461],
["459","BIG C KLONG 6","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1631.0,60849.63,591],
["738","NUMBER ONE PLAZA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1573.0,57815.51,632],
["3147","Jompol Market","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1816.0,64317.92,754],
["558","THE WALK NAKHONSAWAN","นครสวรรค์","ภาคเหนือ","2026-05-24",1150.0,52903.57,174],
["3156","Lotus’s Uttaradit","อุตรดิตถ์","ภาคเหนือ","2026-05-24",1680.0,63969.56,551],
["979","TESCO ROIET","ร้อยเอ็ด","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1703.0,61790.92,614],
["819","ICON SIAM","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",3081.0,107271.06,2193],
["3015","LOTUS TRAKANPHUETPHON","อุบลราชธานี","ภาคตะวันออกเฉียงเหนือ","2026-05-24",937.0,43213.48,299],
["425","Lotus's Lamai","สุราษฎร์ธานี","ภาคใต้","2026-05-24",1773.0,65274.16,549],
["3322","The Backyard Mahidol","เชียงใหม่","ภาคเหนือ","2026-05-24",3.0,79.02,2],
["387","TESCO AMATA CITY","ชลบุรี","ภาคตะวันออก","2026-05-24",6.0,126.32,0],
["3180","Big C Ratchada","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1433.0,65046.99,518],
["3137","Meet&Eat Nakornpathom","นครปฐม","กรุงเทพมหานครและปริมณฑล","2026-05-24",1793.0,64218.46,575],
["647","BIG C KRANUAN","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1306.0,49294.53,544],
["808","Lotus's Bangpakok","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",954.0,42147.07,275],
["789","Lotus's Ruamchok","เชียงใหม่","ภาคเหนือ","2026-05-24",1636.0,58583.69,430],
["545","TESCO CHAINAT","ชัยนาท","ภาคกลาง","2026-05-24",256.0,6748.78,74],
["953","คลังคืนสินค้าOnlineวังน้อย","Online","Online","2026-05-24",78.0,1848.49,0],
["262","IMPERIAL LADPRAO","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",925.0,43369.19,339],
["741","LITTLE WALK BANGNA","กรุงเทพฯ","กรุงเทพมหานครและปริมณฑล","2026-05-24",1057.0,41364.93,89],
["405","OCEAN CHUMPORN","ชุมพร","ภาคใต้","2026-05-24",600.0,19919.37,123],
["3328","Chombueng Market (Ratchaburi)","ราชบุรี","ภาคตะวันตก","2026-05-24",1801.0,60314.49,353],
["915","Online","Online","Online","2026-05-24",25.0,725.92,258],
["914","Online","Online","Online","2026-05-24",0.0,0.0,45],
["910","e-Store CCE","Online","Online","2026-05-24",0.0,0.0,264],
["3329","Central Khonkaen Campus","ขอนแก่น","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1418.0,47771.45,302],
["3330","Save One Korat","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",1531.0,51809.41,209],
["3333","Bantew Market  (Loei)","เลย","ภาคตะวันออกเฉียงเหนือ","2026-05-24",569.0,19888.52,243],
["3336","Maepimthong Nonsung","นครราชสีมา","ภาคตะวันออกเฉียงเหนือ","2026-05-24",2119.0,75339.15,462],
["717","Thai Watsadu Bangna","Others","Others","2026-05-24",0.0,0.0,0],
["716","Thai Watsadu Srisaman","Others","Others","2026-05-24",0.0,0.0,0],
["714","Robinson Srisaman","นนทบุรี","กรุงเทพมหานครและปริมณฑล","2026-05-24",0.0,0.0,0],
["713","Koh Lanta","กระบี่","ภาคใต้","2026-05-24",0.0,0.0,0],
["711","Karon Beach","ภูเก็ต","ภาคใต้","2026-05-24",0.0,0.0,0],
["173","คลังคืนสินค้าOffline ขอนแก่น","WH/DC","WH/DC","2026-05-24",0.0,0.0,0],
["171","คลังคืนสินค้าOffline บางนา","WH/DC","WH/DC","2026-05-24",0.0,0.0,0]];


// ── CLIENT B MASTER BRANCH DATA ────────────────────────────────────────────────
const CLIENT_B_BRANCHES = [
["1","DHL SUPPLY CHAIN (THAILAND)","","","2026-05-31",18580.0,859237.15,0.0],
["2","DHL (RETURN TO SUPPLIER)","","","2026-05-31",1649.0,79950.65,0.0],
["3","DHL (DAMAGE STOCK)","","","2026-05-31",2.0,78.04,0.0],
["3940","SCB Park Plaza","","","2026-05-31",120.0,6034.84,4.0],
["3942","Royal Garden Pattaya","","","2026-05-31",531.0,26924.97,8.0],
["3947","Central Ladprao","","","2026-05-31",465.0,24107.12,23.0],
["3948","Central Rama 3","","","2026-05-31",514.0,27275.62,25.0],
["3954","Times Square","","","2026-05-31",9.0,486.02,6.0],
["3963","Pacific Park Sriracha","","","2026-05-31",554.0,29962.69,16.0],
["3965","Lotus's Suratthani","","","2026-05-31",797.0,38484.49,7.0],
["3966","Market Place Sukhapiban 3","","","2026-05-31",613.0,33053.75,10.0],
["3971","Silom Soi 2","","","2026-05-31",519.0,25394.12,19.0],
["3977","Central Udonthani","","","2026-05-31",529.0,27938.36,44.0],
["3978","Chiang Rai Night Bazaar","","","2026-05-31",574.0,31246.13,1.0],
["3979","The Mall Thapra","","","2026-05-31",191.0,8173.26,5.0],
["3985","Passione Rayong","","","2026-05-31",580.0,29750.83,7.0],
["3986","MBK (2 Fl)","","","2026-05-31",84.0,3671.73,11.0],
["3990","Lee Gardens Hat Yai","","","2026-05-31",455.0,24262.82,83.0],
["3992","The Mall Korat","","","2026-05-31",603.0,33037.8,16.0],
["4006","La Villa Phaholyothin","","","2026-05-31",51.0,2327.02,7.0],
["4007","Exchange Tower","","","2026-05-31",109.0,4404.73,14.0],
["4009","Esplanade Ratchada","","","2026-05-31",675.0,34767.05,3.0],
["4011","The Avenue Pattaya","","","2026-05-31",113.0,5472.6,1.0],
["4013","Tha Phae Gate Chiang Mai","","","2026-05-31",538.0,27858.25,24.0],
["4016","Lotus Ban Fah","","","2026-05-31",479.0,24913.65,0.0],
["4017","Silom Soi 4","","","2026-05-31",556.0,28654.88,16.0],
["4021","SVB Airport (Dometic BF)","","","2026-05-31",202.0,8092.37,12.0],
["4024","SVB Airport (Inter TE)","","","2026-05-31",81.0,4031.03,7.0],
["4025","SVB Airport (Landside4F)","","","2026-05-31",167.0,7556.37,22.0],
["4026","SVB Airport (Inter TW)","","","2026-05-31",537.0,28275.25,81.0],
["4028","The Crystal Park","","","2026-05-31",112.0,5762.56,2.0],
["4030","Lotus's Rangsit","","","2026-05-31",402.0,21536.07,8.0],
["4033","Asawann Nongkhai 1","","","2026-05-31",62.0,3766.18,4.0],
["4034","MBK (3 Fl)","","","2026-05-31",10.0,342.75,0.0],
["4036","Big C Samui","","","2026-05-31",83.0,3909.29,2.0],
["4037","Big C Hang Dong","","","2026-05-31",642.0,29468.42,10.0],
["4039","The Avenue Ratchayothin","","","2026-05-31",615.0,31404.38,13.0],
["4044","Silom Complex","","","2026-05-31",31.0,1466.88,6.0],
["4047","Siam Paragon (G Fl)","","","2026-05-31",439.0,23332.61,74.0],
["4068","Central Chaengwattana(3F)","","","2026-05-31",640.0,33822.41,16.0],
["4069","Big C Phuket","","","2026-05-31",90.0,5156.7,2.0],
["4071","Central Pattaya (3F)","","","2026-05-31",502.0,27761.3,33.0],
["4072","Central Chonburi","","","2026-05-31",568.0,28833.18,14.0],
["4073","Lotus's Srinakarin","","","2026-05-31",526.0,25887.31,12.0],
["4074","Market Village Phuket","","","2026-05-31",68.0,3170.69,1.0],
["4075","Interchange 21","","","2026-05-31",82.0,3314.84,4.0],
["4077","Fashion Island (2 Fl)","","","2026-05-31",565.0,29486.73,14.0],
["4078","Central Khonkaen (2 Fl)","","","2026-05-31",619.0,32284.07,10.0],
["4079","Esplanade Rattanatibet","","","2026-05-31",27.0,1313.03,5.0],
["4097","Energy Complex","","","2026-05-31",5.0,292.75,0.0],
["4100","K Village","","","2026-05-31",91.0,3314.65,7.0],
["4101","Market Place Wongsawang","","","2026-05-31",81.0,4546.34,4.0],
["4103","Lotus's Laksi","","","2026-05-31",12.0,609.44,0.0],
["4104","Seacon Square (2 Fl)","","","2026-05-31",507.0,28294.95,23.0],
["4107","Supreme Samsen","","","2026-05-31",23.0,1032.35,0.0],
["4108","Central Khonkaen (G Fl)","","","2026-05-31",205.0,11221.81,4.0],
["4112","RSU Sukhumvit 31","","","2026-05-31",29.0,1411.93,0.0],
["4119","Sun Tower","","","2026-05-31",141.0,6798.98,7.0],
["4121","The Crystal Design Center","","","2026-05-31",539.0,27490.18,6.0],
["4122","Central Chiangrai","","","2026-05-31",534.0,28741.58,34.0],
["4123","The Nine Rama 9","","","2026-05-31",445.0,24583.0,36.0],
["4126","Muang Thai Phatra Complex","","","2026-05-31",15.0,763.25,0.0],
["4127","Siam Square","","","2026-05-31",569.0,30322.17,23.0],
["4129","Terminal 21 Asok","","","2026-05-31",123.0,4535.63,28.0],
["4130","Sunee Ubon Ratchathani","","","2026-05-31",117.0,4363.46,1.0],
["4134","Laemthong Bangsaen","","","2026-05-31",537.0,28094.62,17.0],
["4135","Central Phitsanulok","","","2026-05-31",116.0,6298.97,4.0],
["4136","Lotus's Thalang","","","2026-05-31",14.0,681.31,3.0],
["4137","Park Ventures","","","2026-05-31",10.0,514.86,0.0],
["4138","Central Rama 9","","","2026-05-31",478.0,26435.45,30.0],
["4140","Lotus's Khonkaen","","","2026-05-31",9.0,476.35,2.0],
["4143","The Mall Bangkapi","","","2026-05-31",97.0,3728.18,13.0],
["4146","Big C Extra Chaengwattana","","","2026-05-31",58.0,3751.19,1.0],
["4147","Int Intersect Rama3","","","2026-05-31",27.0,1184.1,0.0],
["4149","Mega Bangna 1","","","2026-05-31",25.0,1318.92,4.0],
["4151","The Promenade","","","2026-05-31",543.0,28406.11,14.0],
["4152","Gateway Ekkamai","","","2026-05-31",30.0,1457.16,1.0],
["4154","Bangna Tower","","","2026-05-31",489.0,24124.68,21.0],
["4157","Porto Chino","","","2026-05-31",15.0,693.95,6.0],
["4162","Seacon Bangkae (2Fl)","","","2026-05-31",86.0,4593.52,4.0],
["4163","The Scene Town In Town","","","2026-05-31",347.0,18510.12,5.0],
["4165","DMK Airport (Inter1)","","","2026-05-31",363.0,20083.34,53.0],
["4171","Central Suratthani","","","2026-05-31",204.0,11102.89,0.0],
["4177","Sermthai Mahasarakham","","","2026-05-31",587.0,29765.87,7.0],
["4184","Central Ubon Ratchathani","","","2026-05-31",544.0,29619.98,26.0],
["4185","The Walk Kaset Nawamin","","","2026-05-31",582.0,30874.96,11.0],
["4187","Belle Grand Rama 9","","","2026-05-31",18.0,913.88,0.0],
["4191","Big C Extra Pattaya Klang","","","2026-05-31",12.0,550.8,0.0],
["4192","DMK Airport  (Domestic 1)","","","2026-05-31",447.0,23606.22,5.0],
["4194","Big C Extra Ladprao","","","2026-05-31",600.0,30792.54,11.0],
["4195","Mercury Ville","","","2026-05-31",638.0,32928.19,12.0],
["4197","Harbor Mall Laemchabang","","","2026-05-31",14.0,680.47,2.0],
["4198","Maya Chiang Mai (B Fl)","","","2026-05-31",130.0,6504.68,8.0],
["4199","Central Chiangmai (GF)","","","2026-05-31",537.0,26866.04,9.0],
["4201","Robinson Ratchaburi","","","2026-05-31",503.0,26723.34,3.0],
["4202","Robinson Saraburi","","","2026-05-31",2.0,92.28,0.0],
["4205","Central Festival Hatyai","","","2026-05-31",481.0,25380.38,169.0],
["4206","Taweekit Buriram","","","2026-05-31",40.0,2077.01,2.0],
["4208","The Crystal Chaiyapruek","","","2026-05-31",658.0,33668.37,6.0],
["4210","Marketvillage Huahin (BF)","","","2026-05-31",366.0,14883.56,3.0],
["4212","Lotus Thewalk Nakhonsawan","","","2026-05-31",224.0,7317.19,0.0],
["4213","River City","","","2026-05-31",429.0,23123.48,5.0],
["4215","Central Samui (2Fl)","","","2026-05-31",173.0,7300.38,6.0],
["4216","The Paseo Lat Krabang","","","2026-05-31",456.0,24052.45,3.0],
["4220","Beehive Muangthongthani","","","2026-05-31",629.0,31869.48,14.0],
["4222","I'M Park Chula","","","2026-05-31",490.0,25280.05,4.0],
["4223","Thonglor","","","2026-05-31",92.0,4724.49,17.0],
["4224","Plearnary Mall","","","2026-05-31",513.0,26701.03,2.0],
["4225","Lotus Bangyai","","","2026-05-31",16.0,750.96,1.0],
["4227","Ayutthaya City Park","","","2026-05-31",99.0,5202.21,3.0],
["4229","Central Salaya","","","2026-05-31",604.0,32069.27,8.0],
["4230","Robinson Chachoengsao","","","2026-05-31",111.0,5798.87,0.0],
["4233","Tukcom South Pattaya","","","2026-05-31",21.0,1057.95,0.0],
["4237","The Crystal Ratchapreuk","","","2026-05-31",681.0,35201.74,6.0],
["4238","Asawann Nongkhai 2","","","2026-05-31",532.0,27673.32,18.0],
["4240","Future Park Rangsit (BF)","","","2026-05-31",131.0,6853.35,21.0],
["4241","Sathorn Square","","","2026-05-31",63.0,2491.35,3.0],
["4242","Victoria Gardens","","","2026-05-31",680.0,34199.7,2.0],
["4243","Marketvillage Suvarnabhum","","","2026-05-31",550.0,27859.42,33.0],
["4244","Emquartier","","","2026-05-31",661.0,33502.63,69.0],
["4245","Central Rayong","","","2026-05-31",574.0,31070.79,40.0],
["4246","Lotus Chiangmai Kamthieng","","","2026-05-31",761.0,36098.48,11.0],
["4248","Central Lampang","","","2026-05-31",526.0,26908.73,23.0],
["4250","Central Westgate","","","2026-05-31",638.0,34771.01,18.0],
["4251","Robinson Buriram","","","2026-05-31",98.0,5088.56,3.0],
["4252","Bangrak","","","2026-05-31",857.0,46718.55,19.0],
["4253","Lotus's Banpong","","","2026-05-31",20.0,910.64,4.0],
["4254","Lotus's Hat Yai","","","2026-05-31",337.0,17658.48,20.0],
["4257","Riverside Plaza","","","2026-05-31",384.0,20847.33,28.0],
["4262","Robinson Srisaman","","","2026-05-31",82.0,4463.59,0.0],
["4264","Fashion Island (B Fl)","","","2026-05-31",113.0,5829.16,6.0],
["4265","Central Eastville","","","2026-05-31",547.0,30170.66,12.0],
["4266","The Street Ratchada","","","2026-05-31",58.0,3801.72,0.0],
["4267","Central Pinklao","","","2026-05-31",550.0,30164.66,16.0],
["4268","Big C Aranyaprathet","","","2026-05-31",141.0,8292.44,1.0],
["4269","Robinson Mae Sot","","","2026-05-31",85.0,4472.59,11.0],
["4270","Suanplern Market","","","2026-05-31",556.0,30205.02,13.0],
["4271","Big C Lopburi","","","2026-05-31",711.0,35058.65,13.0],
["4272","Emporium","","","2026-05-31",508.0,26199.39,29.0],
["4275","Union Mall","","","2026-05-31",416.0,21202.07,8.0],
["4278","Sindhorn Building","","","2026-05-31",458.0,23918.57,22.0],
["4279","Habito Mall","","","2026-05-31",32.0,1403.05,0.0],
["4280","DMK Airport (Domestic 2)","","","2026-05-31",21.0,962.34,0.0],
["4282","Phuket Airport (Airside)","","","2026-05-31",141.0,6078.66,16.0],
["4284","Khaosan Road","","","2026-05-31",43.0,2086.37,2.0],
["4286","Central Nakhonsithammarat","","","2026-05-31",522.0,26298.82,15.0],
["4287","Central World (3 Fl)","","","2026-05-31",466.0,24272.76,151.0],
["4289","Bluport Hua Hin","","","2026-05-31",624.0,32478.61,15.0],
["4290","Robinson Chantaburi","","","2026-05-31",712.0,36979.6,13.0],
["4292","Central Marina Pattaya","","","2026-05-31",228.0,9398.33,1.0],
["4293","Terminal 21 Korat","","","2026-05-31",573.0,31060.83,14.0],
["4297","G Tower","","","2026-05-31",627.0,31081.35,16.0],
["4298","Tha Phae Chiang Mai 2","","","2026-05-31",604.0,32398.67,20.0],
["4299","Jungceylon (B Fl)","","","2026-05-31",590.0,31332.94,33.0],
["4303","Robinson Phetchaburi","","","2026-05-31",539.0,28895.79,3.0],
["4304","Lotus's Songkhla","","","2026-05-31",13.0,593.47,0.0],
["4307","Lotus's Pakchong","","","2026-05-31",567.0,27784.26,9.0],
["4311","Central Korat","","","2026-05-31",139.0,8081.43,7.0],
["4312","Central Mahachai","","","2026-05-31",520.0,27208.82,12.0],
["4315","Maya Chiang Mai (2 Fl)","","","2026-05-31",125.0,6868.65,6.0],
["4317","Jungceylon (G Fl)","","","2026-05-31",74.0,3788.36,11.0],
["4319","Central Samui (1Fl)","","","2026-05-31",81.0,3831.82,4.0],
["4320","Central World (4 Fl)","","","2026-05-31",621.0,31571.52,53.0],
["4321","Lazada","","","2026-05-31",3073.0,90754.02,193.0],
["4323","Yaowarat Road","","","2026-05-31",81.0,4010.15,13.0],
["4326","CentralPhuket Floresta-GF","","","2026-05-31",96.0,4446.56,6.0],
["4327","CentralPhuket Floresta-3F","","","2026-05-31",678.0,33098.65,16.0],
["4328","Bali Hai Pattaya","","","2026-05-31",603.0,28321.82,14.0],
["4329","Terminal 21 Pattaya","","","2026-05-31",529.0,27093.43,56.0],
["4330","Central Phuket (4F)","","","2026-05-31",572.0,29727.55,60.0],
["4333","Central Chiangmai (4F)","","","2026-05-31",616.0,31326.62,34.0],
["4334","True Digital Park","","","2026-05-31",573.0,28491.52,13.0],
["4335","Central Patong","","","2026-05-31",29.0,1370.78,0.0],
["4339","Nimman Road","","","2026-05-31",28.0,1298.82,0.0],
["4340","Grand 5 Sukhumvit","","","2026-05-31",122.0,5656.89,5.0],
["4345","Central Village","","","2026-05-31",549.0,28656.5,45.0],
["4346","Wat Phra Singh Chiang Mai","","","2026-05-31",21.0,1129.29,3.0],
["4348","Central Pattaya Beach(GF)","","","2026-05-31",459.0,25258.85,32.0],
["4349","Samyan Mitrtown","","","2026-05-31",166.0,8882.61,6.0],
["4350","Robinson Trang","","","2026-05-31",486.0,25822.04,26.0],
["4353","Lotus North Pattaya","","","2026-05-31",377.0,20002.43,4.0],
["4354","Lotus's Salaya","","","2026-05-31",498.0,23327.33,8.0],
["4357","Sukhumvit Soi 17","","","2026-05-31",73.0,3790.25,10.0],
["4358","The Mall Ngamwongwan","","","2026-05-31",423.0,22077.09,14.0],
["4360","Terminal21 Rama3","","","2026-05-31",16.0,780.13,2.0],
["4361","Robinson Thalang","","","2026-05-31",91.0,4752.72,0.0],
["4362","ICS","","","2026-05-31",460.0,23288.5,42.0],
["4363","DMK Airport (Inter2)","","","2026-05-31",329.0,17301.99,15.0],
["4364","DMK Airport (Landside)","","","2026-05-31",43.0,2277.79,1.0],
["4365","The Palladium","","","2026-05-31",19.0,857.37,2.0],
["4367","Porto de Phuket","","","2026-05-31",206.0,8126.44,12.0],
["4368","Central Ayutthaya","","","2026-05-31",24.0,1160.97,2.0],
["4369","Wave Place Building","","","2026-05-31",17.0,774.02,2.0],
["4370","Jewelry Trade Center","","","2026-05-31",25.0,1100.8,0.0],
["4371","itr Tower","","","2026-05-31",108.0,3753.15,2.0],
["4372","own","","","2026-05-31",536.0,26717.29,15.0],
["4373","Kubon","","","2026-05-31",588.0,29073.65,3.0],
["4374","re Park Rangsit 1FL","","","2026-05-31",474.0,24633.15,20.0],
["4375","le Walk Pattaya","","","2026-05-31",356.0,19403.32,6.0],
["4376","Silom","","","2026-05-31",399.0,21296.0,7.0],
["4378","ral Westville","","","2026-05-31",562.0,29602.35,9.0],
["4379","here","","","2026-05-31",147.0,5733.84,23.0],
["4383","Asiatique","","","2026-05-31",185.0,6504.03,2.0],
["4384","Siam Premium Outlet","","","2026-05-31",654.0,34276.41,40.0],
["4385","Dannok","","","2026-05-31",464.0,22666.97,64.0],
["4386","Empire Tower","","","2026-05-31",452.0,23556.42,25.0],
["4387","Lagoon Road Phuket","","","2026-05-31",110.0,4024.5,1.0],
["4388","One Bangkok","","","2026-05-31",523.0,26284.29,9.0],
["4389","Pattaya Sai 2 Road","","","2026-05-31",90.0,4430.88,0.0],
["4390","Platinum","","","2026-05-31",606.0,29390.99,80.0],
["4391","Central Hadyai 3rd Flr","","","2026-05-31",515.0,24086.53,95.0],
["4392","Marketplace Ekkamai","","","2026-05-31",19.0,867.77,0.0],
["4393","Robinson Latkrabang","","","2026-05-31",66.0,3655.9,22.0],
["4394","Zuellig House Saladeang","","","2026-05-31",329.0,17894.28,12.0],
["4395","ral Nakhonpathom","","","2026-05-31",413.0,23304.14,0.0],
["4396","Centerpoint Siam Square","","","2026-05-31",335.0,18849.22,0.0],
["4442","Shopee","","","2026-05-31",-135.0,-5193.98,153.0],
["4455","TikTok","","","2026-05-31",-50.0,-1517.59,0.0],
["4503","Vogue Krabi","","","2026-05-31",19.0,870.94,3.0],
["4504","Major Cineplex Sukhumvit","","","2026-05-31",416.0,21810.12,21.0],
["4511","The Mall Bangkae","","","2026-05-31",459.0,25179.84,7.0],
["4515","JC Town Square Wanghin","","","2026-05-31",625.0,31669.18,3.0],
["4518","All Season Place","","","2026-05-31",15.0,687.79,2.0],
["4519","Siam Paragon (4 Fl)","","","2026-05-31",448.0,20061.7,16.0],
["4520","Central Rama 2","","","2026-05-31",494.0,25591.46,14.0],
["4521","Central CM Airport (2F)","","","2026-05-31",536.0,27619.74,7.0],
["4522","Big C South Pattaya","","","2026-05-31",23.0,1089.9,1.0],
["4523","Big C Ratchadamri","","","2026-05-31",566.0,29115.01,42.0],
["4525","V Sqaure Nakhon Sawan","","","2026-05-31",805.0,38208.18,6.0],
["4528","Fortune Town","","","2026-05-31",14.0,679.62,2.0],
["4529","Imperial World Ladprao","","","2026-05-31",19.0,942.07,2.0],
["4535","Marketvillage Huahin (GF)","","","2026-05-31",494.0,26038.35,23.0],
["4537","Q House Lumpini","","","2026-05-31",110.0,4217.29,6.0],
["4541","Jasmine City","","","2026-05-31",69.0,2840.48,0.0],
["4543","Lotus's Phuket","","","2026-05-31",623.0,32128.84,16.0],
["4545","Marketplace Thonglor","","","2026-05-31",521.0,26266.42,20.0],
["4547","Paradise Park","","","2026-05-31",566.0,28114.03,10.0],
["4550","Big C Hat Yai 2","","","2026-05-31",529.0,26739.56,21.0],
["4558","Big C Aomyai","","","2026-05-31",420.0,22414.53,17.0],
["4561","Ploenchit Center","","","2026-05-31",91.0,3498.04,4.0],
["4566","Central Bangna","","","2026-05-31",541.0,29599.63,3.0],
["4574","Century Rangnam","","","2026-05-31",20.0,942.09,1.0],
["4578","Big C Rama 4","","","2026-05-31",549.0,26978.38,15.0],
["4805","Pop up Eight Thonglor","","","2026-05-31",74.0,4024.76,5.0],
["4807","Pop-up Gateway Bangsue","","","2026-05-31",3.0,120.45,0.0],
["4809","Pop Up Singha Complex","","","2026-05-31",495.0,25490.2,3.0],
["4811","POP UP Chamchuri Sq.","","","2026-05-31",2.0,80.3,2.0],
["4815","Pop Up The Fourth Sai 4","","","2026-05-31",621.0,30997.24,5.0],
["4999","Warehouse BMA","","","2026-05-31",1832.0,78130.04,48.0]];

// ── CLIENT C MASTER BRANCH DATA ────────────────────────────────────────────────
const CLIENT_C_BRANCHES = [
["1","","","","2026-05-31",5209.0,186635.31,null],
["DC002","","","","2026-05-31",38526.0,1228575.62,null],
["ST001","สาขาเซ็นเตอร์ พ้อยส์ ออฟ สยาม (ST001)","Bangkok","Central","2026-05-31",1799.0,65290.72,null],
["ST002","สาขาซีคอนบางแค (ST002)","Bangkok","Central","2026-05-31",2641.0,84758.77,null],
["ST003","สาขาจีทาวเวอร์ (ST003)","Bangkok","Central","2026-05-31",1544.0,55667.05,null],
["ST009","สาขาสยามสแควร์ (ST009)","Bangkok","Central","2026-05-31",6044.0,199685.57,null],
["ST010","สาขาเซ็นทรัลปิ่นเกล้า (ST010)","Bangkok","Central","2026-05-31",2727.0,86966.71,null],
["ST011","สาขาเซ็นทรัลเวสต์เกต (ST011)","Nonthaburi ","Central","2026-05-31",3695.0,122174.1,null],
["ST012","สาขาเซ็นทรัลชลบุรี (ST012)","Chonburi","Eastern","2026-05-31",2267.0,70284.82,null],
["ST013","สาขาโรบินสันลาดกระบัง (ST013)","Bangkok","Central","2026-05-31",1910.0,64766.27,null],
["ST014","สาขาเซ็นทรัลพระราม2 (ST014)","Bangkok","Central","2026-05-31",2104.0,73094.41,null],
["ST015","สาขาเซ็นทรัลพัทยาบีช (ST015)","Chonburi","Eastern","2026-05-31",2143.0,74223.58,null],
["ST016","สาขาเซ็นทรัลระยอง (ST016)","Rayong ","Eastern","2026-05-31",1858.0,68216.33,null],
["ST017","สาขาเซ็นทรัลหาดใหญ่ (ST017)","Songkhla ","Southern","2026-05-31",4098.0,139094.29,null],
["ST018","สาขาเซ็นทรัลขอนแก่น (ST018)","Khon Kaen ","Northeastern","2026-05-31",2672.0,90834.2,null],
["ST019","สาขาเซ็นทรัลศรีราชา (ST019)","Chonburi","Eastern","2026-05-31",1829.0,64637.28,null],
["ST020","สาขาเซ็นทรัลอยุธยา (ST020)","Phra Nakhon Si Ayutthaya ","Central","2026-05-31",1860.0,63023.27,null],
["ST021","สาขาเซ็นทรัลอุดรธานี (ST021)","Udon Thani ","Northeastern","2026-05-31",1988.0,71124.6,null],
["ST022","สาขาเซ็นทรัลวิลเลจ (ST022)","Samut Prakan ","Central","2026-05-31",1684.0,59902.71,null],
["ST023","สาขาโรบินสันฉะเชิงเทรา (ST023)","Chachoengsao ","Eastern","2026-05-31",1572.0,58626.83,null],
["ST024","สาขาเซ็นทรัลศาลายา (ST024)","Nakhon Pathom ","Central","2026-05-31",1793.0,64673.55,null],
["ST025","สาขาเซ็นทรัลจันทบุรี (ST025)","Chanthaburi ","Eastern","2026-05-31",2141.0,75684.23,null],
["ST026","สาขาเซ็นทรัลสุราษฎร์ (ST026)","Surat Thani ","Southern","2026-05-31",1955.0,67341.48,null],
["ST027","สาขาเซ็นทรัลพระราม9 (ST027)","Bangkok","Central","2026-05-31",2519.0,78945.99,null],
["ST028","สาขาโรบินสันถลาง (ST028)","Phuket ","Southern","2026-05-31",1704.0,61584.23,null],
["ST029","สาขาโรบินสันราชพฤกษ์ (ST029)","Nonthaburi ","Central","2026-05-31",1351.0,43992.81,null],
["ST030","สาขาไอคอนสยาม (ST030)","Bangkok","Central","2026-05-31",2040.0,72018.92,null],
["ST031","สาขาฟิวเจอร์พาร์ครังสิต(ST031)","Bangkok","Central","2026-05-31",2003.0,70152.97,null],
["ST032","สาขาเซ็นทรัลรามอินทรา (ST032)","Bangkok","Central","2026-05-31",1815.0,64431.15,null],
["ST033","สาขาเซ็นทรัลพระราม3 (ST033)","Bangkok","Central","2026-05-31",1833.0,63822.89,null],
["ST034","สาขาศูนย์การค้าจังซีลอน (ST034)","Phuket ","Southern","2026-05-31",2276.0,76279.09,null],
["ST035","สาขาเซ็นทรัลลาดพร้าว (ST035)","Bangkok","Central","2026-05-31",3800.0,127669.72,null],
["ST036","สาขาเอสพลานาดรัชดาภิเษก (ST036)","Bangkok","Central","2026-05-31",1616.0,57523.62,null],
["ST037","สาขาโรบินสันฉลองภูเก็ต (ST037)","Phuket ","Southern","2026-05-31",1538.0,55844.92,null],
["ST038","สาขาโรบินสันสุพรรณบุรี (ST038)","Suphan buri","Central","2026-05-31",1578.0,58635.11,null],
["ST039","สาขาเซ็นทรัลโคราช (ST039)","Nakhon Ratchasima","Northeastern","2026-05-31",1642.0,58956.82,null],
["ST040","สาขาเซ็นทรัลเวสต์วิลล์ (ST040)","Nonthaburi ","Central","2026-05-31",1595.0,59769.37,null],
["ST041","สาขาเซ็นทรัลเวิลด์ (ST041)","Bangkok","Central","2026-05-31",3642.0,130038.09,null],
["ST042","สาขาเซ็นทรัลพิษณุโลก (ST042)","Phitsanulok","Central","2026-05-31",2174.0,65851.78,null],
["ST043","สาขาเซ็นทรัลนครสวรรค์ (ST043)","Nakhon Sawan","Central","2026-05-31",1638.0,60682.56,null],
["ST044","สาขามาร์เก็ตวิลเลจหัวหิน (ST044)","Prachuap Khiri Khan","Western","2026-05-31",1823.0,66899.4,null],
["ST045","สาขาเซ็นทรัลนครปฐม (ST045)","Nakhon Pathom ","Central","2026-05-31",1907.0,68917.24,null],
["ST046","สาขาโรบินสันตรัง (ST046)","Trang","Southern","2026-05-31",1557.0,57787.46,null],
["ST047","สาขาอัศวรรณหนองคาย (ST047)","Nong Khai","Northeastern","2026-05-31",1265.0,40349.98,null],
["ST048","สาขามาร์เก็ตวิลเลจสุวรรณภูมิ (ST048)","Samut Prakan ","Central","2026-05-31",1688.0,60168.56,null],
["ST049","สาขาโลตัส สงขลา (ST049)","Songkhla ","Southern","2026-05-31",1564.0,55688.64,null],
["ST050","สาขาศูนย์การค้าเอเชียทีค (ST050)","Bangkok","Central","2026-05-31",1800.0,61236.17,null],
["ST051","สาขาเซ็นทรัลแจ้งวัฒนะ (ST051)","Bangkok","Central","2026-05-31",1870.0,63340.12,null],
["ST052","สาขาศูนย์การค้าสีลมคอมเพล็กซ์ (ST052)","Bangkok","Central","2026-05-31",1270.0,39028.81,null],
["ST053","สาขาโลตัส ยะลา (ST053)","Yala","Southern","2026-05-31",1089.0,34328.08,null],
["ST054","สาขาศูนย์การค้าเกตเวย์บางซื่อ (ST054)","Bangkok","Central","2026-05-31",1697.0,63129.69,null],
["ST055","สาขาโอเชี่ยน ชุมพร (ST055)","Chumphon","Southern","2026-05-31",1298.0,41033.45,null],
["ST056","สาขาเซ็นทรัลมหาชัย (ST056)","Samut sakhon","Central","2026-05-31",1648.0,60108.62,null],
["ST057","สาขาโคลีเซี่ยม ยะลา (ST057)","Yala","Southern","2026-05-31",1474.0,57044.29,null],
["ST058","สาขาเซ็นทรัล สมุย (ST058)","Surat Thani ","Southern","2026-05-31",1826.0,68862.89,null],
["ST059","สาขาทวีกิจ บุรีรัมย์ (ST059)","Buriram ","Northeastern","2026-05-31",1293.0,39636.53,null],
["ST060","สาขาโลตัส สามกอง ภูเก็ต(ST060)","Phuket ","Southern","2026-05-31",1602.0,57497.14,null],
["ST061","สาขาโรบินสัน บ่อวิน (ST061)","Chonburi","Eastern","2026-05-31",1738.0,63586.94,null],
["ST062","สาขาไดอาน่า ปัตตานี (ST062)","Pattani","Southern","2026-05-31",1901.0,64772.62,null],
["ST063","สาขาโรบินสัน บ้านฉาง (ST063)","Rayong ","Eastern","2026-05-31",1574.0,59276.28,null],
["ST064","สาขาบิ๊กซี ปัตตานี (ST064)","Pattani","Southern","2026-05-31",1152.0,36594.72,null],
["ST065","สาขาแม็คโคร สาทร (ST065)","Bangkok","Central","2026-05-31",1142.0,34759.47,null],
["ST066","สาขาแม็คโคร ถ.ศรีอยุธยา(ST066)","Bangkok","Central","2026-05-31",1669.0,59727.35,null],
["ST067","สาขามาร์เก็ตเพลส เทพรักษ์(ST067)","Bangkok","Central","2026-05-31",1190.0,38972.04,null],
["ST068","สาขาเซ็นทรัล มารีนา (ST068)","Chonburi","Eastern","2026-05-31",1689.0,57857.94,null],
["ST069","สาขาโรบินสันไลฟ์สไตล์ สระบุรี(ST069)","Saraburi","Central","2026-05-31",1451.0,55927.22,null],
["ST070","สาขาโลตัส สระบุรี(ST070)","Saraburi","Central","2026-05-31",1483.0,56476.21,null],
["ST071","สาขาหัวหมาก ทาวน์เซ็นเตอร์ (ST071)","Bangkok","Central","2026-05-31",1097.0,35660.68,null],
["ST072","สาขาอเกณฑ์ พัทยา(ST072)","Chonburi","Eastern","2026-05-31",1584.0,59256.1,null],
["ST073","สาขาเซ็นทรัล ลำปาง(ST073)","Lampang","Northern","2026-05-31",1614.0,58594.32,null],
["ST074","สาขาโรบินสัน ไลฟ์สไตล์ ราชบุรี (ST074)","Ratchaburi","Western","2026-05-31",1280.0,42168.96,null],
["ST075","สาขาเซ็นทรัล เชียงใหม่แอร์พอร์ต(ST075)","Chiang Mai","Northern","2026-05-31",1665.0,59757.72,null],
["ST076","สาขาโลตัส สมุย (ST076)","Surat Thani ","Southern","2026-05-31",1781.0,61432.51,null],
["ST077","สาขาโรบินสัน ไลฟ์สไตล์ สุรินทร์ (ST077)","Surin","Northeastern","2026-05-31",1292.0,41496.9,null],
["ST078","สาขาโรบินสัน เพชรบุรี (ST078)","Phetchaburi ","Central","2026-05-31",1700.0,61760.81,null],
["ST079","สาขาเมกาบางนา (ST079)","Samut Prakan ","Central","2026-05-31",2293.0,81163.42,null],
["ST080","สาขาดุสิต เซ็นทรัล พาร์ค(ST080)","Bangkok","Central","2026-05-31",1857.0,72351.52,null],
["ST081","สาขาโลตัสบ่อวิน(ST081)","Chonburi","Eastern","2026-05-31",1550.0,60101.87,null],
["ST082","สาขาโรบินสัน ไลฟ์สไตล์ ศรีสมาน(ST082)","Nonthaburi ","Central","2026-05-31",1646.0,59465.45,null],
["ST083","สาขาโลตัส นครศรีธรรมราช(ST083)","Nakhon si thammarat","Southern","2026-05-31",1629.0,60962.1,null],
["ST084","สาขาเซ็นทรัลเฟสติวัล อีสต์วิลล์(ST084)","Bangkok","Central","2026-05-31",1707.0,62263.12,null],
["ST085","สาขาโรบินสันไลฟ์สไตล์ บุรีรัมย์(ST085)","Buriram ","Northeastern","2026-05-31",1740.0,63038.47,null],
["ST086","สาขาโรบินสัน ไลฟ์สไตล์ กาญจนบุรี(ST086)","Kanchanaburi","Central","2026-05-31",1591.0,58053.52,null],
["ST087","สาขาเซ็นทรัล กระบี่(ST087)","Krabi","Southern","2026-05-31",1662.0,56807.31,null],
["ST088","สาขาโลตัส บ้านบึง(ST088)","Chonburi","Eastern","2026-05-31",1862.0,65840.18,null],
["ST089","สาขาเซ็นทรัล นครศรีธรรมราช(ST089)","Nakhon si thammarat","Southern","2026-05-31",1798.0,64487.84,null],
["ST090","สาขาเซ็นทรัลพลาซา บางนา(ST090)","Bangkok","Central","2026-05-31",1835.0,63439.92,null],
["ST091","สาขาแปซิฟิค พาร์ค ศรีราชา(ST091)","Chonburi","Eastern","2026-05-31",1352.0,42005.34,null],
["ST092","สาขาบิ๊กซี นครปฐม(ST092)","Nakhon Pathom ","Central","2026-05-31",1207.0,39503.67,null],
["ST093","สาขาเดอะพาซิโอ กาญจนา(ST093)","Kanchanaburi","Central","2026-05-31",1669.0,63712.16,null],
["ST094","สาขาแม็คโคร บ้านไผ่(ST094)","Khon Kaen","Northeastern","2026-05-31",1341.0,40595.25,null],
["ST095","สาขาโลตัส ศรีราชา(ST095)","Chonburi","Eastern","2026-05-31",1487.0,45750.11,null],
["ST096","สาขาโรบินสันสมุทรปราการ(ST096)","Samut Prakan ","Central","2026-05-31",1200.0,36453.63,null],
["ST097","สาขาบิ๊กซี ฉะเชิงเทรา(ST097)","Chachoengsao ","Eastern","2026-05-31",1164.0,36442.62,null],
["ST099","สาขาเซ็นทรัล เชียงใหม่(ST099)","Chiang Mai","Northern","2026-05-31",1920.0,67791.52,null],
["ST100","สาขาโรบินสัน ชลบุรี(ST100)","Chonburi","Eastern","2026-05-31",2006.0,66888.74,null],
["ST101","สาขาโลตัส อมตะนคร(ST101)","Chonburi","Eastern","2026-05-31",1083.0,35069.04,null],
["ST102","สาขาสุรินทร์ พลาซ่า(ST102)","Surin","Northeastern","2026-05-31",1313.0,42096.96,null],
["ST103","สาขาโลตัส พัทลุง(ST103)","Phatthalung","Southern","2026-05-31",1843.0,55011.02,null],
["ST104","สาขาโรบินสัน กำแพงเพชร(ST104)","Kamphaeng Phet","Northern","2026-05-31",1319.0,42355.58,null],
["ST105","สาขาโลตัส ชุมพร(ST105)","Chumphon","Southern","2026-05-31",1685.0,48297.86,null],
["ST107","สาขาโลตัส แกลง ระยอง(ST107)","Rayong ","Eastern","2026-05-31",1233.0,40359.8,null],
["ST108","สาขาไชยแสงดีพาร์ทเม้นท์(ST108)","Singburi","Central","2026-05-31",1628.0,57442.6,null],
["ST109","สาขาเซ็นเตอร์ วัน(ST109)","Bangkok","Central","2026-05-31",1534.0,56176.06,null],
["ST110","สาขาสยามสเคป(ST110)","Bangkok","Central","2026-05-31",1282.0,44393.65,null],
["ST112","สาขาเซ็นทรัล ขอนแก่น แคมปัส(ST112)","Khonkaen","Northeastern","2026-05-31",2562.0,93671.0,null],
["WH008","online","online","online","2026-05-31",3106.0,120486.09,null],
["WH009","","","","2026-05-31",2389.0,70894.38,null],
["ST004","สาขาซีคอน ศรีนครินทร์ (ST004)","Bangkok","Central","2026-05-31",0.0,0.0,null],
["WH006","online","online","online","2026-05-31",0.0,0.0,null],
["WH010","online","online","online","2026-05-31",0.0,0.0,null]];

// ── CLIENT D MASTER BRANCH DATA ────────────────────────────────────────────────
const CLIENT_D_BRANCHES = [
["00_WH1","","","","2026-05-31",0.0,0.0,null],
["02_KKU","KHON KAEN UNIVERSITY","","","2026-05-31",3399.0,102652.5,null],
["05_ZPL","ZPELL (RUNGSIT)","","","2026-05-31",7641.0,222958.4,null],
["06_MGB","MEGA BANGNA","","","2026-05-31",8594.0,271945.2,null],
["07_KRT","TERMINAL 21 KORAT","","","2026-05-31",3326.0,100687.1,null],
["08_SQ1","SIAM SQUARE 1","","","2026-05-31",11186.0,358555.3,null],
["09_M08","THE MALL BANGKAPI","","","2026-05-31",7433.0,209062.3,null],
["10_M07","THE MALL BANGKAE","","","2026-05-31",4023.0,125392.6,null],
["11_FSH","FASHION ISLAND","","","2026-05-31",8635.0,268575.8,null],
["12_ASK","TERMINAL 21 ASOK","","","2026-05-31",10454.0,314451.8,null],
["13_PTY","TERMINAL 21 PATTAYA","","","2026-05-31",11399.0,305519.1,null],
["14_MYA","MAYA CHIANG MAI","","","2026-05-31",6107.0,182603.8,null],
["15_SPO","SIAM PREMIUM OUTLET","","","2026-05-31",2721.0,89123.25,null],
["16_SMT","SAMYAN MITR TOWN","","","2026-05-31",3977.0,127957.6,null],
["17_M06","THE MALL NGAMWONGWAN","","","2026-05-31",3863.0,117461.6,null],
["18_M05","THE MALL THAPRA","","","2026-05-31",3898.0,121925.5,null],
["19_MBK","MBK CENTER","","","2026-05-31",10645.0,301288.3,null],
["20_SCS","SEACON SQUARE","","","2026-05-31",4408.0,123429.3,null],
["21_RM3","TERMINAL 21 RAMA3","","","2026-05-31",2997.0,99641.7,null],
["22_PTN","PLATINUM FASHION MALL","","","2026-05-31",11955.0,337512.6,null],
["23_PAS","PASSIONE RAYONG","","","2026-05-31",3546.0,105277.8,null],
["24_UDT","UD TOWNUDON THANI","","","2026-05-31",3288.0,100596.8,null],
["26_EMS","EMSPHERE","","","2026-05-31",3290.0,112874.3,null],
["27_IWS","IMPERIAL WORLD SAMRONG","","","2026-05-31",5966.0,158040.6,null],
["28_GAS","GAYSORN","","","2026-05-31",3015.0,98796.42,null],
["29_BCR","BIG C RATCHADAMRI","","","2026-05-31",4025.0,125649.3,null],
["30_ACP","AYUTTHAYA CITY PARK","","","2026-05-31",4227.0,113641.9,null],
["31_STC","SERMTHAI COMPLEX","","","2026-05-31",3559.0,110450.3,null],
["32_CHA","CHARN AT THE AVENUE","","","2026-05-31",2842.0,93327.39,null],
["33_HYV","HADYAI VILLAGE","","","2026-05-31",7396.0,206876.7,null],
["34_CSS","CENTER POINT SIAM SQUARE","","","2026-05-31",8117.0,282616.0,null],
["35_TSR","THE STREET RATCHADA","","","2026-05-31",3908.0,114844.9,null],
["36_UNM","UNION MALL","","","2026-05-31",6562.0,181157.6,null],
["37_TUD","TU DOME PLAZA","","","2026-05-31",2642.0,84990.63,null],
["38_MJP","MAJOR PINKLAO","","","2026-05-31",4036.0,113843.7,null],
["39_TDP","TRUE DIGITAL PARK","","","2026-05-31",2351.0,74816.41,null],
["40_SHS","SAHATHAI GARDEN PLAZA","","","2026-05-31",2919.0,91303.77,null],
["41_M10","THE MALL KORAT(M10)","","","2026-05-31",3484.0,111755.0,null],
["42_VSN","NAKONSAWAN V SQ(VSN)","","","2026-05-31",2725.0,88012.17,null],
["43_SHN","SAHATHAI PLAZA NAKHON SI THAMMARAT","","","2026-05-31",3516.0,108764.1,null],
["44_ACM","AROUND COMMUNITY MALL(ACM)","","","2026-05-31",2179.0,71176.03,null],
["45_PTC","THE PANTIP LIFESTYLE HUB - CHIANGMAI","","","2026-05-31",2705.0,87169.61,null],
["46_FKK","FAIRY PLAZA KHON KAEN_FKK","","","2026-05-31",3120.0,98246.82,null],
["47_PBN","PARC BANGNA_PBN","","","2026-05-31",2611.0,85084.8,null],
["48_BPH","BLUPORT HUAHIN_BPH","","","2026-05-31",2792.0,91764.34,null],
["49_PDP","PARADISE PARK_PDP","","","2026-05-31",2266.0,75201.63,null],
["50_JRT","THE JAS RAMINDRA","","","2026-05-31",3032.0,91788.16,null],
["51_TSP","THE SPHERES","","","2026-05-31",2642.0,83786.86,null],
["52_SNT","SUNEE TOWER UBON","","","2026-05-31",3089.0,104712.8,null],
["53_LTB","LAEMTHONG BANGSEAN","","","2026-05-31",2861.0,93974.0,null],
["54_LTN","LOTUS NORHT RATCHAPREUK","","","2026-05-31",2639.0,91685.44,null],
["55_PNR","PLEARNARY MALL WATCHARAPOL","","","2026-05-31",2176.0,75030.98,null],
["56_OCP","OCEAN SHOPPING MALL (CHUMPORN)","","","2026-05-31",3277.0,104070.2,null],
["57_CBL","COSMO BAZAAR LIFESTYLE MALL","","","2026-05-31",2726.0,90503.62,null],
["58_JSP","JUNGCELON SHOPPING CENTER (PHUKET)","","","2026-05-31",3441.0,112109.0,null],
["59_LLP","LIME LIGHT (PHUKET)","","","2026-05-31",3438.0,106932.4,null],
["60_BCT","BIG C RAMA 2","","","2026-05-31",2520.0,80357.44,null],
["61_TMP","THONBURI MARKET PLACE","","","2026-05-31",2667.0,88810.36,null],
["62_BCA","BIG C AMNAT CHAREON","","","2026-05-31",2800.0,91960.32,null],
["63_JPS","J-PARK SRI RACHA NIHON MURA","","","2026-05-31",2890.0,97601.97,null],
["64_SKS","SK SQUARE","","","2026-05-31",2779.0,86099.37,null],
["65_TKB","TAVEEKIT BURIRUM","","","2026-05-31",3006.0,95596.88,null],
["66_JKB","THE JAS KHUBON","","","2026-05-31",3422.0,117763.0,null],
["67_ONE","ONE BANGKOK","","","2026-05-31",5386.0,174380.1,null],
["68_BCP","BIG C PHUKET","","","2026-05-31",3837.0,119269.1,null],
["70_LSK","LOTUS SRINAKRARIN","","","2026-05-31",2940.0,98913.28,null],
["71_SLM","SILOM COMPLEX","","","2026-05-31",2733.0,88881.02,null],
["72_MJR","MAJOR CINEPLEX RANGSIT","","","2026-05-31",2488.0,86613.81,null],
["73_BTT","BANTHAT TONG","","","2026-05-31",1954.0,81474.46,null],
["74_JCM","JC MALL NAWAMIN","","","2026-05-31",1999.0,83151.89,null],
["75_PTP","","","","2026-05-31",3591.0,125945.0,null],
["ONLINE","ONLINE","","","2026-05-31",23809.0,762666.7,null]];

// ── CLIENT A SCHEDULE DATA (May, Jun, Jul 2026) ────────────────────────────────
const MAY_DATA=[
["May 2026","W1","Tuesday","28-Apr","3276","Bang Prong","BKK-06","Bangkok","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","3299","Lotus Srinakarin","BKK-06","Bangkok","Night"],
  ["May 2026","W1","Tuesday","28-Apr","798","BIG C SAKAEO","E-01","East","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","3284","Udomsuk Market Kabinburi","E-01","East","Morning"],
  ["May 2026","W1","Thursday","30-Apr","3115","Jas Ramindra","BKK-02","Bangkok","Morning"],
  ["May 2026","W1","Saturday","2-May","3061","Metro Mall Phra Ram 9","BKK-05","Bangkok","Morning"],
  ["May 2026","W1","Tuesday","28-Apr","3168","Little Walk Krungthep Kreetha","BKK-03","Bangkok","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","358","UNION MALL","BKK-01","Bangkok","Night"],
  ["May 2026","W1","Tuesday","28-Apr","3213","Century Sukhumvit","BKK-05","Bangkok","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","3042","IM CHAINATOWN","BKK-16","Bangkok","Morning"],
  ["May 2026","W1","Tuesday","28-Apr","3191","Danchang Suphanburi","N-07","North","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","795","TESCO U TONG","N-07","North","Morning"],
  ["May 2026","W1","Tuesday","28-Apr","782","BIG C CHANTHABURI","E-02","East","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","3320","Chann 24 Mapyangphon Rayong","E-02","East","Morning"],
  ["May 2026","W1","Saturday","2-May","537","NATION TOWER","BKK-07","Bangkok","Morning"],
  ["May 2026","W1","Tuesday","28-Apr","3078","JAS GREEN VILLAGE KHUBON","BKK-04","Bangkok","Morning"],
  ["May 2026","W1","Wednesday","29-Apr","446","FORETUNE TOWN","BKK-12","Bangkok","Night"],
  ["May 2026","W1","Wednesday","29-Apr","759","AMPORN AYUTTHAYA","C-01","Central","Morning"],
  ["May 2026","W2","Tuesday","5-May","995","AO NANG KRABI 2","S-05","South","Morning"],
  ["May 2026","W2","Wednesday","6-May","730","AO NANG KRABI","S-05","South","Morning"],
  ["May 2026","W2","Thursday","7-May","472","VOGUE KRABI","S-05","South","Morning"],
  ["May 2026","W2","Monday","4-May","3322","The Backyard Mahidol","N-01","North","Morning"],
  ["May 2026","W2","Wednesday","6-May","732","PAI","N-01","North","Morning"],
  ["May 2026","W2","Thursday","7-May","777","KADKAM PLAZA","N-01","North","Morning"],
  ["May 2026","W2","Friday","8-May","551","MAYA CHIANG MAI","N-01","North","Night"],
  ["May 2026","W2","Monday","4-May","3291","Lasalles Avenue","BKK-05","Bangkok","Morning"],
  ["May 2026","W2","Wednesday","6-May","3217","Big C Yala","S-06","South","Morning"],
  ["May 2026","W2","Thursday","7-May","3234","Lotus Yala","S-06","South","Morning"],
  ["May 2026","W2","Friday","8-May","315","COLISEUM YALA","S-06","South","Morning"],
  ["May 2026","W2","Tuesday","5-May","3280","WIANG PA PAO","N-02","North","Morning"],
  ["May 2026","W2","Wednesday","6-May","868","TESCO PHAYAO","N-02","North","Morning"],
  ["May 2026","W2","Thursday","7-May","783","TOP PLAZA PHAYAO","N-02","North","Morning"],
  ["May 2026","W2","Friday","8-May","3044","TESCO CHIANGKHAM","N-02","North","Morning"],
  ["May 2026","W2","Monday","4-May","478","TUK COM SRIRACHA","E-04","East","Morning"],
  ["May 2026","W2","Wednesday","6-May","3090","ATARA MALL","E-04","East","Morning"],
  ["May 2026","W2","Thursday","7-May","395","HARBOR MALL","E-04","East","Morning"],
  ["May 2026","W2","Friday","8-May","3154","Paradise Park","BKK-05","Bangkok","Morning"],
  ["May 2026","W2","Saturday","9-May","201","Maneeya","BKK-20","Bangkok","Night"],
  ["May 2026","W2","Tuesday","5-May","3164","Naresuan University","N-04","North","Morning"],
  ["May 2026","W2","Wednesday","6-May","3121","TOPLAND PHITSANULOK","N-04","North","Morning"],
  ["May 2026","W2","Thursday","7-May","552","BIG C PHITSANULOK","N-04","North","Night"],
  ["May 2026","W2","Friday","8-May","443","CENTRAL PHITSANULOK","N-04","North","Night"],
  ["May 2026","W2","Monday","4-May","269","LOTUS CHARANSANITWONG","BKK-18","Bangkok","Morning"],
  ["May 2026","W2","Wednesday","6-May","3156","Lotus Uttaradit","N-03","North","Morning"],
  ["May 2026","W2","Thursday","7-May","786","SRIPONG PARK","N-03","North","Morning"],
  ["May 2026","W2","Friday","8-May","3236","Friday Uttaradit","N-03","North","Morning"],
  ["May 2026","W2","Tuesday","5-May","3324","Lotus Amata City","E-05","East","Morning"],
  ["May 2026","W2","Wednesday","6-May","820","BIG C CHACHOENGSAO","E-05","East","Morning"],
  ["May 2026","W2","Thursday","7-May","3109","Robinson Chachoengsao","E-05","East","Night"],
  ["May 2026","W2","Friday","8-May","470","MEGA BANGNA 1","BKK-07","Bangkok","Night"],
  ["May 2026","W2","Tuesday","5-May","420","TESCO CHUMPARE","NE-03","Northeast","Morning"],
  ["May 2026","W2","Wednesday","6-May","3082","Lotus Nongruea","NE-03","Northeast","Morning"],
  ["May 2026","W2","Thursday","7-May","3238","Mancha Khiri","NE-03","Northeast","Morning"],
  ["May 2026","W2","Friday","8-May","233","FAIRY PLAZA KHON-KAEN","NE-03","Northeast","Night"],
  ["May 2026","W2","Tuesday","5-May","3237","Phangkhon","NE-02","Northeast","Morning"],
  ["May 2026","W2","Wednesday","6-May","622","TESCO THATPANOM","NE-02","Northeast","Morning"],
  ["May 2026","W2","Thursday","7-May","899","BIG C NAKHON PHANOM","NE-02","Northeast","Morning"],
  ["May 2026","W2","Friday","8-May","3263","Nakhonphanom Market","NE-02","Northeast","Morning"],
  ["May 2026","W2","Tuesday","5-May","583","TESCO SALAYA","C-04","Central","Morning"],
  ["May 2026","W2","Wednesday","6-May","3211","Big C Omyai","C-05","Central","Morning"],
  ["May 2026","W2","Thursday","7-May","567","ROBINSON SUKHUMVIT","BKK-15","Bangkok","Night"],
  ["May 2026","W2","Tuesday","5-May","745","BIG C PETCHABURI","C-07","Central","Morning"],
  ["May 2026","W2","Wednesday","6-May","742","ROBINSON PETCHBURI","S-06","South","Morning"],
  ["May 2026","W3","Monday","11-May","3045","BIG C UBONRATCHATHANI","NE-05","Northeast","Morning"],
  ["May 2026","W3","Tuesday","12-May","752","TESCO UBONRATCHATHANI","NE-05","Northeast","Morning"],
  ["May 2026","W3","Wednesday","13-May","3270","Det Udom Market","NE-05","Northeast","Morning"],
  ["May 2026","W3","Thursday","14-May","3118","UBON SQUARE","NE-05","Northeast","Morning"],
  ["May 2026","W3","Friday","15-May","486","TESCO WARINCHAMRAP","NE-05","Northeast","Morning"],
  ["May 2026","W3","Monday","11-May","3286","Hatyai Village","S-04","South","Morning"],
  ["May 2026","W3","Tuesday","12-May","291","SINAEHA PLAZA HADYAI","S-04","South","Morning"],
  ["May 2026","W3","Wednesday","13-May","650","CENTRAL HATYAI 2","S-04","South","Night"],
  ["May 2026","W3","Thursday","14-May","548","CENTRAL HATYAI","S-04","South","Night"],
  ["May 2026","W3","Monday","11-May","615","ROBINSON BURIRUM","NE-06","Northeast","Morning"],
  ["May 2026","W3","Tuesday","12-May","3067","Lotus Uthumphonphisai","NE-06","Northeast","Morning"],
  ["May 2026","W3","Wednesday","13-May","3024","LOTUS KHUKHAN","NE-06","Northeast","Morning"],
  ["May 2026","W3","Thursday","14-May","658","TESCO SANGKHA","NE-06","Northeast","Morning"],
  ["May 2026","W3","Friday","15-May","3004","TESCO PRAKONCHAI","NE-06","Northeast","Morning"],
  ["May 2026","W3","Monday","11-May","3268","Kosum Phisai","NE-04","Northeast","Morning"],
  ["May 2026","W3","Tuesday","12-May","787","TESCO PHONTHONG","NE-04","Northeast","Morning"],
  ["May 2026","W3","Wednesday","13-May","979","TESCO ROIET","NE-04","Northeast","Morning"],
  ["May 2026","W3","Thursday","14-May","576","ROBINSON ROI ET","NE-04","Northeast","Night"],
  ["May 2026","W3","Monday","11-May","3126","Homepro Rayong","E-02","East","Morning"],
  ["May 2026","W3","Tuesday","12-May","268","LEAMTHONG SHOPPING PLAZA RAYONG","E-02","East","Night"],
  ["May 2026","W3","Wednesday","13-May","3100","Robinson Banchang","E-02","East","Night"],
  ["May 2026","W3","Thursday","14-May","217","MAHBOONKRONG","BKK-20","Bangkok","Night"],
  ["May 2026","W3","Friday","15-May","3198","The Mall Bangkapi","BKK-03","Bangkok","Night"],
  ["May 2026","W3","Monday","11-May","681","ASAWANN PHONPHISAI","NE-01","Northeast","Morning"],
  ["May 2026","W3","Tuesday","12-May","3323","Asawan Shopping Complex 2","NE-01","Northeast","Morning"],
  ["May 2026","W3","Wednesday","13-May","600","UD TOWN 2","NE-01","Northeast","Morning"],
  ["May 2026","W3","Thursday","14-May","3127","Tops Plaza Nonghan","NE-01","Northeast","Morning"],
  ["May 2026","W3","Friday","15-May","464","CENTRAL PLAZA UDONTHANI","NE-01","Northeast","Night"],
  ["May 2026","W3","Monday","11-May","3047","TESCO BUENGSAMPHAN","N-05","North","Morning"],
  ["May 2026","W3","Tuesday","12-May","776","BIG C WICHIENBURI","N-05","North","Morning"],
  ["May 2026","W3","Wednesday","13-May","417","TOPLAND PETCHABOON","N-05","North","Morning"],
  ["May 2026","W3","Thursday","14-May","605","TESCO PETCHABOON","N-05","North","Morning"],
  ["May 2026","W3","Friday","15-May","3048","TESCO LOMSAK","N-05","North","Morning"],
  ["May 2026","W3","Monday","11-May","3218","Maejo University","N-06","North","Morning"],
  ["May 2026","W3","Tuesday","12-May","3292","Green Park Chiangmai","N-06","North","Morning"],
  ["May 2026","W3","Wednesday","13-May","541","CENTRAL CHIANGMAI 2","N-06","North","Night"],
  ["May 2026","W3","Thursday","14-May","867","CENTRAL CHIANGMAI 2 4F","N-06","North","Night"],
  ["May 2026","W3","Friday","15-May","239","AIRPORT PLAZA CHIANGMAI","N-01","North","Night"],
  ["May 2026","W3","Monday","11-May","688","TESCO CHO HO","NE-07","Northeast","Morning"],
  ["May 2026","W3","Monday","11-May","812","Central Floresta","S-02","South","Morning"],
  ["May 2026","W3","Tuesday","12-May","3235","R.N.Yard","NE-07","Northeast","Morning"],
  ["May 2026","W3","Tuesday","12-May","3117","Robinson Thalang","S-02","South","Morning"],
  ["May 2026","W3","Wednesday","13-May","756","CENTRAL NAKRONRATCHASRIMA 3F","NE-07","Northeast","Morning"],
  ["May 2026","W3","Wednesday","13-May","3240","Tilokuthit Road","S-02","South","Morning"],
  ["May 2026","W3","Thursday","14-May","598","THE MALL KORAT","NE-07","Northeast","Morning"],
  ["May 2026","W3","Thursday","14-May","813","CENTRAL FESTIVAL PHUKET","S-02","South","Morning"],
  ["May 2026","W3","Friday","15-May","3182","Home Garden Ville","NE-07","Northeast","Morning"],
  ["May 2026","W3","Friday","15-May","3171","Robinson Chalong","S-02","South","Morning"],
  ["May 2026","W3","Monday","11-May","3181","Suwan Kleaw Thong Market","C-01","Central","Morning"],
  ["May 2026","W3","Tuesday","12-May","3215","Porto Go Bang Pa-In","C-01","Central","Morning"],
  ["May 2026","W3","Wednesday","13-May","388","AYUTTHAYA PARK","C-01","Central","Night"],
  ["May 2026","W3","Friday","15-May","976","ROBINSON LADKRABANG","BKK-07","Bangkok","Night"],
  ["May 2026","W3","Friday","15-May","245","LOTUS SUKUMVIT 50","BKK-05","Bangkok","Night"],
  ["May 2026","W4","Monday","18-May","3214","AC Saimai","BKK-02","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","3325","Asoke Tower","BKK-15","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","418","BIG C RAJDAMRI","BKK-20","Bangkok","Night"],
  ["May 2026","W4","Tuesday","19-May","671","WANGLANG","BKK-19","Bangkok","Night"],
  ["May 2026","W4","Monday","18-May","3298","BaanFah Le Marche","BKK-08","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","865","TESCO BANFAH","BKK-08","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","593","THE PASEO PARK","BKK-19","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","557","TESCO ROJANA","C-01","Central","Morning"],
  ["May 2026","W4","Monday","18-May","3224","Asiatique","BKK-16","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","670","BIG C BANGYAI","BKK-10","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","662","BIG C SUKSAWAT","BKK-17","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","638","THE BRIGHT RAMA 2","BKK-17","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","3294","Lotus Lamnarai","C-02","Central","Morning"],
  ["May 2026","W4","Tuesday","19-May","433","TESCO PAKKRED","BKK-13","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","360","N-MARK","BKK-03","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","596","BIG C ROMKLAO","BKK-04","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","3232","Thai Thanee Navanakorn","BKK-09","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","3152","The Forrest Rangsit","BKK-09","Bangkok","Morning"],
  ["May 2026","W4","Monday","18-May","259","LOTUS LADPRAO","BKK-01","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","648","UNION MALL BFL","BKK-01","Bangkok","Morning"],
  ["May 2026","W4","Tuesday","19-May","619","TESCO SAMUTSONGKRAM","C-03","Central","Morning"],
];
const JUN_DATA=[
["Jun 2026","W1","Tuesday","26-May","3139","Lotus Bangpakong","E-05","East","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","384","CENTRAL CHANGWATTANA","BKK-13","Bangkok","Night"],
  ["Jun 2026","W1","Thursday","28-May","3053","IMPERIAL WORLD SAMRONG GF","BKK-06","Bangkok","Night"],
  ["Jun 2026","W1","Tuesday","26-May","364","PATTAYA KLANG","E-03","East","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","3220","Lotus Pattaya 53","E-03","East","Morning"],
  ["Jun 2026","W1","Thursday","28-May","817","LITTLE WALK PATTAYA","E-03","East","Morning"],
  ["Jun 2026","W1","Friday","29-May","292","NANA SQUARE","BKK-15","Bangkok","Morning"],
  ["Jun 2026","W1","Tuesday","26-May","3302","Nongbualamphu","NE-01","Northeast","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","3266","THAI SIRI 3 MARKET","NE-01","Northeast","Morning"],
  ["Jun 2026","W1","Friday","29-May","3202","Ramindra 40","BKK-02","Bangkok","Morning"],
  ["Jun 2026","W1","Thursday","28-May","807","BIG C ONNUT","BKK-05","Bangkok","Morning"],
  ["Jun 2026","W1","Friday","29-May","700","MARKET WALK BANGYAI","BKK-10","Bangkok","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","448","INDRA PRATUNAM","BKK-20","Bangkok","Night"],
  ["Jun 2026","W1","Thursday","28-May","570","CENTRAL SALAYA","C-04","Central","Night"],
  ["Jun 2026","W1","Tuesday","26-May","3166","Movenpick Chiangmai","N-01","North","Morning"],
  ["Jun 2026","W1","Tuesday","26-May","3179","Lotus Simahaphot","E-01","East","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","3272","Pong Pai Market","E-01","East","Morning"],
  ["Jun 2026","W1","Thursday","28-May","456","HUAY KWANG MARKET","BKK-12","Bangkok","Morning"],
  ["Jun 2026","W1","Friday","29-May","3242","PTT Borom 97","BKK-18","Bangkok","Morning"],
  ["Jun 2026","W1","Tuesday","26-May","3261","Thong Pha Phum","C-03","Central","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","562","BIG C KANCHANABURI","C-03","Central","Morning"],
  ["Jun 2026","W1","Thursday","28-May","318","KANOKPETCH KANCHANABURI","C-03","Central","Morning"],
  ["Jun 2026","W1","Friday","29-May","3025","TESCO LADKRABANG","BKK-07","Bangkok","Morning"],
  ["Jun 2026","W1","Tuesday","26-May","3273","Seka Buengkan","NE-02","Northeast","Morning"],
  ["Jun 2026","W1","Wednesday","27-May","3269","Wanon Niwat","NE-02","Northeast","Morning"],
  ["Jun 2026","W1","Friday","29-May","769","TAWANNA BANGKAPI","BKK-03","Bangkok","Morning"],
  ["Jun 2026","W1","Friday","29-May","863","THAMMASAT HOSPITAL","BKK-09","Bangkok","Night"],
  ["Jun 2026","W1","Tuesday","26-May","3161","Makro Sanambinnam","BKK-11","Bangkok","Morning"],
  ["Jun 2026","W1","Thursday","28-May","484","J MARKET","BKK-01","Bangkok","Morning"],
  ["Jun 2026","W1","Friday","29-May","699","THE PLATFORM WONG WIEN YAI","BKK-19","Bangkok","Morning"],
  ["Jun 2026","W1","Saturday","30-May","3148","Chamchuri square","BKK-20","Bangkok","Morning"],
  ["Jun 2026","W2","Tuesday","2-Jun","362","UDOMSUK","BKK-05","Bangkok","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","3271","PTT U Park Mahasarakham","NE-04","Northeast","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","297","SERMTHAI MAHASARAKHAM","NE-04","Northeast","Morning"],
  ["Jun 2026","W2","Tuesday","2-Jun","3150","Central Ramindra","BKK-02","Bangkok","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","3265","NONG BUA DAENG","NE-08","Northeast","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","870","ROBINSON CHAIYAPHUM","NE-08","Northeast","Night"],
  ["Jun 2026","W2","Thursday","4-Jun","641","THE MALL KORAT B FLOOR","NE-07","Northeast","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","3155","U Ville","NE-07","Northeast","Morning"],
  ["Jun 2026","W2","Tuesday","2-Jun","465","ISARAPHAP ROAD","BKK-19","Bangkok","Morning"],
  ["Jun 2026","W2","Wednesday","3-Jun","693","NAMPU PLAZA","C-05","Central","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","287","TAD CHAN HUA HIN","S-07","South","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","678","BLU PORT HUA HIN","S-07","South","Morning"],
  ["Jun 2026","W2","Sunday","7-Jun","3281","Government Complex Chaengwattana","BKK-13","Bangkok","Morning"],
  ["Jun 2026","W2","Tuesday","2-Jun","864","THE CIRCLE","BKK-18","Bangkok","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","3231","La-Ngu","S-03","South","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","668","CENTRAL NAKORNSRITHAMMARAT","S-03","South","Night"],
  ["Jun 2026","W2","Tuesday","2-Jun","3293","Big C Bangpakok","BKK-17","Bangkok","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","3054","THE MALL THAPRA","BKK-17","Bangkok","Night"],
  ["Jun 2026","W2","Friday","5-Jun","377","MAJOR RATCHAYOTIN","BKK-01","Bangkok","Night"],
  ["Jun 2026","W2","Tuesday","2-Jun","3086","Lotus Langsuan","S-01","South","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","3169","The Mall Bangkapi GF","BKK-03","Bangkok","Night"],
  ["Jun 2026","W2","Friday","5-Jun","3120","TERMINAL 21 RAMA 3","BKK-16","Bangkok","Night"],
  ["Jun 2026","W2","Tuesday","2-Jun","785","BIG C RAMA 4","BKK-15","Bangkok","Morning"],
  ["Jun 2026","W2","Tuesday","2-Jun","3244","Rongkluea Market Navanakhon","C-01","Central","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","400","SEACON SQUARE 1","BKK-05","Bangkok","Night"],
  ["Jun 2026","W2","Friday","5-Jun","342","CENTURY THE MOVIE PLAZA","BKK-12","Bangkok","Night"],
  ["Jun 2026","W2","Tuesday","2-Jun","553","PASEO SUKHAPIBAN 3","BKK-03","Bangkok","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","889","BIG C SAIMAI","BKK-02","Bangkok","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","3203","DD Marche Market","BKK-08","Bangkok","Morning"],
  ["Jun 2026","W2","Tuesday","2-Jun","3051","MIXT CHATUCHAK","BKK-21","Bangkok","Morning"],
  ["Jun 2026","W2","Thursday","4-Jun","3039","TESCO LAMPANG","N-03","North","Morning"],
  ["Jun 2026","W2","Friday","5-Jun","498","CENTRAL LAMPANG","N-03","North","Night"],
  ["Jun 2026","W3","Monday","8-Jun","3055","ONNUT PLAZA","BKK-05","Bangkok","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","3287","Central Village","BKK-07","Bangkok","Night"],
  ["Jun 2026","W3","Wednesday","10-Jun","788","CENTRAL WORLD 4FL","BKK-20","Bangkok","Night"],
  ["Jun 2026","W3","Thursday","11-Jun","435","CENTRAL LADPRAO","BKK-01","Bangkok","Night"],
  ["Jun 2026","W3","Friday","12-Jun","442","TERMINAL 21","BKK-15","Bangkok","Night"],
  ["Jun 2026","W3","Monday","8-Jun","984","SURATTHANI HOSPITAL","S-01","South","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","390","COLESIUM SURATTHANI","S-01","South","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","3123","Sahathai Garden","S-01","South","Morning"],
  ["Jun 2026","W3","Thursday","11-Jun","490","CENTRAL SURATTHANI","S-01","South","Night"],
  ["Jun 2026","W3","Friday","12-Jun","3199","Central Samui","S-01","South","Night"],
  ["Jun 2026","W3","Monday","8-Jun","489","PLAZA LAGOON","BKK-01","Bangkok","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","883","MARKET PLACE DUSIT","BKK-21","Bangkok","Morning"],
  ["Jun 2026","W3","Thursday","11-Jun","960","IMPERIAL WORLD SAMRONG 2F","BKK-06","Bangkok","Night"],
  ["Jun 2026","W3","Monday","8-Jun","3321","Betong Yala","S-06","South","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","3087","Diana Pattni","S-06","South","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","3241","Big C Pattani","S-06","South","Morning"],
  ["Jun 2026","W3","Thursday","11-Jun","3072","BIG C NARATHIWAT","S-06","South","Morning"],
  ["Jun 2026","W3","Saturday","13-Jun","784","ZUELLIG HOUSE","BKK-16","Bangkok","Morning"],
  ["Jun 2026","W3","Monday","8-Jun","3183","Big C Saraburi","C-01","Central","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","543","ROBINSON SARABURI","C-01","Central","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","687","ROBINSON LOPBURI","C-02","Central","Night"],
  ["Jun 2026","W3","Thursday","11-Jun","485","SEACON BANGKAE","BKK-19","Bangkok","Night"],
  ["Jun 2026","W3","Friday","12-Jun","3119","Robinson Ratchapruek","BKK-11","Bangkok","Night"],
  ["Jun 2026","W3","Monday","8-Jun","594","BIG C SRISAKET","NE-06","Northeast","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","330","SURIN PLAZA","NE-06","Northeast","Night"],
  ["Jun 2026","W3","Wednesday","10-Jun","555","ROBINSON SURIN","NE-06","Northeast","Night"],
  ["Jun 2026","W3","Friday","12-Jun","692","BIG C MAHACHAI","C-05","Central","Morning"],
  ["Jun 2026","W3","Sunday","14-Jun","3255","Muangthai Phatra Complex","BKK-12","Bangkok","Morning"],
  ["Jun 2026","W3","Monday","8-Jun","3320","Chann 24 Mapyangphon Rayong","E-02","East","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","609","CENTRAL RAYONG","E-02","East","Night"],
  ["Jun 2026","W3","Wednesday","10-Jun","224","ROBINSON CHANTABURI","E-02","East","Night"],
  ["Jun 2026","W3","Thursday","11-Jun","3112","Central Chantaburi","E-02","East","Night"],
  ["Jun 2026","W3","Monday","8-Jun","3267","DREAMLAND MARKET","N-05","North","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","3327","K Avenue Nakornsawan","N-05","North","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","3205","Central Nakhonsawan","N-05","North","Night"],
  ["Jun 2026","W3","Monday","8-Jun","3296","Loeng Nok Tha","NE-05","Northeast","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","3297","Big C Amnatcharoeng","NE-05","Northeast","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","520","CENTRAL UBONRATCHATHANI","NE-05","Northeast","Night"],
  ["Jun 2026","W3","Friday","12-Jun","695","TESCO SUKAPHIBAN 3","BKK-04","Bangkok","Morning"],
  ["Jun 2026","W3","Monday","8-Jun","3304","Phu Wiang Khonkaen","NE-03","Northeast","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","3147","Jompol Market","NE-03","Northeast","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","3326","Bangchak KKU","NE-03","Northeast","Morning"],
  ["Jun 2026","W3","Thursday","11-Jun","389","CENTRAL PLAZA KHONKAEN","NE-03","Northeast","Night"],
  ["Jun 2026","W3","Friday","12-Jun","780","CENTRAL KHONKAEN 2","NE-03","Northeast","Night"],
  ["Jun 2026","W3","Monday","8-Jun","3133","Saphan 4 Market","E-04","East","Morning"],
  ["Jun 2026","W3","Tuesday","9-Jun","3081","Big C BORWIN","E-04","East","Morning"],
  ["Jun 2026","W3","Wednesday","10-Jun","289","SAVELAND BANGSEAN","E-04","East","Morning"],
  ["Jun 2026","W3","Thursday","11-Jun","797","ROBINSON CHONBURI","E-05","East","Night"],
  ["Jun 2026","W4","Monday","15-Jun","412","BIG C KHAMPHAENGPETCH","N-04","North","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","765","ROBINSON KAMPHAENGPHET","N-04","North","Night"],
  ["Jun 2026","W4","Wednesday","17-Jun","503","TESCO PHITSANULOK 2","N-04","North","Night"],
  ["Jun 2026","W4","Friday","19-Jun","403","PURE PLACE SAMMAKORN","BKK-03","Bangkok","Morning"],
  ["Jun 2026","W4","Monday","15-Jun","3132","The Seasons Mall","BKK-21","Bangkok","Morning"],
  ["Jun 2026","W4","Monday","15-Jun","3027","MARKET VILLAGE RANGSIT KLONG 4","BKK-08","Bangkok","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","212","CENTRAL RAMA III","B-16","Bangkok","Night"],
  ["Jun 2026","W4","Tuesday","16-Jun","603","FUTURE PARK RANGSIT","BKK-09","Bangkok","Night"],
  ["Jun 2026","W4","Wednesday","17-Jun","618","CENTRAL WEST GATE","BKK-10","Bangkok","Night"],
  ["Jun 2026","W4","Thursday","18-Jun","3209","The Mall Bangkae GF","BKK-19","Bangkok","Night"],
  ["Jun 2026","W4","Thursday","18-Jun","608","MARKET VILLAGE SUVARNABHUMI","BKK-07","Bangkok","Night"],
  ["Jun 2026","W4","Friday","19-Jun","349","ZEER RANGSIT 1","B-09","Bangkok","Night"],
  ["Jun 2026","W4","Friday","19-Jun","3076","Robinson Srisaman","BKK-13","Bangkok","Night"],
  ["Jun 2026","W4","Monday","15-Jun","694","THE CHILLED PATTAYA","E-03","East","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","3262","ROYAL GARDEN PATTAYA","E-03","East","Morning"],
  ["Jun 2026","W4","Wednesday","17-Jun","3005","CENTRAL PATTAYA BEACH","E-03","East","Night"],
  ["Jun 2026","W4","Thursday","18-Jun","599","SIAM SQUARE 2","BKK-20","Bangkok","Night"],
  ["Jun 2026","W4","Friday","19-Jun","3305","EmQuartier","BKK-15","Bangkok","Night"],
  ["Jun 2026","W4","Saturday","20-Jun","554","CP TOWER","C-02","Central","Night"],
  ["Jun 2026","W4","Monday","15-Jun","3328","Chombueng Market","C-03","Central","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","365","RATCHABURI","C-03","Central","Morning"],
  ["Jun 2026","W4","Wednesday","17-Jun","636","FASHION ISLAND 4","BKK-04","Bangkok","Night"],
  ["Jun 2026","W4","Monday","15-Jun","3092","Lotus Borom Sai 2","BKK-18","Bangkok","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","278","LOTUS RAMA I CHAROENPOL","BKK-20","Bangkok","Night"],
  ["Jun 2026","W4","Wednesday","17-Jun","819","ICON SIAM","BKK-17","Bangkok","Night"],
  ["Jun 2026","W4","Thursday","18-Jun","579","ROBINSON SAMUTPRAKARN","BKK-06","Bangkok","Night"],
  ["Jun 2026","W4","Friday","19-Jun","762","CENTRAL MAHACHAI","C-05","Central","Night"],
  ["Jun 2026","W4","Monday","15-Jun","3071","Big C Krabi","S-05","South","Morning"],
  ["Jun 2026","W4","Monday","15-Jun","876","BIG C CHIANGRAI 2","N-02","North","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","3312","Central Krabi","S-05","South","Night"],
  ["Jun 2026","W4","Tuesday","16-Jun","422","CENTRAL CHIANGRAI","N-02","North","Night"],
  ["Jun 2026","W4","Thursday","18-Jun","646","CENTRAL PINKLAO","BKK-18","Bangkok","Night"],
  ["Jun 2026","W4","Monday","15-Jun","351","GRAND PLAZA NGAMWONGWAN","BKK-11","Bangkok","Morning"],
  ["Jun 2026","W4","Monday","15-Jun","3204","PTT Bangyai","BKK-10","Bangkok","Morning"],
  ["Jun 2026","W4","Thursday","18-Jun","3009","The mall Ngamwongwan","BKK-11","Bangkok","Night"],
  ["Jun 2026","W4","Monday","15-Jun","3300","Bangli Suphanburi","C-02","Central","Morning"],
  ["Jun 2026","W4","Tuesday","16-Jun","394","TESCO SUPHANBURI","C-02","Central","Morning"],
  ["Jun 2026","W4","Wednesday","17-Jun","988","ROBINSON SUPHANBURI 1F","C-02","Central","Night"],
  ["Jun 2026","W5","Monday","22-Jun","803","TESCO PHANATNIKHOM","E-05","East","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","3200","Mingle Hill Minburi","BKK-04","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","881","TESCO BANSUAN","E-05","East","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","589","TESCO KLONGLUANG","BKK-09","Bangkok","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","3176","Ladsawai Market","BKK-08","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","466","TIME SQUARE","BKK-15","Bangkok","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","257","BANG RAK","BKK-16","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","569","TESCO SAMUTSAKHON","C-05","Central","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","3245","Banna Nakornnayok","E-01","East","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","565","TESCO NAKONNAYOK","E-01","East","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","3011","SALAYA MARKET","C-04","Central","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","568","TESCO BANGYAI","BKK-10","Bangkok","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","738","NUMBER ONE PLAZA","BKK-07","Bangkok","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","735","TESCO PATHUMTHANI","BKK-13","Bangkok","Morning"],
  ["Jun 2026","W5","Monday","22-Jun","3249","Big C Kanlapapruek","BKK-17","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","3246","Dernlen Market","BKK-17","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","528","PURE RATCHAPRUEK","BKK-11","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","3130","Infinite Mall","BKK-06","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","871","MAX VALU LAKSI","BKK-02","Bangkok","Morning"],
  ["Jun 2026","W5","Tuesday","23-Jun","620","TESCO BANGKAE","BKK-19","Bangkok","Morning"],
];
const JUL_DATA=[
["Jul 2026","W1","Thursday","2-Jul","624","TESCO BANGPOO","BKK-06","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","300","RAMKHAMHAENG 53","BKK-03","Bangkok","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","3143","Jampha 1 Lamphun","N-03","North","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","664","BIG C LAMPHUN","N-03","North","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","3091","Rungroj Market","BKK-08","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","774","TESCO RANGSIT","BKK-09","Bangkok","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","3189","Phan Chiangrai","N-02","North","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","3160","Mae Fah Luang University","N-02","North","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","808","TESCO BANGPAKOK","BKK-17","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","896","TESCO BANGLEN","C-04","Central","Morning"],
  ["Jul 2026","W1","Saturday","4-Jul","665","MID TOWN PLAZA","BKK-15","Bangkok","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","489","PLAZA LAGOON","BKK-01","Bangkok","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","375","BIG C AYUDHAYA","C-01","Central","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","519","THE TREE PATHUMTHANI","BKK-13","Bangkok","Morning"],
  ["Jul 2026","W1","Thursday","2-Jul","3146","INT INTERSECT","BKK-16","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","572","PINKLAO","BKK-18","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","3135","Phahonyothin 54","BKK-02","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","3065","Lotus Bangna-Trad","BKK-07","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","501","TESCO BANG PA-IN","C-01","Central","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","3144","I'm Park","BKK-20","Bangkok","Morning"],
  ["Jul 2026","W1","Friday","3-Jul","3192","Thiandat Market","C-05","Central","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","393","TESCO KANCHANABURI","C-03","Central","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","805","TESCO RATCHABURI","C-03","Central","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","538","ROBINSON RATCHABURI","C-03","Central","Night"],
  ["Jul 2026","W2","Thursday","9-Jul","3003","BIG C KHEHAROMKLAO","BKK-04","Bangkok","Night"],
  ["Jul 2026","W2","Monday","6-Jul","530","BIG C KALASIN","NE-02","Northeast","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","3026","Tesco Kalasin","NE-02","Northeast","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","3290","Somdet Kalasin","NE-02","Northeast","Morning"],
  ["Jul 2026","W2","Saturday","11-Jul","975","TESCO RAMINDRA","BKK-01","Bangkok","Morning"],
  ["Jul 2026","W2","Sunday","12-Jul","243","PHAHOLYOTHIN PLACE","BKK-21","Bangkok","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","492","BIG C RANGSIT","BKK-09","Bangkok","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","460","TESCO PRACHINBURI","E-01","East","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","585","ROBINSON DEPARTMENT PRACHINBURI","E-01","East","Night"],
  ["Jul 2026","W2","Thursday","9-Jul","345","IT SQUARE","BKK-13","Bangkok","Night"],
  ["Jul 2026","W2","Monday","6-Jul","3247","Chumphae","NE-03","Northeast","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","617","TESCO KHONKAEN","NE-03","Northeast","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","3128","Sentosa Maliwan","NE-03","Northeast","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","3096","Sentosa Srichan Khonkaen","NE-03","Northeast","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","790","BIG C SATTAHIP","E-03","East","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","866","TESCO U TAPAO","E-03","East","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","961","TESCO CHONBURI","E-05","East","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","982","BIG C CHONBURI 1","E-05","East","Morning"],
  ["Jul 2026","W2","Friday","10-Jul","3307","Burapha Market","E-05","East","Morning"],
  ["Jul 2026","W2","Saturday","11-Jul","3303","Susco Square Phutthabucha","BKK-17","Bangkok","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","3083","Big C Nongbua","N-05","North","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","3080","Big C Thatako","N-05","North","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","558","THE WALK NAKHONSAWAN","N-05","North","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","575","PLERNNARY MALL","BKK-02","Bangkok","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","3036","TESCO KANTHARALAK","NE-06","Northeast","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","380","LOTUS SRISAKET","NE-06","Northeast","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","3250","Big C Surin","NE-06","Northeast","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","758","TESCO PRASAT SURIN","NE-06","Northeast","Morning"],
  ["Jul 2026","W2","Friday","10-Jul","3275","Ban Kruat Buriram","NE-06","Northeast","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","536","TESCO KLANG RAYONG","E-02","East","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","3206","Mangmee Market","E-02","East","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","645","TESCO RAYONG","E-02","East","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","649","V MARKET","B-07","Bangkok","Morning"],
  ["Jul 2026","W2","Friday","10-Jul","992","BIG C BANGNA","BKK-05","Bangkok","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","545","TESCO CHAINAT","C-02","Central","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","3074","C Mall Chainat","C-02","Central","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","3252","Makro Chainat","C-02","Central","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","524","NANA CHAROEN MARKET","BKK-08","Bangkok","Morning"],
  ["Jul 2026","W2","Friday","10-Jul","535","THE WALK RATCHAPRUEK","BKK-10","Bangkok","Morning"],
  ["Jul 2026","W2","Saturday","11-Jul","3032","METRO MALL PHETCHABURI","BKK-15","Bangkok","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","267","ASAWANN","NE-01","Northeast","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","463","ASAWANN 1","NE-01","Northeast","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","493","UD TOWN","NE-01","Northeast","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","3257","Wang Sam Mo","NE-01","Northeast","Morning"],
  ["Jul 2026","W2","Monday","6-Jul","3336","Maepimthong Nonsung","NE-07","Northeast","Morning"],
  ["Jul 2026","W2","Tuesday","7-Jul","679","TESCO PHIMAI","NE-07","Northeast","Morning"],
  ["Jul 2026","W2","Wednesday","8-Jul","3330","Save One Korat","NE-07","Northeast","Morning"],
  ["Jul 2026","W2","Thursday","9-Jul","755","CENTRAL NAKRONRATCHASRIMA GF","NE-07","Northeast","Morning"],
  ["Jul 2026","W2","Friday","10-Jul","3317","Siri Plaza Market","NE-07","Northeast","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","522","MAX VALUE SUKHUMVIT 71","BKK-05","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","544","PASEO LADKRABANG","BKK-07","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","3075","Maruay Market Hathairat","BKK-02","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","3022","BIG C MAHACHAI 2","C-05","Central","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","327","BANG HUNSRI","BKK-18","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","261","BIG C TIWANON","BKK-21","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","3282","PTT Kluaynamthai","BKK-15","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","885","Lotus Angthong","C-01","Central","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","479","BIG C LUMLUKKA","BKK-08","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","794","VIBHAVADEE SOI 2","BKK-12","Bangkok","Morning"],
  ["Jul 2026","W3","Monday","13-Jul","974","BIG C SAINOI","BKK-10","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","3264","Gemmo Market","BKK-05","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","3063","MARKET PLACE KRUNGTHEP KREETHA","BKK-03","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","3258","Makro Rangsit","BKK-09","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","254","TESCO RAMA III","BKK-16","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","373","SOUTHERN BUS TERMINAL","BKK-18","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","539","SAPANKWAI","BKK-21","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","241","LOTUS RAMA IV","BKK-15","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","606","TESCO NONGJOK","BKK-04","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","604","ESPLANADE RATCHADA","BKK-12","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","556","CRYTAL CHAIYAPRUEK","BKK-11","Bangkok","Morning"],
  ["Jul 2026","W3","Tuesday","14-Jul","3170","Foody Farm","BKK-10","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","791","BIG C RAMINDRA","BKK-02","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","507","BIG C SUKHAPIBAN 3","BKK-04","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","3145","NAWAMIN CITY AVENUE","BKK-01","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","346","CHAN ROAD","BKK-16","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","3125","Seri Market Phuttamonthon 5","C-04","Central","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","3306","Market Place Prachauthit","BKK-17","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","673","TESCO CHANGWATTANA","BKK-13","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","3021","TESCO SAMUTPRAKARN","BKK-06","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","459","BIG C KLONG 6","BKK-08","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","792","THE ONE PARK","BKK-10","Bangkok","Morning"],
  ["Jul 2026","W3","Wednesday","15-Jul","540","THE TREE BANGBON","BKK-19","Bangkok","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","506","BIG C SUANLUANG","BKK-05","Bangkok","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","623","FUTURE MART RAMA3","BKK-16","Bangkok","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","977","Space@304","E-01","East","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","3149","Pattaya Avenue","E-03","East","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","3256","Kad Farang Village Maerim","N-06","North","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","3163","Big C Itsaraphap","BKK-19","Bangkok","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","610","BIG C SUWINTHAWONG","BKK-04","Bangkok","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","429","TU DOME","BKK-09","Bangkok","Morning"],
  ["Jul 2026","W4","Monday","20-Jul","811","TESCO SINGBURI","C-02","Central","Morning"],
  ["Jul 2026","W4","Tuesday","21-Jul","3130","Infinite Mall","BKK-06","Bangkok","Morning"],
  ["Jul 2026","W4","Tuesday","21-Jul","3223","The Prom Dindaeng","BKK-12","Bangkok","Morning"],
  ["Jul 2026","W4","Tuesday","21-Jul","564","BIG C SRIMAHAPOH","E-01","East","Morning"],
  ["Jul 2026","W4","Tuesday","21-Jul","467","BIG C PATTAYA KLANG","E-03","East","Morning"],
  ["Jul 2026","W4","Tuesday","21-Jul","274","LOTUS KHAMTIENG CHIENGMAI","N-06","North","Morning"],
  ["Jul 2026","W4","Tuesday","21-Jul","332","CHAI SAENG SINGHABURI","C-02","Central","Morning"],
  ["Jul 2026","W4","Wednesday","22-Jul","495","@CURVE CHIANGMAI","N-01","North","Morning"],
  ["Jul 2026","W4","Wednesday","22-Jul","505","TESCO HANGDONG","N-01","North","Morning"],
];
const SAMPLE_SCHEDULE=[...MAY_DATA,...JUN_DATA,...JUL_DATA];


function buildInitialBranches(rawArr){
  return rawArr.map(r=>{
    const [no,name,province,region,dataDate,stockQty,stockBalance,salesQty]=r;
    // Preserve nulls as genuinely missing (don't coerce to 0) so completeness tracking is accurate
    const sq = (stockQty===null||stockQty===undefined||stockQty==='') ? null : +stockQty;
    const sb = (stockBalance===null||stockBalance===undefined||stockBalance==='') ? null : +stockBalance;
    const saq = (salesQty===null||salesQty===undefined||salesQty==='') ? null : +salesQty;
    return {no:String(no), name:name||'', province:province||'', region:region||'', dataDate:dataDate||'',
            stockQty:sq, stockBalance:sb, salesQty:saq};
  });
}

// ── APP STATE ─────────────────────────────────────────────────────────────────
let clients=[
  { id:1, name:"Client A", label:"Client A", color:COLORS[0],
    branches: buildInitialBranches(CLIENT_A_BRANCHES),
    schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
  { id:2, name:"Client B", label:"Client B", color:COLORS[1],
    branches: buildInitialBranches(CLIENT_B_BRANCHES),
    schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
  { id:3, name:"Client C", label:"Client C", color:COLORS[2],
    branches: buildInitialBranches(CLIENT_C_BRANCHES),
    schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
  { id:4, name:"Client D", label:"Client D", color:COLORS[3],
    branches: buildInitialBranches(CLIENT_D_BRANCHES),
    schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
];
let activeClientId=1;
let editingClientId=null, selectedColor=COLORS[0];
let editingBranchNo=null;
let mainView='home';

const getClient=()=>clients.find(c=>c.id===activeClientId);

// ── PERSISTENCE (localStorage) ─────────────────────────────────────────────────
const STORAGE_KEY='inv_cyclecount_dashboard_v3';

function saveState(){
  try{
    const state=clients.map(c=>({
      id:c.id, name:c.name, label:c.label, color:c.color,
      branches:c.branches, schedule:c.schedule,
      selected:[...c.selected.entries()],
      counted:[...c.counted.entries()],
      threshold: (c.threshold===undefined?null:c.threshold),
      haircutPct: (c.haircutPct===undefined?100:c.haircutPct),
      thresholdSelected: [...(c.thresholdSelected||new Map()).entries()],
      excludedBranches: [...(c.excludedBranches||new Map()).entries()],
      monthOrder: c.monthOrder||[],
      selectedMonths: [...(c.selectedMonths||new Set())],
      manuallyMatched: [...(c.manuallyMatched||new Map()).entries()],
      handoffLog: c.handoffLog||[],
    }));
    localStorage.setItem(STORAGE_KEY, JSON.stringify({clients:state,activeClientId,savedAt:new Date().toISOString()}));
  }catch(e){console.warn('Save failed',e);}
}

function loadState(){
  try{
    const raw=localStorage.getItem(STORAGE_KEY);
    if(!raw) return false;
    const data=JSON.parse(raw);
    if(!data.clients||!data.clients.length) return false;
    clients=data.clients.map(c=>({
      id:c.id, name:c.name, label:c.label, color:c.color,
      branches:c.branches||[], schedule:c.schedule||[],
      selected:new Map(c.selected||[]),
      counted:new Map(c.counted||[]),
      threshold: (c.threshold===undefined?null:c.threshold),
      haircutPct: (c.haircutPct===undefined?100:c.haircutPct),
      thresholdSelected: new Map(c.thresholdSelected||[]),
      excludedBranches: new Map(c.excludedBranches||[]),
      monthOrder: c.monthOrder||[],
      selectedMonths: new Set(c.selectedMonths||[]),
      manuallyMatched: new Map(c.manuallyMatched||[]),
      handoffLog: c.handoffLog||[],
    }));
    activeClientId=data.activeClientId||clients[0].id;
    return true;
  }catch(e){console.warn('Load failed',e); return false;}
}

function clearSavedState(){ localStorage.removeItem(STORAGE_KEY); }

// ── CLIENT TABS ───────────────────────────────────────────────────────────────
function renderClientTabs(){
  document.getElementById('client-tabs').innerHTML=clients.map(c=>
    `<button class="ctab ${c.id===activeClientId?'active':''}" onclick="switchClient(${c.id})">
      <span class="dot" style="background:${c.color}"></span>${c.label}
      <span style="font-size:10px;opacity:.5;margin-left:2px">(${c.branches.length})</span>
    </button>`).join('');
}
function switchClient(id){
  activeClientId=id;
  renderClientTabs();
  refreshAllViews();
}
window.switchClient=switchClient;

function showAddClient(){
  editingClientId=null;
  selectedColor=COLORS[clients.length%COLORS.length];
  const nextLetter=String.fromCharCode(65+clients.length);
  document.getElementById('client-modal-title').textContent='Add new client';
  document.getElementById('cl-name').value='Client '+nextLetter;
  document.getElementById('cl-label').value='Client '+nextLetter;
  document.getElementById('cl-color-row').innerHTML=buildColorRow(selectedColor);
  document.getElementById('client-modal-bg').style.display='flex';
  setTimeout(()=>{const n=document.getElementById('cl-name');n.focus();n.select();},50);
}
function showEditClient(){
  const c=getClient();
  editingClientId=c.id; selectedColor=c.color;
  document.getElementById('client-modal-title').textContent='Edit client';
  document.getElementById('cl-name').value=c.name;
  document.getElementById('cl-label').value=c.label;
  document.getElementById('cl-color-row').innerHTML=buildColorRow(selectedColor);
  document.getElementById('client-modal-bg').style.display='flex';
}
function buildColorRow(sel){
  return COLORS.map(c=>`<span class="color-swatch${c===sel?' selected':''}" style="background:${c}" onclick="pickClientColor('${c}')"></span>`).join('');
}
function pickClientColor(c){selectedColor=c;document.getElementById('cl-color-row').innerHTML=buildColorRow(c);}
window.pickClientColor=pickClientColor;
function closeClientModal(){document.getElementById('client-modal-bg').style.display='none';}
function saveClientModal(){
  const name=document.getElementById('cl-name').value.trim();
  const label=document.getElementById('cl-label').value.trim();
  if(!name||!label){alert('Please fill in both fields.');return;}
  if(editingClientId){
    const c=clients.find(cl=>cl.id===editingClientId);
    c.name=name;c.label=label;c.color=selectedColor;
  }else{
    const id=Date.now();
    clients.push({id,name,label,color:selectedColor,branches:[],schedule:[],selected:new Map(),counted:new Map(),threshold:null,haircutPct:100,thresholdSelected:new Map(),excludedBranches:new Map(),monthOrder:[],selectedMonths:new Set(),manuallyMatched:new Map(),handoffLog:[]});
    activeClientId=id;
  }
  closeClientModal();
  renderClientTabs();
  refreshAllViews();
  saveState();
}
function deleteClient(){
  if(clients.length===1){alert('At least one client must remain.');return;}
  if(!confirm(`Remove "${getClient().name}"? All its data will be lost.`))return;
  clients=clients.filter(c=>c.id!==activeClientId);
  activeClientId=clients[0].id;
  renderClientTabs();
  refreshAllViews();
  saveState();
}

// ── MAIN VIEW SWITCHING ────────────────────────────────────────────────────────
function switchMainView(view){
  mainView=view;
  ['home','dashboard','branches','schedule'].forEach(v=>{
    document.getElementById('view-'+v).classList.toggle('active', v===view);
    document.getElementById('mt-'+v).classList.toggle('active', v===view);
  });
  refreshAllViews();
}

function refreshAllViews(){
  renderHome();
  renderDashboard();
  renderBranchMaster();
  renderSchedule();
  updateTopBadges();
}

function updateTopBadges(){
  const c=getClient();
  document.getElementById('badge-branches').textContent=c.branches.length+' branches';
  const total=c.branches.length;
  const complete=c.branches.filter(b=>isBranchComplete(b)).length;
  const pct=total? Math.round(complete/total*100):0;
  document.getElementById('badge-complete').textContent=pct+'% complete';
}

function isBranchComplete(b){
  return !!(b.no && b.name && b.province && b.region && b.dataDate &&
    b.stockQty!==null && b.stockQty!==undefined &&
    b.stockBalance!==null && b.stockBalance!==undefined &&
    b.salesQty!==null && b.salesQty!==undefined);
}

// ── DUPLICATE SCHEDULE RESOLUTION ─────────────────────────────────────────────
// When a branch appears on more than one scheduled date (e.g. 19-Feb and 20-Feb),
// only ONE of those dates counts as "the" visit, per this rule:
//   - If exactly one of the dates already has a recorded count result \u2192 use that date.
//   - If BOTH dates have a result \u2192 use the LATEST date (and its result).
//   - If NEITHER date has a result yet \u2192 use the LATEST date (still pending).
// Returns the single schedule row that should represent this branch.
function resolveScheduleRowForBranch(c, branchNo){
  const rowsForBranch=c.schedule.filter(r=>r.branchNo===branchNo);
  if(!rowsForBranch.length) return null;
  if(rowsForBranch.length===1) return rowsForBranch[0];

  const sorted=[...rowsForBranch].sort((a,b)=>compareDateLabels(a.date,b.date));
  const cr=c.counted.get(branchNo);
  const hasResult = r => !!cr && cr.date===r.date;

  const withResult=sorted.filter(hasResult);
  if(withResult.length===1) return withResult[0];
  // Both have a result, or neither does \u2014 either way, latest date wins.
  return sorted[sorted.length-1];
}

// Returns a Set of branchNos that are "duplicated" (scheduled on 2+ dates) along with
// which specific schedule row is the resolved/effective one for each.
// A branch counts as "Matched" for Dashboard/Schedule purposes if it's EITHER
// threshold-Selected OR has been manually flagged as Manually Match in Branch Master Data.
function isBranchMatched(c, branchNo){
  return c.thresholdSelected.has(branchNo) || c.manuallyMatched.has(branchNo);
}
function totalMatchedCount(c){
  const all=new Set([...c.thresholdSelected.keys(), ...c.manuallyMatched.keys()]);
  return all.size;
}

function getEffectiveScheduleMap(c){
  const byBranch=new Map();
  c.schedule.forEach(r=>{
    if(!byBranch.has(r.branchNo)) byBranch.set(r.branchNo, []);
    byBranch.get(r.branchNo).push(r);
  });
  const effective=new Map(); // branchNo -> resolved row
  byBranch.forEach((rows, no)=>{
    effective.set(no, resolveScheduleRowForBranch(c, no));
  });
  return effective;
}


// ── DASHBOARD VIEW ──────────────────────────────────────────────────────────────
let dashSortCol=-1, dashSortAsc=true;

function getDashboardRows(){
  const c=getClient();
  return c.branches.map(b=>{
    const cr=c.counted.get(b.no);
    return {
      no:b.no, name:b.name, province:b.province, region:b.region,
      stockQty:b.stockQty, stockBalance:b.stockBalance,
      countDate: cr?cr.date:null,
      diff: cr?cr.diff:null,
      level: cr?cr.level:null,
    };
  });
}

function getFilteredDashboardRows(){
  const rows=getDashboardRows();
  const fRisk=document.getElementById('dash-f-risk').value;
  const fRegion=document.getElementById('dash-f-region').value;
  const srch=document.getElementById('dash-f-search').value.toLowerCase().trim();
  return rows.filter(r=>{
    if(fRisk==='none'&&r.level) return false;
    if(fRisk&&fRisk!=='none'&&r.level!==fRisk) return false;
    if(fRegion&&r.region!==fRegion) return false;
    if(srch&&!r.no.toLowerCase().includes(srch)&&!r.name.toLowerCase().includes(srch)&&!(r.province||'').toLowerCase().includes(srch)) return false;
    return true;
  });
}

function updateDashRegionFilter(){
  const c=getClient();
  const regions=[...new Set(c.branches.map(b=>b.region).filter(Boolean))].sort();
  const sel=document.getElementById('dash-f-region');
  const cur=sel.value;
  sel.innerHTML='<option value="">All regions</option>'+regions.map(r=>`<option${r===cur?' selected':''}>${r}</option>`).join('');
}

function dashSortBy(col){
  if(dashSortCol===col) dashSortAsc=!dashSortAsc; else{dashSortCol=col;dashSortAsc=true;}
  document.querySelectorAll('[id^="dth"]').forEach((th,i)=>{
    th.classList.remove('sort-asc','sort-desc');
    if(i===col) th.classList.add(dashSortAsc?'sort-asc':'sort-desc');
  });
  renderDashboard();
}
window.dashSortBy=dashSortBy;

// ── HOME (all-client overview) ─────────────────────────────────────────────
function renderHome(){
  if(mainView!=='home') return;

  const pct=(n,d)=> d? Math.round(n/d*100) : 0;

  let grandTotalBranches=0, grandTotalQty=0, grandTotalCost=0;
  let grandCountedBranches=0, grandCountedQty=0, grandCountedCost=0;
  let grandPendingBranches=0;
  let grandDiffSum=0, grandDiffCount=0;
  const perClientRows=[];

  clients.forEach(c=>{
    const branches=c.branches||[];
    const counted=c.counted||new Map();
    const thresholdSelected=c.thresholdSelected||new Map();
    const manuallyMatched=c.manuallyMatched||new Map();

    const totalBranches=branches.length;
    const totalQty=branches.reduce((s,b)=>s+(typeof b.stockQty==='number'?b.stockQty:0),0);
    const totalCost=branches.reduce((s,b)=>s+(typeof b.stockBalance==='number'?b.stockBalance:0),0);

    const countedBranchList=branches.filter(b=>counted.has(b.no));
    const countedBranches=countedBranchList.length;
    const countedQty=countedBranchList.reduce((s,b)=>s+(typeof b.stockQty==='number'?b.stockQty:0),0);
    const countedCost=countedBranchList.reduce((s,b)=>s+(typeof b.stockBalance==='number'?b.stockBalance:0),0);

    const matchedNos=new Set([...thresholdSelected.keys(), ...manuallyMatched.keys()]);
    const pendingBranches=[...matchedNos].filter(no=>!counted.has(no)).length;

    const diffs=[...counted.values()].filter(cr=>cr.diff!==null&&cr.diff!==undefined&&!isNaN(cr.diff)).map(cr=>cr.diff);
    const avgDiff=diffs.length? diffs.reduce((a,b)=>a+b,0)/diffs.length : 0;
    const highRisk=[...counted.values()].filter(cr=>cr.level==='High').length;

    grandTotalBranches+=totalBranches;
    grandTotalQty+=totalQty;
    grandTotalCost+=totalCost;
    grandCountedBranches+=countedBranches;
    grandCountedQty+=countedQty;
    grandCountedCost+=countedCost;
    grandPendingBranches+=pendingBranches;
    grandDiffSum+=diffs.reduce((a,b)=>a+b,0);
    grandDiffCount+=diffs.length;

    perClientRows.push({
      id:c.id, name:c.name, color:c.color,
      totalBranches, totalQty, totalCost,
      countedBranches, countedQty, countedCost,
      avgDiff, diffCount:diffs.length, highRisk,
    });
  });

  const grandAvgDiff=grandDiffCount? grandDiffSum/grandDiffCount : 0;

  document.getElementById('home-total-clients').textContent=clients.length;
  document.getElementById('home-total-branches').textContent=grandTotalBranches.toLocaleString();
  document.getElementById('home-total-qty').textContent=grandTotalQty.toLocaleString();
  document.getElementById('home-total-cost').textContent='\u0e3f'+grandTotalCost.toLocaleString(undefined,{maximumFractionDigits:0});

  document.getElementById('home-counted-branches').textContent=grandCountedBranches.toLocaleString();
  document.getElementById('home-counted-pct').textContent='('+pct(grandCountedBranches,grandTotalBranches)+'%)';
  document.getElementById('home-counted-qty').textContent=grandCountedQty.toLocaleString();
  document.getElementById('home-counted-qty-pct').textContent='('+pct(grandCountedQty,grandTotalQty)+'%)';
  document.getElementById('home-counted-cost').textContent='\u0e3f'+grandCountedCost.toLocaleString(undefined,{maximumFractionDigits:0});
  document.getElementById('home-counted-cost-pct').textContent='('+pct(grandCountedCost,grandTotalCost)+'%)';
  document.getElementById('home-pending-branches').textContent=grandPendingBranches.toLocaleString();
  document.getElementById('home-avg-diff').textContent=grandAvgDiff.toFixed(2)+'%';

  const tbody=document.getElementById('home-client-tbl-body');
  if(!perClientRows.length){
    tbody.innerHTML=`<tr><td colspan="9"><div class="empty-state"><p>No clients yet.</p></div></td></tr>`;
    return;
  }
  tbody.innerHTML=perClientRows.map(r=>{
    const isActive=r.id===activeClientId;
    const avgDiffCell=r.diffCount? `${r.avgDiff.toFixed(2)}%` : '<span style="color:#ccc">\u2014</span>';
    return `<tr style="${isActive?'background:#f0f6fb':''};cursor:pointer" onclick="switchClient(${r.id});switchMainView('dashboard')" title="Click to open this client's Dashboard">
      <td style="font-weight:700;color:${r.color}"><span class="dot" style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${r.color};margin-right:6px"></span>${r.name}</td>
      <td class="num-cell">${r.totalBranches.toLocaleString()}</td>
      <td class="num-cell">${r.totalQty.toLocaleString()}</td>
      <td class="num-cell">\u0e3f${r.totalCost.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
      <td class="num-cell">${r.countedBranches.toLocaleString()} <span style="color:#888;font-size:10px">(${pct(r.countedBranches,r.totalBranches)}%)</span></td>
      <td class="num-cell">${r.countedQty.toLocaleString()}</td>
      <td class="num-cell">\u0e3f${r.countedCost.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
      <td class="num-cell">${avgDiffCell}</td>
      <td class="num-cell">${r.highRisk>0?`<span class="risk-pill risk-high">${r.highRisk}</span>`:'<span style="color:#ccc">0</span>'}</td>
    </tr>`;
  }).join('');
}

function renderDashboard(){
  if(mainView!=='dashboard') { renderDashboardStats(); return; }
  updateDashRegionFilter();
  renderDashboardStats();

  let rows=getFilteredDashboardRows();
  if(dashSortCol>=0){
    const keys=['no','name','province','region','countDate','diff'];
    rows.sort((a,b)=>{
      let va,vb;
      if(dashSortCol===6){ // risk level - sort by severity
        const order={null:-1,Low:0,Medium:1,High:2};
        va=order[a.level]; vb=order[b.level];
      } else {
        va=a[keys[dashSortCol]]; vb=b[keys[dashSortCol]];
        va=(va===null||va===undefined)?'':va;
        vb=(vb===null||vb===undefined)?'':vb;
        if(typeof va==='string') va=va.toLowerCase();
        if(typeof vb==='string') vb=vb.toLowerCase();
      }
      if(va<vb) return dashSortAsc?-1:1;
      if(va>vb) return dashSortAsc?1:-1;
      return 0;
    });
  }

  document.getElementById('dash-row-count').textContent=rows.length+' of '+getDashboardRows().length+' branches';
  const tbody=document.getElementById('dash-tbl-body');
  if(!rows.length){
    tbody.innerHTML=`<tr><td colspan="7"><div class="empty-state"><p>No branches match current filters.</p></div></td></tr>`;
    return;
  }
  tbody.innerHTML=rows.map(r=>{
    const lvl=r.level;
    const pillCls=riskPillClass(lvl);
    const pillText=lvl||'Not counted';
    const diffText=r.diff!==null?r.diff.toFixed(2)+'%':'—';
    const dateText=r.countDate||'—';
    return `<tr>
      <td class="td-num">${r.no}</td>
      <td class="td-name" title="${r.name}">${r.name}</td>
      <td>${r.province||'—'}</td>
      <td>${r.region||'—'}</td>
      <td style="font-size:11px">${dateText}</td>
      <td class="num-cell" style="font-weight:700">${diffText}</td>
      <td><span class="risk-pill ${pillCls}">${pillText}</span></td>
    </tr>`;
  }).join('');
}

function renderDashboardStats(){
  const c=getClient();
  const rows=getDashboardRows();
  const total=rows.length;
  const pct=(n,d)=> d? Math.round(n/d*100) : 0;

  // Card 1: Total branches + Selected (threshold-Selected OR Manually Match) with % of total
  const totalSelected=totalMatchedCount(c);
  document.getElementById('dash-total-branches').textContent=total;
  document.getElementById('dash-total-selected').textContent=totalSelected;
  document.getElementById('dash-selected-pct').textContent='('+pct(totalSelected,total)+'%)';

  // Card 2: Counted / Pending \u2014 scoped to threshold-Selected branches only, % of total branches
  const selectedRows=rows.filter(r=>isBranchMatched(c,r.no));
  const selCounted=selectedRows.filter(r=>r.level).length;
  const selPending=selectedRows.length-selCounted;
  document.getElementById('dash-counted').textContent=selCounted;
  document.getElementById('dash-counted-pct').textContent='('+pct(selCounted,total)+'%)';
  document.getElementById('dash-pending').textContent=selPending;
  document.getElementById('dash-pending-pct').textContent='('+pct(selPending,total)+'%)';

  // Card 3: Coverage of result counted \u2014 Qty and Cost (THB), measured against ALL branches' totals,
  // not just the threshold-Selected ones. "Counted" here means any branch with a recorded result.
  const countedRows=rows.filter(r=>r.level);
  const sumQty=(arr)=>arr.reduce((s,r)=>s+(typeof r.stockQty==='number'?r.stockQty:0),0);
  const sumBal=(arr)=>arr.reduce((s,r)=>s+(typeof r.stockBalance==='number'?r.stockBalance:0),0);
  const totalQty=sumQty(rows), countedQty=sumQty(countedRows);
  const totalBal=sumBal(rows), countedBal=sumBal(countedRows);
  const qtyPct=pct(countedQty,totalQty), costPct=pct(countedBal,totalBal);

  document.getElementById('dash-qty-coverage').textContent=qtyPct+'%';
  document.getElementById('dash-cost-coverage').textContent=costPct+'%';
  document.getElementById('dash-coverage-detail').innerHTML=
    `${countedQty.toLocaleString()} / ${totalQty.toLocaleString()} qty &middot; `+
    `\u0e3f${countedBal.toLocaleString(undefined,{maximumFractionDigits:0})} / \u0e3f${totalBal.toLocaleString(undefined,{maximumFractionDigits:0})}`;

  // Row 2: risk breakdown, % of total Counted (any branch with a recorded result)
  const totalCounted=countedRows.length;
  const lowRisk=rows.filter(r=>r.level==='Low').length;
  const mediumRisk=rows.filter(r=>r.level==='Medium').length;
  const highRisk=rows.filter(r=>r.level==='High').length;

  document.getElementById('risk-low-count').textContent=lowRisk;
  document.getElementById('risk-low-pct').textContent='('+pct(lowRisk,totalCounted)+'%)';
  document.getElementById('risk-medium-count').textContent=mediumRisk;
  document.getElementById('risk-medium-pct').textContent='('+pct(mediumRisk,totalCounted)+'%)';
  document.getElementById('risk-high-count').textContent=highRisk;
  document.getElementById('risk-high-pct').textContent='('+pct(highRisk,totalCounted)+'%)';
}

function resetDashFilters(){
  document.getElementById('dash-f-risk').value='';
  document.getElementById('dash-f-region').value='';
  document.getElementById('dash-f-search').value='';
  dashSortCol=-1; dashSortAsc=true;
  document.querySelectorAll('[id^="dth"]').forEach(th=>th.classList.remove('sort-asc','sort-desc'));
  renderDashboard();
}

function exportDashboard(){
  const c=getClient();
  const rows=getDashboardRows();
  if(!rows.length){alert('No branches to export.');return;}
  const H=['Branch No.','Branch Name','Province','Region','Count Date','Diff %','Risk Level'];
  const data=[H,...rows.map(r=>[r.no,r.name,r.province||'',r.region||'',r.countDate||'',r.diff!==null?r.diff:'',r.level||'Not counted'])];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws['!cols']=[{wch:12},{wch:36},{wch:16},{wch:14},{wch:12},{wch:10},{wch:14}];
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Dashboard');
  XLSX.writeFile(wb,`${c.label.replace(/[^a-z0-9]/gi,'_')}_Dashboard_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ── COUNT ENTRY MODAL ─────────────────────────────────────────────────────────
function updateCmDiffPreview(){
  const diffEl=document.getElementById('cm-diff');
  const bookEl=document.getElementById('cm-bookqty');
  const actualEl=document.getElementById('cm-actualqty');
  const book=parseFloat(bookEl.value), actual=parseFloat(actualEl.value);
  if(!isNaN(book)&&!isNaN(actual)&&book!==0){
    const autoDiff=Math.abs((actual-book)/book*100);
    diffEl.value=autoDiff.toFixed(2);
  }
  const v=diffEl.value;
  const preview=document.getElementById('cm-diff-preview');
  if(v===''){preview.style.display='none';return;}
  const lvl=calcRiskLevel(v);
  const cls=riskPillClass(lvl);
  preview.style.display='block';
  preview.className='risk-pill '+cls;
  preview.style.display='block';
  preview.style.padding='8px 12px';
  preview.textContent=parseFloat(v).toFixed(2)+'% difference — '+lvl+' Risk';
}

function openCountEntryFor(no){
  const c=getClient();
  const b=c.branches.find(b=>b.no===no);
  const existing=c.counted.get(no);
  document.getElementById('cm-branchno').value=no;
  document.getElementById('cm-branchname').value=b?b.name:(existing?existing.branchName:'');
  document.getElementById('cm-date').value=existing?existing.date:new Date().toISOString().slice(0,10);
  document.getElementById('cm-bookqty').value=existing?(existing.bookQty||''):'';
  document.getElementById('cm-actualqty').value=existing?(existing.actualQty||''):'';
  document.getElementById('cm-diff').value=existing?existing.diff:'';
  updateCmDiffPreview();
  document.getElementById('count-modal-bg').style.display='flex';
}
window.openCountEntryFor=openCountEntryFor;

function openCountEntryBlank(){
  document.getElementById('cm-branchno').value='';
  document.getElementById('cm-branchname').value='';
  document.getElementById('cm-date').value=new Date().toISOString().slice(0,10);
  document.getElementById('cm-bookqty').value='';
  document.getElementById('cm-actualqty').value='';
  document.getElementById('cm-diff').value='';
  document.getElementById('cm-diff-preview').style.display='none';
  document.getElementById('count-modal-bg').style.display='flex';
  setTimeout(()=>document.getElementById('cm-branchno').focus(),50);
}

function saveCountEntry(){
  try{
    const no=document.getElementById('cm-branchno').value.trim();
    const name=document.getElementById('cm-branchname').value.trim();
    const date=document.getElementById('cm-date').value;
    const bookQty=document.getElementById('cm-bookqty').value;
    const actualQty=document.getElementById('cm-actualqty').value;
    const diff=document.getElementById('cm-diff').value;
    if(!no){alert('Enter a branch number.');return;}
    if(!date){alert('Enter a count date.');return;}
    if(diff===''){alert('Enter a difference %.');return;}
    const pct=parseFloat(diff);
    if(isNaN(pct)||pct<0){alert(`"${diff}" is not a valid difference %. Please enter a number (e.g. 0.75).`);return;}
    const c=getClient();
    const lvl=calcRiskLevel(pct);
    const b=c.branches.find(b=>b.no===no);
    c.counted.set(no,{
      branchNo:no, branchName:name||(b?b.name:no), date,
      bookQty:bookQty!==''?parseFloat(bookQty):null,
      actualQty:actualQty!==''?parseFloat(actualQty):null,
      diff:pct, level:lvl,
    });
    closeCountModal();
    refreshAllViews();
    saveState();
  }catch(err){
    console.error('saveCountEntry failed', err);
    alert(`Could not save the count result.\n\nError: ${err.message}`);
  }
}

function closeCountModal(){ document.getElementById('count-modal-bg').style.display='none'; }


// ── BRANCH MASTER DATA VIEW ────────────────────────────────────────────────────
let bmSortCol=-1, bmSortAsc=true;

function getFilteredBranches(){
  const c=getClient();
  const fRegion=document.getElementById('bm-f-region').value;
  const fComplete=document.getElementById('bm-f-complete').value;
  const fStatus=document.getElementById('bm-f-status').value;
  const srch=document.getElementById('bm-f-search').value.toLowerCase().trim();
  return c.branches.filter(b=>{
    if(fRegion&&b.region!==fRegion) return false;
    if(fComplete==='complete'&&!isBranchComplete(b)) return false;
    if(fComplete==='incomplete'&&isBranchComplete(b)) return false;
    if(fStatus==='selected'&&!c.thresholdSelected.has(b.no)) return false;
    if(fStatus==='manual'&&!c.manuallyMatched.has(b.no)) return false;
    if(fStatus==='excluded'&&!c.excludedBranches.has(b.no)) return false;
    if(fStatus==='watchlisted'&&!c.selected.has(b.no)) return false;
    if(fStatus==='none'&&(c.thresholdSelected.has(b.no)||c.excludedBranches.has(b.no)||c.manuallyMatched.has(b.no))) return false;
    if(srch&&!b.no.toLowerCase().includes(srch)&&!b.name.toLowerCase().includes(srch)&&!(b.province||'').toLowerCase().includes(srch)) return false;
    return true;
  });
}

function updateBmRegionFilter(){
  const c=getClient();
  const regions=[...new Set(c.branches.map(b=>b.region).filter(Boolean))].sort();
  const sel=document.getElementById('bm-f-region');
  const cur=sel.value;
  sel.innerHTML='<option value="">All regions</option>'+regions.map(r=>`<option${r===cur?' selected':''}>${r}</option>`).join('');
}

function bmSortBy(col){
  if(bmSortCol===col) bmSortAsc=!bmSortAsc; else{bmSortCol=col;bmSortAsc=true;}
  document.querySelectorAll('[id^="bth"]').forEach((th,i)=>{
    th.classList.remove('sort-asc','sort-desc');
    if(i===col) th.classList.add(bmSortAsc?'sort-asc':'sort-desc');
  });
  renderBranchMaster();
}
window.bmSortBy=bmSortBy;

function renderBranchMaster(){
  if(mainView!=='branches') { renderBmStats(); return; }
  updateBmRegionFilter();

  let rows=getFilteredBranches();
  renderBmStats(rows);
  if(bmSortCol>=0){
    const keys=['no','name','province','region','dataDate','stockQty','stockBalance','salesQty','_selected'];
    const c=getClient();
    rows=[...rows].sort((a,b)=>{
      let va,vb;
      if(bmSortCol===8){ va=c.thresholdSelected.has(a.no)?1:0; vb=c.thresholdSelected.has(b.no)?1:0; }
      else{ va=a[keys[bmSortCol]]; vb=b[keys[bmSortCol]];
        va=(va===null||va===undefined)?'':va; vb=(vb===null||vb===undefined)?'':vb;
        if(typeof va==='string') va=va.toLowerCase();
        if(typeof vb==='string') vb=vb.toLowerCase();
      }
      if(va<vb) return bmSortAsc?-1:1;
      if(va>vb) return bmSortAsc?1:-1;
      return 0;
    });
  }

  document.getElementById('bm-row-count').textContent=rows.length+' of '+getClient().branches.length+' branches';
  const tbody=document.getElementById('bm-tbl-body');
  if(!rows.length){
    tbody.innerHTML=`<tr><td colspan="10"><div class="empty-state"><p>No branches match current filters.<br>Click <strong>+ Add branch</strong> to get started.</p></div></td></tr>`;
    return;
  }
  const c2=getClient();
  tbody.innerHTML=rows.map(b=>{
    const complete=isBranchComplete(b);
    const excluded=c2.excludedBranches.get(b.no);
    const isSelected=c2.thresholdSelected.has(b.no);
    const isManualMatch=c2.manuallyMatched.has(b.no);
    const isWatchlisted=c2.selected.has(b.no);
    const watchPin=isWatchlisted?`<span title="${(c2.selected.get(b.no)||{}).note||'Watchlisted'}">&#128204; </span>`:'';
    const rowStyle=excluded?'background:#fdf3f0':complete?'':'background:#fffbf0';
    const fmt=(v,dec=0)=>v===null||v===undefined||v===''?'<span style="color:#ccc">—</span>':Number(v).toLocaleString(undefined,{minimumFractionDigits:dec,maximumFractionDigits:dec});

    let statusCell, actionCell;
    if(excluded){
      statusCell=`<span class="risk-pill" style="background:#fde8e8;color:#a32020" title="${excluded.note}">&#10005; Excluded</span>`;
      actionCell=`<button class="btn btn-sm btn-outline" onclick="reincludeBranch('${b.no}')" style="padding:3px 8px;font-size:10px" title="Remove exclusion">&#8635; Re-include</button>`;
    } else if(isManualMatch){
      statusCell=`<span class="risk-pill" style="background:#e8e3fb;color:#4a2db5" title="Manually flagged as matched, independent of the threshold calculation">&#10003; Manually Match</span>`;
      actionCell=`<button class="btn btn-sm btn-outline" onclick="toggleManualMatch('${b.no}')" style="padding:3px 8px;font-size:10px" title="Remove manual match">&#10005; Unmatch</button>`;
    } else if(isSelected){
      statusCell=`<span class="risk-pill" style="background:#e3f6e8;color:#1a6b2e">&#10003; Selected</span>`;
      actionCell=`<button class="btn btn-sm btn-outline" onclick="openExcludeBranch('${b.no}')" style="padding:3px 8px;font-size:10px" title="Exclude from counting">&#10005; Exclude</button>`;
    } else {
      statusCell=`<span style="color:#ccc;font-size:11px">—</span>`;
      actionCell=`<button class="btn btn-sm btn-outline" onclick="toggleManualMatch('${b.no}')" style="padding:3px 8px;font-size:10px" title="Manually flag this branch as matched">&#10003; Manual Match</button>`;
    }
    const watchTag=isWatchlisted?` <span class="risk-pill" style="background:#fff3cd;color:#8a6500;font-size:9px;padding:1px 5px" title="${(c2.selected.get(b.no)||{}).note||'Watchlisted'}">&#128204; Watchlist</span>`:'';
    statusCell=statusCell+watchTag;

    return `<tr style="${rowStyle}">
      <td class="td-num">${b.no}</td>
      <td class="td-name" title="${b.name}">${watchPin}${b.name||'<span style="color:#ccc">—</span>'}</td>
      <td>${b.province||'<span style="color:#ccc">—</span>'}</td>
      <td>${b.region||'<span style="color:#ccc">—</span>'}</td>
      <td style="font-size:11px">${b.dataDate||'<span style="color:#ccc">—</span>'}</td>
      <td class="num-cell">${fmt(b.stockQty)}</td>
      <td class="num-cell">${fmt(b.stockBalance,2)}</td>
      <td class="num-cell">${fmt(b.salesQty)}</td>
      <td style="text-align:center">${statusCell}</td>
      <td style="text-align:center;display:flex;gap:3px;justify-content:center">
        <button class="btn btn-sm btn-outline" onclick="openEditBranch('${b.no}')" style="padding:3px 7px">&#9998;</button>
        ${actionCell}
      </td>
    </tr>`;
  }).join('');
}

function renderExcludedList(){
  const c=getClient();
  const badge=document.getElementById('ex-count-badge');
  const listEl=document.getElementById('ex-list');
  if(!badge||!listEl) return;
  badge.textContent=c.excludedBranches.size;
  if(!c.excludedBranches.size){
    listEl.innerHTML='<p class="hint" style="text-align:center;color:#bbb">No exclusions yet.</p>';
    return;
  }
  listEl.innerHTML=[...c.excludedBranches.entries()].map(([no,info])=>{
    const b=c.branches.find(b=>b.no===no);
    return `<div style="display:flex;align-items:flex-start;gap:6px;padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:11px">
      <div style="flex:1">
        <div style="font-weight:700;color:#1a3a5c">${no}${b?' \u2014 '+b.name:''}</div>
        <div style="color:#888;margin-top:2px">${info.note}</div>
        <div style="color:#bbb;font-size:10px;margin-top:1px">Excluded ${info.date}</div>
      </div>
      <button onclick="reincludeBranch('${no}')" style="background:none;border:none;cursor:pointer;color:#2980b9;font-size:10px;font-weight:700;white-space:nowrap" title="Remove exclusion">&#8635; Re-include</button>
    </div>`;
  }).join('');
}

function renderBmStats(filteredRows){
  const c=getClient();
  // Summary cards (Total / Complete / Incomplete / Avg balance / Total balance) reflect the CURRENT FILTER.
  // Falls back to the full client branch list when called without filtered rows (e.g. on other tabs).
  const rowsForCards = filteredRows || c.branches;
  const isFiltered = !!filteredRows && filteredRows.length !== c.branches.length;

  const total=rowsForCards.length;
  const complete=rowsForCards.filter(isBranchComplete).length;
  const incomplete=total-complete;
  const completePct=total?Math.round(complete/total*100):0;
  const incompletePct=total?Math.round(incomplete/total*100):0;

  document.getElementById('bm-total').textContent=total;
  document.getElementById('bm-complete').textContent=complete;
  document.getElementById('bm-complete-pct').textContent='('+completePct+'%)';
  document.getElementById('bm-incomplete').textContent=incomplete;
  document.getElementById('bm-incomplete-pct').textContent='('+incompletePct+'%)';

  renderExcludedList();

  const fill=document.getElementById('bm-progress');
  fill.style.width=completePct+'%';
  fill.className='progress-fill'+(completePct<50?' danger':completePct<85?' warn':'');

  // Stock balance metrics within the current filter (only counting branches with a known balance)
  const balances=rowsForCards.map(b=>b.stockBalance).filter(v=>v!==null&&v!==undefined&&!isNaN(v));
  const totalBalance=balances.length? balances.reduce((a,b)=>a+b,0) : 0;
  const avgBalance=balances.length? totalBalance/balances.length : 0;
  document.getElementById('bm-avg-balance').textContent='\u0e3f'+avgBalance.toLocaleString(undefined,{maximumFractionDigits:0});
  document.getElementById('bm-total-balance').textContent='\u0e3f'+totalBalance.toLocaleString(undefined,{maximumFractionDigits:0});

  // Stock quantity metrics within the current filter (only counting branches with a known quantity)
  const qtys=rowsForCards.map(b=>b.stockQty).filter(v=>v!==null&&v!==undefined&&!isNaN(v));
  const totalQty=qtys.length? qtys.reduce((a,b)=>a+b,0) : 0;
  const avgQty=qtys.length? totalQty/qtys.length : 0;
  document.getElementById('bm-avg-qty').textContent=avgQty.toLocaleString(undefined,{maximumFractionDigits:0});
  document.getElementById('bm-total-qty').textContent=totalQty.toLocaleString(undefined,{maximumFractionDigits:0});

  // Visible "(filtered)" tag on every card label when a filter is active, so it's unmistakable
  // that these numbers are scoped to the current view, not the full client.
  const filterTag = isFiltered ? ' <span style="color:#2980b9;font-weight:700">(filtered)</span>' : '';
  const totalLbl=document.getElementById('bm-total-lbl');
  if(totalLbl) totalLbl.innerHTML = 'Total branches' + filterTag;
  const avgLbl=document.getElementById('bm-avg-balance-lbl');
  if(avgLbl) avgLbl.innerHTML = 'Avg. / branch' + filterTag;
  const totalBalLbl=document.getElementById('bm-total-balance-lbl');
  if(totalBalLbl) totalBalLbl.innerHTML = 'Total balance' + filterTag;
  const avgQtyLbl=document.getElementById('bm-avg-qty-lbl');
  if(avgQtyLbl) avgQtyLbl.innerHTML = 'Avg. qty / branch' + filterTag;
  const totalQtyLbl=document.getElementById('bm-total-qty-lbl');
  if(totalQtyLbl) totalQtyLbl.innerHTML = 'Total qty' + filterTag;

  // Reflect this client's saved haircut % value in the input
  const haircutInput=document.getElementById('bm-threshold-haircut');
  if(haircutInput && document.activeElement!==haircutInput){
    haircutInput.value = (c.haircutPct!==null && c.haircutPct!==undefined) ? c.haircutPct : 100;
  }
  updateThresholdResultPreview();

  // Missing fields summary panel
  const missingCounts={};
  c.branches.forEach(b=>{
    ['name','province','region','dataDate','stockQty','stockBalance','salesQty'].forEach(f=>{
      const v=b[f];
      if(v===null||v===undefined||v==='') missingCounts[f]=(missingCounts[f]||0)+1;
    });
  });
  const hintEl=document.getElementById('bm-incomplete-hint');
  const fieldsEl=document.getElementById('bm-missing-fields');
  const labels={name:'Branch name',province:'Province',region:'Region',dataDate:'Data date',stockQty:'Stock qty',stockBalance:'Stock balance',salesQty:'Sales qty'};
  const entries=Object.entries(missingCounts).filter(([k,v])=>v>0);
  if(!entries.length){
    hintEl.textContent='All required fields present across all branches.';
    fieldsEl.innerHTML='';
  } else {
    hintEl.textContent=incomplete+' branch(es) have missing data:';
    fieldsEl.innerHTML=entries.sort((a,b)=>b[1]-a[1]).map(([k,v])=>
      `<div style="display:flex;justify-content:space-between;font-size:11px;padding:3px 0;border-bottom:1px solid #f5f5f5">
        <span>${labels[k]||k}</span><span style="font-weight:700;color:#e67e22">${v} missing</span>
      </div>`
    ).join('');
  }
}

function resetBmFilters(){
  document.getElementById('bm-f-region').value='';
  document.getElementById('bm-f-complete').value='';
  document.getElementById('bm-f-status').value='';
  document.getElementById('bm-f-search').value='';
  renderBranchMaster();
}

// ── ADD / EDIT BRANCH MODAL ────────────────────────────────────────────────────
function openAddBranch(){
  editingBranchNo=null;
  document.getElementById('bmodal-title').textContent='Add branch';
  ['bm-no','bm-name','bm-province','bm-datadate','bm-stockqty','bm-stockbal','bm-salesqty'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('bm-region').value='';
  document.getElementById('btn-bm-delete').style.display='none';
  document.getElementById('bm-no').removeAttribute('readonly');
  document.getElementById('branch-modal-bg').style.display='flex';
  setTimeout(()=>document.getElementById('bm-no').focus(),50);
}
function openEditBranch(no){
  const c=getClient();
  const b=c.branches.find(b=>b.no===no);
  if(!b) return;
  editingBranchNo=no;
  document.getElementById('bmodal-title').textContent='Edit branch';
  document.getElementById('bm-no').value=b.no;
  document.getElementById('bm-name').value=b.name||'';
  document.getElementById('bm-province').value=b.province||'';
  document.getElementById('bm-region').value=b.region||'';
  document.getElementById('bm-datadate').value=b.dataDate||'';
  document.getElementById('bm-stockqty').value=b.stockQty!==null&&b.stockQty!==undefined?b.stockQty:'';
  document.getElementById('bm-stockbal').value=b.stockBalance!==null&&b.stockBalance!==undefined?b.stockBalance:'';
  document.getElementById('bm-salesqty').value=b.salesQty!==null&&b.salesQty!==undefined?b.salesQty:'';
  document.getElementById('btn-bm-delete').style.display='inline-flex';
  document.getElementById('bm-no').setAttribute('readonly','readonly');
  document.getElementById('branch-modal-bg').style.display='flex';
}
window.openEditBranch=openEditBranch;

function closeBranchModal(){
  document.getElementById('branch-modal-bg').style.display='none';
  editingBranchNo=null;
}

function saveBranchModal(){
  const no=document.getElementById('bm-no').value.trim();
  const name=document.getElementById('bm-name').value.trim();
  if(!no){alert('Branch number is required.');return;}
  if(!name){alert('Branch name is required.');return;}
  const province=document.getElementById('bm-province').value.trim();
  const region=document.getElementById('bm-region').value;
  const dataDate=document.getElementById('bm-datadate').value;
  const sq=document.getElementById('bm-stockqty').value;
  const sb=document.getElementById('bm-stockbal').value;
  const saq=document.getElementById('bm-salesqty').value;

  const branchObj={
    no,name,province,region,dataDate,
    stockQty:sq!==''?parseFloat(sq):null,
    stockBalance:sb!==''?parseFloat(sb):null,
    salesQty:saq!==''?parseFloat(saq):null,
  };

  const c=getClient();
  if(editingBranchNo){
    const idx=c.branches.findIndex(b=>b.no===editingBranchNo);
    if(idx>=0) c.branches[idx]=branchObj;
  } else {
    if(c.branches.find(b=>b.no===no)){alert(`Branch ${no} already exists.`);return;}
    c.branches.push(branchObj);
  }
  closeBranchModal();
  refreshAllViews();
  saveState();
}

function deleteBranchModal(){
  if(!editingBranchNo) return;
  if(!confirm(`Delete branch ${editingBranchNo}? This cannot be undone.`)) return;
  const c=getClient();
  c.branches=c.branches.filter(b=>b.no!==editingBranchNo);
  closeBranchModal();
  refreshAllViews();
  saveState();
}

// ── BULK UPDATE (key feature: update incomplete data over time) ──────────────
function applyBulkBranchUpdate(rows, sourceLabel){
  const c=getClient();
  const existingMap=new Map(c.branches.map(b=>[b.no,b]));
  let added=0, updated=0;
  const debugLog=[];

  rows.forEach((r,rowIdx)=>{
    const no=String(r[0]||'').trim(); if(!no) return;
    const name=r[1]!==undefined&&r[1]!==''?String(r[1]).trim():null;
    const province=r[2]!==undefined&&r[2]!==''?String(r[2]).trim():null;
    const region=r[3]!==undefined&&r[3]!==''?String(r[3]).trim():null;
    const dataDate=r[4]!==undefined&&r[4]!==''?String(r[4]).trim():null;
    const stockQty=r[5]!==undefined&&r[5]!==''?parseFloat(r[5]):null;
    const stockBalance=r[6]!==undefined&&r[6]!==''?parseFloat(r[6]):null;
    const salesQty=r[7]!==undefined&&r[7]!==''?parseFloat(r[7]):null;

    const found=existingMap.has(no);
    let beforeBalance=null, afterBalance=null;

    if(found){
      // Update only fields that were actually provided (non-empty) — preserve existing otherwise
      const b=existingMap.get(no);
      beforeBalance=b.stockBalance;
      if(name!==null) b.name=name;
      if(province!==null) b.province=province;
      if(region!==null) b.region=region;
      if(dataDate!==null) b.dataDate=dataDate;
      if(stockQty!==null && !isNaN(stockQty)) b.stockQty=stockQty;
      if(stockBalance!==null && !isNaN(stockBalance)) b.stockBalance=stockBalance;
      if(salesQty!==null && !isNaN(salesQty)) b.salesQty=salesQty;
      afterBalance=b.stockBalance;
      updated++;
    } else {
      c.branches.push({no,name:name||'',province:province||'',region:region||'',dataDate:dataDate||'',
        stockQty:isNaN(stockQty)?null:stockQty, stockBalance:isNaN(stockBalance)?null:stockBalance,
        salesQty:isNaN(salesQty)?null:salesQty});
      existingMap.set(no,c.branches[c.branches.length-1]);
      afterBalance=isNaN(stockBalance)?null:stockBalance;
      added++;
    }
    debugLog.push(`Row ${rowIdx+1} [${no}]: found=${found}, rawCell="${r[6]}", parsedBalance=${stockBalance}, before=${beforeBalance}, after=${afterBalance}`);
  });

  console.log('Bulk update debug log:\n'+debugLog.join('\n'));

  refreshAllViews();
  saveState();

  // Verify the in-memory branches array actually reflects the new values after render+save
  const verifyLog=rows.map(r=>{
    const no=String(r[0]||'').trim();
    const stockBalance=r[6]!==undefined&&r[6]!==''?parseFloat(r[6]):null;
    const liveBranch=getClient().branches.find(b=>b.no===no);
    const match = liveBranch && stockBalance!==null && Math.abs(liveBranch.stockBalance-stockBalance)<0.001;
    return `[${no}] expected=${stockBalance}, actualInMemory=${liveBranch?liveBranch.stockBalance:'NOT FOUND'}, match=${match}`;
  });
  console.log('Post-save verification:\n'+verifyLog.join('\n'));

  const mismatches=verifyLog.filter(l=>l.includes('match=false'));
  let msg=`Update complete from "${sourceLabel}":\n${added} branch(es) added, ${updated} branch(es) updated.`;
  if(mismatches.length){
    msg+=`\n\n\u26a0\ufe0f WARNING: ${mismatches.length} branch(es) may not have saved correctly. Open the browser console (F12) to see exactly which ones and why.`;
  }
  alert(msg);
}

function parseBulkBranchText(){
  const txt=document.getElementById('bm-bulk-text').value;
  const lines=txt.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
  if(!lines.length){alert('Paste some rows first.');return;}
  const rows=lines.map(l=>l.split(',').map(c=>c.trim()));
  applyBulkBranchUpdate(rows,'pasted text');
  document.getElementById('bm-bulk-text').value='';
}

function parseBulkBranchCSVFile(file){
  const r=new FileReader();
  r.onload=ev=>{
    const lines=ev.target.result.split(/\r?\n/).filter(l=>l.trim());
    // Skip header row if it looks like one
    const startIdx=/branch|store|no\.?$/i.test(lines[0].split(',')[0])?1:0;
    const rows=lines.slice(startIdx).map(l=>l.split(',').map(c=>c.trim().replace(/^"|"$/g,'')));
    applyBulkBranchUpdate(rows,file.name);
  };
  r.readAsText(file);
}

function importBranchFile(file){
  const ext=file.name.split('.').pop().toLowerCase();
  if(ext==='csv'||ext==='txt'){
    parseBulkBranchCSVFile(file);
  } else if(ext==='xlsx'||ext==='xls'){
    const r=new FileReader();
    r.onload=ev=>{
      const wb2=XLSX.read(ev.target.result,{type:'binary'});
      const ws=wb2.Sheets[wb2.SheetNames[0]];
      const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
      const startIdx=/branch|store|no\\.?$/i.test(String(data[0][0]))?1:0;
      applyBulkBranchUpdate(data.slice(startIdx), file.name);
    };
    r.readAsBinaryString(file);
  } else {
    alert('Please upload CSV or XLSX.');
  }
}


// ── STOCK BALANCE THRESHOLD (drives count-selection) ──────────────────────────
function getClientAvgBalance(c){
  const balances=c.branches.map(b=>b.stockBalance).filter(v=>v!==null&&v!==undefined&&!isNaN(v));
  return balances.length? balances.reduce((a,b)=>a+b,0)/balances.length : 0;
}

function getComputedThresholdAmount(c){
  const avg=getClientAvgBalance(c);
  const haircut=(c.haircutPct===null||c.haircutPct===undefined)?100:c.haircutPct;
  return avg*(haircut/100);
}

function updateThresholdResultPreview(){
  const c=getClient();
  const avgEl=document.getElementById('bm-threshold-avg-display');
  const computedEl=document.getElementById('bm-threshold-computed');
  const resultEl=document.getElementById('bm-threshold-result');
  if(!resultEl) return;

  const avg=getClientAvgBalance(c);
  const computed=getComputedThresholdAmount(c);
  if(avgEl) avgEl.value='\u0e3f'+avg.toLocaleString(undefined,{maximumFractionDigits:0});
  if(computedEl) computedEl.value='\u0e3f'+computed.toLocaleString(undefined,{maximumFractionDigits:0});

  const qualifying=c.branches.filter(b=>b.stockBalance!==null&&b.stockBalance!==undefined&&b.stockBalance>=computed&&!c.excludedBranches.has(b.no));
  const qualifyingNos=new Set(qualifying.map(b=>b.no));
  const alreadySelected=qualifying.filter(b=>c.thresholdSelected.has(b.no)).length;
  const haircut=(c.haircutPct===null||c.haircutPct===undefined)?100:c.haircutPct;

  // Staleness check: branches now qualifying-but-not-Selected, or Selected-but-no-longer-qualifying.
  // This happens whenever stock balance, haircut %, or branch data changes after the last Apply click.
  const newlyQualifying=[...qualifyingNos].filter(no=>!c.thresholdSelected.has(no)).length;
  const noLongerQualifying=[...c.thresholdSelected.keys()].filter(no=>!qualifyingNos.has(no)).length;
  const isStale = c.thresholdSelected.size>0 && (newlyQualifying>0 || noLongerQualifying>0);
  const staleBadge=document.getElementById('bm-threshold-stale-badge');
  if(staleBadge) staleBadge.style.display=isStale?'inline-block':'none';

  let html=`At <strong>${haircut}%</strong> of average, minimum balance = \u0e3f${computed.toLocaleString(undefined,{maximumFractionDigits:0})}.<br><strong>${qualifying.length}</strong> of ${c.branches.length} branches qualify. ${alreadySelected} already marked Selected.`;

  if(isStale){
    html+=`<div style="margin-top:6px;padding:6px 9px;background:#fff3cd;border:1px solid #f0c040;border-radius:6px;color:#8a6500;font-weight:700">
      &#9888; Selected is out of date with current data.<br>
      <span style="font-weight:400">${newlyQualifying} branch(es) now qualify but aren\u2019t Selected; ${noLongerQualifying} Selected branch(es) no longer qualify.<br>
      Click <strong>Apply &amp; select</strong> again to refresh.</span>
    </div>`;
  } else if(c.thresholdSelected.size>0){
    html+=`<div style="margin-top:4px;color:#1a6b2e;font-size:10px">&#10003; Selected matches current data.</div>`;
  }

  resultEl.innerHTML=html;
}

function applyThresholdSelection(){
  const c=getClient();
  const haircutVal=document.getElementById('bm-threshold-haircut').value;
  if(haircutVal===''){alert('Enter a haircut % first.');return;}
  const haircut=parseFloat(haircutVal);
  if(isNaN(haircut)||haircut<0){alert('Enter a valid non-negative percentage.');return;}
  c.haircutPct=haircut;

  const computed=getComputedThresholdAmount(c);
  c.threshold=computed; // keep in sync for any legacy reference

  // Threshold selection is tracked SEPARATELY from the manual "Branches to count" list.
  // It drives the "Selected" status in Branch Master Data and the Match status on Schedule & Count.
  // Branches manually excluded (e.g. WH/DC, online warehouse) are skipped even if they qualify.
  c.thresholdSelected=new Map();
  const qualifying=c.branches.filter(b=>b.stockBalance!==null&&b.stockBalance!==undefined&&b.stockBalance>=computed);
  let excludedSkipped=0;
  qualifying.forEach(b=>{
    if(c.excludedBranches.has(b.no)){ excludedSkipped++; return; }
    c.thresholdSelected.set(b.no, b.name||'');
  });

  updateThresholdResultPreview();
  refreshAllViews();
  saveState();
  let msg=`Threshold applied at ${haircut}% of average (\u0e3f${computed.toLocaleString(undefined,{maximumFractionDigits:0})}):\n${c.thresholdSelected.size} branch(es) marked as Selected for counting.`;
  if(excludedSkipped>0) msg+=`\n${excludedSkipped} qualifying branch(es) skipped due to manual exclusion.`;
  alert(msg);
}

function clearThresholdSelection(){
  try{
    const c=getClient();
    const sizeBefore=c.thresholdSelected.size;
    if(!sizeBefore){ alert('No threshold-Selected branches to clear.'); return; }
    if(!confirm(`Clear all ${sizeBefore} threshold-Selected branch(es) for this client?\n\n(This only clears the Selected status — it does not affect your manual Watchlist or Manually Match flags.)`)) return;
    c.thresholdSelected.clear();
    console.log('clearThresholdSelection: size after clear =', c.thresholdSelected.size, 'for client', c.id, c.name);
    updateThresholdResultPreview();
    refreshAllViews();
    saveState();
    // Verify the clear actually stuck after the full render+save cycle
    const sizeAfter=getClient().thresholdSelected.size;
    if(sizeAfter!==0){
      console.error('clearThresholdSelection FAILED \u2014 still', sizeAfter, 'entries after clear');
      alert(`Something went wrong \u2014 ${sizeAfter} branch(es) are still marked Selected. Please try refreshing the page and clearing again, or report this.`);
    }
  }catch(err){
    console.error('clearThresholdSelection threw an error:', err);
    alert(`Could not clear the selection.\n\nError: ${err.message}\n\nPlease check the browser console (F12) for details.`);
  }
}


// ── MANUAL EXCLUSION (e.g. WH/DC, online warehouse — branches that can't be counted) ──
let excludingBranchNo=null;

function openExcludeBranch(no){
  const c=getClient();
  const b=c.branches.find(b=>b.no===no);
  excludingBranchNo=no;
  document.getElementById('ex-branchno-display').textContent=no+(b?' \u2014 '+b.name:'');
  document.getElementById('ex-note').value='';
  document.getElementById('exclude-modal-bg').style.display='flex';
  setTimeout(()=>document.getElementById('ex-note').focus(),50);
}
window.openExcludeBranch=openExcludeBranch;

function closeExcludeModal(){
  document.getElementById('exclude-modal-bg').style.display='none';
  excludingBranchNo=null;
}

function saveExcludeBranch(){
  const note=document.getElementById('ex-note').value.trim();
  if(!note){alert('Please enter a reason (e.g. "WH/DC \u2014 not a sellable branch").');return;}
  const c=getClient();
  c.excludedBranches.set(excludingBranchNo, {note, date:new Date().toISOString().slice(0,10)});
  // Immediately remove from threshold-Selected and Manually Match if it was there \u2014
  // a branch can't be both excluded and counted as matched at the same time.
  c.thresholdSelected.delete(excludingBranchNo);
  c.manuallyMatched.delete(excludingBranchNo);
  closeExcludeModal();
  refreshAllViews();
  saveState();
}

function reincludeBranch(no){
  const c=getClient();
  if(!confirm(`Remove the exclusion on branch ${no}? It will become eligible for threshold selection again (re-apply the threshold to pick it up).`)) return;
  c.excludedBranches.delete(no);
  refreshAllViews();
  saveState();
}
window.reincludeBranch=reincludeBranch;

// ── MANUALLY MATCH (override the threshold for a single branch, independent of avg-balance calc) ──
function toggleManualMatch(no){
  const c=getClient();
  if(c.manuallyMatched.has(no)){
    c.manuallyMatched.delete(no);
  } else {
    const b=c.branches.find(b=>b.no===no);
    c.manuallyMatched.set(no, b?b.name:'');
  }
  refreshAllViews();
  saveState();
}
window.toggleManualMatch=toggleManualMatch;

// ── HANDOFF NOTE ───────────────────────────────────────────────────────────────
// Wraps the actual export functions: opens a small modal asking who's sending the
// file and what the next person should know, then stamps that onto a cover sheet
// in the exported workbook and logs it in c.handoffLog for later reference.
let pendingExportType=null; // 'branches' | 'schedule'

function exportBranchMaster(){ openHandoffModal('branches'); }
function exportScheduleFiltered(){ openHandoffModal('schedule'); }

function openHandoffModal(type){
  pendingExportType=type;
  document.getElementById('ho-from').value='';
  document.getElementById('ho-to').value='';
  document.getElementById('ho-note').value='';
  document.getElementById('handoff-modal-bg').style.display='flex';
  setTimeout(()=>document.getElementById('ho-from').focus(),50);
}
function closeHandoffModal(){
  document.getElementById('handoff-modal-bg').style.display='none';
  pendingExportType=null;
}
function confirmHandoffExport(){
  const from=document.getElementById('ho-from').value.trim();
  const to=document.getElementById('ho-to').value.trim();
  const note=document.getElementById('ho-note').value.trim();
  if(!from){alert('Please enter your name \u2014 the next person needs to know who sent this.');return;}

  const c=getClient();
  const handoff={
    from, to, note,
    date: new Date().toISOString().slice(0,10),
    time: new Date().toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}),
    type: pendingExportType,
  };

  if(pendingExportType==='branches') _doExportBranchMaster(handoff);
  else if(pendingExportType==='schedule') _doExportScheduleFiltered(handoff);

  if(!c.handoffLog) c.handoffLog=[];
  c.handoffLog.unshift(handoff); // newest first
  saveState();
  closeHandoffModal();
}

function addHandoffCoverSheet(wb, c, handoff, exportLabel){
  const ws=XLSX.utils.aoa_to_sheet([
    ['HANDOFF NOTE'],
    [''],
    ['Client', c.name],
    ['Export type', exportLabel],
    ['From', handoff.from],
    ['To', handoff.to||'(not specified)'],
    ['Date', handoff.date],
    ['Time', handoff.time],
    [''],
    ['Note for next person:'],
    [handoff.note||'(no note added)'],
  ]);
  ws['!cols']=[{wch:18},{wch:50}];
  XLSX.utils.book_append_sheet(wb, ws, 'Handoff Note');
}

function openHandoffHistory(){
  const c=getClient();
  document.getElementById('ho-history-client-name').textContent='\u2014 '+c.name;
  const listEl=document.getElementById('ho-history-list');
  const log=c.handoffLog||[];
  if(!log.length){
    listEl.innerHTML='<p class="hint" style="text-align:center;color:#bbb;padding:20px 0">No handoffs recorded yet for this client.</p>';
  } else {
    listEl.innerHTML=log.map(h=>{
      const typeLabel = h.type==='branches' ? 'Branch Master Data' : 'Schedule & Count';
      const typeColor = h.type==='branches' ? '#2980b9' : '#16a085';
      return `<div style="padding:10px 0;border-bottom:1px solid #f0f2f5">
        <div style="display:flex;justify-content:space-between;align-items:baseline">
          <span style="font-weight:700;color:${typeColor};font-size:11px">${typeLabel}</span>
          <span style="color:#bbb;font-size:10px">${h.date} ${h.time||''}</span>
        </div>
        <div style="font-size:12px;margin-top:3px"><strong>${h.from}</strong>${h.to?' &rarr; '+h.to:''}</div>
        ${h.note?`<div style="font-size:11px;color:#666;margin-top:3px;background:#f8fafd;padding:6px 8px;border-radius:5px">${h.note}</div>`:''}
      </div>`;
    }).join('');
  }
  document.getElementById('handoff-history-modal-bg').style.display='flex';
}

function _doExportBranchMaster(handoff){
  const c=getClient();
  if(!c.branches.length){alert('No branches to export.');return;}
  const H=['Branch No.','Branch Name','Province','Region','Data Date','Stock Qty','Stock Balance','Sales Qty'];
  const data=[H,...c.branches.map(b=>[b.no,b.name||'',b.province||'',b.region||'',b.dataDate||'',
    b.stockQty??'', b.stockBalance??'', b.salesQty??''])];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws['!cols']=[{wch:12},{wch:36},{wch:16},{wch:14},{wch:12},{wch:12},{wch:14},{wch:12}];
  const wb=XLSX.utils.book_new();
  if(handoff) addHandoffCoverSheet(wb, c, handoff, 'Branch Master Data export');
  XLSX.utils.book_append_sheet(wb,ws,'Branch Master');
  XLSX.writeFile(wb,`${c.label.replace(/[^a-z0-9]/gi,'_')}_BranchMaster_${new Date().toISOString().slice(0,10)}.xlsx`);
}


// ── SCHEDULE & COUNT VIEW ──────────────────────────────────────────────────────
let schSortCol=-1, schSortAsc=true;

function getFilteredSchedule(){
  const c=getClient();
  const mo=document.getElementById('sch-f-month').value;
  const wk=document.getElementById('sch-f-week').value;
  const rg=document.getElementById('sch-f-region').value;
  const st=document.getElementById('sch-f-status').value;
  const srch=document.getElementById('sch-f-search').value.toLowerCase().trim();
  // If any month checkboxes are ticked in the sidebar, scope the table to just those months.
  const checkedMonths=c.selectedMonths;
  return c.schedule.filter(r=>{
    const isMatch=isBranchMatched(c,r.branchNo);
    if(st==='match'&&!isMatch) return false;
    if(st==='nomatch'&&isMatch) return false;
    if(st==='watchlisted'&&!c.selected.has(r.branchNo)) return false;
    if(checkedMonths&&checkedMonths.size>0&&!checkedMonths.has(r.month)) return false;
    if(mo&&r.month!==mo) return false;
    if(wk&&r.week!==wk) return false;
    if(rg&&r.region!==rg) return false;
    if(srch&&!r.branchNo.toLowerCase().includes(srch)&&!r.branchName.toLowerCase().includes(srch)) return false;
    return true;
  });
}

function updateSchMonthWeekOptions(){
  const c=getClient();
  const months=[...new Set(c.schedule.map(r=>r.month))];
  const weeks=[...new Set(c.schedule.map(r=>r.week))].sort();
  const fm=document.getElementById('sch-f-month'), fw=document.getElementById('sch-f-week');
  const cm=fm.value, cw=fw.value;
  fm.innerHTML='<option value="">All months</option>'+months.map(m=>`<option${m===cm?' selected':''}>${m}</option>`).join('');
  fw.innerHTML='<option value="">All weeks</option>'+weeks.map(w=>`<option${w===cw?' selected':''}>${w}</option>`).join('');
}

function schSortByCol(col){
  if(schSortCol===col) schSortAsc=!schSortAsc; else{schSortCol=col;schSortAsc=true;}
  document.querySelectorAll('[id^="sth"]').forEach((th,i)=>{
    th.classList.remove('sort-asc','sort-desc');
    if(i===col) th.classList.add(schSortAsc?'sort-asc':'sort-desc');
  });
  renderSchedule();
}
window.schSortByCol=schSortByCol;

function renderRecordedCountsList(){
  const c=getClient();
  const badge=document.getElementById('rc-count-badge');
  const listEl=document.getElementById('rc-list');
  if(!badge||!listEl) return;
  badge.textContent=c.counted.size;

  const srch=(document.getElementById('rc-search')||{value:''}).value.toLowerCase().trim();
  let entries=[...c.counted.values()];
  if(srch){
    entries=entries.filter(e=>String(e.branchNo||'').toLowerCase().includes(srch)||String(e.branchName||'').toLowerCase().includes(srch));
  }
  // Most recently recorded first (Map preserves insertion order; reverse for newest-first)
  entries=entries.reverse();

  if(!entries.length){
    listEl.innerHTML=c.counted.size
      ? '<p class="hint" style="text-align:center;color:#bbb">No matches for that search.</p>'
      : '<p class="hint" style="text-align:center;color:#bbb">No count results recorded yet.</p>';
    return;
  }

  listEl.innerHTML=entries.map(e=>{
    const pillCls=riskPillClass(e.level);
    const diffNum=Number(e.diff);
    const diffDisplay=isNaN(diffNum)?'\u2014':diffNum.toFixed(2)+'%';
    const safeNo=String(e.branchNo).replace(/'/g,"&#39;");
    const safeName=String(e.branchName||'').replace(/</g,'&lt;');
    return `<div style="display:flex;align-items:flex-start;gap:6px;padding:7px 0;border-bottom:1px solid #f5f5f5;font-size:11px">
      <div style="flex:1;min-width:0">
        <div style="font-weight:700;color:#1a3a5c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${e.branchNo} \u2014 ${safeName}</div>
        <div style="color:#888;margin-top:2px">${e.date||'\u2014'} &middot; <span class="risk-pill ${pillCls}" style="font-size:10px">${diffDisplay} ${e.level||''}</span></div>
      </div>
      <div style="display:flex;gap:6px;flex-shrink:0;align-items:center">
        <button type="button" data-action="edit-count" data-branchno="${safeNo}" style="background:#eef4fa;border:1px solid #d8e4f0;cursor:pointer;color:#2980b9;font-size:12px;padding:4px 7px;border-radius:5px;line-height:1" title="Edit">&#9998;</button>
        <button type="button" data-action="clear-count" data-branchno="${safeNo}" style="background:#fdeeee;border:1px solid #f5cfcf;cursor:pointer;color:#c0392b;font-size:12px;padding:4px 7px;border-radius:5px;line-height:1" title="Clear this record">&#10005;</button>
      </div>
    </div>`;
  }).join('');
}

function clearRecordedCount(no){
  const c=getClient();
  if(!confirm(`Clear the recorded count result for branch ${no}? This cannot be undone.`)) return;
  c.counted.delete(no);
  refreshAllViews();
  saveState();
}
window.clearRecordedCount=clearRecordedCount;

function clearAllRecordedCounts(){
  try{
    const c=getClient();
    if(!c.counted.size){alert('No recorded counts to clear.');return;}
    const countBefore=c.counted.size;
    if(!confirm(`Clear ALL ${countBefore} recorded count result(s) for ${c.name}? This cannot be undone.`)) return;
    c.counted.clear();
    saveState();
    refreshAllViews();
    // Verify the clear actually took effect
    if(getClient().counted.size!==0){
      console.error('Clear all failed: counted size is still', getClient().counted.size);
      alert(`Something went wrong \u2014 the list still shows ${getClient().counted.size} record(s). Please try refreshing the page and clearing again.`);
    }
  }catch(err){
    console.error('clearAllRecordedCounts failed', err);
    alert(`Could not clear recorded counts.\n\nError: ${err.message}`);
  }
}

function renderSchedule(){
  const c=getClient();

  // Sidebar: months list — ordered by c.monthOrder, with any new months appended at the end.
  // Manually-added entries (note:"Manual") are excluded from this list — they aren't a real uploaded month.
  const allMonths=[...new Set(c.schedule.filter(r=>r.note!=='Manual').map(r=>r.month))];
  syncMonthOrder(c, allMonths);
  const months=c.monthOrder.filter(m=>allMonths.includes(m));

  const ml=document.getElementById('sch-month-list');
  if(months.length){
    ml.innerHTML=months.map((m,idx)=>{
      const cnt=c.schedule.filter(r=>r.month===m).length;
      const checked=c.selectedMonths.has(m);
      return `<div style="display:flex;align-items:center;gap:5px;background:${checked?'#eaf4fd':'#f5f8fb'};border:1px solid ${checked?'#aac6dd':'#d8dde5'};border-radius:6px;padding:4px 6px;font-size:11px">
        <input type="checkbox" data-month-checkbox="${m}" ${checked?'checked':''} style="cursor:pointer;flex-shrink:0" title="Select this month">
        <span style="flex:1;font-weight:700;color:#1a3a5c;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">&#128197; ${m}</span>
        <span style="color:#888;font-size:10px;flex-shrink:0">${cnt}</span>
        <button data-month-up="${m}" ${idx===0?'disabled':''} style="background:none;border:none;cursor:${idx===0?'default':'pointer'};color:${idx===0?'#ddd':'#2980b9'};font-size:12px;padding:0 2px;flex-shrink:0" title="Move up">&#9650;</button>
        <button data-month-down="${m}" ${idx===months.length-1?'disabled':''} style="background:none;border:none;cursor:${idx===months.length-1?'default':'pointer'};color:${idx===months.length-1?'#ddd':'#2980b9'};font-size:12px;padding:0 2px;flex-shrink:0" title="Move down">&#9660;</button>
        <button data-month-remove="${m}" style="background:none;border:none;cursor:pointer;color:#c0392b;font-size:12px;padding:0 0 0 2px;flex-shrink:0" title="Remove this month">&#10005;</button>
      </div>`;
    }).join('');
    document.getElementById('sch-months-summary').textContent=months.join(' \u00b7 ');
  } else {
    ml.innerHTML='';
    document.getElementById('sch-months-summary').textContent='No schedule loaded';
  }
  renderMonthSelectionBar(c, months);

  renderWatchlist();

  if(mainView!=='schedule') { renderScheduleStats(); return; }

  updateSchMonthWeekOptions();
  renderScheduleStats();

  let rows=getFilteredSchedule();
  if(schSortCol>=0){
    const keys=['month','week','_d','day','branchNo','branchName','area','region','countTime'];
    rows=[...rows].sort((a,b)=>{
      let va,vb;
      if(schSortCol===2){ va=dateSortKey(a.date); vb=dateSortKey(b.date); }
      else if(schSortCol===9){ va=isBranchMatched(c,a.branchNo)?0:1; vb=isBranchMatched(c,b.branchNo)?0:1; }
      else if(schSortCol===10){ const ca=c.counted.get(a.branchNo),cb=c.counted.get(b.branchNo); va=ca?ca.date:''; vb=cb?cb.date:''; }
      else if(schSortCol===11){ const ca=c.counted.get(a.branchNo),cb=c.counted.get(b.branchNo); va=(ca&&ca.diff!==null)?ca.diff:-1; vb=(cb&&cb.diff!==null)?cb.diff:-1; }
      else if(schSortCol===12){ const order={null:-1,undefined:-1,Low:0,Medium:1,High:2}; const ca=c.counted.get(a.branchNo),cb=c.counted.get(b.branchNo); va=order[ca?ca.level:null]; vb=order[cb?cb.level:null]; }
      else{ va=(a[keys[schSortCol]]||'').toString().toLowerCase(); vb=(b[keys[schSortCol]]||'').toString().toLowerCase(); }
      if(va<vb) return schSortAsc?-1:1;
      if(va>vb) return schSortAsc?1:-1;
      return 0;
    });
  }

  document.getElementById('sch-row-count').textContent=rows.length+' of '+c.schedule.length+' rows';
  const tbody=document.getElementById('sch-tbl-body');
  if(!c.schedule.length){
    tbody.innerHTML=`<tr><td colspan="11"><div class="empty-state"><p>No schedule loaded.<br>Upload a CSV/XLSX or load the sample.</p></div></td></tr>`;
    return;
  }
  if(!rows.length){
    tbody.innerHTML=`<tr><td colspan="11"><div class="empty-state"><p>No entries match current filters.</p></div></td></tr>`;
    return;
  }
  const WBG={W1:'#f7fafd',W2:'#f5fbf6',W3:'#fdfdf5',W4:'#fdf8f5'};
  const effectiveMap=getEffectiveScheduleMap(c); // branchNo -> the one row that "counts" when duplicated
  tbody.innerHTML=rows.map(r=>{
    const isMatch=isBranchMatched(c,r.branchNo);
    const cr=c.counted.get(r.branchNo);
    const isManual=(r.note||'').startsWith('Manual');
    const isUnregistered=(r.note||'').includes('Not Registered');
    const isWatchlisted=c.selected.has(r.branchNo);
    const effectiveRow=effectiveMap.get(r.branchNo);
    const isDuplicate=c.schedule.filter(x=>x.branchNo===r.branchNo).length>1;
    const isEffective=!effectiveRow||effectiveRow===r||effectiveRow.date===r.date;

    const bg=cr?'#eaf4fd':isMatch?'#eafaef':isManual?'#fdf6e8':(WBG[r.week]||'#fff');
    const timePill=r.countTime==='Night'?'<span class="risk-pill" style="background:#fff3cd;color:#8a6500">Night</span>':'<span style="color:#999;font-size:11px">Morning</span>';
    const statusPill=(isMatch?'<span class="risk-pill risk-low">&#10004; Match</span>':'<span class="risk-pill" style="background:#e8f2fb;color:#1a3a5c">Scheduled</span>')
      +(isWatchlisted?` <span class="risk-pill" style="background:#fff3cd;color:#8a6500;font-size:9px;padding:1px 5px" title="${(c.selected.get(r.branchNo)||{}).note||'Watchlisted'}">&#128204; Watchlist</span>`:'');
    const manualBadge=isManual?' <span class="risk-pill" style="background:#fde9c8;color:#8a5a00;font-size:9px;padding:1px 5px" title="Manually added \u2014 not in any uploaded schedule">Manual</span>':'';
    const unregBadge=isUnregistered?' <span class="risk-pill" style="background:#fde8e8;color:#a32020;font-size:9px;padding:1px 5px" title="Branch not found in Branch Master Data \u2014 register it to fill in the name">Not Registered</span>':'';
    const dupBadge=(isDuplicate&&!isEffective)?` <span class="risk-pill" style="background:#eee;color:#888;font-size:9px;padding:1px 5px" title="This branch also appears on ${effectiveRow.date} \u2014 that date is used for counting">See ${effectiveRow.date}</span>`:'';
    const watchPin=isWatchlisted?` <span title="${(c.selected.get(r.branchNo)||{}).note||'Watchlisted'}">&#128204;</span>`:'';
    const nameDisplay=r.branchName||`<span style="color:#bbb;font-style:italic">(unregistered)</span>`;

    const countDateVal=cr?cr.date:'';
    const diffVal=(cr&&cr.diff!==null&&cr.diff!==undefined)?cr.diff:'';
    const lvlCell=cr&&cr.level?`<span class="risk-pill ${riskPillClass(cr.level)}">${cr.level}</span>`:'<span style="color:#ccc;font-size:11px">\u2014</span>';
    const rid=r.branchNo+'_'+r.date.replace(/[^a-z0-9]/gi,'');
    const rowKey=r.branchNo+'|'+r.date;

    // Only the effective (resolved) row for a duplicated branch gets editable count cells.
    // The other duplicate(s) show the same resolved result, locked, with a pointer to the effective date.
    const dateCell = isEffective
      ? `<input type="date" id="sc-date-${rid}" value="${countDateVal}" style="width:120px;padding:3px 5px;font-size:11px;border:1px solid #d8dde5;border-radius:4px" onchange="inlineSaveCount('${r.branchNo}','${r.branchName.replace(/'/g,"\\'")}')">`
      : `<span style="font-size:11px;color:#aaa">${countDateVal||'\u2014'}</span>`;
    const diffCell = isEffective
      ? `<input type="number" id="sc-diff-${rid}" value="${diffVal}" step="0.01" min="0" placeholder="\u2014" style="width:70px;padding:3px 5px;font-size:11px;border:1px solid #d8dde5;border-radius:4px;text-align:right" onkeydown="if(event.key==='Enter')this.blur()" onchange="inlineSaveCount('${r.branchNo}','${r.branchName.replace(/'/g,"\\'")}')">`
      : `<span style="font-size:11px;color:#aaa">${diffVal!==''?diffVal+'%':'\u2014'}</span>`;

    return `<tr style="background:${bg}">
      <td>${r.month}</td><td style="text-align:center;font-weight:700;color:#1a3a5c">${r.week}</td>
      <td style="text-align:center">${r.date}</td><td style="text-align:center">${r.day.substring(0,3)}</td>
      <td class="td-num">${r.branchNo}</td><td class="td-name" title="${r.branchName}">${watchPin}${nameDisplay}${manualBadge}${unregBadge}${dupBadge}</td>
      <td style="text-align:center">${r.area}</td><td>${r.region}</td>
      <td style="text-align:center">${timePill}</td><td style="text-align:center">${statusPill}</td>
      <td style="text-align:center">${dateCell}</td>
      <td style="text-align:center">${diffCell}</td>
      <td style="text-align:center" id="sc-lvl-${rid}">${lvlCell}</td>
      <td style="text-align:center"><button data-action="delete-sched-row" data-rowkey="${rowKey}" style="background:none;border:none;cursor:pointer;color:#c0392b;font-size:13px;padding:2px" title="Delete this schedule row">&#10005;</button></td>
    </tr>`;
  }).join('');
}

// Inline save: triggered when the Count date or Diff % cell changes directly in the schedule table.
// Updates c.counted using whatever date+diff currently sit in those two inputs for that row.
function inlineSaveCount(branchNo, branchName){
  try{
    const c=getClient();
    // Escape branchNo for safe use inside a CSS attribute selector (handles any special characters)
    const escaped = window.CSS && CSS.escape ? CSS.escape(branchNo) : branchNo.replace(/[^a-zA-Z0-9_-]/g,'\\$&');
    const dateInputs=document.querySelectorAll(`input[id^="sc-date-${escaped}_"]`);
    const diffInputs=document.querySelectorAll(`input[id^="sc-diff-${escaped}_"]`);

    let date='', diff='';
    for(const el of dateInputs){ if(el.value){ date=el.value; break; } }
    for(const el of diffInputs){ if(el.value!==''){ diff=el.value; break; } }

    if(diff===''){
      if(c.counted.has(branchNo)){
        c.counted.delete(branchNo);
        saveState();
        refreshAllViews();
      }
      return;
    }
    if(!date) date=new Date().toISOString().slice(0,10);

    const pct=parseFloat(diff);
    if(isNaN(pct)||pct<0){
      alert(`"${diff}" is not a valid difference %. Please enter a number (e.g. 0.75).`);
      return;
    }
    const lvl=calcRiskLevel(pct);
    const existing=c.counted.get(branchNo)||{};
    c.counted.set(branchNo,{
      branchNo, branchName: branchName||existing.branchName||branchNo, date,
      bookQty: existing.bookQty!==undefined?existing.bookQty:null,
      actualQty: existing.actualQty!==undefined?existing.actualQty:null,
      diff:pct, level:lvl,
    });
    saveState();
    refreshAllViews();
  }catch(err){
    console.error('inlineSaveCount failed for branch', branchNo, err);
    alert(`Could not save the count result for branch ${branchNo}.\n\nError: ${err.message}\n\nPlease try the "+ Detailed entry form" button in the sidebar instead, or report this with the branch number.`);
  }
}
window.inlineSaveCount=inlineSaveCount;

function renderScheduleStats(){
  const c=getClient();
  const tot=c.schedule.length, sel=totalMatchedCount(c);
  const matchV=c.schedule.filter(r=>isBranchMatched(c,r.branchNo)).length;
  const matchedNos=new Set(c.schedule.filter(r=>isBranchMatched(c,r.branchNo)).map(r=>r.branchNo));
  const allMatched=new Set([...c.thresholdSelected.keys(), ...c.manuallyMatched.keys()]);
  const noShow=sel>0?[...allMatched].filter(n=>!matchedNos.has(n)).length:0;
  // Counted = branches IN THIS SCHEDULE that have a recorded count result (not the client's entire count history).
  const scheduledBranchNos=new Set(c.schedule.map(r=>r.branchNo));
  const counted=[...scheduledBranchNos].filter(no=>c.counted.has(no)).length;
  document.getElementById('sch-total').textContent=tot;
  document.getElementById('sch-sel').textContent=sel;
  document.getElementById('sch-match').textContent=matchV;
  document.getElementById('sch-counted').textContent=counted;
  document.getElementById('sch-noshow').textContent=noShow;
}

function resetSchFilters(){
  ['sch-f-month','sch-f-week','sch-f-region','sch-f-status'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('sch-f-search').value='';
  schSortCol=-1; schSortAsc=true;
  document.querySelectorAll('[id^="sth"]').forEach(th=>th.classList.remove('sort-asc','sort-desc'));
  renderSchedule();
}

function deleteScheduleRow(rowKey){
  const c=getClient();
  const [branchNo, date]=rowKey.split('|');
  const idx=c.schedule.findIndex(r=>r.branchNo+'|'+r.date===rowKey);
  if(idx<0) return;
  if(!confirm(`Delete the schedule entry for branch ${branchNo} on ${date}?`)) return;
  c.schedule.splice(idx,1);
  refreshAllViews();
  saveState();
}
window.deleteScheduleRow=deleteScheduleRow;

function removeScheduleMonth(month){
  if(!confirm(`Remove all "${month}" schedule data?`))return;
  const c=getClient();
  c.schedule=c.schedule.filter(r=>r.month!==month);
  c.monthOrder=(c.monthOrder||[]).filter(m=>m!==month);
  c.selectedMonths.delete(month);
  refreshAllViews();
  saveState();
}
window.removeScheduleMonth=removeScheduleMonth;

// ── MONTH ORDERING & SELECTION ────────────────────────────────────────────────
// Keeps c.monthOrder in sync with whatever months actually exist in the schedule:
// preserves the user's chosen order, appends any newly-uploaded months at the end,
// and drops any months that no longer have data (e.g. after a remove).
function syncMonthOrder(c, allMonths){
  if(!c.monthOrder) c.monthOrder=[];
  if(!c.selectedMonths) c.selectedMonths=new Set();
  // Append new months not yet in the order
  allMonths.forEach(m=>{ if(!c.monthOrder.includes(m)) c.monthOrder.push(m); });
  // Drop months no longer present
  c.monthOrder=c.monthOrder.filter(m=>allMonths.includes(m));
}

function moveMonth(month, direction){
  const c=getClient();
  const allMonths=[...new Set(c.schedule.filter(r=>r.note!=='Manual').map(r=>r.month))];
  syncMonthOrder(c, allMonths);
  const idx=c.monthOrder.indexOf(month);
  const newIdx=idx+direction;
  if(idx<0||newIdx<0||newIdx>=c.monthOrder.length) return;
  [c.monthOrder[idx], c.monthOrder[newIdx]]=[c.monthOrder[newIdx], c.monthOrder[idx]];
  renderSchedule();
  saveState();
}
window.moveMonth=moveMonth;

function toggleMonthSelection(month, isChecked){
  const c=getClient();
  if(isChecked) c.selectedMonths.add(month);
  else c.selectedMonths.delete(month);
  renderSchedule();
  saveState();
}
window.toggleMonthSelection=toggleMonthSelection;

function renderMonthSelectionBar(c, months){
  const bar=document.getElementById('sch-month-selection-bar');
  if(!bar) return;
  const n=c.selectedMonths.size;
  if(!months.length){ bar.style.display='none'; return; }
  bar.style.display='flex';
  document.getElementById('sch-month-sel-count').textContent = n>0 ? `${n} month(s) selected — table filtered` : 'None selected';
  document.getElementById('btn-sch-remove-selected-months').disabled = n===0;
  document.getElementById('btn-sch-remove-selected-months').style.opacity = n===0?'.4':'1';
  document.getElementById('btn-sch-remove-selected-months').style.cursor = n===0?'default':'pointer';
}

function selectAllMonths(){
  const c=getClient();
  const allMonths=[...new Set(c.schedule.filter(r=>r.note!=='Manual').map(r=>r.month))];
  allMonths.forEach(m=>c.selectedMonths.add(m));
  renderSchedule();
  saveState();
}

function deselectAllMonths(){
  const c=getClient();
  c.selectedMonths.clear();
  renderSchedule();
  saveState();
}

function removeSelectedMonths(){
  const c=getClient();
  const n=c.selectedMonths.size;
  if(!n) return;
  const list=[...c.selectedMonths].join(', ');
  if(!confirm(`Remove schedule data for ${n} selected month(s)?\n\n${list}\n\nThis cannot be undone.`)) return;
  c.schedule=c.schedule.filter(r=>!c.selectedMonths.has(r.month));
  c.monthOrder=(c.monthOrder||[]).filter(m=>!c.selectedMonths.has(m));
  c.selectedMonths.clear();
  refreshAllViews();
  saveState();
}

function loadSampleSchedule(){
  const c=getClient();
  const months=['May 2026','Jun 2026','Jul 2026'];
  const already=months.some(m=>c.schedule.some(r=>r.month===m));
  if(already&&!confirm('Sample schedule (May–Jul 2026) already loaded. Replace it?'))return;
  c.schedule=c.schedule.filter(r=>!months.includes(r.month));
  c.schedule=[...c.schedule,...SAMPLE_SCHEDULE.map(r=>({month:r[0],week:r[1],day:r[2],date:r[3],branchNo:r[4],branchName:r[5],area:r[6],region:r[7],countTime:r[8],note:''}))];
  refreshAllViews();
  saveState();
}

// Looks up a branch number against Branch Master Data. Returns {name, registered}.
// If not registered, name is left blank per spec — caller is responsible for tagging
// the row with the "Not Registered" note so the user knows to go register it.
function resolveBranchInfo(c, branchNo){
  const b=c.branches.find(b=>b.no===branchNo);
  if(b) return {name:b.name||'', registered:true};
  return {name:'', registered:false};
}

function parseScheduleRows(rows, filename){
  const c=getClient();
  const newRows=rows.filter(r=>r.length>=8&&r[0]).map(r=>{
    const branchNo=String(r[4]||'');
    const info=resolveBranchInfo(c, branchNo);
    return {
      month:String(r[0]||''), week:String(r[1]||''), day:String(r[2]||''), date:String(r[3]||''),
      branchNo, branchName:info.registered?info.name:'', area:String(r[6]||''), region:String(r[7]||''), countTime:String(r[8]||'Morning'),
      note: info.registered?'':'Not Registered',
    };
  });
  if(!newRows.length){alert(`No valid rows found in "${filename}".`);return;}
  const existingKeys=new Set(c.schedule.map(r=>r.branchNo+'|'+r.date));
  const added=newRows.filter(r=>!existingKeys.has(r.branchNo+'|'+r.date));
  const dupes=newRows.length-added.length;
  const unregistered=added.filter(r=>r.note==='Not Registered').length;
  c.schedule=[...c.schedule,...added];
  refreshAllViews();
  saveState();
  let msg=`Added ${added.length} entries from "${filename}".`;
  if(dupes) msg+=` ${dupes} duplicates skipped.`;
  if(unregistered) msg+=`\n\n\u26a0\ufe0f ${unregistered} branch(es) are not yet in Branch Master Data \u2014 tagged "Not Registered". Register them there to fill in branch names automatically.`;
  alert(msg);
}

function parseScheduleCSV(txt,fname){
  const lines=txt.split(/\r?\n/).filter(l=>l.trim());
  parseScheduleRows(lines.slice(1).map(l=>l.split(',').map(c=>c.trim().replace(/^"|"$/g,''))), fname);
}

function handleScheduleFileUpload(files){
  Array.from(files).forEach(f=>{
    const ext=f.name.split('.').pop().toLowerCase();
    if(ext==='csv'||ext==='txt'){
      const r=new FileReader();
      r.onload=ev=>parseScheduleCSV(ev.target.result,f.name);
      r.readAsText(f);
    } else if(ext==='xlsx'||ext==='xls'){
      const r=new FileReader();
      r.onload=ev=>{
        const wb2=XLSX.read(ev.target.result,{type:'binary'});
        const ws=wb2.Sheets[wb2.SheetNames[0]];
        parseScheduleRows(XLSX.utils.sheet_to_json(ws,{header:1}).slice(1), f.name);
      };
      r.readAsBinaryString(f);
    } else {
      alert(`"${f.name}": Unsupported type. Use CSV or XLSX.`);
    }
  });
}

// ── MANUAL SCHEDULE ENTRY (special counts not in any uploaded file) ──────────
// Tagged note:"Manual" so they're identifiable in the table/export, and deliberately
// excluded from the month-order/reorder list since they aren't a real "month" of data.
const MONTH_ABBR=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const DAY_NAMES=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

function deriveDateParts(isoDate){
  // isoDate is "YYYY-MM-DD" from <input type=date>. Returns {month, week, day, date}
  const d=new Date(isoDate+'T00:00:00');
  if(isNaN(d.getTime())) return null;
  const monthLabel=`${MONTH_ABBR[d.getMonth()]} ${d.getFullYear()}`;
  const dayName=DAY_NAMES[d.getDay()];
  const dateLabel=`${d.getDate()}-${MONTH_ABBR[d.getMonth()]}`;
  // Week-of-month label (W1..W5) based on day-of-month, consistent with how uploaded schedules are bucketed
  const weekNum=Math.ceil(d.getDate()/7);
  return { month:monthLabel, week:'W'+weekNum, day:dayName, date:dateLabel };
}

function updateManualBranchPreview(){
  const c=getClient();
  const no=document.getElementById('man-branchno').value.trim();
  const previewEl=document.getElementById('man-branchname-preview');
  if(!previewEl) return;
  if(!no){ previewEl.textContent=''; return; }
  const info=resolveBranchInfo(c, no);
  if(info.registered){
    previewEl.innerHTML=`<span style="color:#1a6b2e">&#10003; ${info.name}</span>`;
  } else {
    previewEl.innerHTML=`<span style="color:#a32020">&#9888; Not in Branch Master Data \u2014 will be tagged "Not Registered"</span>`;
  }
}

function addManualScheduleEntry(){
  const no=document.getElementById('man-branchno').value.trim();
  const isoDate=document.getElementById('man-date').value;
  const countTime=document.getElementById('man-counttime').value;
  const area=document.getElementById('man-area').value.trim();
  const region=document.getElementById('man-region').value.trim();

  if(!no){alert('Enter a branch number.');return;}
  if(!isoDate){alert('Pick a count date.');return;}

  const parts=deriveDateParts(isoDate);
  if(!parts){alert('Invalid date.');return;}

  const c=getClient();
  const key=no+'|'+parts.date;
  if(c.schedule.some(r=>r.branchNo+'|'+r.date===key)){
    alert(`Branch ${no} already has an entry on ${parts.date}. No new record was added.`);
    return;
  }

  const info=resolveBranchInfo(c, no);

  c.schedule.push({
    month:parts.month, week:parts.week, day:parts.day, date:parts.date,
    branchNo:no, branchName:info.registered?info.name:'', area:area||'', region:region||'', countTime,
    note: info.registered?'Manual':'Manual, Not Registered',
  });

  // Clear the form
  ['man-branchno','man-area','man-region'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('man-date').value='';
  document.getElementById('man-counttime').value='Morning';

  refreshAllViews();
  saveState();

  if(!info.registered){
    alert(`Added. \u26a0\ufe0f Branch ${no} is not yet in Branch Master Data \u2014 tagged "Not Registered". Register it there to fill in the branch name automatically.`);
  }
}

function addManualScheduleBulk(){
  const txt=document.getElementById('man-bulk-text').value;
  const lines=txt.split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
  if(!lines.length){alert('Paste at least one row first.');return;}

  const c=getClient();
  let added=0, skipped=0, unregistered=0;
  const errors=[];

  lines.forEach((line,i)=>{
    const parts=line.split(',').map(s=>s.trim());
    const no=parts[0]||'';
    const isoDate=parts[1]||'';
    const countTime=parts[2]||'Morning';
    const area=parts[3]||'';
    const region=parts[4]||'';

    if(!no||!isoDate){ errors.push(`Line ${i+1}: missing branch no. or date \u2014 skipped`); skipped++; return; }
    const dateParts=deriveDateParts(isoDate);
    if(!dateParts){ errors.push(`Line ${i+1}: invalid date "${isoDate}" \u2014 skipped`); skipped++; return; }

    const key=no+'|'+dateParts.date;
    if(c.schedule.some(r=>r.branchNo+'|'+r.date===key)){ skipped++; return; }

    const info=resolveBranchInfo(c, no);
    if(!info.registered) unregistered++;

    c.schedule.push({
      month:dateParts.month, week:dateParts.week, day:dateParts.day, date:dateParts.date,
      branchNo:no, branchName:info.registered?info.name:'', area, region,
      countTime: (countTime==='Night'?'Night':'Morning'),
      note: info.registered?'Manual':'Manual, Not Registered',
    });
    added++;
  });

  document.getElementById('man-bulk-text').value='';
  refreshAllViews();
  saveState();

  let msg=`Added ${added} manual entr${added===1?'y':'ies'}.`;
  if(skipped) msg+=` ${skipped} skipped (duplicate or invalid).`;
  if(unregistered) msg+=`\n\n\u26a0\ufe0f ${unregistered} branch(es) not yet in Branch Master Data \u2014 tagged "Not Registered".`;
  if(errors.length) msg+='\n\n'+errors.slice(0,5).join('\n')+(errors.length>5?`\n...and ${errors.length-5} more.`:'');
  alert(msg);
}

// ── BRANCHES TO COUNT (selection list) ─────────────────────────────────────────
// ── WATCHLIST (personal follow-up flags, separate from any matching logic) ───
function addToWatchlist(){
  const no=document.getElementById('wl-inp-no').value.trim();
  const note=document.getElementById('wl-inp-note').value.trim();
  if(!no){alert('Enter a branch number.');return;}
  const c=getClient();
  const existing=c.selected.get(no);
  c.selected.set(no, {note: note||(existing?existing.note:''), addedDate: existing?existing.addedDate:new Date().toISOString().slice(0,10)});
  document.getElementById('wl-inp-no').value='';
  document.getElementById('wl-inp-note').value='';
  refreshAllViews();
  saveState();
}
function removeFromWatchlist(no){
  getClient().selected.delete(no);
  refreshAllViews();
  saveState();
}
window.removeFromWatchlist=removeFromWatchlist;
function clearWatchlist(){
  if(!getClient().selected.size){return;}
  if(!confirm(`Clear all ${getClient().selected.size} watchlisted branch(es)?`)) return;
  getClient().selected.clear();
  refreshAllViews();
  saveState();
}
function renderWatchlist(){
  const c=getClient();
  const badge=document.getElementById('watchlist-count-badge');
  const listEl=document.getElementById('watchlist-list');
  if(!badge||!listEl) return;
  badge.textContent=c.selected.size;
  if(!c.selected.size){
    listEl.innerHTML='<p class="hint" style="text-align:center;color:#bbb">No branches watchlisted yet.</p>';
    return;
  }
  listEl.innerHTML=[...c.selected.entries()].map(([no,rawInfo])=>{
    // Normalize legacy shapes (plain string, or {name} from the old scratch-list) into {note, addedDate}.
    const info = (rawInfo && typeof rawInfo==='object') ? rawInfo : {note:'', addedDate:''};
    const b=c.branches.find(b=>b.no===no);
    const name=b?b.name:'';
    const safeNo=String(no).replace(/'/g,"&#39;");
    return `<div style="display:flex;align-items:flex-start;gap:6px;padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:11px">
      <div style="flex:1;min-width:0">
        <div style="font-weight:700;color:#1a3a5c">&#128204; ${no}${name?' \u2014 '+name:''}</div>
        ${info.note?`<div style="color:#888;margin-top:2px">${info.note}</div>`:''}
        <div style="color:#bbb;font-size:9px;margin-top:1px">Added ${info.addedDate||''}</div>
      </div>
      <button data-action="remove-watchlist" data-branchno="${safeNo}" style="background:none;border:none;cursor:pointer;color:#c0392b;font-size:13px;padding:2px" title="Remove from watchlist">&#10005;</button>
    </div>`;
  }).join('');
}

function _doExportScheduleFiltered(handoff){
  const c=getClient();
  if(!c.schedule.length){alert('No schedule loaded.');return;}
  const rows=getFilteredSchedule();
  if(!rows.length){alert('No rows match the current filters \u2014 nothing to export.');return;}

  const H=['Month','Week','Date','Day','Branch No.','Branch Name','Area Code','Region','Count Time','Status','Note','Count Date','Diff %','Risk Level'];
  const data=[H,...rows.map(r=>{
    const cr=c.counted.get(r.branchNo);
    return [
      r.month, r.week, r.date, r.day, r.branchNo, r.branchName, r.area, r.region, r.countTime,
      isBranchMatched(c,r.branchNo)?'Match':'Scheduled',
      r.note||'',
      cr?cr.date:'', (cr&&cr.diff!==null&&cr.diff!==undefined)?cr.diff:'', cr?cr.level||'':'',
    ];
  })];
  const ws=XLSX.utils.aoa_to_sheet(data);
  ws['!cols']=[{wch:12},{wch:8},{wch:10},{wch:12},{wch:12},{wch:36},{wch:10},{wch:12},{wch:12},{wch:12},{wch:18},{wch:12},{wch:8},{wch:10}];
  const wb=XLSX.utils.book_new();
  if(handoff) addHandoffCoverSheet(wb, c, handoff, 'Schedule & Count export (filtered view)');
  XLSX.utils.book_append_sheet(wb,ws,'Schedule (filtered)');
  XLSX.writeFile(wb,`${c.label.replace(/[^a-z0-9]/gi,'_')}_Schedule_Filtered_${new Date().toISOString().slice(0,10)}.xlsx`);
}


// ── INIT ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function(){

  // Expose functions for dynamically-generated HTML onclick handlers
  window.switchClient=switchClient;
  window.switchMainView=switchMainView;
  window.pickClientColor=pickClientColor;
  window.openCountEntryFor=openCountEntryFor;
  window.openEditBranch=openEditBranch;
  window.removeScheduleMonth=removeScheduleMonth;
  window.dashSortBy=dashSortBy;
  window.bmSortBy=bmSortBy;
  window.schSortByCol=schSortByCol;

  // ── Client management ───────────────────────────────────────────────────────
  document.getElementById('btn-add-client').addEventListener('click', showAddClient);
  document.getElementById('btn-cl-cancel').addEventListener('click', closeClientModal);
  document.getElementById('btn-cl-save').addEventListener('click', saveClientModal);
  document.getElementById('client-modal-bg').addEventListener('click', function(e){ if(e.target===this) closeClientModal(); });

  // ── Main view tabs ────────────────────────────────────────────────────────
  document.getElementById('mt-home').addEventListener('click', ()=>switchMainView('home'));
  document.getElementById('mt-dashboard').addEventListener('click', ()=>switchMainView('dashboard'));
  document.getElementById('mt-branches').addEventListener('click', ()=>switchMainView('branches'));
  document.getElementById('mt-schedule').addEventListener('click', ()=>switchMainView('schedule'));

  // ── DASHBOARD view ────────────────────────────────────────────────────────
  document.getElementById('dash-f-risk').addEventListener('change', renderDashboard);
  document.getElementById('dash-f-region').addEventListener('change', renderDashboard);
  document.getElementById('dash-f-search').addEventListener('input', renderDashboard);
  document.getElementById('btn-dash-reset').addEventListener('click', resetDashFilters);
  document.getElementById('btn-export-dashboard').addEventListener('click', exportDashboard);
  [0,1,2,3,4,5,6].forEach(i=>{
    const el=document.getElementById('dth'+i);
    if(el) el.addEventListener('click', ()=>dashSortBy(i));
  });

  // Count entry modal
  document.getElementById('btn-cm-cancel').addEventListener('click', closeCountModal);
  document.getElementById('btn-cm-save').addEventListener('click', saveCountEntry);
  document.getElementById('count-modal-bg').addEventListener('click', function(e){ if(e.target===this) closeCountModal(); });
  document.getElementById('cm-diff').addEventListener('input', updateCmDiffPreview);
  document.getElementById('cm-bookqty').addEventListener('input', updateCmDiffPreview);
  document.getElementById('cm-actualqty').addEventListener('input', updateCmDiffPreview);
  document.getElementById('cm-branchno').addEventListener('input', function(){
    const no=this.value.trim();
    const c=getClient();
    const found=c.branches.find(b=>b.no===no);
    if(found && !document.getElementById('cm-branchname').value) document.getElementById('cm-branchname').value=found.name;
  });

  // ── BRANCH MASTER DATA view ──────────────────────────────────────────────
  document.getElementById('bm-f-region').addEventListener('change', renderBranchMaster);
  document.getElementById('bm-f-complete').addEventListener('change', renderBranchMaster);
  document.getElementById('bm-f-status').addEventListener('change', renderBranchMaster);
  document.getElementById('bm-f-search').addEventListener('input', renderBranchMaster);
  document.getElementById('btn-add-branch-master').addEventListener('click', openAddBranch);
  document.getElementById('btn-import-branches').addEventListener('click', ()=>document.getElementById('import-branch-file').click());
  document.getElementById('import-branch-file').addEventListener('change', function(e){
    const f=e.target.files[0]; if(!f) return;
    importBranchFile(f); e.target.value='';
  });
  document.getElementById('btn-bm-reset').addEventListener('click', resetBmFilters);
  document.getElementById('btn-export-branches').addEventListener('click', exportBranchMaster);
  document.getElementById('btn-handoff-history-branches').addEventListener('click', openHandoffHistory);
  document.getElementById('btn-apply-threshold').addEventListener('click', applyThresholdSelection);
  document.getElementById('btn-clear-threshold-sel').addEventListener('click', clearThresholdSelection);
  document.getElementById('bm-threshold-haircut').addEventListener('input', function(){
    const c=getClient();
    const v=this.value;
    c.haircutPct = v===''?null:parseFloat(v);
    updateThresholdResultPreview();
  });
  document.getElementById('bm-threshold-haircut').addEventListener('keydown', function(e){ if(e.key==='Enter') applyThresholdSelection(); });
  [0,1,2,3,4,5,6,7,8].forEach(i=>{
    const el=document.getElementById('bth'+i);
    if(el) el.addEventListener('click', ()=>bmSortBy(i));
  });

  // Bulk update panel
  document.getElementById('btn-bulk-csv-pick').addEventListener('click', ()=>document.getElementById('bulk-branch-csv').click());
  document.getElementById('bulk-branch-csv').addEventListener('change', function(e){
    const f=e.target.files[0]; if(!f) return;
    parseBulkBranchCSVFile(f); e.target.value='';
  });
  document.getElementById('btn-bulk-apply').addEventListener('click', parseBulkBranchText);

  // Branch modal
  document.getElementById('btn-bm-cancel').addEventListener('click', closeBranchModal);
  document.getElementById('btn-bm-save').addEventListener('click', saveBranchModal);
  document.getElementById('btn-bm-delete').addEventListener('click', deleteBranchModal);
  document.getElementById('branch-modal-bg').addEventListener('click', function(e){ if(e.target===this) closeBranchModal(); });

  // Exclude branch modal
  document.getElementById('btn-ex-cancel').addEventListener('click', closeExcludeModal);
  document.getElementById('btn-ex-save').addEventListener('click', saveExcludeBranch);
  document.getElementById('exclude-modal-bg').addEventListener('click', function(e){ if(e.target===this) closeExcludeModal(); });
  document.getElementById('ex-note').addEventListener('keydown', function(e){ if(e.key==='Enter' && e.ctrlKey) saveExcludeBranch(); });

  // ── SCHEDULE & COUNT view ────────────────────────────────────────────────
  document.getElementById('btn-record-count').addEventListener('click', openCountEntryBlank);
  ['sch-f-month','sch-f-week','sch-f-region','sch-f-status'].forEach(id=>
    document.getElementById(id).addEventListener('change', renderSchedule));
  document.getElementById('sch-f-search').addEventListener('input', renderSchedule);
  document.getElementById('btn-sch-reset').addEventListener('click', resetSchFilters);
  [0,1,2,3,4,5,6,7,8,9,10,11,12].forEach(i=>{
    const el=document.getElementById('sth'+i);
    if(el) el.addEventListener('click', ()=>schSortByCol(i));
  });

  document.getElementById('btn-sch-upload').addEventListener('click', ()=>document.getElementById('sch-file-input').click());
  document.getElementById('sch-file-input').addEventListener('change', function(e){
    if(e.target.files.length) handleScheduleFileUpload(e.target.files);
    e.target.value='';
  });
  document.getElementById('btn-sch-load-sample').addEventListener('click', loadSampleSchedule);

  // Month reorder / select / deselect
  document.getElementById('btn-sch-select-all-months').addEventListener('click', selectAllMonths);

  // Manual schedule entry
  document.getElementById('btn-man-add').addEventListener('click', addManualScheduleEntry);
  document.getElementById('btn-man-bulk-add').addEventListener('click', addManualScheduleBulk);
  document.getElementById('man-branchno').addEventListener('input', updateManualBranchPreview);
  ['man-branchno','man-area','man-region'].forEach(id=>{
    document.getElementById(id).addEventListener('keydown', e=>{ if(e.key==='Enter') addManualScheduleEntry(); });
  });
  document.getElementById('btn-sch-deselect-all-months').addEventListener('click', deselectAllMonths);
  document.getElementById('btn-sch-remove-selected-months').addEventListener('click', removeSelectedMonths);

  // Event delegation for the month list (checkbox, up/down, remove) — survives every re-render
  document.getElementById('sch-month-list').addEventListener('click', function(e){
    const upBtn=e.target.closest('button[data-month-up]');
    const downBtn=e.target.closest('button[data-month-down]');
    const rmBtn=e.target.closest('button[data-month-remove]');
    if(upBtn && !upBtn.disabled) moveMonth(upBtn.getAttribute('data-month-up'), -1);
    else if(downBtn && !downBtn.disabled) moveMonth(downBtn.getAttribute('data-month-down'), 1);
    else if(rmBtn) removeScheduleMonth(rmBtn.getAttribute('data-month-remove'));
  });
  document.getElementById('sch-month-list').addEventListener('change', function(e){
    const cb=e.target.closest('input[data-month-checkbox]');
    if(cb) toggleMonthSelection(cb.getAttribute('data-month-checkbox'), cb.checked);
  });

  // Event delegation for per-row delete buttons in the schedule table
  document.getElementById('sch-tbl-body').addEventListener('click', function(e){
    const delBtn=e.target.closest('button[data-action="delete-sched-row"]');
    if(delBtn) deleteScheduleRow(delBtn.getAttribute('data-rowkey'));
  });

  document.getElementById('btn-wl-add').addEventListener('click', addToWatchlist);
  document.getElementById('wl-inp-no').addEventListener('keydown', e=>{ if(e.key==='Enter') addToWatchlist(); });
  document.getElementById('wl-inp-note').addEventListener('keydown', e=>{ if(e.key==='Enter') addToWatchlist(); });
  document.getElementById('btn-wl-clear').addEventListener('click', clearWatchlist);
  document.getElementById('watchlist-list').addEventListener('click', function(e){
    const btn=e.target.closest('button[data-action="remove-watchlist"]');
    if(btn) removeFromWatchlist(btn.getAttribute('data-branchno'));
  });

  document.getElementById('btn-sch-export-filtered').addEventListener('click', exportScheduleFiltered);
  document.getElementById('btn-handoff-history-schedule').addEventListener('click', openHandoffHistory);

  // ── Restore built-in data button ─────────────────────────────────────────
  document.getElementById('btn-ho-cancel').addEventListener('click', closeHandoffModal);
  document.getElementById('btn-ho-confirm').addEventListener('click', confirmHandoffExport);
  document.getElementById('handoff-modal-bg').addEventListener('click', function(e){ if(e.target===this) closeHandoffModal(); });
  document.getElementById('ho-from').addEventListener('keydown', e=>{ if(e.key==='Enter') confirmHandoffExport(); });

  document.getElementById('btn-ho-history-close').addEventListener('click', function(){
    document.getElementById('handoff-history-modal-bg').style.display='none';
  });
  document.getElementById('handoff-history-modal-bg').addEventListener('click', function(e){
    if(e.target===this) this.style.display='none';
  });

  document.getElementById('btn-help').addEventListener('click', function(){
    document.getElementById('help-modal-bg').style.display='flex';
  });
  document.getElementById('btn-help-close').addEventListener('click', function(){
    document.getElementById('help-modal-bg').style.display='none';
  });
  document.getElementById('help-modal-bg').addEventListener('click', function(e){
    if(e.target===this) this.style.display='none';
  });

  document.getElementById('btn-restore-builtin').addEventListener('click', function(){
    if(!confirm('Restore the original built-in data for all 4 clients (A, B, C, D)?\n\nThis will REPLACE all current branches, schedules, selections, and count results in this browser with the original dataset. This cannot be undone.')) return;
    clearSavedState();
    clients=[
      { id:1, name:"Client A", label:"Client A", color:COLORS[0],
        branches: buildInitialBranches(CLIENT_A_BRANCHES),
        schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
      { id:2, name:"Client B", label:"Client B", color:COLORS[1],
        branches: buildInitialBranches(CLIENT_B_BRANCHES),
        schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
      { id:3, name:"Client C", label:"Client C", color:COLORS[2],
        branches: buildInitialBranches(CLIENT_C_BRANCHES),
        schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
      { id:4, name:"Client D", label:"Client D", color:COLORS[3],
        branches: buildInitialBranches(CLIENT_D_BRANCHES),
        schedule: [], selected: new Map(), counted: new Map(), threshold: null, haircutPct: 100, thresholdSelected: new Map(), excludedBranches: new Map(), monthOrder: [], selectedMonths: new Set(), manuallyMatched: new Map(), handoffLog: [] },
    ];
    activeClientId=1;
    loadSampleSchedule();
    renderClientTabs();
    refreshAllViews();
    saveState();
    alert('Built-in data restored for all 4 clients.');
  });

  // ── Boot — restore from localStorage or use built-in 4-client data ────────
  const restored=loadState();
  renderClientTabs();
  if(!restored){
    // First-time load: seed Client A with sample schedule too
    loadSampleSchedule();
  }
  refreshAllViews();
});

</script>
</body>
</html>
