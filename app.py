import pandas as pd, json, os, warnings, time, threading, pickle
from flask import Flask, render_template_string, jsonify, request, Response
from flask_httpauth import HTTPBasicAuth
warnings.filterwarnings("ignore")

DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)
auth = HTTPBasicAuth()

USER = "VIVA 1960"
PASS = "iMlWoJv1HpeD6fGCUH0UclM6Jvo=3(JK"

@auth.verify_password
def verify_pw(u, p):
    return u == USER and p == PASS

_cache = {"data": None, "ts": 0}
_CACHE_FILE = os.path.join(DIR, "dashboard_cache.json")
_DF_CACHE = os.path.join(DIR, "dataframe.pkl")
_EXCEL_PATH = os.path.join(DIR, "Viva Financial model 2026 (6).xlsx")

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Financial Dashboard</title>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<link href="/static/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="/static/all.min.css">
<script src="/static/plotly.min.js"></script>
<style>
:root {
  --sidebar-bg: #0a0a0a;
  --sidebar-width: 220px;
  --header-bg: #0f0f0f;
  --accent: #b08d57;
  --accent2: #8a6d3b;
  --card-bg: #161616;
  --card-border: #2a2a2a;
  --text-primary: #f0f0f0;
  --text-secondary: #7a7a7a;
  --green: #4caf50;
  --red: #e74c3c;
  --teal: #b08d57;
  --blue: #5c7cfa;
  --purple: #9b7ed8;
  --orange: #d4a843;
  --pink: #d4709f;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:#0a0a0a;color:var(--text-primary);overflow-x:hidden}
.header{background:var(--header-bg);color:#fff;height:100px;display:flex;align-items:center;padding:0 20px;position:fixed;top:0;left:0;right:0;z-index:1000;border-bottom:1px solid #222}
.header .brand{display:flex;align-items:center;gap:14px;width:var(--sidebar-width);font-size:17px;font-weight:700;letter-spacing:.5px;color:#fff}
.header .brand img{height:72px;width:auto;flex-shrink:0}
.header .top-right{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:12px;color:var(--text-secondary)}
.header .top-right .status-dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--green);margin-right:4px;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.header .currency-selector{display:flex;align-items:center;gap:6px;margin-left:14px;padding-left:14px;border-left:1px solid #333}
.header .currency-selector label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)}
.header .currency-selector select{background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:4px 28px 4px 10px;color:#f1f5f9;font-size:12px;font-weight:500;cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%237a8ba3' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center}
.sidebar{position:fixed;top:100px;left:0;bottom:0;width:var(--sidebar-width);background:var(--sidebar-bg);color:var(--text-secondary);overflow-y:auto;z-index:999;padding-top:0;border-right:1px solid #222}
.sidebar .nav-item{padding:10px 20px;display:flex;align-items:center;gap:11px;cursor:pointer;font-size:13px;transition:.15s;border-left:3px solid transparent}
.sidebar .nav-item:hover{background:rgba(176,141,87,0.08);color:#fff}
.sidebar .nav-item.active{background:rgba(176,141,87,0.12);color:#fff;border-left-color:var(--accent);font-weight:600}
.sidebar .nav-item i{width:18px;text-align:center;font-size:14px;color:var(--accent)}
.sidebar .nav-section{font-size:10px;text-transform:uppercase;letter-spacing:1.2px;color:#555;padding:16px 20px 5px;font-weight:600}
.main{margin-left:var(--sidebar-width);margin-top:100px;padding:16px 20px;min-height:calc(100vh - 100px)}
.page-title{font-size:18px;font-weight:600;color:#f1f5f9;margin-bottom:16px}
.kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:18px}
@media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:600px){.kpi-grid{grid-template-columns:1fr 1fr}}
.kpi-card{background:var(--card-bg);border-radius:8px;padding:14px 10px;border:1px solid var(--card-border);border-left:3px solid #333;display:flex;flex-direction:row;align-items:center;gap:10px;transition:.2s}
.kpi-card:hover{border-color:#444;transform:translateY(-1px)}
.kpi-card .kpi-icon{width:38px;height:38px;min-width:38px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;color:#fff}
.kpi-card .kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary);margin-bottom:2px}
.kpi-card .kpi-value{font-size:20px;font-weight:700;color:#f1f5f9;line-height:1.1}
.kpi-card .kpi-sub{font-size:10px;color:var(--text-secondary);margin-top:1px}
.kpi-card.kpi-gs{border-left-color:var(--green)}.kpi-card.kpi-gs .kpi-icon{background:var(--green)}
.kpi-card.kpi-ns{border-left-color:var(--blue)}.kpi-card.kpi-ns .kpi-icon{background:var(--blue)}
.kpi-card.kpi-cogs{border-left-color:var(--orange)}.kpi-card.kpi-cogs .kpi-icon{background:var(--orange)}
.kpi-card.kpi-gp{border-left-color:var(--purple)}.kpi-card.kpi-gp .kpi-icon{background:var(--purple)}
.kpi-card.kpi-exp{border-left-color:var(--red)}.kpi-card.kpi-exp .kpi-icon{background:var(--red)}
.kpi-card.kpi-ni{border-left-color:var(--teal)}.kpi-card.kpi-ni .kpi-icon{background:var(--teal)}
.kpi-card .kpi-value.green{color:var(--green)}.kpi-card .kpi-value.red{color:var(--red)}
#pnlKpiGrid{grid-template-columns:repeat(5,1fr)!important;gap:8px!important;margin-bottom:14px!important}
#pnlKpiGrid .kpi-card{padding:6px 6px!important}
#pnlKpiGrid .kpi-icon{width:26px!important;height:26px!important;min-width:26px!important;font-size:12px!important}
#pnlKpiGrid .kpi-card .kpi-value{font-size:14px!important}
#pnlKpiGrid .kpi-card .kpi-sub{font-size:8px!important}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.chart-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}
@media(max-width:1100px){.chart-grid-3{grid-template-columns:1fr 1fr}}
@media(max-width:900px){.chart-grid,.chart-grid-3{grid-template-columns:1fr}}
.chart-card{background:var(--card-bg);border-radius:8px;border:1px solid var(--card-border);overflow:hidden}
.chart-card .chart-header{display:flex;align-items:center;justify-content:space-between;padding:12px 16px 0}
.chart-card .chart-header h5{font-size:13px;font-weight:600;color:#f1f5f9;margin:0}
.chart-card .chart-body{padding:4px 4px 6px;min-height:260px}
.chart-card .pnl-body{padding:0;min-height:220px;display:flex;align-items:stretch}
.chart-card.full{grid-column:1/-1}
.pnl-card{background:var(--card-bg);border-radius:8px;border:1px solid var(--card-border);padding:0;height:100%;display:flex;flex-direction:column;width:100%}
.pnl-card .pnl-header{padding:12px 20px 0;display:flex;align-items:center;justify-content:space-between}
.pnl-card .pnl-header h5{font-size:13px;font-weight:600;color:#f1f5f9;margin:0}
.pnl-card .pnl-header .pnl-period{font-size:11px;color:var(--text-secondary)}
.pnl-card .pnl-inner{padding:0;flex:1;overflow-x:auto}
.pnl-card table{width:100%;border-collapse:collapse;font-size:13px}
.pnl-card td{padding:7px 20px;border:0}
.pnl-card td:last-child{text-align:right;font-weight:700;font-variant-numeric:tabular-nums;letter-spacing:.3px;width:170px;padding-right:24px}
.pnl-card .pnl-label{color:var(--text-secondary);padding-left:22px}
.pnl-card .pnl-label.sub{padding-left:40px;font-size:12px}
.pnl-card .pnl-label.sub2{padding-left:52px;font-size:11.5px}
.pnl-card .pnl-label.section-title{color:#f1f5f9;font-weight:600;font-size:12px;padding-left:22px;text-transform:uppercase;letter-spacing:.5px}
.pnl-card tr.section-row td{height:1px;padding:0;border:0}
.pnl-card tr.section-row td::before{content:"";display:block;width:100%;height:1px;background:rgba(255,255,255,0.06)}
.pnl-card tr.section-divider td{height:4px;padding:0;border:0}
.pnl-card tr.subtotal td{padding-top:8px}
.pnl-card tr.subtotal .pnl-label{color:#7a8ba3;font-weight:600}
.pnl-card tr.subtotal td:last-child{border-top:1px solid #333;padding-top:8px;font-weight:700;color:#f0f0f0}
.pnl-card tr.grand-total td{padding-top:12px}
.pnl-card tr.grand-total .pnl-label{color:#f1f5f9;font-size:14px;font-weight:700}
.pnl-card tr.grand-total td:last-child{border-top:3px double var(--accent);padding-top:12px;font-size:16px;font-weight:800;color:var(--accent)}
.pnl-card .neg{color:var(--red)}.pnl-card .pos{color:var(--green)}
.pnl-card tr.deduction td:last-child{color:var(--red)}
.pnl-card tr.revenue-header td{padding-top:4px}
.chart-grid+.chart-grid{margin-top:14px}
.tab-content{display:none}.tab-content.active{display:block}
body.blur-mode .main .kpi-value,
body.blur-mode .main .kpi-sub,
body.blur-mode .main td.num,
body.blur-mode .main .num,
body.blur-mode .main .plotly-num,
body.blur-mode .main .js-plotly-plot svg text{filter:blur(12px);-webkit-filter:blur(12px);user-select:none;transition:filter .3s}
body.blur-mode .main .js-plotly-plot svg g.legend text,
body.blur-mode .main .js-plotly-plot svg .legendtext,
body.blur-mode .main .js-plotly-plot svg .gtitle,
body.blur-mode .main .js-plotly-plot svg g.title text{filter:blur(0);}
body:not(.blur-mode) .main .kpi-value,
body:not(.blur-mode) .main .kpi-sub,
body:not(.blur-mode) .main td:last-child,
body:not(.blur-mode) .main .value-cell,
body:not(.blur-mode) .main .plotly-num,
body:not(.blur-mode) .main .js-plotly-plot svg text{filter:blur(0);transition:filter .3s}
.footer{text-align:center;padding:16px 0;font-size:11px;color:var(--text-secondary)}
</style>
</head>
<body>

<!-- Header -->
<div class="header">
    <div class="brand"><img src="/static/logo.png" alt="VIVA 1960">VIVA 1960 Dashboard</div>
  <div class="top-right">
    <span><span class="status-dot"></span>Live</span>
    <span id="tsDisplay"><i class="far fa-clock"></i> <span id="timestamp"></span></span>
    <div class="currency-selector">
      <label>Currency</label>
      <select id="currencySelect" onchange="changeCurrency(this.value)">
        <option value="EGP">EGP — Egyptian Pound</option>
        <option value="USD">USD — US Dollar</option>
        <option value="EUR">EUR — Euro</option>
        <option value="GBP">GBP — British Pound</option>
        <option value="CHF">CHF — Swiss Franc</option>
        <option value="KWD">KWD — Kuwaiti Dinar</option>
        <option value="BHD">BHD — Bahraini Dinar</option>
        <option value="OMR">OMR — Omani Rial</option>
        <option value="JOD">JOD — Jordanian Dinar</option>
        <option value="SAR">SAR — Saudi Riyal</option>
      </select>
    </div>
    <button onclick="refreshData()" class="btn btn-sm btn-outline-light" style="font-size:12px;padding:3px 12px"><i class="fas fa-sync-alt"></i> Refresh</button>
    <button id="blurToggle" onclick="toggleBlur()" class="btn btn-sm btn-outline-light" style="font-size:12px;padding:3px 10px" title="Toggle number visibility"><i class="fas fa-eye"></i></button>
  </div>
</div>

<!-- Sidebar -->
<div class="sidebar">
  <div class="nav-section">Overview</div>
  <div class="nav-item active" onclick="switchTab('dashboard')"><i class="fas fa-tachometer-alt"></i> Dashboard</div>
  <div class="nav-item" onclick="switchTab('pnl')"><i class="fas fa-file-invoice-dollar"></i> Actuals Vs Forecast</div>
  <div class="nav-item" onclick="switchTab('sales')"><i class="fas fa-chart-bar"></i> Sales</div>
  <div class="nav-item" onclick="switchTab('expenses')"><i class="fas fa-receipt"></i> Expenses</div>
  <div class="nav-item" onclick="switchTab('style')"><i class="fas fa-tshirt"></i> Style Analysis</div>
  <div class="nav-item" onclick="switchTab('investment')"><i class="fas fa-coins"></i> Investment Insights</div>
  <div class="nav-item" onclick="switchTab('cashflow')"><i class="fas fa-money-bill-wave"></i> Cash Flow</div>
  <div class="nav-section">Data</div>
  <div class="nav-item"><i class="fas fa-table"></i> Raw Data</div>
  <div class="nav-item" style="cursor:pointer" onclick="document.getElementById('fileInput').click()"><i class="fas fa-upload"></i> Upload Excel</div>
  <input type="file" id="fileInput" accept=".xlsx" style="display:none" onchange="uploadExcel(this)">
</div>

<!-- Main -->
<div class="main">

<div id="tab-dashboard" class="tab-content active">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
    <div class="page-title" style="margin-bottom:0"><i class="fas fa-tachometer-alt" style="color:var(--accent);margin-right:8px"></i>Dashboard</div>
    <div style="display:flex;align-items:center;gap:10px">
      <div style="display:flex;align-items:center;gap:5px;background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:5px 10px">
        <i class="fas fa-calendar" style="font-size:11px;color:var(--accent)"></i>
        <label style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">Period</label>
        <select id="dashPeriod" onchange="loadData()" style="background:transparent;border:none;color:#f1f5f9;font-size:12px;font-weight:500;cursor:pointer;appearance:none;padding-right:14px;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%237a8ba3' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E&quot;);background-repeat:no-repeat;background-position:right 2px center">
          <option value="ytd">YTD</option>
          <option value="q1">Q1</option>
          <option value="q2">Q2</option>
          <option value="q3">Q3</option>
        </select>
      </div>
      <div style="display:flex;align-items:center;gap:5px;background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:5px 10px">
        <i class="fas fa-building" style="font-size:11px;color:var(--accent)"></i>
        <label style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">B.U</label>
        <select id="dashBU" onchange="document.getElementById('dashMonth').value='all';loadData()" style="background:transparent;border:none;color:#f1f5f9;font-size:12px;font-weight:500;cursor:pointer;appearance:none;padding-right:14px;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%237a8ba3' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E&quot;);background-repeat:no-repeat;background-position:right 2px center">
          <option value="all">All</option>
          <option value="Export">Export</option>
          <option value="B2B">B2B</option>
          <option value="B2C">B2C</option>
          <option value="CM">CM</option>
        </select>
      </div>
      <div style="display:flex;align-items:center;gap:5px;background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:5px 10px">
        <i class="fas fa-calendar-day" style="font-size:11px;color:var(--accent)"></i>
        <label style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">Month</label>
        <select id="dashMonth" onchange="document.getElementById('dashPeriod').value='ytd';loadData()" style="background:transparent;border:none;color:#f1f5f9;font-size:12px;font-weight:500;cursor:pointer;appearance:none;padding-right:14px;background-image:url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%237a8ba3' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E&quot;);background-repeat:no-repeat;background-position:right 2px center">
          <option value="all">All</option>
          <option value="1">January</option>
          <option value="2">February</option>
          <option value="3">March</option>
          <option value="4">April</option>
          <option value="5">May</option>
          <option value="6">June</option>
          <option value="7">July</option>
          <option value="8">August</option>
        </select>
      </div>
    </div>
  </div>
  <div class="kpi-grid" id="kpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-university" style="color:var(--accent);margin-right:6px"></i>CBE Official Exchange Rates</h5><span style="font-size:11px;color:var(--text-secondary)">Base: EGP</span></div><div class="chart-body" id="currencyRateChart" style="min-height:280px"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#8a6d3b;margin-right:6px"></i>Monthly Gross Sales</h5></div><div class="chart-body" id="salesChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-receipt" style="color:#c0392b;margin-right:6px"></i>Monthly Expenses</h5></div><div class="chart-body" id="expChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-arrows-alt-h" style="color:#8a6d3b;margin-right:6px"></i>Monthly Net Sales vs Expenses</h5></div><div class="chart-body" id="monthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#7a6faa;margin-right:6px"></i>Expenses by Department</h5></div><div class="chart-body" id="deptChart"></div></div>
    <div class="chart-card"><div class="pnl-card"><div class="pnl-header"><h5><i class="fas fa-file-invoice-dollar" style="color:var(--accent);margin-right:6px"></i>Profit &amp; Loss Statement</h5><span class="pnl-period">Jan - Jun 2026</span></div><div class="pnl-inner" id="pnlCard"></div></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-trophy" style="color:#a07830;margin-right:6px"></i>Top 5 <span id="topSlicerLabel">Customers</span> by Sales</h5><div class="slicer-selector" style="display:flex;align-items:center;gap:6px"><label style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">View</label><select id="topSlicer" onchange="changeTopSlicer(this.value)" style="background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:4px 28px 4px 10px;color:#f1f5f9;font-size:12px;font-weight:500;cursor:pointer;appearance:none;background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238899aa' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E\");background-repeat:no-repeat;background-position:right 8px center"><option value="customers">Customers</option><option value="types">Types</option><option value="fabrics">Fabrics</option></select></div></div><div class="chart-body" id="topCustomersChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-cubes" style="color:#5a8a5e;margin-right:6px"></i>Monthly Quantity Sold</h5></div><div class="chart-body" id="qtyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-balance-scale" style="color:#a07830;margin-right:6px"></i>Monthly Net Sales vs COGS</h5></div><div class="chart-body" id="cogsChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#8a6d3b;margin-right:6px"></i>BU Sales %</h5></div><div class="chart-body" id="buSalesChart"></div></div>
  </div>
</div>

<div id="tab-pnl" class="tab-content">
  <div class="page-title"><i class="fas fa-file-invoice-dollar" style="color:var(--accent);margin-right:8px"></i>P&amp;L Actuals vs Forecast</div>
  <div class="kpi-grid" id="pnlKpiGrid" style="grid-template-columns:repeat(5,1fr)"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#8a6d3b;margin-right:6px"></i>Monthly Net Sales — Actual vs Forecast</h5></div><div class="chart-body" id="pnlSalesChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#5a8a5e;margin-right:6px"></i>Monthly Gross Profit — Actual vs Forecast</h5></div><div class="chart-body" id="pnlGpChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#6a9b6e;margin-right:6px"></i>Monthly Net Income — Actual vs Forecast</h5></div><div class="chart-body" id="pnlNiChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#c9a96e;margin-right:6px"></i>Gross Profit Margin % Trend</h5></div><div class="chart-body" id="pnlMarginChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#8a7fb8;margin-right:6px"></i>Department Expenses — Actual vs Forecast</h5></div><div class="chart-body" id="pnlDeptChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-arrows-alt" style="color:#b08d57;margin-right:6px"></i>P&amp;L Waterfall — YTD Actual vs Forecast</h5></div><div class="chart-body" id="pnlWaterfallChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#f1f5f9;margin-right:6px"></i>P&amp;L Summary — Actual vs Forecast</h5></div><div class="pnl-body" id="pnlSummaryTable"></div></div>
  </div>
</div>

<div id="tab-sales" class="tab-content">
  <div class="page-title"><i class="fas fa-chart-bar" style="color:var(--accent);margin-right:8px"></i>Sales Analytics</div>
  <div class="kpi-grid" id="salesKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#8a6d3b;margin-right:6px"></i>Monthly Gross Sales — Actual vs Budget</h5></div><div class="chart-body" id="salesGSChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-area" style="color:#5a8a5e;margin-right:6px"></i>Monthly Returns &amp; Discounts</h5></div><div class="chart-body" id="salesDedChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-cubes" style="color:#a07830;margin-right:6px"></i>Monthly Quantity Sold</h5></div><div class="chart-body" id="salesQtyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#7a6faa;margin-right:6px"></i>Sales by Business Unit</h5></div><div class="chart-body" id="salesBUChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-trophy" style="color:#c9a96e;margin-right:6px"></i>Top 5 Customers</h5></div><div class="chart-body" id="salesTopChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-list" style="color:#b08d57;margin-right:6px"></i>Top 5 Types</h5></div><div class="chart-body" id="salesTypeChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-tshirt" style="color:#5c7cfa;margin-right:6px"></i>Top 5 Fabrics</h5></div><div class="chart-body" id="salesFabricChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-percent" style="color:#6a9b6e;margin-right:6px"></i>Net Sales Conversion Rate %</h5></div><div class="chart-body" id="salesConversionChart"></div></div>
  </div>
</div>

<div id="tab-expenses" class="tab-content">
  <div class="page-title"><i class="fas fa-receipt" style="color:var(--accent);margin-right:8px"></i>Expense Analytics</div>
  <div class="kpi-grid" id="expKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#c0392b;margin-right:6px"></i>Monthly Expenses — Actual vs Budget</h5></div><div class="chart-body" id="expMonthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#8a7fb8;margin-right:6px"></i>Department Expenses — Actual vs Forecast</h5></div><div class="chart-body" id="expDeptChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#c9a96e;margin-right:6px"></i>Monthly Expense Variance</h5></div><div class="chart-body" id="expVarChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#b08d57;margin-right:6px"></i>Expense Distribution</h5></div><div class="chart-body" id="expPieChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-area" style="color:#5c7cfa;margin-right:6px"></i>Top Departments Monthly Trend</h5></div><div class="chart-body" id="expTrendChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-table" style="color:#f1f5f9;margin-right:6px"></i>Department Expense Detail — Variance Analysis</h5></div><div class="pnl-body" id="expTable"></div></div>
  </div>
</div>

<div id="tab-style" class="tab-content">
  <div class="page-title"><i class="fas fa-tshirt" style="color:var(--accent);margin-right:8px"></i>32 Degree Style Analysis — 168 Soft Bra</div>
  <div class="kpi-grid" id="styleKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-water" style="color:#b08d57;margin-right:6px"></i>Profit Waterfall — Planned to Actual</h5></div><div class="chart-body" id="styleWaterfall"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-hand-holding-usd" style="color:#6a9b6e;margin-right:6px"></i>Net Profit by Purchase Order</h5></div><div class="chart-body" id="styleNpChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-pie" style="color:#8a7fb8;margin-right:6px"></i>COGS % by Purchase Order</h5></div><div class="chart-body" id="styleCogsPie"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-th-list" style="color:#c9a96e;margin-right:6px"></i>PO-Level Cost Breakdown</h5></div><div class="pnl-body" id="styleTable"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#6a9b6e;margin-right:6px"></i>Margin Trend Across POs</h5></div><div class="chart-body" id="styleMarginChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#5c7cfa;margin-right:6px"></i>Cost Per Item Comparison</h5></div><div class="chart-body" id="styleCostBar"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#c9a96e;margin-right:6px"></i>Variance Analysis</h5></div><div class="pnl-body" id="styleVariance"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-lightbulb" style="color:#6a9b6e;margin-right:6px"></i>Recommendations</h5></div><div class="pnl-body" id="styleRecommendations"></div></div>
  </div>
</div>

<div id="tab-investment" class="tab-content">
  <div class="page-title"><i class="fas fa-coins" style="color:var(--accent);margin-right:8px"></i>Investment Insights — 2025 vs 2026 Analysis</div>
  <div class="kpi-grid" id="invKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#c9a96e;margin-right:6px"></i>Return Rate Comparison — Gold / Silver / Swiss Frank</h5></div><div class="chart-body" id="invReturnChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#8a7fb8;margin-right:6px"></i>Risk-Adjusted Returns — Sharpe / Sortino / Calmar</h5></div><div class="chart-body" id="invRiskChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-shield-alt" style="color:#c0392b;margin-right:6px"></i>Risk Metrics — MDD / Downside Deviation / Volatility</h5></div><div class="chart-body" id="invRiskMetrics"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-project-diagram" style="color:#b08d57;margin-right:6px"></i>Correlation Matrix</h5></div><div class="chart-body" id="invCorrChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#5a8a5e;margin-right:6px"></i>Portfolio Allocation</h5></div><div class="chart-body" id="invPieChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#a07830;margin-right:6px"></i>Investment Performance — Actual vs Forecast</h5></div><div class="chart-body" id="invPerfChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-lightbulb" style="color:#6a9b6e;margin-right:6px"></i>Recommendation</h5></div><div class="pnl-body" id="invRecommendation"></div></div>
  </div>
</div>

<div id="tab-cashflow" class="tab-content">
  <div class="page-title"><i class="fas fa-money-bill-wave" style="color:var(--accent);margin-right:8px"></i>Cash Flow Analysis</div>
  <div class="kpi-grid" id="cfKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-exchange-alt" style="color:#b08d57;margin-right:6px"></i>Monthly Cash In vs Out vs Net</h5></div><div class="chart-body" id="cfMonthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-building" style="color:#8a7fb8;margin-right:6px"></i>Collections by Business Unit</h5></div><div class="chart-body" id="cfBuChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-credit-card" style="color:#c9a96e;margin-right:6px"></i>Payment Status Breakdown</h5></div><div class="chart-body" id="cfPayChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#5a8a5e;margin-right:6px"></i>Spending by Department</h5></div><div class="chart-body" id="cfDeptChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-tags" style="color:#c0392b;margin-right:6px"></i>Top Spending Categories</h5></div><div class="chart-body" id="cfCatChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#b08d57;margin-right:6px"></i>Collection vs Spending by Month (Stacked)</h5></div><div class="chart-body" id="cfStackChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-clock" style="color:#a07830;margin-right:6px"></i>Aging Analysis — Receivables</h5></div><div class="chart-body" id="cfAgingChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-users" style="color:#6a9b6e;margin-right:6px"></i>Top 10 Customers</h5></div><div class="chart-body" id="cfCustChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-list-ol" style="color:#8a7fb8;margin-right:6px"></i>Top 10 Spending Categories</h5></div><div class="chart-body" id="cfCatBarChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#a07830;margin-right:6px"></i>Collections vs Forecast (Actuals)</h5></div><div class="chart-body" id="cfCollForecastChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#c0392b;margin-right:6px"></i>Spending vs Forecast (Actuals)</h5></div><div class="chart-body" id="cfSpendForecastChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-lightbulb" style="color:#6a9b6e;margin-right:6px"></i>Cash Flow Insights</h5></div><div class="pnl-body" id="cfInsights"></div></div>
  </div>
</div>

<div class="footer">VIVA 1960 Dashboard &mdash; Data refreshes every 30 seconds</div>
</div>

<script>
// Currency rates (base: EGP) - CBE Official Rates (29 Jul 2026) - Average of Buy & Sell
let CURRENCIES = {
  EGP: {rate: 1, symbol: 'EGP', locale: 'en-US'},
  USD: {rate: 0.019721, symbol: '$', locale: 'en-US'},
  EUR: {rate: 0.017317, symbol: '\u20AC', locale: 'de-DE'},
  GBP: {rate: 0.014831, symbol: '\u00A3', locale: 'en-GB'},
  CHF: {rate: 0.016161, symbol: 'CHF', locale: 'de-CH'},
  KWD: {rate: 0.006071, symbol: 'KWD', locale: 'ar-KW'},
  BHD: {rate: 0.007439, symbol: 'BHD', locale: 'ar-BH'},
  OMR: {rate: 0.007673, symbol: 'OMR', locale: 'ar-OM'},
  JOD: {rate: 0.014085, symbol: 'JOD', locale: 'ar-JO'},
  SAR: {rate: 0.074034, symbol: 'SAR', locale: 'ar-SA'},
  AED: {rate: 0.072428, symbol: 'AED', locale: 'ar-AE'}
};
let currentCurrency = 'EGP';
let _data = null;
let topSlicer = 'customers';
const AUTH = 'Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL';
function authFetch(url, opts) {
  opts = opts || {};
  opts.headers = opts.headers || {};
  opts.headers['Authorization'] = AUTH;
  return fetch(url, opts);
}

function getCurrency() {
  return CURRENCIES[currentCurrency];
}

function fmt(v, sign) {
  const c = getCurrency();
  const converted = Math.abs(v) * c.rate;
  const s = converted.toLocaleString(c.locale, {minimumFractionDigits: 2, maximumFractionDigits: 2});
  const sym = c.symbol;
  if (sign) return v < 0 ? '<span class="neg">(' + s + ' ' + sym + ')</span>' : '<span class="pos">' + s + ' ' + sym + '</span>';
  return v < 0 ? '(' + s + ' ' + sym + ')' : s + ' ' + sym;
}

function fmtRaw(v) {
  const c = getCurrency();
  const converted = Math.abs(v) * c.rate;
  return converted.toLocaleString(c.locale, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function fmtNoSymbol(v) {
  const c = getCurrency();
  const converted = Math.abs(v) * c.rate;
  return converted.toLocaleString(c.locale, {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function fmtShort(v) {
  const c = getCurrency();
  const n = Math.abs(v) * c.rate;
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M '+c.symbol;
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K '+c.symbol;
  return n.toFixed(0)+' '+c.symbol;
}

function buildKPI(label, value, icon, color, cls, sub) {
  const isNeg = value < 0;
  const c = getCurrency();
  const n = Math.abs(value) * c.rate;
  const valStr = fmtShort(value);
  return '<div class="kpi-card kpi-' + cls + '">' +
    '<div class="kpi-icon"><i class="fas fa-' + icon + '"></i></div>' +
    '<div class="kpi-label">' + label + '</div>' +
    '<div class="kpi-value ' + (isNeg ? 'red' : 'green') + '">' + (isNeg ? '-' : '') + valStr + '</div>' +
    (sub ? '<div class="kpi-sub">' + sub + '</div>' : '') +
    '</div>';
}

function changeCurrency(curr) {
  currentCurrency = curr;
  const select = document.getElementById('currencySelect');
  if (select) select.value = curr;
  renderCurrencyChart();
  loadData();
  window._pnlLoaded = false;
  window._salesLoaded = false;
  window._expLoaded = false;
  var active = document.querySelector('.tab-content.active');
  if (active) {
    var id = active.id.replace('tab-', '');
    if (id !== 'dashboard') switchTab(id);
  }
}

function renderCurrencyChart() {
  const allCurs = ['USD','EUR','GBP','CHF','KWD','BHD','OMR','JOD','SAR','AED'];
  const sel = currentCurrency;
  const vals = allCurs.map(n => {
    const rate = CURRENCIES[n] ? CURRENCIES[n].rate : 0;
    return rate > 0 ? (1 / rate) : 0;
  });
  const barColors = allCurs.map(n => n === sel ? '#b08d57' : '#333');
  Plotly.newPlot('currencyRateChart', [{
    type: 'bar', x: allCurs, y: vals,
    marker: {color: barColors, line: {color: '#b08d57', width: allCurs.map(n => n === sel ? 2 : 0)}},
    text: vals.map(v => v.toFixed(2) + ' EGP'),
    textposition: 'outside', textfont: {size: 14, color: '#ffffff'}, cliponaxis: false
  }], {
    margin: {t: 30, b: 45, l: 70, r: 25},
    paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
    font: {size: 14, color: '#a0b4c8'},
    yaxis: {title: 'EGP per 1 unit', gridcolor: 'rgba(255,255,255,0.04)', rangemode: 'tozero'},
    height: 280, hovermode: 'x unified', showlegend: false
  }, {responsive: true, displayModeBar: false});
  const ts = document.getElementById('rateTimestamp');
  if (ts) ts.textContent = 'Last updated: ' + new Date().toLocaleString();
}

function renderTopChart(d, c) {
  let items;
  let label;
  if (topSlicer === 'types') { items = d.top_types; label = 'Types'; }
  else if (topSlicer === 'fabrics') { items = d.top_fabrics; label = 'Fabrics'; }
  else { items = d.top_customers; label = 'Customers'; }
  document.getElementById('topSlicerLabel').textContent = label;
  const names = items.map(x => x.name);
  const vals = items.map(x => x.sales * c.rate);
  Plotly.newPlot('topCustomersChart', [{type:'bar', x:names, y:vals,
    marker:{color:['#b08d57','#c4a265','#d4b87a','#e0c88f','#ecdaa3']},
    text:vals.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:14,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:50,b:55,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#a0b4c8'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}, height:340,
     hovermode:'x unified', showlegend:false}, {responsive:true, displayModeBar:false});
}

function changeTopSlicer(val) {
  topSlicer = val;
  if (_data) renderTopChart(_data, getCurrency());
  setTimeout(updatePlotlyBlur, 200);
}

async function loadData() {
  const periodEl = document.getElementById('dashPeriod');
  const buEl = document.getElementById('dashBU');
  const monthEl = document.getElementById('dashMonth');
  const period = periodEl ? periodEl.value : 'ytd';
  const bu = buEl ? buEl.value : 'all';
  const month = monthEl ? monthEl.value : 'all';
  const r = await authFetch('/api/data?period=' + encodeURIComponent(period) + '&bu=' + encodeURIComponent(bu) + '&month=' + encodeURIComponent(month));
  const d = await r.json();
  _data = d;
  const now = new Date();
  document.getElementById('timestamp').textContent = now.toLocaleString();
  const T = d.totals;
  document.getElementById('kpiGrid').innerHTML =
    buildKPI('Gross Sales', T.gs, 'chart-line', '#2ecc71', 'gs', 'Total Revenue') +
    buildKPI('Net Sales', T.ns, 'shopping-cart', '#3498db', 'ns', 'After Returns &amp; Discounts') +
    buildKPI('COGS', T.cogs, 'truck', '#f39c12', 'cogs', 'Cost of Goods Sold') +
    buildKPI('Gross Profit', T.gp, 'arrow-up', '#9b59b6', 'gp') +
    buildKPI('Expenses', T.exp, 'receipt', '#e74c3c', 'exp') +
    buildKPI('Net Income', T.ni, 'wallet', '#1abc9c', 'ni');

  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const M = d.months;

  const c = getCurrency();
  const gsV = months.map((_,i) => M[i].gs * c.rate);
  Plotly.newPlot('salesChart', [{type:'bar', x:months, y:gsV, marker:{color:'#b08d57'}, text:gsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:14,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:40,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#a0b4c8'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'},
     height:320, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  const expV = months.map((_,i) => M[i].exp * c.rate);
  Plotly.newPlot('expChart', [{type:'bar', x:months, y:expV, marker:{color:'#c0392b'},     text:expV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:14,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:40,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#a0b4c8'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'},
     height:320, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  const nsV = months.map((_,i) => M[i].ns * c.rate);
  Plotly.newPlot('monthlyChart', [
    {type:'bar', name:'Net Sales', x:months, y:nsV, marker:{color:'#b08d57'}, text:nsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false},
    {type:'bar', name:'Expenses', x:months, y:expV, marker:{color:'#c0392b'}, text:expV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false}
  ], {margin:{t:50,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:14,color:'#a0b4c8'}, barmode:'group', height:330, hovermode:'x unified',
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:14,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  const depts = Object.keys(d.depts);
  const dVals = Object.values(d.depts).map(v => v * c.rate);
  Plotly.newPlot('deptChart', [{type:'bar', orientation:'h', x:dVals, y:depts,
    marker:{color:'#8a7fb8'}, text:dVals.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:15,b:25,l:160,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#a0b4c8'}, height:Math.max(300, depts.length*45), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  document.getElementById('pnlCard').innerHTML =
    '<table>' +
    '<tr class="section-divider"><td colspan="2"></td></tr>' +
    '<tr class="revenue-header"><td class="pnl-label section-title">Revenue</td><td></td></tr>' +
    '<tr><td class="pnl-label">Gross Sales</td><td class="pos num">' + fmt(T.gs) + '</td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Less: Returns</td><td class="neg num">' + fmt(-T.ret) + '</td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Less: Discounts</td><td class="neg num">' + fmt(-T.disc) + '</td></tr>' +
    '<tr class="section-row"><td colspan="2"></td></tr>' +
    '<tr class="subtotal"><td class="pnl-label">Net Sales</td><td class="num">' + fmt(T.ns) + '</td></tr>' +
    '<tr class="section-divider"><td colspan="2"></td></tr>' +
    '<tr class="revenue-header"><td class="pnl-label section-title">Cost of Goods Sold</td><td></td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Less: COGS</td><td class="neg num">' + fmt(-T.cogs) + '</td></tr>' +
    '<tr class="section-row"><td colspan="2"></td></tr>' +
    '<tr class="subtotal"><td class="pnl-label">Gross Profit</td><td class="pos num">' + fmt(T.gp) + '</td></tr>' +
    '<tr class="section-divider"><td colspan="2"></td></tr>' +
    '<tr class="revenue-header"><td class="pnl-label section-title">Operating Expenses</td><td></td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Operating Expenses</td><td class="neg num">' + fmt(-T.exp) + '</td></tr>' +
    '<tr class="section-row"><td colspan="2"></td></tr>' +
    '<tr class="grand-total"><td class="pnl-label">Net Income</td><td class="pos num">' + fmt(T.ni) + '</td></tr>' +
    '</table>';

  renderTopChart(d, c);

  // Business Units
  const busUnits = d.business_units;
  const buNames = busUnits.map(x => x.name);
  const buColors = ['#8a6d3b','#5a8a5e','#a07830','#7a6faa','#c0392b'];

  const buSalesPct = busUnits.map(x => x.sales_pct);
  Plotly.newPlot('buSalesChart', [{
    type:'pie', labels:buNames, values:buSalesPct,
    text:buSalesPct.map(v => v.toFixed(1) + '%'), textinfo:'label+percent', textfont:{size:20,color:'#ffffff'},
    marker:{colors:buColors.slice(0,buNames.length),line:{color:'#fff',width:3}},
    hovertemplate:'%{label}<br>%{value:.1f}%<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:320, showlegend:true,
     legend:{orientation:'h',y:-0.12,font:{size:14,color:'#a0b4c8'}}},
    {responsive:true, displayModeBar:false});

  // Monthly Quantity Sold
  const qtyVals = months.map((_,i) => M[i].qty);
  const qtyLabels = qtyVals.map(v => v > 0 ? v.toLocaleString() : '');
  Plotly.newPlot('qtyChart', [{type:'bar', x:months, y:qtyVals,
    marker:{color:'#5a8a5e'}, text:qtyLabels, textposition:'outside', textfont:{size:16,color:'#ffffff',family:'Arial'}, cliponaxis:false}],
    {margin:{t:35,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:13,color:'#a0b4c8'}, yaxis:{rangemode:'tozero',gridcolor:'rgba(255,255,255,0.04)',automargin:true},
     height:340, hovermode:'x unified', showlegend:false, bargap:0.3},
    {responsive:true, displayModeBar:false});

  // Monthly Net Sales vs COGS
  const ncogsV = months.map((_,i) => M[i].cogs * c.rate);
  const nnsV = months.map((_,i) => M[i].ns * c.rate);
  const nsLabels = nnsV.map(v => v > 0 ? fmtNoSymbol(v)+' '+c.symbol : '');
  const cogsLabels = ncogsV.map(v => v > 0 ? fmtNoSymbol(v)+' '+c.symbol : '');
  Plotly.newPlot('cogsChart', [
    {type:'bar', name:'Net Sales', x:months, y:nnsV, marker:{color:'#b08d57'}, text:nsLabels, textposition:'outside', textfont:{size:14,color:'#ffffff',family:'Arial'}, cliponaxis:false},
    {type:'bar', name:'COGS', x:months, y:ncogsV, marker:{color:'#a07830'}, text:cogsLabels, textposition:'outside', textfont:{size:14,color:'#ffffff',family:'Arial'}, cliponaxis:false}
  ], {margin:{t:45,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:13,color:'#a0b4c8'}, barmode:'group', height:380, hovermode:'x unified', bargap:0.25, bargroupgap:0.1,
      legend:{orientation:'h',y:1.12,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)',automargin:true}}, {responsive:true, displayModeBar:false});
  renderCurrencyChart();
  setTimeout(updatePlotlyBlur, 100);
}

async function loadStyleAnalysis() {
  try {
    const r = await authFetch('/api/style_analysis');
    if (!r.ok) return;
    const s = await r.json();
    if (s.error) return;
    const c = getCurrency();
    const cb = s.cost_breakdown;
    const pf = s.profitability;
    const wf = s.waterfall;
    const po = s.po_names;
    const fmtS = v => { const a=Math.abs(v); if(a>=1e6)return (v/1e6).toFixed(1)+'M'; if(a>=1e3)return (v/1e3).toFixed(1)+'K'; return v.toFixed(2); };
    const fmtPct = v => (v*100).toFixed(1)+'%';
    const totalRev = cb.revenue.total;
    const totalCogs = cb.total_cogs.total;
    const totalGP = cb.total_gross_profit.total;
    const totalNP = pf.total_net_profit.values.reduce((a,b)=>a+b,0);
    const avgGPM = cb.gross_profit_margin.total;
    const avgNPM = totalNP / (totalRev || 1);
    const kpis = [
      {label:'Total Revenue', value:fmtS(totalRev*c.rate)+' '+c.symbol, color:'#6a9b6e'},
      {label:'Total COGS', value:fmtS(totalCogs*c.rate)+' '+c.symbol, color:'#c9a96e'},
      {label:'Gross Profit', value:fmtS(totalGP*c.rate)+' '+c.symbol, color:'#8a7fb8'},
      {label:'Net Profit', value:fmtS(totalNP*c.rate)+' '+c.symbol, color:'#6a9b6e'},
      {label:'Avg GP Margin', value:fmtPct(avgGPM), color:'#6a9b6e'},
      {label:'Avg NP Margin', value:fmtPct(avgNPM), color:'#6a9b6e'}
    ];
    document.getElementById('styleKpiGrid').innerHTML = kpis.map(k =>
      '<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas fa-chart-line" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+'">'+k.value+'</div></div>'
    ).join('');
    // Waterfall
    const wfLabels = ['Planned Profit','Labor\\nAdjustment','Overtime','Packing\\nSavings','Clearance\\nSavings','Finance\\nSavings','SG&A\\nSavings','Actual Profit'];
    const wfVals = [wf.planned_profit, wf.labor_cost_adjustment, wf.overtime, wf.saving_packing, wf.saving_clearance, wf.saving_finance, wf.saving_sga, wf.actual_profit];
    Plotly.newPlot('styleWaterfall', [{
      type:'waterfall', orientation:'v',
      x: wfLabels, y: wfVals,
      connector:{line:{color:'#475569',width:1,dash:'dot'}},
      decreasing:{marker:{color:'#c0392b'}},
      increasing:{marker:{color:'#6a9b6e'}},
      totals:{marker:{color:'#b08d57'}},
      text: wfVals.map(v => fmtS(v*c.rate)),
      textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:40,b:70,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:400, showlegend:false,
      yaxis:{ticksuffix:' '+c.symbol, gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Net Profit by PO (bar chart)
    const npVals = pf.total_net_profit.values.map(v => v * c.rate);
    const npColors = npVals.map(v => v >= 0 ? '#6a9b6e' : '#c0392b');
    Plotly.newPlot('styleNpChart', [{
      type:'bar', x:po, y:npVals, marker:{color:npColors},
      text:npVals.map(v=>fmtS(v)+' '+c.symbol), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:40,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:350, showlegend:false, xaxis:{tickangle:-30},
      yaxis:{ticksuffix:' '+c.symbol, gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // COGS % pie
    Plotly.newPlot('styleCogsPie', [{
      type:'pie', labels:po, values:cb.total_cogs.values,
      text:cb.total_cogs.values.map(v=>fmtS(v*c.rate)), textinfo:'label+percent',
      textfont:{size:14,color:'#ffffff'},
      marker:{colors:['#b08d57','#c9a96e','#6a9b6e','#8a7fb8','#c0392b'],line:{color:'#fff',width:2}},
      hovertemplate:'%{label}<br>'+c.symbol+': %{text}<br>%{percent}<extra></extra>'
    }], {
      margin:{t:10,b:10,l:10,r:10}, paper_bgcolor:'rgba(0,0,0,0)', height:350,
      showlegend:true, legend:{orientation:'h',y:-0.1,font:{size:12,color:'#a0b4c8'}}
    }, {responsive:true, displayModeBar:false});
    // Table
    const tableRows = [
      {label:'Order Quantity', key:'order_quantity', fmt:v=>v.toLocaleString()},
      {label:'Selling Price', key:'selling_price', fmt:v=>v.toFixed(2)},
      {label:'Revenue', key:'revenue', fmt:v=>fmtS(v*c.rate)+' '+c.symbol},
      {label:'Labor Cost/Item', key:'labor_cost_per_item', fmt:v=>v.toFixed(2)},
      {label:'Overtime/Item', key:'overtime_per_item', fmt:v=>v.toFixed(2)},
      {label:'Packing/Item', key:'packing_per_item', fmt:v=>v.toFixed(2)},
      {label:'COGS %', key:'cogs_pct', fmt:v=>fmtPct(v)},
      {label:'GP Margin', key:'gross_profit_margin', fmt:v=>fmtPct(v)},
      {label:'Production Time (days)', key:'production_time', fmt:v=>v.toLocaleString()}
    ];
    let tbl = '<table style="width:100%;border-collapse:collapse;font-size:13px;color:#cbd5e1"><thead><tr style="border-bottom:1px solid #334155"><th style="padding:8px 12px;text-align:left;color:#94a3b8;font-weight:500">Metric</th>';
    po.forEach(p => { tbl += '<th style="padding:8px 12px;text-align:right;color:#94a3b8;font-weight:500">'+p+'</th>'; });
    tbl += '<th style="padding:8px 12px;text-align:right;color:#f1f5f9;font-weight:600">Total</th><th style="padding:8px 12px;text-align:right;color:#b08d57;font-weight:600">Forecast</th><th style="padding:8px 12px;text-align:right;color:#94a3b8;font-weight:600">Variance</th></tr></thead><tbody>';
    tableRows.forEach((row, i) => {
      const d = cb[row.key];
      const bg = i%2===0 ? 'rgba(255,255,255,0.02)' : 'transparent';
      tbl += '<tr style="border-bottom:1px solid #1e293b;background:'+bg+'"><td style="padding:8px 12px;color:#f1f5f9;font-weight:500">'+row.label+'</td>';
      d.values.forEach(v => { tbl += '<td style="padding:8px 12px;text-align:right">'+row.fmt(v)+'</td>'; });
      tbl += '<td style="padding:8px 12px;text-align:right;color:#f1f5f9;font-weight:600">'+row.fmt(d.total)+'</td>';
      tbl += '<td style="padding:8px 12px;text-align:right;color:#b08d57">'+row.fmt(d.forecasted||0)+'</td>';
      const vv = d.variance||0;
      tbl += '<td style="padding:8px 12px;text-align:right;color:'+(vv>0?'#6a9b6e':vv<0?'#c0392b':'#94a3b8')+'">'+row.fmt(vv)+'</td></tr>';
    });
    tbl += '</tbody></table>';
    document.getElementById('styleTable').innerHTML = tbl;
    // Margin trend
    const gpMargins = cb.gross_profit_margin.values.map(v => v * 100);
    const npMargins = pf.net_profit_margin.values.map(v => v * 100);
    Plotly.newPlot('styleMarginChart', [
      {type:'scatter', mode:'lines+markers', name:'GP Margin', x:po, y:gpMargins, line:{color:'#8a7fb8',width:3}, marker:{size:10,color:'#8a7fb8'}, text:gpMargins.map(v=>v.toFixed(1)+'%'), textposition:'top center', textfont:{size:13,color:'#8a7fb8'}},
      {type:'scatter', mode:'lines+markers', name:'NP Margin', x:po, y:npMargins, line:{color:'#6a9b6e',width:3}, marker:{size:10,color:'#6a9b6e'}, text:npMargins.map(v=>v.toFixed(1)+'%'), textposition:'top center', textfont:{size:13,color:'#6a9b6e'}}
    ], {
      margin:{t:50,b:50,l:60,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:350, showlegend:true, xaxis:{tickangle:-30},
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:'%', gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Cost per item stacked bar
    const laborV = cb.labor_cost_per_item.values;
    const otV = cb.overtime_per_item.values;
    const packV = cb.packing_per_item.values;
    Plotly.newPlot('styleCostBar', [
      {type:'bar', name:'Labor', x:po, y:laborV, marker:{color:'#b08d57'}},
      {type:'bar', name:'Overtime', x:po, y:otV, marker:{color:'#c9a96e'}},
      {type:'bar', name:'Packing', x:po, y:packV, marker:{color:'#6a9b6e'}}
    ], {
      margin:{t:40,b:45,l:60,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'stack', height:350, xaxis:{tickangle:-30},
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Variance text
    if (s.variance_detail) {
      document.getElementById('styleVariance').innerHTML = '<div style="padding:18px 22px"><div style="font-size:13px;font-weight:600;color:#c9a96e;margin-bottom:12px">'+s.variance_title+'</div><div style="font-size:13px;color:#94a3b8;line-height:1.8;white-space:pre-line">'+s.variance_detail+'</div></div>';
    }
    // Recommendations
    if (s.recommendations) {
      document.getElementById('styleRecommendations').innerHTML = '<div style="padding:18px 22px"><div style="font-size:13px;font-weight:600;color:#6a9b6e;margin-bottom:12px">'+s.rec_title+'</div><div style="font-size:13px;color:#94a3b8;line-height:1.8;white-space:pre-line">'+s.recommendations+'</div></div>';
    }
    setTimeout(updatePlotlyBlur, 100);
  } catch(e) { console.error('Style analysis error:', e); }
}

async function loadInvestmentData() {
  try {
    const r = await authFetch('/api/investment');
    if (!r.ok) return;
    const s = await r.json();
    if (s.error) { console.error(s.error); return; }
    const c = getCurrency();
    const pct = v => (v * 100).toFixed(1) + '%';
    const fmtM = v => { const a = Math.abs(v); if (a >= 1e6) return (v/1e6).toFixed(2) + 'M'; if (a >= 1e3) return (v/1e3).toFixed(1) + 'K'; return v.toFixed(2); };
    const fmtR = v => (v * 100).toFixed(2) + '%';
    const p = s.portfolio;
    const inv = s.investment;
    const kpis = [
      {label:'Portfolio Return 2025', value:fmtR(p.return_rate.y2025), color:p.return_rate.y2025>=0?'#6a9b6e':'#c0392b'},
      {label:'Portfolio Return 2026A', value:fmtR(p.return_rate.y2026a), color:p.return_rate.y2026a>=0?'#6a9b6e':'#c0392b'},
      {label:'Portfolio Return 2026F', value:fmtR(p.return_rate.y2026f), color:p.return_rate.y2026f>=0?'#6a9b6e':'#c0392b'},
      {label:'Sharpe Ratio 2026A', value:p.sharpe_ratio.y2026a.toFixed(3), color:'#8a7fb8'},
      {label:'Max Drawdown 2026A', value:pct(p.mdd.y2026a), color:'#c0392b'},
      {label:'Initial Investment', value:fmtM(inv.initial)+' '+c.symbol, color:'#b08d57'}
    ];
    document.getElementById('invKpiGrid').innerHTML = kpis.map(k =>
      '<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas fa-chart-line" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+';font-size:18px">'+k.value+'</div></div>'
    ).join('');
    const assets = ['Gold','Silver','Swiss Frank'];
    const metrics = ['return_rate','avg_monthly_return','avg_monthly_real_return'];
    const metricNames = ['Annual Return','Avg Monthly Return','Avg Monthly Real Return'];
    const yrs = ['y2025','y2026a','y2026f'];
    const yrLabels = ['2025','2026 Actual','2026 Forecast'];
    const traces_return = [];
    const colors = {'Gold':'#c9a96e','Silver':'#94a3b8','Swiss Frank':'#b08d57'};
    assets.forEach((asset, ai) => {
      const data = asset === 'Gold' ? s.gold : asset === 'Silver' ? s.silver : s.swiss;
      traces_return.push({
        type:'bar', name:asset+' Return',
        x:yrLabels,
        y:metrics.map(m => data[m].y2025*100),
        marker:{color:colors[asset],opacity:0.9},
        text:metrics.map(m => pct(data[m].y2025)),
        textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false,
        offset: ai * 0.25
      });
    });
    Plotly.newPlot('invReturnChart', traces_return, {
      margin:{t:50,b:70,l:70,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'group', height:400,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:'%',gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    const riskMetrics = ['sharpe_ratio','sortino','calmar'];
    const riskNames = ['Sharpe Ratio','Sortino Ratio','Calmar Ratio'];
    const traces_risk = [];
    assets.forEach((asset, ai) => {
      const data = asset === 'Gold' ? s.gold : asset === 'Silver' ? s.silver : s.swiss;
      traces_risk.push({
        type:'scatter', mode:'lines+markers', name:asset,
        x:yrLabels, y:riskMetrics.map(m => data[m].y2026a),
        line:{color:colors[asset],width:3}, marker:{size:10}
      });
    });
    Plotly.newPlot('invRiskChart', traces_risk, {
      margin:{t:50,b:60,l:60,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:380,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    const riskLabels2 = ['Max Drawdown','Downside Deviation','Monthly Volatility'];
    const riskKeys = ['mdd','downside_dev','monthly_risk'];
    const traces_rm = [];
    assets.forEach((asset) => {
      const data = asset === 'Gold' ? s.gold : asset === 'Silver' ? s.silver : s.swiss;
      traces_rm.push({
        type:'bar', name:asset,
        x:riskLabels2,
        y:riskKeys.map(k => Math.abs(data[k].y2026a)*100),
        marker:{color:colors[asset]},
        text:riskKeys.map(k => pct(Math.abs(data[k].y2026a))),
        textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false
      });
    });
    Plotly.newPlot('invRiskMetrics', traces_rm, {
      margin:{t:50,b:60,l:60,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'group', height:380,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:'%',gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    const corrLabels = ['Gold','Silver','Swiss Frank'];
    const corrMatrix = [
      [1, s.correlations.gold_silver, s.correlations.gold_swiss],
      [s.correlations.gold_silver, 1, s.correlations.silver_swiss],
      [s.correlations.gold_swiss, s.correlations.silver_swiss, 1]
    ];
    const corrText = corrMatrix.map(row => row.map(v => v.toFixed(3)));
    Plotly.newPlot('invCorrChart', [{
      type:'heatmap', z:corrMatrix, x:corrLabels, y:corrLabels,
      text:corrText, texttemplate:'%{text}', textfont:{size:16,color:'#ffffff'},
      colorscale:[[0,'#c0392b'],[0.5,'#1e293b'],[1,'#6a9b6e']],
      zmin:-1, zmax:1,
      hovertemplate:'%{x} vs %{y}<br>Correlation: %{z:.3f}<extra></extra>'
    }], {
      margin:{t:30,b:50,l:80,r:30}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:350,
      xaxis:{side:'bottom'}, yaxis:{autorange:'reversed'}
    }, {responsive:true, displayModeBar:false});
    const pw = s.portfolio.weights || {gold:0.3,silver:0.2,swiss:0.5};
    Plotly.newPlot('invPieChart', [{
      type:'pie', labels:['Gold','Silver','Swiss Frank'],
      values:[pw.gold*100,pw.silver*100,pw.swiss*100],
      text:[(pw.gold*100).toFixed(0)+'%',(pw.silver*100).toFixed(0)+'%',(pw.swiss*100).toFixed(0)+'%'],
      textinfo:'label+percent', textfont:{size:18,color:'#ffffff'},
      marker:{colors:['#c9a96e','#94a3b8','#b08d57'],line:{color:'#fff',width:3}},
      hovertemplate:'%{label}<br>%{value:.0f}%<extra></extra>'
    }], {
      margin:{t:10,b:10,l:10,r:10}, paper_bgcolor:'rgba(0,0,0,0)', height:350,
      showlegend:true, legend:{orientation:'h',y:-0.1,font:{size:13,color:'#a0b4c8'}}
    }, {responsive:true, displayModeBar:false});
    const perfLabels = ['Annual Return','Return After Loss','Present Value'];
    const perfActual = [inv.annual_return.actual, inv.return_after_loss.actual, inv.pv.actual];
    const perfForecast = [inv.annual_return.forecast, inv.return_after_loss.forecast, inv.pv.forecast];
    Plotly.newPlot('invPerfChart', [
      {type:'bar', name:'2026 Actual', x:perfLabels, y:perfActual, marker:{color:'#b08d57'},
       text:perfActual.map(v=>fmtM(v*c.rate)+' '+c.symbol), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false},
      {type:'bar', name:'2026 Forecast', x:perfLabels, y:perfForecast, marker:{color:'#c9a96e'},
       text:perfForecast.map(v=>fmtM(v*c.rate)+' '+c.symbol), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false}
    ], {
      margin:{t:50,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'group', height:380,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    if (s.recommendation) {
      document.getElementById('invRecommendation').innerHTML = '<div style="padding:18px 22px;font-size:13px;color:#94a3b8;line-height:1.8;white-space:pre-line">'+s.recommendation+'</div>';
    }
    setTimeout(updatePlotlyBlur, 100);
  } catch(e) { console.error('Investment data error:', e); }
}

async function loadCashflowData() {
  try {
    const r = await authFetch('/api/cashflow');
    if (!r.ok) return;
    const s = await r.json();
    if (s.error) { console.error(s.error); return; }
    const c = getCurrency();
    const fmtM = v => { const a = Math.abs(v*c.rate); if (a >= 1e6) return (v*c.rate/1e6).toFixed(2) + 'M'; if (a >= 1e3) return (v*c.rate/1e3).toFixed(1) + 'K'; return (v*c.rate).toFixed(2); };
    const fmtFull = v => (v * c.rate).toLocaleString(c.locale, {minimumFractionDigits: 0, maximumFractionDigits: 0}) + ' ' + c.symbol;
    const monthNames = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const months = Object.keys(s.monthly_collections).sort((a,b)=>Number(a)-Number(b));
    // KPIs
    const netCF = s.total_collections - s.total_spending;
    const kpis = [
      {label:'Total Collections', value:fmtFull(s.total_collections), color:'#6a9b6e', icon:'fa-arrow-down'},
      {label:'Total Spending', value:fmtFull(s.total_spending), color:'#c0392b', icon:'fa-arrow-up'},
      {label:'Net Cash Flow', value:fmtFull(netCF), color: netCF>=0?'#b08d57':'#c9a96e', icon:'fa-balance-scale'},
      {label:'Days Receivable', value:s.days_receivable.toFixed(1) + ' days', color:'#8a7fb8', icon:'fa-clock'},
      {label:'Days Payable', value:s.days_payable.toFixed(1) + ' days', color:'#a07830', icon:'fa-clock'}
    ];
    document.getElementById('cfKpiGrid').innerHTML = kpis.map(k =>
      '<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas '+k.icon+'" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+';font-size:18px">'+k.value+'</div></div>'
    ).join('');
    // Monthly Cash In vs Out vs Net
    const inVals = months.map(m => s.monthly_collections[m] * c.rate);
    const outVals = months.map(m => -s.monthly_spending[m] * c.rate);
    const netVals = months.map(m => s.monthly_net[m] * c.rate);
    Plotly.newPlot('cfMonthlyChart', [
      {type:'bar', name:'Collections', x:months.map(m=>monthNames[m]), y:inVals, marker:{color:'#6a9b6e'}, text:inVals.map(v=>fmtFull(v/c.rate)), textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false},
      {type:'bar', name:'Spending', x:months.map(m=>monthNames[m]), y:outVals, marker:{color:'#c0392b'}, text:outVals.map(v=>'(-'+fmtFull(-v/c.rate)+')'), textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false},
      {type:'scatter', mode:'lines+markers', name:'Net Cash Flow', x:months.map(m=>monthNames[m]), y:netVals, line:{color:'#b08d57',width:3}, marker:{size:10}}
    ], {
      margin:{t:50,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'relative', height:400,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Collections by BU (from dashboard - authoritative)
    const buKeys = Object.keys(s.bu_collections_actual).filter(k=>k!=='TOTAL');
    Plotly.newPlot('cfBuChart', [{
      type:'pie', labels:buKeys, values:buKeys.map(k=>s.bu_collections_actual[k]*c.rate),
      textinfo:'label+percent', textfont:{size:14,color:'#ffffff'},
      marker:{colors:['#b08d57','#8a7fb8','#6a9b6e','#c9a96e'],line:{color:'#fff',width:3}},
      hovertemplate:'%{label}<br>'+c.symbol+' %{value:,.0f}<br>%{percent}<extra></extra>'
    }], {
      margin:{t:10,b:10,l:10,r:10}, paper_bgcolor:'rgba(0,0,0,0)', height:350,
      showlegend:true, legend:{orientation:'h',y:-0.1,font:{size:13,color:'#a0b4c8'}}
    }, {responsive:true, displayModeBar:false});
    // Payment status
    const payKeys = Object.keys(s.payment_status);
    const payColors = {'Bank Transfer':'#b08d57','Cash':'#6a9b6e','Cheque':'#c9a96e'};
    Plotly.newPlot('cfPayChart', [{
      type:'doughnut', labels:payKeys, values:payKeys.map(k=>s.payment_status[k]*c.rate),
      textinfo:'label+percent', textfont:{size:14,color:'#ffffff'},
      marker:{colors:payKeys.map(k=>payColors[k]||'#8a7fb8'),line:{color:'#fff',width:3}},
      hole:0.5, hovertemplate:'%{label}<br>'+c.symbol+' %{value:,.0f}<br>%{percent}<extra></extra>'
    }], {
      margin:{t:10,b:10,l:10,r:10}, paper_bgcolor:'rgba(0,0,0,0)', height:350,
      showlegend:true, legend:{orientation:'h',y:-0.1,font:{size:13,color:'#a0b4c8'}}
    }, {responsive:true, displayModeBar:false});
    // Spending by department (from dashboard - authoritative)
    const deptKeys = Object.keys(s.dept_spending_actual).filter(k=>k!=='TOTAL');
    const deptColors = {'Production':'#c9a96e','G&A':'#b08d57','Assets':'#8a7fb8','Financing Expenses':'#c0392b','S&M':'#6a9b6e','R&D':'#a07830'};
    Plotly.newPlot('cfDeptChart', [{
      type:'bar', x:deptKeys, y:deptKeys.map(k=>s.dept_spending_actual[k]*c.rate),
      marker:{color:deptKeys.map(k=>deptColors[k]||'#6a9b6e')},
      text:deptKeys.map(k=>fmtFull(s.dept_spending_actual[k])),
      textposition:'outside', textfont:{size:12,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:30,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:350,
      xaxis:{tickangle:-30}, yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Top spending categories
    const catKeys = Object.keys(s.category_spending).slice().reverse();
    Plotly.newPlot('cfCatChart', [{
      type:'bar', y:catKeys, x:catKeys.map(k=>s.category_spending[k]*c.rate),
      orientation:'h', marker:{color:'#c0392b'},
      text:catKeys.map(k=>fmtFull(s.category_spending[k])),
      textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:10,b:40,l:120,r:60}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:380,
      xaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Stacked collection vs spending by month
    const buList = [...new Set(Object.values(s.monthly_bu).flatMap(v=>Object.keys(v)))];
    const deptList = [...new Set(Object.values(s.monthly_dept).flatMap(v=>Object.keys(v)))];
    const stackTraces = [];
    const buColors2 = {'CM':'#6a9b6e','Export':'#b08d57','B2B':'#8a7fb8','B2C':'#c9a96e'};
    buList.forEach(bu => {
      stackTraces.push({
        type:'bar', name:'In: '+bu,
        x:months.map(m=>monthNames[m]),
        y:months.map(m=>(s.monthly_bu[m]&&s.monthly_bu[m][bu]||0)*c.rate),
        marker:{color:buColors2[bu]||'#666'}
      });
    });
    const deptColors2 = {'Production':'#c0392b','G&A':'#a07830','Assets':'#94a3b8'};
    deptList.forEach(dept => {
      stackTraces.push({
        type:'bar', name:'Out: '+dept,
        x:months.map(m=>monthNames[m]),
        y:months.map(m=>-(s.monthly_dept[m]&&s.monthly_dept[m][dept]||0)*c.rate),
        marker:{color:deptColors2[dept]||'#666'}
      });
    });
    Plotly.newPlot('cfStackChart', stackTraces, {
      margin:{t:50,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'relative', height:420,
      legend:{orientation:'h',y:1.2,x:.5,xanchor:'center',font:{size:12,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Aging chart
    const agingKeys = Object.keys(s.aging);
    if (agingKeys.length) {
      Plotly.newPlot('cfAgingChart', [{
        type:'bar', x:agingKeys, y:agingKeys.map(k=>s.aging[k].amount*c.rate),
        marker:{color:agingKeys.map((_,i)=>['#6a9b6e','#c9a96e','#c0392b'][i]||'#666')},
        text:agingKeys.map(k=>s.aging[k].aging+'<br>'+fmtFull(s.aging[k].amount)),
        textposition:'outside', textfont:{size:12,color:'#ffffff'}, cliponaxis:false
      }], {
        margin:{t:30,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
        font:{size:12,color:'#a0b4c8'}, height:350,
        yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
      }, {responsive:true, displayModeBar:false});
    }
    // Top 10 customers
    const custKeys = Object.keys(s.customer_collections).slice().reverse();
    Plotly.newPlot('cfCustChart', [{
      type:'bar', y:custKeys, x:custKeys.map(k=>s.customer_collections[k]*c.rate),
      orientation:'h', marker:{color:'#6a9b6e'},
      text:custKeys.map(k=>fmtFull(s.customer_collections[k])),
      textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:10,b:40,l:120,r:60}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:400,
      xaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Top 10 spending categories bar
    const catKeysBar = Object.keys(s.category_spending).sort((a,b)=>s.category_spending[b]-s.category_spending[a]);
    Plotly.newPlot('cfCatBarChart', [{
      type:'bar', x:catKeysBar, y:catKeysBar.map(k=>s.category_spending[k]*c.rate),
      marker:{color:'#8a7fb8'},
      text:catKeysBar.map(k=>fmtFull(s.category_spending[k])),
      textposition:'outside', textfont:{size:11,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:30,b:80,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, height:400,
      xaxis:{tickangle:-40}, yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Collections vs Forecast
    const cfBUs = Object.keys(s.bu_collections_forecast).filter(k=>k!=='TOTAL');
    Plotly.newPlot('cfCollForecastChart', [
      {type:'bar', name:'Actual', x:cfBUs, y:cfBUs.map(bu=>(s.bu_collections_actual[bu]||0)*c.rate),
       marker:{color:'#b08d57'}, text:cfBUs.map(bu=>fmtFull(s.bu_collections_actual[bu]||0)), textposition:'outside', textfont:{size:12,color:'#ffffff'}, cliponaxis:false},
      {type:'bar', name:'Forecast', x:cfBUs, y:cfBUs.map(bu=>s.bu_collections_forecast[bu]*c.rate),
       marker:{color:'#c9a96e'}, text:cfBUs.map(bu=>fmtFull(s.bu_collections_forecast[bu])), textposition:'outside', textfont:{size:12,color:'#ffffff'}, cliponaxis:false}
    ], {
      margin:{t:50,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'group', height:380,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Spending vs Forecast
    const cfDepts = Object.keys(s.dept_spending_forecast).filter(k=>k!=='TOTAL');
    Plotly.newPlot('cfSpendForecastChart', [
      {type:'bar', name:'Actual', x:cfDepts, y:cfDepts.map(d=>(s.dept_spending_actual[d]||0)*c.rate),
       marker:{color:'#c0392b'}, text:cfDepts.map(d=>fmtFull(s.dept_spending_actual[d]||0)), textposition:'outside', textfont:{size:12,color:'#ffffff'}, cliponaxis:false},
      {type:'bar', name:'Forecast', x:cfDepts, y:cfDepts.map(d=>s.dept_spending_forecast[d]*c.rate),
       marker:{color:'#c9a96e'}, text:cfDepts.map(d=>fmtFull(s.dept_spending_forecast[d])), textposition:'outside', textfont:{size:12,color:'#ffffff'}, cliponaxis:false}
    ], {
      margin:{t:50,b:60,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#a0b4c8'}, barmode:'group', height:380,
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)'}
    }, {responsive:true, displayModeBar:false});
    // Insights
    const collVar = s.total_forecast_coll ? ((s.total_collections - s.total_forecast_coll)/s.total_forecast_coll*100) : 0;
    const spendVar = s.total_forecast_spend ? ((s.total_spending - s.total_forecast_spend)/s.total_forecast_spend*100) : 0;
    const insights = [
      '<div style="margin-bottom:14px"><b style="color:#b08d57;font-size:14px">Cash Flow Summary</b></div>',
      '<div style="margin-bottom:10px"><b style="color:#6a9b6e">Days Receivable:</b> ' + s.days_receivable.toFixed(1) + ' days — ' + (s.days_receivable <= 30 ? '<span style="color:#6a9b6e">Healthy</span>' : s.days_receivable <= 60 ? '<span style="color:#c9a96e">Moderate — consider tighter credit terms</span>' : '<span style="color:#c0392b">High — collections are slow, review receivables process</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#a07830">Days Payable:</b> ' + s.days_payable.toFixed(1) + ' days — ' + (s.days_payable >= 30 ? '<span style="color:#6a9b6e">Good leverage of payment terms</span>' : '<span style="color:#c9a96e">Paying too fast — consider optimizing cash retention</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#8a7fb8">Collection vs Forecast:</b> ' + (collVar>=0?'+':'') + collVar.toFixed(1) + '% variance — ' + (collVar>=0?'<span style="color:#6a9b6e">Above target</span>':'<span style="color:#c0392b">Below target — review sales pipeline</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#c0392b">Spending vs Forecast:</b> ' + (spendVar>=0?'+':'') + spendVar.toFixed(1) + '% variance — ' + (spendVar<=0?'<span style="color:#6a9b6e">Under budget</span>':'<span style="color:#c0392b">Over budget — review departmental spend</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#b08d57">Net Cash Position:</b> ' + fmtFull(netCF) + ' — ' + (netCF>=0?'<span style="color:#6a9b6e">Positive cash generation</span>':'<span style="color:#c0392b">Cash burn — monitor closely</span>') + '</div>'
    ].join('');
    document.getElementById('cfInsights').innerHTML = '<div style="padding:18px 22px;font-size:13px;color:#94a3b8;line-height:1.8">' + insights + '</div>';
    setTimeout(updatePlotlyBlur, 100);
  } catch(e) { console.error('Cash flow data error:', e); }
}

async function uploadExcel(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const r = await authFetch('/upload', {method:'POST', body:fd});
  const j = await r.json();
  if (j.ok) { loadData(); alert('Excel updated successfully!'); }
  else { alert('Error: ' + j.error); }
  input.value = '';
}

async function refreshData() {
  await authFetch('/api/refresh');
  loadData();
}

// ---- Blur Toggle ----
let _blurred = localStorage.getItem('dashBlur') === 'true';
if (_blurred) document.body.classList.add('blur-mode');
function toggleBlur() {
  _blurred = !_blurred;
  document.body.classList.toggle('blur-mode', _blurred);
  localStorage.setItem('dashBlur', _blurred);
  document.querySelector('#blurToggle i').className = _blurred ? 'fas fa-eye-slash' : 'fas fa-eye';
  updatePlotlyBlur();
}
function updatePlotlyBlur() {
  const blurred = document.body.classList.contains('blur-mode');
  var found = 0, marked = 0;
  document.querySelectorAll('.js-plotly-plot svg text').forEach(el => {
    let t = el.textContent.trim();
    if (!t) return;
    found++;
    t = t.replace(/\s+(EGP|CHF|KWD|BHD|OMR|JOD|SAR|EUR|GBP|AED|USD|\$|€|£)$/, '');
    const isNum = /^[\+\-]?\$?[\d,]+\.?[\d]*[KMB]?%?$/.test(t) ||
                  /^\(\$?[\d,]+\.?[\d]*\)$/.test(t);
    if (isNum) marked++;
    el.classList.toggle('plotly-num', blurred && isNum);
  });
  document.querySelector('#blurToggle i').className = blurred
    ? (marked > 0 ? 'fas fa-eye-slash' : 'fas fa-exclamation-triangle')
    : 'fas fa-eye';
}
// Poll every second to catch any new chart renders
setInterval(function(){ if (document.body.classList.contains('blur-mode')) updatePlotlyBlur(); }, 1000);

document.querySelector('#blurToggle i').className = _blurred ? 'fas fa-eye-slash' : 'fas fa-eye';

loadData();

// ---- Tab Switching ----
function switchTab(tab) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  document.querySelectorAll('.sidebar .nav-item').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.sidebar .nav-item').forEach(el => {
    if (el.textContent.trim().toLowerCase().includes(tab === 'pnl' ? 'p&l' : tab)) {
      el.classList.add('active');
    }
  });
  if (tab === 'pnl' && !window._pnlLoaded) { window._pnlLoaded = true; loadPnlData(); }
  if (tab === 'sales' && !window._salesLoaded) { window._salesLoaded = true; loadSalesData(); }
  if (tab === 'expenses' && !window._expLoaded) { window._expLoaded = true; loadExpensesData(); }
  if (tab === 'style' && !window._styleLoaded) { window._styleLoaded = true; loadStyleAnalysis(); }
  if (tab === 'investment' && !window._invLoaded) { window._invLoaded = true; loadInvestmentData(); }
  if (tab === 'cashflow' && !window._cfLoaded) { window._cfLoaded = true; loadCashflowData(); }
  setTimeout(updatePlotlyBlur, 300);
}

// ---- P&L Actual vs Forecast ----
const pnlMonths = ['Jan','Feb','Mar','Apr','May','Jun'];

function pnlFmt(v) { const c=getCurrency(); const n=Math.abs(v);
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M '+c.symbol;
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K '+c.symbol;
  return n.toFixed(0)+' '+c.symbol;
}

function pnlFmtShort(v) { return pnlFmt(v); }

async function loadPnlData() {
  const r = await authFetch('/api/pnl_forecast');
  const d = await r.json();
  const ytd = d.ytd;
  const monthly = d.monthly;

  // KPI cards
  const pnlKpis = [
    {label:'Net Sales', a:ytd.net_sales.actual, f:ytd.net_sales.forecast, icon:'shopping-cart', cls:'ns'},
    {label:'COGS', a:ytd.cogs.actual, f:ytd.cogs.forecast, icon:'truck', cls:'cogs'},
    {label:'Gross Profit', a:ytd.gross_profit.actual, f:ytd.gross_profit.forecast, icon:'arrow-up', cls:'gp'},
    {label:'Expenses', a:ytd.expenses.actual, f:ytd.expenses.forecast, icon:'receipt', cls:'exp'},
    {label:'Net Income', a:ytd.net_income.actual, f:ytd.net_income.forecast, icon:'wallet', cls:'ni'}
  ];
  const cr = getCurrency();
  document.getElementById('pnlKpiGrid').innerHTML = pnlKpis.map(k => {
    const a = k.a * cr.rate, f = k.f * cr.rate;
    const varAmt = a - f;
    const varPct = k.f !== 0 ? ((varAmt / f) * 100).toFixed(1) : 'N/A';
    const isPos = varAmt >= 0;
    const isNegKpi = (k.label === 'COGS' || k.label === 'Expenses');
    const good = isNegKpi ? !isPos : isPos;
    return '<div class="kpi-card kpi-' + k.cls + '">' +
      '<div class="kpi-icon"><i class="fas fa-' + k.icon + '"></i></div>' +
      '<div class="kpi-label">' + k.label + '</div>' +
      '<div class="kpi-value">' + pnlFmt(a) + '</div>' +
      '<div class="kpi-sub">F: ' + pnlFmt(f) + '</div>' +
      '<div class="kpi-sub" style="font-weight:600;color:' + (good ? '#4caf50' : '#e74c3c') + '">' +
      (isPos ? '+' : '') + pnlFmtShort(varAmt) + ' (' + varPct + '%)</div></div>';
  }).join('');

  const colAct = '#b08d57';
  const colFct = '#94a3b8';

  // 1. Net Sales chart
  const nsAct = monthly.map(m => m.net_sales.actual * cr.rate);
  const nsFct = monthly.map(m => m.net_sales.forecast * cr.rate);
  buildGroupedBar('pnlSalesChart', pnlMonths, [
    {name:'Actual', y:nsAct, color:colAct},
    {name:'Forecast', y:nsFct, color:colFct}
  ], 'Net Sales');

  // 2. Gross Profit chart
  const gpAct = monthly.map(m => m.gross_profit.actual * cr.rate);
  const gpFct = monthly.map(m => m.gross_profit.forecast * cr.rate);
  buildGroupedBar('pnlGpChart', pnlMonths, [
    {name:'Actual', y:gpAct, color:'#5a8a5e'},
    {name:'Forecast', y:gpFct, color:colFct}
  ], 'Gross Profit');

  // 3. Net Income chart
  const niAct = monthly.map(m => m.net_income.actual * cr.rate);
  const niFct = monthly.map(m => m.net_income.forecast * cr.rate);
  buildGroupedBar('pnlNiChart', pnlMonths, [
    {name:'Actual', y:niAct, color:'#6a9b6e'},
    {name:'Forecast', y:niFct, color:colFct}
  ], 'Net Income');

  // 4. GP Margin trend
  const gpMarginAct = monthly.map(m => m.net_sales.actual > 0 ? (m.gross_profit.actual / m.net_sales.actual * 100) : 0);
  const gpMarginFct = monthly.map(m => m.net_sales.forecast > 0 ? (m.gross_profit.forecast / m.net_sales.forecast * 100) : 0);
  Plotly.newPlot('pnlMarginChart', [
    {type:'scatter', mode:'lines+markers', name:'Actual Margin', x:pnlMonths, y:gpMarginAct, line:{color:'#6a9b6e', width:3}, marker:{size:8, color:'#6a9b6e'}},
    {type:'scatter', mode:'lines+markers', name:'Forecast Margin', x:pnlMonths, y:gpMarginFct, line:{color:'#94a3b8', width:3, dash:'dot'}, marker:{size:8, color:'#94a3b8'}}
  ], {margin:{t:15,b:35,l:50,r:15}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:12},
      yaxis:{ticksuffix:'%', gridcolor:'rgba(255,255,255,0.04)', rangemode:'tozero'}, height:260,
      hovermode:'x unified', legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:11,color:'#a0b4c8'}}},
    {responsive:true, displayModeBar:false});

  // 5. Department Expenses
  const depts = Object.keys(d.dept_expenses);
  const deptAct = depts.map(k => d.dept_expenses[k].actual * cr.rate);
  const deptFct = depts.map(k => d.dept_expenses[k].forecast * cr.rate);
  Plotly.newPlot('pnlDeptChart', [
    {type:'bar', orientation:'h', name:'Actual', x:deptAct, y:depts, marker:{color:'#8a7fb8'}, text:deptAct.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}},
    {type:'bar', orientation:'h', name:'Forecast', x:deptFct, y:depts, marker:{color:'#c4b5fd'}, text:deptFct.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:13,color:'#94a3b8'}}
  ], {margin:{t:15,b:25,l:160,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
      barmode:'group', height:Math.max(250, depts.length*50), hovermode:'y unified',
      legend:{orientation:'h',y:1.05,x:.5,xanchor:'center',font:{size:12,color:'#a0b4c8'}},
      xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 6. P&L Waterfall
  const pnlLabels = ['Net Sales', 'COGS', 'Gross Profit', 'Expenses', 'Net Income'];
  const pnlActVals = [ytd.net_sales.actual, -ytd.cogs.actual, ytd.gross_profit.actual, -ytd.expenses.actual, ytd.net_income.actual].map(v=>v*cr.rate);
  const pnlFctVals = [ytd.net_sales.forecast, -ytd.cogs.forecast, ytd.gross_profit.forecast, -ytd.expenses.forecast, ytd.net_income.forecast].map(v=>v*cr.rate);
  const waterfallMeasures = ['relative', 'relative', 'total', 'relative', 'total'];
  Plotly.newPlot('pnlWaterfallChart', [
    {type:'waterfall', name:'Actual', x:pnlLabels, y:pnlActVals, measure:waterfallMeasures,
     decreasing:{marker:{color:'#c0392b'}}, increasing:{marker:{color:'#6a9b6e'}}, totals:{marker:{color:'#b08d57'}},
     connector:{line:{color:'rgba(255,255,255,0.2)', width:2}},
     text:pnlActVals.map(v => pnlFmtShort(v)), textposition:'outside', textfont:{size:11,color:'#e2e8f0'}},
    {type:'waterfall', name:'Forecast', x:pnlLabels, y:pnlFctVals, measure:waterfallMeasures,
     decreasing:{marker:{color:'#fca5a5'}}, increasing:{marker:{color:'#86efac'}}, totals:{marker:{color:'#d4b87a'}},
     connector:{line:{color:'rgba(255,255,255,0.1)', width:1}},
     text:pnlFctVals.map(v => pnlFmtShort(v)), textposition:'outside', textfont:{size:11,color:'#94a3b8'},
     opacity:0.7}
  ], {margin:{t:50,b:40,l:55,r:15}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:12},
      height:280, hovermode:'x unified',
      legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:11,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 7. P&L Summary Table
  const pnlLines = [
    {label:'Net Sales', a:ytd.net_sales.actual*cr.rate, f:ytd.net_sales.forecast*cr.rate},
    {label:'COGS', a:-ytd.cogs.actual*cr.rate, f:ytd.cogs.forecast*cr.rate, neg:true},
    {label:'Gross Profit', a:ytd.gross_profit.actual*cr.rate, f:ytd.gross_profit.forecast*cr.rate, bold:true},
    {label:'Expenses', a:-ytd.expenses.actual*cr.rate, f:ytd.expenses.forecast*cr.rate, neg:true},
    {label:'Net Income', a:ytd.net_income.actual*cr.rate, f:ytd.net_income.forecast*cr.rate, bold:true}
  ];
  document.getElementById('pnlSummaryTable').innerHTML =
    '<div style="padding:16px 24px;width:100%"><table style="width:100%;border-collapse:collapse;font-size:14px">' +
    '<tr style="color:var(--text-secondary);font-size:11px;text-transform:uppercase;letter-spacing:.5px">' +
    '<td style="padding:8px 16px">Line Item</td><td style="padding:8px 16px;text-align:right">Actual</td>' +
    '<td style="padding:8px 16px;text-align:right">Forecast</td><td style="padding:8px 16px;text-align:right">Variance</td>' +
    '<td style="padding:8px 16px;text-align:right">Var %</td></tr>' +
    pnlLines.map(l => {
      const varAmt = l.a - (l.neg ? -l.f : l.f);
      const varPct = l.f !== 0 ? ((l.a - (l.neg ? -l.f : l.f)) / (l.neg ? -l.f : l.f) * 100).toFixed(1) : 'N/A';
      const isGood = l.neg ? (varAmt >= 0) : (varAmt >= 0);
      return '<tr' + (l.bold ? ' style="border-top:2px solid #475569;font-weight:700"' : '') + '>' +
        '<td style="padding:10px 16px;color:' + (l.bold ? '#f1f5f9' : 'var(--text-secondary)') + '">' + l.label + '</td>' +
        '<td class="num" style="padding:10px 16px;text-align:right;font-weight:600">' + pnlFmt(l.a) + '</td>' +
        '<td class="num" style="padding:10px 16px;text-align:right;color:#94a3b8">' + pnlFmt(l.neg ? -l.f : l.f) + '</td>' +
        '<td class="num" style="padding:10px 16px;text-align:right;color:' + (isGood ? 'var(--green)' : 'var(--red)') + ';font-weight:600">' +
        (varAmt >= 0 ? '+' : '') + pnlFmt(varAmt) + '</td>' +
        '<td class="num" style="padding:10px 16px;text-align:right;color:' + (isGood ? 'var(--green)' : 'var(--red)') + '">' +
        (varPct !== 'N/A' ? (varAmt >= 0 ? '+' : '') + varPct + '%' : 'N/A') + '</td></tr>';
    }).join('') + '</table></div>';
  setTimeout(updatePlotlyBlur, 100);
}

function buildGroupedBar(divId, labels, series, title) {
  const c = getCurrency();
  const traces = series.map(s => ({
    type:'bar', name:s.name, x:labels, y:s.y, marker:{color:s.color},
    text:s.y.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:12, color:'#e2e8f0'}, cliponaxis:false
  }));
  Plotly.newPlot(divId, traces, {
    margin:{t:50,b:35,l:55,r:15}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
    font:{color:'#a0b4c8',size:12}, barmode:'group', height:280, hovermode:'x unified',
    legend:{orientation:'h',y:1.12,x:.5,xanchor:'center',font:{size:11,color:'#a0b4c8'}},
    yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.04)',automargin:true}
  }, {responsive:true, displayModeBar:false});
}

// ---- Sales Tab ----
function salesFmt(v) { const c=getCurrency(); const n=Math.abs(v);
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M '+c.symbol;
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K '+c.symbol;
  return n.toFixed(0)+' '+c.symbol;
}
function salesFmtShort(v) { return salesFmt(v); }

async function loadSalesData() {
  const [d, sf] = await Promise.all([
    authFetch('/api/data').then(r=>r.json()),
    authFetch('/api/sales_forecast').then(r=>r.json())
  ]);
  const T = d.totals;
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const M = d.months;
  const actMonths = months.slice(0,6);
  const cr = getCurrency();

  // KPI cards
  const nsPct = T.gs > 0 ? ((T.ns / T.gs) * 100).toFixed(1) : '0';
  document.getElementById('salesKpiGrid').innerHTML =
    '<div class="kpi-card kpi-gs"><div class="kpi-icon"><i class="fas fa-chart-line"></i></div><div class="kpi-label">Gross Sales</div><div class="kpi-value">'+salesFmt(T.gs*cr.rate)+'</div><div class="kpi-sub">Budget: '+salesFmt(sf.ytd.gross_sales.forecast*cr.rate)+'</div></div>' +
    '<div class="kpi-card kpi-ns"><div class="kpi-icon"><i class="fas fa-shopping-cart"></i></div><div class="kpi-label">Net Sales</div><div class="kpi-value">'+salesFmt(T.ns*cr.rate)+'</div><div class="kpi-sub">'+nsPct+'% of Gross</div></div>' +
    '<div class="kpi-card kpi-cogs"><div class="kpi-icon"><i class="fas fa-undo"></i></div><div class="kpi-label">Returns</div><div class="kpi-value" style="color:var(--red)">'+salesFmt(T.ret*cr.rate)+'</div><div class="kpi-sub">'+(T.gs>0?((T.ret/T.gs)*100).toFixed(2):'0')+'% of Sales</div></div>' +
    '<div class="kpi-card kpi-gp"><div class="kpi-icon"><i class="fas fa-percent"></i></div><div class="kpi-label">Discounts</div><div class="kpi-value" style="color:var(--red)">'+salesFmt(T.disc*cr.rate)+'</div><div class="kpi-sub">'+(T.gs>0?((T.disc/T.gs)*100).toFixed(2):'0')+'% of Sales</div></div>' +
    '<div class="kpi-card kpi-ni"><div class="kpi-icon"><i class="fas fa-box"></i></div><div class="kpi-label">Qty Sold</div><div class="kpi-value">'+d.months.slice(0,6).reduce((s,m)=>s+m.qty,0).toLocaleString()+'</div><div class="kpi-sub">Jan-Jun total</div></div>';

  // 1. Gross Sales Actual vs Budget
  const gsAct = actMonths.map((_,i)=>M[i].gs*cr.rate);
  const gsFct = sf.monthly.map(m=>m.gross_sales.forecast*cr.rate);
  buildGroupedBar('salesGSChart', actMonths, [
    {name:'Actual', y:gsAct, color:'#8a6d3b'},
    {name:'Budget', y:gsFct, color:'#94a3b8'}
  ]);

  // 2. Returns & Discounts
  const retVals = sf.monthly.map(m=>m.returns * cr.rate);
  const discVals = sf.monthly.map(m=>m.discounts * cr.rate);
  Plotly.newPlot('salesDedChart', [
    {type:'bar', name:'Returns', x:actMonths, y:retVals, marker:{color:'#c0392b'}, text:retVals.map(v=>salesFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#e2e8f0'}},
    {type:'bar', name:'Discounts', x:actMonths, y:discVals, marker:{color:'#c9a96e'}, text:discVals.map(v=>salesFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#e2e8f0'}}
  ], {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
      barmode:'group', height:300, hovermode:'x unified',
      legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#a0b4c8'}},
      yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 3. Quantity Sold
  const qtyV = actMonths.map((_,i)=>M[i].qty);
  Plotly.newPlot('salesQtyChart', [{type:'bar', x:actMonths, y:qtyV, marker:{color:'#a07830'}, text:qtyV.map(v=>v.toLocaleString()), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}, cliponaxis:false}],
    {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
     yaxis:{rangemode:'tozero',gridcolor:'rgba(255,255,255,0.04)'}, height:300, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  // 4. Sales by BU (pie)
  const bu = d.business_units;
  const buColors = ['#8a6d3b','#5a8a5e','#a07830','#7a6faa'];
  Plotly.newPlot('salesBUChart', [{
    type:'pie', labels:bu.map(x=>x.name), values:bu.map(x=>x.sales_pct),
    text:bu.map(x=>x.sales_pct.toFixed(1)+'%'), textinfo:'label+percent', textfont:{size:15,color:'#ffffff'},
    marker:{colors:buColors.slice(0,bu.length),line:{color:'#fff',width:2}},
    hovertemplate:'%{label}<br>%{value:.1f}%<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:290, showlegend:true,
     legend:{orientation:'h',y:-0.08,font:{size:12,color:'#a0b4c8'}}},
    {responsive:true, displayModeBar:false});

  // 5. Top Customers
  const tc = d.top_customers;
  Plotly.newPlot('salesTopChart', [{type:'bar', orientation:'h', x:tc.map(x=>x.sales*cr.rate), y:tc.map(x=>x.name),
    marker:{color:['#b08d57','#c4a265','#d4b87a','#e0c88f','#ecdaa3']},
    text:tc.map(x=>salesFmtShort(x.sales*cr.rate)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}}],
    {margin:{t:10,b:20,l:100,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
     height:Math.max(200, tc.length*40), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 6. Top Types
  const tt = d.top_types;
  Plotly.newPlot('salesTypeChart', [{type:'bar', orientation:'h', x:tt.map(x=>x.sales*cr.rate), y:tt.map(x=>x.name),
    marker:{color:'#b08d57'}, text:tt.map(x=>salesFmtShort(x.sales*cr.rate)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}}],
    {margin:{t:10,b:20,l:120,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
     height:Math.max(200, tt.length*40), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 7. Top Fabrics
  const tf = d.top_fabrics;
  Plotly.newPlot('salesFabricChart', [{type:'bar', orientation:'h', x:tf.map(x=>x.sales*cr.rate), y:tf.map(x=>x.name),
    marker:{color:'#5c7cfa'}, text:tf.map(x=>salesFmtShort(x.sales*cr.rate)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}}],
    {margin:{t:10,b:20,l:120,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
     height:Math.max(200, tf.length*40), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 8. Net Sales Conversion Rate %
  const convAct = actMonths.map((_,i)=>M[i].gs > 0 ? ((M[i].ns / M[i].gs) * 100) : 0);
  const convFct = sf.monthly.map(m=>m.gross_sales.forecast > 0 ? ((m.net_sales / m.gross_sales.forecast) * 100) : 0);
  Plotly.newPlot('salesConversionChart', [
    {type:'scatter', mode:'lines+markers', name:'Actual', x:actMonths, y:convAct, line:{color:'#6a9b6e', width:3}, marker:{size:8, color:'#6a9b6e'}},
    {type:'scatter', mode:'lines+markers', name:'Budget', x:actMonths, y:convFct, line:{color:'#94a3b8', width:3, dash:'dot'}, marker:{size:8, color:'#94a3b8'}}
  ], {margin:{t:15,b:40,l:55,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
      yaxis:{ticksuffix:'%', gridcolor:'rgba(255,255,255,0.04)', rangemode:'tozero'}, height:300,
      hovermode:'x unified', legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#a0b4c8'}}},
    {responsive:true, displayModeBar:false});
  setTimeout(updatePlotlyBlur, 100);
}

// ---- Expenses Tab ----
function expFmt(v) { const c=getCurrency(); const n=Math.abs(v);
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M '+c.symbol;
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K '+c.symbol;
  return n.toFixed(0)+' '+c.symbol;
}
function expFmtShort(v) { return expFmt(v); }

async function loadExpensesData() {
  const p = await authFetch('/api/pnl_forecast').then(r=>r.json());
  const d = await authFetch('/api/data').then(r=>r.json());
  const months = ['Jan','Feb','Mar','Apr','May','Jun'];
  const ytd = p.ytd.expenses;
  const depts = Object.keys(p.dept_expenses);
  const monthly = p.monthly;
  const cr = getCurrency();

  // KPI cards
  const topDept = depts.reduce((a,b)=>p.dept_expenses[a].actual > p.dept_expenses[b].actual ? a : b);
  const avgMonthly = ytd.actual * cr.rate / 6;
  const budgetUtil = ytd.forecast > 0 ? ((ytd.actual / ytd.forecast) * 100).toFixed(1) : 'N/A';
  document.getElementById('expKpiGrid').innerHTML =
    '<div class="kpi-card kpi-exp"><div class="kpi-icon"><i class="fas fa-receipt"></i></div><div class="kpi-label">Total Expenses</div><div class="kpi-value">'+expFmt(ytd.actual*cr.rate)+'</div><div class="kpi-sub">Budget: '+expFmt(ytd.forecast*cr.rate)+'</div></div>' +
    '<div class="kpi-card kpi-ni"><div class="kpi-icon"><i class="fas fa-building"></i></div><div class="kpi-label">Largest Dept</div><div class="kpi-value" style="font-size:16px">'+topDept+'</div><div class="kpi-sub">'+expFmt(p.dept_expenses[topDept].actual*cr.rate)+'</div></div>' +
    '<div class="kpi-card kpi-gs"><div class="kpi-icon"><i class="fas fa-calendar"></i></div><div class="kpi-label">Monthly Avg</div><div class="kpi-value">'+expFmt(avgMonthly)+'</div><div class="kpi-sub">Jan-Jun average</div></div>' +
    '<div class="kpi-card kpi-gp"><div class="kpi-icon"><i class="fas fa-percent"></i></div><div class="kpi-label">Budget Util.</div><div class="kpi-value" style="color:'+(ytd.actual<=ytd.forecast?'var(--green)':'var(--red)')+'">'+budgetUtil+'%</div><div class="kpi-sub">of budget used</div></div>';

  // 1. Monthly Expenses Actual vs Budget
  const expAct = monthly.map(m=>m.expenses.actual*cr.rate);
  const expFct = monthly.map(m=>m.expenses.forecast*cr.rate);
  buildGroupedBar('expMonthlyChart', months, [
    {name:'Actual', y:expAct, color:'#c0392b'},
    {name:'Budget', y:expFct, color:'#94a3b8'}
  ]);

  // 2. Department Expenses
  const deptAct = depts.map(k=>p.dept_expenses[k].actual*cr.rate);
  const deptFct = depts.map(k=>p.dept_expenses[k].forecast*cr.rate);
  Plotly.newPlot('expDeptChart', [
    {type:'bar', orientation:'h', name:'Actual', x:deptAct, y:depts, marker:{color:'#8a7fb8'}, text:deptAct.map(v=>expFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#e2e8f0'}},
    {type:'bar', orientation:'h', name:'Forecast', x:deptFct, y:depts, marker:{color:'#c4b5fd'}, text:deptFct.map(v=>expFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#94a3b8'}}
  ], {margin:{t:15,b:25,l:165,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
      barmode:'group', height:Math.max(250, depts.length*50), hovermode:'y unified',
      legend:{orientation:'h',y:1.05,x:.5,xanchor:'center',font:{size:12,color:'#a0b4c8'}},
      xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}}, {responsive:true, displayModeBar:false});

  // 3. Monthly Variance (Actual - Forecast)
  const varVals = monthly.map(m=>(m.expenses.actual - m.expenses.forecast)*cr.rate);
  const varColors = varVals.map(v=>v <= 0 ? '#6a9b6e' : '#c0392b');
  Plotly.newPlot('expVarChart', [{type:'bar', x:months, y:varVals, marker:{color:varColors},
    text:varVals.map(v=>(v<=0?'':'+')+expFmtShort(v)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}, cliponaxis:false}],
    {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
     yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}, height:300, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  // 4. Expense Distribution Pie
  const pieColors = ['#c0392b','#c9a96e','#8a7fb8','#b08d57','#5c7cfa','#6a9b6e','#d4709f'];
  Plotly.newPlot('expPieChart', [{
    type:'pie', labels:depts, values:deptAct,
    text:deptAct.map(v=>expFmtShort(v)), textinfo:'label+percent', textfont:{size:13,color:'#ffffff'},
    marker:{colors:pieColors.slice(0,depts.length),line:{color:'#fff',width:2}},
    hovertemplate:'%{label}<br>%{value:,.0f} '+cr.symbol+'<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:300, showlegend:true,
     legend:{orientation:'h',y:-0.1,font:{size:11,color:'#a0b4c8'}}},
    {responsive:true, displayModeBar:false});

  // 5. Top 3 Departments Monthly Trend
  const top3 = depts.sort((a,b)=>p.dept_expenses[b].actual - p.dept_expenses[a].actual).slice(0,3);
  const trendColors = ['#c0392b','#c9a96e','#8a7fb8'];
  const trendTraces = top3.map((d,i)=>({
    type:'scatter', mode:'lines+markers', name:d,
    x:months, y:p.dept_expenses[d].monthly.map(m=>m.actual*cr.rate),
    line:{color:trendColors[i], width:3}, marker:{size:7, color:trendColors[i]}
  }));
  Plotly.newPlot('expTrendChart', trendTraces,
    {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#a0b4c8',size:13},
     yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.04)'}, height:300, hovermode:'x unified',
     legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:11,color:'#a0b4c8'}}},
    {responsive:true, displayModeBar:false});

  // 6. Department Detail Table
  const sortedDepts = depts.sort((a,b)=>p.dept_expenses[b].actual - p.dept_expenses[a].actual);
  document.getElementById('expTable').innerHTML =
    '<div style="padding:12px 20px;width:100%"><table style="width:100%;border-collapse:collapse;font-size:13px">' +
    '<tr style="color:var(--text-secondary);font-size:10px;text-transform:uppercase;letter-spacing:.5px">' +
    '<td style="padding:6px 14px">Department</td><td style="padding:6px 14px;text-align:right">Actual</td>' +
    '<td style="padding:6px 14px;text-align:right">Budget</td><td style="padding:6px 14px;text-align:right">Variance</td>' +
    '<td style="padding:6px 14px;text-align:right">Var %</td><td style="padding:6px 14px;text-align:right">% of Total</td></tr>' +
    sortedDepts.map(k => {
      const a = p.dept_expenses[k].actual*cr.rate, f = p.dept_expenses[k].forecast*cr.rate, varAmt = a - f, varPct = f>0?((a-f)/f*100).toFixed(1):'N/A', pct = ytd.actual>0?(a/(ytd.actual*cr.rate)*100).toFixed(1):'0';
      const good = varAmt <= 0;
      return '<tr><td style="padding:8px 14px;color:var(--text-secondary)">' + k + '</td>' +
        '<td class="num" style="padding:8px 14px;text-align:right;font-weight:600">' + expFmt(a) + '</td>' +
        '<td class="num" style="padding:8px 14px;text-align:right;color:#94a3b8">' + expFmt(f) + '</td>' +
        '<td class="num" style="padding:8px 14px;text-align:right;font-weight:600;color:' + (good?'var(--green)':'var(--red)') + '">' + (varAmt>0?'+':'') + expFmt(varAmt) + '</td>' +
        '<td class="num" style="padding:8px 14px;text-align:right;color:' + (good?'var(--green)':'var(--red)') + '">' + (varPct!=='N/A'?(varAmt>0?'+':'')+varPct+'%':'N/A') + '</td>' +
        '<td class="num" style="padding:8px 14px;text-align:right;color:#94a3b8">' + pct + '%</td></tr>';
    }).join('') +
    '<tr style="border-top:2px solid #475569"><td style="padding:8px 14px;font-weight:700;color:#f1f5f9">Total</td>' +
    '<td class="num" style="padding:8px 14px;text-align:right;font-weight:700">' + expFmt(ytd.actual*cr.rate) + '</td>' +
    '<td class="num" style="padding:8px 14px;text-align:right;color:#94a3b8">' + expFmt(ytd.forecast*cr.rate) + '</td>' +
    '<td class="num" style="padding:8px 14px;text-align:right;font-weight:700;color:' + (ytd.actual<=ytd.forecast?'var(--green)':'var(--red)') + '">' + (ytd.actual>ytd.forecast?'+':'') + expFmt((ytd.actual-ytd.forecast)*cr.rate) + '</td>' +
    '<td class="num" style="padding:8px 14px;text-align:right;color:' + (ytd.actual<=ytd.forecast?'var(--green)':'var(--red)') + '">' + (ytd.forecast>0?((ytd.actual-ytd.forecast)/ytd.forecast*100).toFixed(1):'N/A') + '%</td>' +
    '<td class="num" style="padding:8px 14px;text-align:right;color:#94a3b8">100%</td></tr></table></div>';
  setTimeout(updatePlotlyBlur, 100);
}
</script>
</body>
</html>'''

def _load_excel():
    mtime = os.path.getmtime(_EXCEL_PATH)
    if os.path.exists(_DF_CACHE):
        try:
            with open(_DF_CACHE, "rb") as f:
                cached = pickle.load(f)
            if cached.get("_mtime") == mtime:
                return cached["sales_raw"], cached["exp_raw"]
        except: pass
    sales_raw = pd.read_excel(_EXCEL_PATH, sheet_name="Sales Raw Data 2026", header=1)
    exp_raw = pd.read_excel(_EXCEL_PATH, sheet_name="Expenses Raw Data 2026", header=1)
    try:
        with open(_DF_CACHE, "wb") as f:
            pickle.dump({"_mtime": mtime, "sales_raw": sales_raw, "exp_raw": exp_raw}, f)
    except: pass
    return sales_raw, exp_raw

def get_data(period="ytd", bu="all", month="all"):
    sales_raw, exp_raw = _load_excel()
    # Apply BU filter
    if bu and bu != "all":
        bu_upper = bu.strip().title()
        sales_raw = sales_raw[sales_raw["Business Unit"].str.strip().str.title() == bu_upper]
    for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
        sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
    sales_raw["Discount Total"] = sales_raw["Sales Amount"] * sales_raw["Discount"]
    sales_raw["Net Sales"] = sales_raw["Sales Amount"] - sales_raw["Return"].abs() - sales_raw["Discount Total"]
    sales_raw["Month_excel"] = pd.to_numeric(sales_raw["Month"], errors="coerce")
    sales_raw["Date_num"] = pd.to_numeric(sales_raw["Date"], errors="coerce")
    sales_raw["Date_dt"] = pd.to_datetime(sales_raw["Date"], format="%d/%m/%Y", errors="coerce")
    m1 = sales_raw["Date_dt"].isna() & sales_raw["Date_num"].notna()
    sales_raw.loc[m1, "Date_dt"] = pd.to_datetime(sales_raw.loc[m1, "Date_num"], origin="1899-12-30", unit="D")
    m2 = sales_raw["Date_dt"].isna()
    sales_raw.loc[m2, "Date_dt"] = pd.to_datetime(sales_raw.loc[m2, "Date"], format="%Y/%d/%m", errors="coerce")
    sales_raw["Month_parsed"] = sales_raw["Date_dt"].dt.month
    sales_raw["Month"] = sales_raw["Month_excel"].fillna(sales_raw["Month_parsed"])
    sales_raw = sales_raw[sales_raw["Month"].notna() & sales_raw["Month"].between(1, 12)]
    sales_raw["Month"] = sales_raw["Month"].astype(int)
    # Apply period or month filter
    month_ranges = {"ytd": range(1, 13), "q1": range(1, 4), "q2": range(4, 7), "q3": range(7, 10)}
    if month and month != "all":
        sales_raw = sales_raw[sales_raw["Month"] == int(month)]
    elif period and period != "ytd" and period in month_ranges:
        valid_months = list(month_ranges[period])
        sales_raw = sales_raw[sales_raw["Month"].isin(valid_months)]
    exp_raw["amount"] = pd.to_numeric(exp_raw["amount"], errors="coerce").fillna(0)
    exp_raw["Department"] = exp_raw["Department"].str.strip().str.title()
    exp_raw["Month"] = pd.to_numeric(exp_raw["Month"], errors="coerce")
    exp_raw = exp_raw[exp_raw["Month"].notna() & exp_raw["Month"].between(1, 12)]
    exp_raw["Month"] = exp_raw["Month"].astype(int)
    # Apply period or month filter to expenses too
    if month and month != "all":
        exp_raw = exp_raw[exp_raw["Month"] == int(month)]
    elif period and period != "ytd" and period in month_ranges:
        valid_months = list(month_ranges[period])
        exp_raw = exp_raw[exp_raw["Month"].isin(valid_months)]
    SALES_STORES_MONTHLY = 10000.0
    gs = sales_raw.groupby("Month")["Sales Amount"].sum()
    ret_raw = sales_raw["Return"].abs().groupby(sales_raw["Month"]).sum()
    disc_raw = sales_raw["Discount Total"].groupby(sales_raw["Month"]).sum()
    ss_mask = exp_raw["Department"] == "Sales Stores"
    exp_no_ss = exp_raw[~ss_mask].copy()
    exp = exp_no_ss.groupby("Month")["amount"].sum()
    filtered_months = sorted([m for m in gs.index if m in range(1, 13)])
    if not filtered_months:
        filtered_months = list(range(1, 7))
    num_filtered = len(filtered_months)
    for m in filtered_months:
        exp[m] = exp.get(m, 0) + SALES_STORES_MONTHLY
    dept_exp = exp_no_ss.groupby("Department")["amount"].sum()
    dept_exp["Sales Stores"] = SALES_STORES_MONTHLY * num_filtered
    total_gs = float(gs.sum())
    total_disc = float(disc_raw.sum())
    total_ret = float(sales_raw["Return"].abs().sum())
    total_exp = float(exp.sum())
    total_cogs = 31223509.0 * (total_gs / 105263158.0) if total_gs else 0
    total_ns = total_gs - total_ret - total_disc
    gp = total_ns - total_cogs
    ni = gp - total_exp
    raw_disc_total = float(disc_raw.sum())
    raw_exp_total = float(exp.sum())
    monthly_qty_raw = sales_raw.groupby("Month")["QTY"].sum()
    month_end = max(filtered_months) if filtered_months else 6
    months_list = []
    for m in range(1, 13):
        if m in filtered_months:
            m_gs = float(gs.get(m, 0))
            m_ret = float(ret_raw.get(m, 0))
            m_disc = float(disc_raw.get(m, 0))
            m_exp = float(exp.get(m, 0))
            m_cogs = total_cogs * (m_gs / total_gs) if total_gs > 0 else 0
            m_qty = int(monthly_qty_raw.get(m, 0))
        else:
            m_gs = 0; m_ret = 0; m_disc = 0; m_exp = 0; m_cogs = 0; m_qty = 0
        months_list.append({"gs": m_gs, "ns": m_gs - m_ret - m_disc, "exp": m_exp, "cogs": m_cogs, "qty": m_qty})
    dept_data = {k: float(v) for k, v in sorted(dept_exp.items(), key=lambda x: x[1], reverse=True)}
    # New: Top 5 customers
    top5 = sales_raw.groupby("Client")["Sales Amount"].sum().sort_values(ascending=False).head(5)
    top_customers = [{"name": str(k), "sales": float(v)} for k, v in top5.items()]
    # Top 5 types
    top5_types = sales_raw.groupby("Type")["Sales Amount"].sum().sort_values(ascending=False).head(5)
    top_types = [{"name": str(k), "sales": float(v)} for k, v in top5_types.items()]
    # Top 5 fabrics
    top5_fabrics = sales_raw.groupby("Fabric")["Sales Amount"].sum().sort_values(ascending=False).head(5)
    top_fabrics = [{"name": str(k), "sales": float(v)} for k, v in top5_fabrics.items()]
    # New: Business Unit analysis
    bu_all = sales_raw.groupby("Business Unit")
    bu_sales = bu_all["Sales Amount"].sum()
    bu_returns = bu_all["Return"].sum().abs()
    bu_discounts = bu_all["Discount Value"].sum()
    bu_list = []
    for bu_name in bu_sales.index:
        s = float(bu_sales[bu_name])
        r = float(bu_returns.get(bu_name, 0))
        d = float(bu_discounts.get(bu_name, 0))
        bu_list.append({"name": str(bu_name), "sales_pct": round(s / total_gs * 100, 2) if total_gs else 0,
                        "return_pct": round(r / total_gs * 100, 2) if total_gs else 0,
                        "discount_pct": round(d / total_gs * 100, 2) if total_gs else 0})
    bu_list.sort(key=lambda x: x["sales_pct"], reverse=True)
    return {
        "totals": {
            "gs": round(total_gs), "ret": round(total_ret),
            "disc": round(total_disc), "ns": round(total_ns),
            "cogs": round(total_cogs), "gp": round(gp),
            "exp": round(total_exp), "ni": round(ni)
        },
        "months": months_list,
        "depts": dept_data,
        "top_customers": top_customers,
        "top_types": top_types,
        "top_fabrics": top_fabrics,
        "business_units": bu_list
    }

_pnl_cache = {"data": None, "ts": 0}

def get_pnl_forecast():
    df = pd.read_excel(_EXCEL_PATH, sheet_name="PNL Dashboard ", header=None)
    # Column offsets: YTD(0), Q1(9), Jan(18), Feb(27), Mar(36), Q2(45), Apr(54), May(63), Jun(72)
    sections = {"ytd": 0, "q1": 9, "jan": 18, "feb": 27, "mar": 36, "q2": 45, "apr": 54, "may": 63, "jun": 72}
    months_order = ["jan", "feb", "mar", "apr", "may", "jun"]
    a_val = lambda r, c: float(df.iloc[r, c+1]) if c+1 < df.shape[1] and pd.notna(df.iloc[r, c+1]) else 0
    f_val = lambda r, c: float(df.iloc[r, c+2]) if c+2 < df.shape[1] and pd.notna(df.iloc[r, c+2]) else 0
    # YTD totals (row 45=Net Sales, 46=COGS, 40=Expenses Total)
    ytd_col = sections["ytd"]
    ytd = {
        "net_sales": {"actual": a_val(45, ytd_col), "forecast": f_val(45, ytd_col)},
        "cogs": {"actual": a_val(46, ytd_col), "forecast": f_val(46, ytd_col)},
        "expenses": {"actual": a_val(40, ytd_col), "forecast": f_val(40, ytd_col)},
    }
    ytd["gross_profit"] = {"actual": ytd["net_sales"]["actual"] - abs(ytd["cogs"]["actual"]), "forecast": ytd["net_sales"]["forecast"] - abs(ytd["cogs"]["forecast"])}
    ytd["net_income"] = {"actual": ytd["gross_profit"]["actual"] - abs(ytd["expenses"]["actual"]), "forecast": ytd["gross_profit"]["forecast"] - abs(ytd["expenses"]["forecast"])}
    # Monthly
    monthly = []
    for i, m in enumerate(months_order):
        col = sections[m]
        ns_a = a_val(45, col); ns_f = f_val(45, col)
        cogs_a = a_val(46, col); cogs_f = f_val(46, col)
        exp_a = a_val(40, col); exp_f = f_val(40, col)
        gp_a = ns_a - abs(cogs_a); gp_f = ns_f - abs(cogs_f)
        ni_a = gp_a - abs(exp_a); ni_f = gp_f - abs(exp_f)
        monthly.append({"month": i+1, "net_sales": {"actual": ns_a, "forecast": ns_f}, "cogs": {"actual": cogs_a, "forecast": cogs_f}, "gross_profit": {"actual": gp_a, "forecast": gp_f}, "expenses": {"actual": exp_a, "forecast": exp_f}, "net_income": {"actual": ni_a, "forecast": ni_f}})
    # Dept expenses (rows 33-39)
    dept_names = ["G&A", "Marketing", "Financing Expenses", "Sales Export", "Sales Stores Expenses", "Sales B2B", "R&D"]
    dept_expenses = {}
    ytd_col = sections["ytd"]
    for i, d in enumerate(dept_names):
        r = 33 + i
        actual = a_val(r, ytd_col)
        forecast = f_val(r, ytd_col)
        monthly_dept = []
        for m in months_order:
            col = sections[m]
            ma = a_val(r, col); mf = f_val(r, col)
            monthly_dept.append({"actual": ma, "forecast": mf})
        dept_expenses[d] = {"actual": actual, "forecast": forecast, "monthly": monthly_dept}
    # Profitability data (rows 53-59)
    return {"ytd": ytd, "monthly": monthly, "dept_expenses": dept_expenses}

_sales_cache = {"data": None, "ts": 0}

def get_sales_forecast():
    df = pd.read_excel(_EXCEL_PATH, sheet_name="PNL Dashboard ", header=None)
    sections = {"ytd": 0, "jan": 18, "feb": 27, "mar": 36, "apr": 54, "may": 63, "jun": 72}
    months_order = ["jan", "feb", "mar", "apr", "may", "jun"]
    a_val = lambda r, c: float(df.iloc[r, c+1]) if c+1 < df.shape[1] and pd.notna(df.iloc[r, c+1]) else 0
    f_val = lambda r, c: float(df.iloc[r, c+2]) if c+2 < df.shape[1] and pd.notna(df.iloc[r, c+2]) else 0
    # Row 6 = Sales TOTAL: label, G.SalesA, G.SalesF, Variance, Return, Discount, NetSales, G.Prof%
    ytd_col = sections["ytd"]
    ytd = {
        "gross_sales": {"actual": a_val(6, ytd_col), "forecast": f_val(6, ytd_col)},
        "returns": float(df.iloc[6, ytd_col+4]) if pd.notna(df.iloc[6, ytd_col+4]) else 0,
        "discounts": float(df.iloc[6, ytd_col+5]) if pd.notna(df.iloc[6, ytd_col+5]) else 0,
        "net_sales": float(df.iloc[6, ytd_col+6]) if pd.notna(df.iloc[6, ytd_col+6]) else 0,
    }
    # Monthly
    monthly = []
    for m in months_order:
        col = sections[m]
        gs_a = a_val(6, col)
        gs_f = f_val(6, col)
        ret = float(df.iloc[6, col+4]) if col+4 < df.shape[1] and pd.notna(df.iloc[6, col+4]) else 0
        disc = float(df.iloc[6, col+5]) if col+5 < df.shape[1] and pd.notna(df.iloc[6, col+5]) else 0
        ns = float(df.iloc[6, col+6]) if col+6 < df.shape[1] and pd.notna(df.iloc[6, col+6]) else 0
        monthly.append({"month": months_order.index(m)+1, "gross_sales": {"actual": gs_a, "forecast": gs_f}, "returns": ret, "discounts": disc, "net_sales": ns})
    return {"ytd": ytd, "monthly": monthly}

def get_sales_cached():
    if _sales_cache["data"] is not None:
        return _sales_cache["data"]
    data = get_sales_forecast()
    _sales_cache["data"] = data
    return _sales_cache["data"]

def get_pnl_cached():
    excel_mtime = os.path.getmtime(_EXCEL_PATH)
    if _pnl_cache["data"] is not None:
        return _pnl_cache["data"]
    data = get_pnl_forecast()
    _pnl_cache["data"] = data
    _pnl_cache["ts"] = time.time()
    return _pnl_cache["data"]

def get_data_cached():
    excel_mtime = os.path.getmtime(_EXCEL_PATH)
    if _cache["data"] is not None:
        return _cache["data"]
    try:
        if os.path.exists(_CACHE_FILE):
            with open(_CACHE_FILE, "r") as f:
                cached = json.load(f)
            if cached.get("_mtime") == excel_mtime:
                _cache["data"] = cached["data"]
                _cache["ts"] = time.time()
                return _cache["data"]
    except: pass
    data = get_data()
    _cache["data"] = data
    _cache["ts"] = time.time()
    try:
        with open(_CACHE_FILE, "w") as f:
            json.dump({"_mtime": excel_mtime, "data": data}, f)
    except: pass
    return _cache["data"]

from flask import send_from_directory

@app.route("/static/<path:filename>")
def serve_static(filename):
    return send_from_directory(os.path.join(DIR, "static"), filename)

@app.route("/")
@auth.login_required
def index():
    return render_template_string(HTML_TEMPLATE)

_rates_cache = {"data": None, "ts": 0}

@app.route("/api/rates")
@auth.login_required
def api_rates():
    import urllib.request as _urllib
    now = time.time()
    if _rates_cache["data"] and (now - _rates_cache["ts"]) < 3600:
        return jsonify(_rates_cache["data"])
    try:
        req = _urllib.Request("https://api.exchangerate-api.com/v4/latest/EGP", headers={"User-Agent": "Mozilla/5.0"})
        resp = _urllib.urlopen(req, timeout=10)
        data = json.loads(resp.read().decode())
        rates = data.get("rates", {})
        result = {
            "EGP": 1,
            "USD": rates.get("USD", 0.02),
            "EUR": rates.get("EUR", 0.017),
            "GBP": rates.get("GBP", 0.015),
            "CHF": rates.get("CHF", 0.016),
            "KWD": rates.get("KWD", 0.006),
            "BHD": rates.get("BHD", 0.007),
            "OMR": rates.get("OMR", 0.007),
            "JOD": rates.get("JOD", 0.014),
            "SAR": rates.get("SAR", 0.075),
            "AED": rates.get("AED", 0.073),
        }
        _rates_cache["data"] = result
        _rates_cache["ts"] = now
        return jsonify(result)
    except:
        return jsonify({"EGP": 1, "USD": 0.02, "EUR": 0.017, "GBP": 0.015, "CHF": 0.016, "KWD": 0.006, "BHD": 0.007, "OMR": 0.007, "JOD": 0.014, "SAR": 0.075, "AED": 0.073})

@app.route("/api/data")
@auth.login_required
def api_data():
    period = request.args.get("period", "ytd")
    bu = request.args.get("bu", "all")
    month = request.args.get("month", "all")
    data = get_data(period=period, bu=bu, month=month)
    return jsonify(data)

@app.route("/api/pnl_forecast")
@auth.login_required
def api_pnl_forecast():
    return jsonify(get_pnl_cached())

@app.route("/api/sales_forecast")
@auth.login_required
def api_sales_forecast():
    return jsonify(get_sales_cached())

@app.route("/api/style_analysis")
@auth.login_required
def api_style_analysis():
    try:
        df = pd.read_excel(_EXCEL_PATH, sheet_name="32 Degree Style Analysis ", header=None)
        result = {}
        # PO columns are at indices 1-5 (PO 79 A, PO 78 A, PO 78-35 A, PO 47 A, PO 49 A)
        # PO names from row 1
        po_names = [str(df.iloc[1, i]).strip() for i in range(1, 6)]
        result["po_names"] = po_names
        # Section 1: Cost breakdown (rows 0-15)
        section1 = {}
        section1_rows = {
            "exchange_rate": 2, "production_time": 3, "order_quantity": 4, "selling_price": 5,
            "revenue": 6, "labor_cost_per_item": 7, "overtime_per_item": 8, "packing_per_item": 9,
            "total_cogs_per_item": 10, "total_cogs": 11, "cogs_pct": 12,
            "gross_profit_per_item": 13, "total_gross_profit": 14, "gross_profit_margin": 15
        }
        for key, row_idx in section1_rows.items():
            vals = []
            for ci in range(1, 6):
                v = df.iloc[row_idx, ci]
                if pd.isna(v):
                    vals.append(0)
                else:
                    try:
                        vals.append(float(v))
                    except (ValueError, TypeError):
                        vals.append(0)
            total_actual = vals[-1] if len(vals) >= 5 else 0
            if row_idx in (12, 15):
                # Percentages
                section1[key] = {"values": vals, "total": total_actual}
            else:
                # Read total from col 6 and forecasted from col 7
                t_val = 0
                f_val = 0
                try:
                    t_val = float(df.iloc[row_idx, 6]) if pd.notna(df.iloc[row_idx, 6]) else 0
                except: pass
                try:
                    f_val = float(df.iloc[row_idx, 7]) if pd.notna(df.iloc[row_idx, 7]) else 0
                except: pass
                try:
                    var_val = float(df.iloc[row_idx, 8]) if pd.notna(df.iloc[row_idx, 8]) else 0
                except: pass
                section1[key] = {"values": vals, "total": t_val, "forecasted": f_val, "variance": var_val}
        result["cost_breakdown"] = section1
        # Section 2: Profitability (rows 19-28)
        section2 = {}
        section2_rows = {
            "order_quantity": 20, "gp_margin_per_item": 21, "total_gross_profit": 22,
            "clearance_per_item": 23, "finance_per_item": 24, "sga_per_item": 25,
            "net_profit_per_item": 26, "total_net_profit": 27, "net_profit_margin": 28
        }
        for key, row_idx in section2_rows.items():
            vals = []
            for ci in range(1, 6):
                v = df.iloc[row_idx, ci]
                if pd.isna(v):
                    vals.append(0)
                else:
                    try:
                        vals.append(float(v))
                    except (ValueError, TypeError):
                        vals.append(0)
            section2[key] = {"values": vals}
        result["profitability"] = section2
        # Variance text (row 30-31)
        try:
            var_text = str(df.iloc[30, 0]) if pd.notna(df.iloc[30, 0]) else ""
            var_detail = str(df.iloc[31, 0]) if pd.notna(df.iloc[31, 0]) else ""
        except:
            var_text = ""; var_detail = ""
        result["variance_title"] = var_text
        result["variance_detail"] = var_detail
        # Recommendations (row 46-47)
        try:
            rec_title = str(df.iloc[46, 0]) if pd.notna(df.iloc[46, 0]) else ""
            rec_detail = str(df.iloc[47, 0]) if pd.notna(df.iloc[47, 0]) else ""
        except:
            rec_title = ""; rec_detail = ""
        result["rec_title"] = rec_title
        result["recommendations"] = rec_detail
        # Waterfall data (rows 49-57)
        waterfall = {}
        wf_rows = {
            "planned_profit": 50, "labor_cost_adjustment": 51, "overtime": 52,
            "saving_packing": 53, "saving_clearance": 54, "saving_finance": 55,
            "saving_sga": 56, "actual_profit": 57
        }
        for key, row_idx in wf_rows.items():
            try:
                v = float(df.iloc[row_idx, 1]) if pd.notna(df.iloc[row_idx, 1]) else 0
            except:
                v = 0
            waterfall[key] = v
        result["waterfall"] = waterfall
        # Cost reduction (rows 53-55)
        cost_reduction = {}
        try:
            cr_header_row = 53
            cr_data_rows = {"saving_packing": 53, "saving_clearance": 54, "saving_finance": 55, "saving_sga": 56}
            for key, row_idx in cr_data_rows.items():
                vals = {}
                for ci in range(1, 6):
                    try:
                        v = float(df.iloc[row_idx, ci]) if pd.notna(df.iloc[row_idx, ci]) else 0
                    except:
                        v = 0
                    vals[po_names[ci-1]] = v
                cost_reduction[key] = vals
        except:
            pass
        result["cost_reduction"] = cost_reduction
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/investment")
@auth.login_required
def api_investment():
    try:
        df = pd.read_excel(_EXCEL_PATH, sheet_name="2025 - 2026 Analysis ", header=None)
        result = {}
        def safe_float(row, col):
            try:
                v = df.iloc[row, col]
                return float(v) if pd.notna(v) else 0
            except: return 0
        def safe_str(row, col):
            try:
                return str(df.iloc[row, col]).strip() if pd.notna(df.iloc[row, col]) else ""
            except: return ""
        # Gold metrics (rows 2-19)
        gold = {}
        gold_labels = {4:"return_rate",5:"monthly_risk",6:"risk_free_rate",7:"avg_monthly_return",8:"avg_monthly_real_return",9:"sharpe_ratio",10:"real_sharpe",11:"inflation_impact",12:"inflation_rate",13:"monthly_inflation",14:"mdd",15:"downside_dev",16:"sortino",17:"calmar",18:"kurtosis",19:"skewness"}
        for row_idx, key in gold_labels.items():
            gold[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["gold"] = gold
        # Silver metrics (rows 21-38)
        silver = {}
        for row_idx, key in gold_labels.items():
            silver[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["silver"] = silver
        # Swiss Frank metrics (rows 40-57)
        swiss = {}
        for row_idx, key in gold_labels.items():
            swiss[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["swiss"] = swiss
        # Correlations (rows 59-65)
        result["correlations"] = {
            "gold_silver": safe_float(59, 1),
            "gold_swiss": safe_float(60, 1),
            "silver_swiss": safe_float(61, 1)
        }
        result["betas"] = {
            "gold_silver": safe_float(63, 1),
            "gold_swiss": safe_float(64, 1),
            "silver_swiss": safe_float(65, 1)
        }
        # Portfolio (rows 67-68)
        result["portfolio"] = {"gold": safe_float(68, 1), "silver": safe_float(68, 2), "swiss": safe_float(68, 3)}
        # Portfolio metrics (rows 70-86)
        portfolio = {}
        portfolio_labels = {71:"return_rate",72:"monthly_risk",73:"risk_free_rate",74:"avg_monthly_return",75:"avg_monthly_real_return",76:"sharpe_ratio",77:"real_sharpe",78:"inflation_impact",79:"inflation_rate",80:"monthly_inflation",81:"mdd",82:"downside_dev",83:"sortino",84:"calmar",85:"kurtosis",86:"skewness"}
        for row_idx, key in portfolio_labels.items():
            portfolio[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["portfolio"] = portfolio
        result["portfolio"]["weights"] = {"gold": 0.3, "silver": 0.2, "swiss": 0.5}
        result["portfolio"]["prices"] = {"current": safe_float(69, 6), "prev": safe_float(69, 7)}
        result["portfolio"]["growth"] = safe_float(70, 7)
        # Investment results (rows 89-96)
        result["investment"] = {
            "initial": safe_float(89, 1),
            "annual_return": {"actual": safe_float(91, 1), "forecast": safe_float(91, 2)},
            "mdd": {"actual": safe_float(92, 1), "forecast": safe_float(92, 2)},
            "return_after_loss": {"actual": safe_float(93, 1), "forecast": safe_float(93, 2)},
            "discount_rate": safe_float(94, 1),
            "pv": {"actual": safe_float(95, 1), "forecast": safe_float(95, 2)},
            "hedge_pct": {"actual": safe_float(96, 1), "forecast": safe_float(96, 2)}
        }
        result["recommendation"] = safe_str(99, 0)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/cashflow")
@auth.login_required
def api_cashflow():
    try:
        dash = pd.read_excel(_EXCEL_PATH, sheet_name="Cash Flow Dashboard ", header=None)
        coll = pd.read_excel(_EXCEL_PATH, sheet_name="Cash Collection Raw Data", header=0)
        spend = pd.read_excel(_EXCEL_PATH, sheet_name="Cash Spending Raw Data ", header=0)
        # --- Dashboard: Authoritative YTD Actuals & Forecasts ---
        # Collection by BU (rows 5-8, col 1=Actual, col 3=Forecast)
        bu_coll_actual = {}
        bu_coll_forecast = {}
        for i, bu in enumerate(["CM", "Export", "B2B", "B2C"]):
            bu_coll_actual[bu] = float(dash.iloc[5 + i, 1]) if pd.notna(dash.iloc[5 + i, 1]) else 0
            bu_coll_forecast[bu] = float(dash.iloc[5 + i, 2]) if pd.notna(dash.iloc[5 + i, 2]) else 0
        bu_coll_actual["TOTAL"] = float(dash.iloc[9, 1]) if pd.notna(dash.iloc[9, 1]) else 0
        bu_coll_forecast["TOTAL"] = float(dash.iloc[9, 2]) if pd.notna(dash.iloc[9, 2]) else 0
        # Spending by Dept (rows 41-46, col 1=Actual, col 3=Forecast)
        dept_spend_actual = {}
        dept_spend_forecast = {}
        for i, dept in enumerate(["Assets", "Production", "G&A", "Financing Expenses", "S&M", "R&D"]):
            dept_spend_actual[dept] = float(dash.iloc[41 + i, 1]) if pd.notna(dash.iloc[41 + i, 1]) else 0
            dept_spend_forecast[dept] = float(dash.iloc[41 + i, 2]) if pd.notna(dash.iloc[41 + i, 2]) else 0
        dept_spend_actual["TOTAL"] = float(dash.iloc[47, 1]) if pd.notna(dash.iloc[47, 1]) else 0
        dept_spend_forecast["TOTAL"] = float(dash.iloc[47, 2]) if pd.notna(dash.iloc[47, 2]) else 0
        # Top Customers from raw data (dashboard only has top 5)
        coll["amount"] = pd.to_numeric(coll["amount"], errors="coerce").fillna(0)
        cust_coll = coll.groupby("Customer")["amount"].sum().sort_values(ascending=False).head(10).to_dict()
        # Payment Status (rows 23-25, col 1=Amount)
        pay_status = {}
        for i in range(23, 26):
            name = str(dash.iloc[i, 0]).strip() if pd.notna(dash.iloc[i, 0]) else ""
            amt = float(dash.iloc[i, 1]) if pd.notna(dash.iloc[i, 1]) else 0
            if name: pay_status[name] = amt
        # Aging (rows 30-33)
        aging = {}
        for i in range(30, 34):
            bu = str(dash.iloc[i, 0]).strip() if pd.notna(dash.iloc[i, 0]) else ""
            amt = float(dash.iloc[i, 1]) if pd.notna(dash.iloc[i, 1]) else 0
            ag = str(dash.iloc[i, 2]).strip() if pd.notna(dash.iloc[i, 2]) else ""
            if bu: aging[bu] = {"amount": amt, "aging": ag}
        # Top Spending Categories (rows 53-57, col 1=Amount)
        cat_spend = {}
        for i in range(53, 58):
            name = str(dash.iloc[i, 0]).strip() if pd.notna(dash.iloc[i, 0]) else ""
            amt = float(dash.iloc[i, 1]) if pd.notna(dash.iloc[i, 1]) else 0
            if name: cat_spend[name] = amt
        # --- Raw Data: Monthly breakdowns ---
        coll["amount"] = pd.to_numeric(coll["amount"], errors="coerce").fillna(0)
        coll["Month"] = pd.to_numeric(coll["Month"], errors="coerce")
        coll = coll[coll["Month"].notna() & coll["Month"].between(1, 12)]
        coll["Month"] = coll["Month"].astype(int)
        spend["amount"] = pd.to_numeric(spend["amount"], errors="coerce").fillna(0)
        spend["Month"] = pd.to_numeric(spend["Month"], errors="coerce")
        spend = spend[spend["Month"].notna() & spend["Month"].between(1, 12)]
        spend["Month"] = spend["Month"].astype(int)
        monthly_coll = coll.groupby("Month")["amount"].sum()
        monthly_spend = spend.groupby("Month")["amount"].sum()
        monthly_bu = coll.groupby(["Month", "B.U"])["amount"].sum().unstack(fill_value=0)
        spend["Department"] = spend["Department"].str.strip().str.title()
        monthly_dept = spend.groupby(["Month", "Department"])["amount"].sum().unstack(fill_value=0)
        # Net cash flow per month
        monthly_net = {}
        all_months = sorted(set(list(monthly_coll.index) + list(monthly_spend.index)))
        for m in all_months:
            monthly_net[m] = float(monthly_coll.get(m, 0)) - float(monthly_spend.get(m, 0))
        # Days Receivable / Payable
        total_coll = bu_coll_actual["TOTAL"]
        total_spend = dept_spend_actual["TOTAL"]
        months_active = max(coll["Month"].nunique(), 1)
        TOTAL_NET_SALES = 62511337.0
        TOTAL_COGS = 31223509.0
        avg_monthly_revenue = TOTAL_NET_SALES / 12 * months_active
        avg_monthly_cogs = TOTAL_COGS / 12 * months_active
        days_receivable = (total_coll / avg_monthly_revenue * 30) if total_coll > 0 else 0
        days_payable = (total_spend / avg_monthly_cogs * 30) if total_spend > 0 else 0
        return jsonify({
            "monthly_collections": {str(k): float(v) for k, v in monthly_coll.items()},
            "monthly_spending": {str(k): float(v) for k, v in monthly_spend.items()},
            "monthly_net": {str(k): float(v) for k, v in monthly_net.items()},
            "bu_collections_actual": bu_coll_actual,
            "bu_collections_forecast": bu_coll_forecast,
            "customer_collections": cust_coll,
            "payment_status": pay_status,
            "dept_spending_actual": dept_spend_actual,
            "dept_spending_forecast": dept_spend_forecast,
            "category_spending": cat_spend,
            "monthly_bu": {str(k): {col: float(v) for col, v in row.items()} for k, row in monthly_bu.iterrows()},
            "monthly_dept": {str(k): {col: float(v) for col, v in row.items()} for k, row in monthly_dept.iterrows()},
            "aging": aging,
            "total_collections": total_coll,
            "total_spending": total_spend,
            "total_forecast_coll": bu_coll_forecast["TOTAL"],
            "total_forecast_spend": dept_spend_forecast["TOTAL"],
            "days_receivable": days_receivable,
            "days_payable": days_payable
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/refresh")
@auth.login_required
def api_refresh():
    _cache["data"] = None
    for p in [_CACHE_FILE, _DF_CACHE]:
        if os.path.exists(p):
            os.remove(p)
    get_data_cached()
    return jsonify({"ok": True})

@app.route("/upload", methods=["POST"])
@auth.login_required
def upload_excel():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"ok": False, "error": "Must be .xlsx file"}), 400
    f.save(_EXCEL_PATH)
    _cache["data"] = None
    for p in [_CACHE_FILE, _DF_CACHE]:
        if os.path.exists(p):
            os.remove(p)
    get_data_cached()
    return jsonify({"ok": True, "msg": "Data updated. Refresh the page."})

PORT = int(os.environ.get("PORT", 8765))
ON_RENDER = os.environ.get("RENDER", "").lower() == "true"

try:
    get_data_cached()
    get_pnl_cached()
    get_sales_cached()
except: pass

if __name__ == "__main__":
    from waitress import serve
    import urllib.request
    from threading import Thread
    def warm_up():
        import time
        for i in range(5):
            try:
                req = urllib.request.Request("http://localhost:" + str(PORT) + "/api/data")
                req.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req, timeout=30)
                req2 = urllib.request.Request("http://localhost:" + str(PORT) + "/api/pnl_forecast")
                req2.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req2, timeout=30)
                req3 = urllib.request.Request("http://localhost:" + str(PORT) + "/api/sales_forecast")
                req3.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req3, timeout=30)
                req4 = urllib.request.Request("http://localhost:" + str(PORT) + "/api/expenses_forecast")
                req4.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req4, timeout=30)
                break
            except:
                time.sleep(2)
    Thread(target=warm_up, daemon=True).start()
    if ON_RENDER:
        print(f"Starting on Render — port {PORT}")
    else:
        print("=" * 50)
        print("  VIVA 1960 Dashboard")
        print(f"  Local:  http://localhost:{PORT}")
        try:
            from pyngrok import ngrok
            tunnel = ngrok.connect(PORT, bind_tls=True)
            public_url = tunnel.public_url.replace("http://", "https://")
            print(f"  Public: {public_url}")
            print(f"  Login:  {USER} / {PASS}")
            url_file = os.path.join(DIR, "dashboard_url.txt")
            with open(url_file, "w") as f:
                f.write(f"VIVA 1960 Dashboard\n")
                f.write(f"URL: {public_url}\n")
                f.write(f"Login: {USER}\nPassword: {PASS}\n")
        except Exception:
            print("  (Ngrok not configured — local only)")
            print()
            print("  For a public URL, deploy to Render (see instructions below)")
        print("=" * 50)
    serve(app, host="0.0.0.0", port=PORT, threads=8)
