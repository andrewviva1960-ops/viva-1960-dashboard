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
_OVERRIDES_FILE = os.path.join(DIR, "dashboard_overrides.json")

def _load_overrides():
    if os.path.exists(_OVERRIDES_FILE):
        try:
            with open(_OVERRIDES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {"values": {}, "labels": {}, "notes": {}, "layout": {}}

def _save_overrides(data):
    with open(_OVERRIDES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

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
<script>
Plotly.setPlotConfig({
  font:{family:'Inter, system-ui, sans-serif',size:12,color:'#72839E'},
  displayModeBar:false,
  responsive:true,
  scrollZoom:false,
  staticPlot:false,
  doubleClick:false,
  displaylogo:false,
  modeBarButtonsToRemove:['zoom2d','pan2d','select2d','lasso2d','zoomIn2d','zoomOut2d','autoScale2d','resetScale2d'],
  layout:{autosize:true,paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',margin:{t:30,b:50,l:60,r:20},dragmode:false,
    xaxis:{gridcolor:'transparent',linecolor:'rgba(0,0,0,0.06)',zerolinecolor:'rgba(0,0,0,0.06)',tickfont:{size:11,color:'#4A5568'},fixedrange:true},
    yaxis:{gridcolor:'transparent',linecolor:'rgba(0,0,0,0.06)',zerolinecolor:'rgba(0,0,0,0.06)',tickfont:{size:11,color:'#4A5568'},fixedrange:true},
    font:{family:'Inter, system-ui, sans-serif',size:12,color:'#72839E'}
  }
});
document.addEventListener('dblclick', function(e) { if (e.target.closest('.js-plotly-plot')) { e.stopPropagation(); e.preventDefault(); } }, true);
document.addEventListener('touchstart', function(e) { if (e.target.closest('.js-plotly-plot') && e.touches.length > 1) { e.preventDefault(); } }, {passive:false, capture:true});
const _CALM = {
  palette: ['#3BD671','#473FE0','#F18C2A','#DF484A','#A0AEC0','#F6AD55','#00B5D8','#F687B3','#2DA356','#63B3ED'],
  bar:    ['#3BD671','#473FE0','#F18C2A','#DF484A','#A0AEC0','#F6AD55','#00B5D8','#F687B3','#2DA356','#63B3ED'],
  line:   ['#3BD671','#473FE0','#F18C2A','#DF484A','#A0AEC0','#F6AD55','#00B5D8','#F687B3','#2DA356','#63B3ED'],
  fill:   ['rgba(59,214,113,0.10)','rgba(71,63,224,0.10)','rgba(241,140,42,0.10)','rgba(223,72,74,0.10)','rgba(160,174,192,0.10)'],
  grid: 'transparent',
  zero: 'rgba(0,0,0,0.06)',
  txt:  '#4A5568',
  txtDark: '#A0AEC0'
};
Plotly.defaultConfig = Object.assign({}, Plotly.defaultConfig || {}, {
  hoverlabel: {
    bgcolor: '#ffffff',
    bordercolor: '#e2e8f0',
    font: { family: 'Inter, system-ui, sans-serif', size: 12, color: '#1e293b', weight: 600 },
    align: 'left'
  }
});
var _origPlotlyNewPlot = Plotly.newPlot;
Plotly.newPlot = function(el, data, layout, config) {
  const dark = document.body.classList.contains('dark-mode');
  const darkHover = {bgcolor:dark?'#1C2434':'#ffffff',bordercolor:dark?'#3BD671':'#3BD671',font:{family:'Inter, system-ui, sans-serif',size:12,color:dark?'#EDF2F7':'#1e293b',weight:600},align:'left'};
  layout = Object.assign({hoverlabel: darkHover}, layout || {});
  if (dark) {
    layout.paper_bgcolor = layout.paper_bgcolor || '#1C2434';
    layout.plot_bgcolor = layout.plot_bgcolor || '#1C2434';
    if (layout.font) layout.font.color = '#EDF2F7';
    if (layout.xaxis && layout.xaxis.tickfont) layout.xaxis.tickfont.color = '#A0AEC0';
    if (layout.yaxis && layout.yaxis.tickfont) layout.yaxis.tickfont.color = '#A0AEC0';
  }
  return _origPlotlyNewPlot(el, data, layout, config);
};
function _calmLayout(extra) {
  const dark = document.body.classList.contains('dark-mode');
  return Object.assign({
    autosize:true,
    paper_bgcolor: dark ? '#1C2434' : 'rgba(0,0,0,0)',
    plot_bgcolor: dark ? '#1C2434' : 'rgba(0,0,0,0)',
    font:{family:'Inter, system-ui, sans-serif',size:12,color:dark?'#EDF2F7':'#2D3748'},
    hoverlabel:{bgcolor:dark?'#1C2434':'#ffffff',bordercolor:dark?'#3BD671':'#3BD671',font:{family:'Inter, system-ui, sans-serif',size:12,color:dark?'#EDF2F7':'#1e293b',weight:600},align:'left',namefont:{size:11,color:dark?'#A0AEC0':'#4A5568'}},
    xaxis:{gridcolor:'transparent',linecolor:dark?'rgba(255,255,255,0.08)':'rgba(0,0,0,0.06)',zerolinecolor:dark?'rgba(255,255,255,0.08)':'rgba(0,0,0,0.06)',tickfont:{size:11,color:dark?'#A0AEC0':'#4A5568'},fixedrange:true},
    yaxis:{gridcolor:'transparent',linecolor:dark?'rgba(255,255,255,0.08)':'rgba(0,0,0,0.06)',zerolinecolor:dark?'rgba(255,255,255,0.08)':'rgba(0,0,0,0.06)',tickfont:{size:11,color:dark?'#A0AEC0':'#4A5568'},fixedrange:true},
    margin:{t:30,b:50,l:60,r:20}
  }, extra||{});
}
</script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
:root {
  --sidebar-bg: #ffffff;
  --sidebar-width: 220px;
  --header-bg: #ffffff;
  --accent: #3BD671;
  --accent2: #2DA356;
  --card-bg: #ffffff;
  --card-border: rgba(0,0,0,0.06);
  --text-primary: #1C2434;
  --text-secondary: #72839E;
  --green: #3BD671;
  --red: #DF484A;
  --teal: #00B5D8;
  --blue: #473FE0;
  --purple: #805AD5;
  --orange: #F18C2A;
  --pink: #D53F8C;
  --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
  --shadow-md: 0 4px 12px rgba(0,0,0,0.06);
  --shadow-lg: 0 8px 24px rgba(0,0,0,0.08);
  --shadow-xl: 0 12px 40px rgba(0,0,0,0.12);
  --radius-sm: 8px;
  --radius-md: 12px;
  --radius-lg: 16px;
  --chart-h: 320px;
  --chart-h-lg: 380px;
}
body.dark-mode {
  --sidebar-bg: #131A25;
  --header-bg: #131A25;
  --card-bg: #1C2434;
  --card-border: rgba(255,255,255,0.06);
  --text-primary: #EDF2F7;
  --text-secondary: #A0AEC0;
  --accent: #3BD671;
  --accent2: #2DA356;
  --green: #3BD671;
  --red: #DF484A;
  --teal: #00B5D8;
  --blue: #63B3ED;
  --purple: #B794F4;
  --orange: #F6AD55;
  --pink: #F687B3;
  --shadow-sm: 0 1px 2px rgba(0,0,0,0.3);
  --shadow-md: 0 4px 12px rgba(0,0,0,0.4);
  --shadow-lg: 0 8px 24px rgba(0,0,0,0.5);
  --shadow-xl: 0 12px 40px rgba(0,0,0,0.6);
  background:#131A25;
  color:#EDF2F7;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',system-ui,-apple-system,sans-serif;background:linear-gradient(135deg,#f8fafc 0%,#f1f5f9 50%,#e8edf5 100%);color:var(--text-primary);overflow-x:hidden;-webkit-font-smoothing:antialiased}
body.dark-mode{background:linear-gradient(135deg,#131A25 0%,#1C2434 50%,#131A25 100%)}
body.dark-mode .header{background:rgba(19,26,37,0.92);border-bottom-color:rgba(255,255,255,0.06)}
body.dark-mode .header::after{opacity:0.6}
body.dark-mode .sidebar{background:rgba(19,26,37,0.95);border-right-color:rgba(255,255,255,0.06)}
body.dark-mode .sidebar .nav-item:hover{background:rgba(255,255,255,0.05)}
body.dark-mode .sidebar .nav-item.active{background:linear-gradient(135deg,rgba(59,214,113,0.15),rgba(59,214,113,0.05))}
body.dark-mode .kpi-card{background:#1C2434;border-color:rgba(255,255,255,0.06)}
body.dark-mode .kpi-card .kpi-value{color:#EDF2F7}
body.dark-mode .kpi-card .kpi-label{color:#A0AEC0}
body.dark-mode .kpi-card .kpi-sub{color:#3BD671}
body.dark-mode .chart-card{background:#1C2434;border-color:rgba(255,255,255,0.06)}
body.dark-mode .chart-card .chart-header h5{color:#EDF2F7}
body.dark-mode .pnl-card{background:#1C2434;border-color:rgba(255,255,255,0.06)}
body.dark-mode .pnl-card .pnl-label{color:#A0AEC0}
body.dark-mode .pnl-card .pnl-label.sub{color:#72839E}
body.dark-mode .pnl-card .pnl-label.section-title{color:#3BD671}
body.dark-mode .pnl-card tr.subtotal .pnl-label{color:#3BD671}
body.dark-mode .pnl-card tr.subtotal td:last-child{color:#EDF2F7;border-top-color:rgba(255,255,255,0.15)}
body.dark-mode .pnl-card tr.grand-total .pnl-label{color:#EDF2F7}
body.dark-mode .pnl-card tr.grand-total td:last-child{color:#EDF2F7;border-top-color:#3BD671}
body.dark-mode .pnl-card .neg{color:#ef9a9a}
body.dark-mode .pnl-card .pos{color:#3BD671}
body.dark-mode td.num{color:#EDF2F7}
body.dark-mode .hoverlayer .hovertext rect{fill:#1C2434!important;stroke:#3BD671!important}
body.dark-mode .hoverlayer .hovertext text{fill:#EDF2F7!important}
body.dark-mode .footer{color:#4a7a4e}
body.dark-mode .currency-selector select{background:#1C2434;color:#EDF2F7;border-color:rgba(255,255,255,0.15)}
body.dark-mode #globalFilters{background:rgba(26,46,30,0.5);border-color:rgba(255,255,255,0.1);border-radius:10px;padding:8px 12px}
body.dark-mode #globalFilters select{color:#EDF2F7}
body.dark-mode #globalFilters label{color:#A0AEC0}
body.dark-mode .hdr-btn{color:#A0AEC0;border-color:rgba(255,255,255,0.2)}
body.dark-mode .hdr-btn:hover{background:rgba(255,255,255,0.12);border-color:rgba(255,255,255,0.3)}
body.dark-mode .page-title{color:#EDF2F7}
body.dark-mode .sidebar-toggle-float{background:linear-gradient(135deg,#3BD671,#2DA356)}
body.dark-mode .edit-modal{background:#1C2434;border-color:rgba(255,255,255,0.15)}
body.dark-mode .edit-modal h4{color:#EDF2F7}
body.dark-mode .edit-modal input,.edit-mode .edit-modal textarea{background:#131A25;color:#EDF2F7;border-color:rgba(255,255,255,0.2)}
body.dark-mode .edit-overlay{background:rgba(0,0,0,0.6)}
body.dark-mode .edit-save-bar .btn-save{background:#3BD671;color:#131A25}
body.dark-mode .mobile-nav{background:rgba(15,26,18,0.95);border-top-color:rgba(255,255,255,0.1)}
body.dark-mode .mobile-nav .mnav-item{color:#A0AEC0}
body.dark-mode .mobile-nav .mnav-item.active{color:#EDF2F7;background:rgba(255,255,255,0.12)}
body.dark-mode .mobile-nav .mnav-item i{color:#3BD671}
body.dark-mode .mobile-nav .mnav-item.active i{color:#3BD671}
body.dark-mode .loading-shimmer{background:linear-gradient(90deg,rgba(255,255,255,0.05) 25%,rgba(255,255,255,0.1) 50%,rgba(255,255,255,0.05) 75%);background-size:200% 100%}
body.dark-mode .pnl-card table tr{border-bottom-color:rgba(255,255,255,0.06)}
body.dark-mode .chart-card .chart-header h5{color:#3BD671}
body.dark-mode .chart-card:hover{box-shadow:0 6px 20px rgba(255,255,255,0.1)}
@keyframes fadeInUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes slideInLeft{from{opacity:0;transform:translateX(-20px)}to{opacity:1;transform:translateX(0)}}
@keyframes scaleIn{from{opacity:0;transform:scale(0.92)}to{opacity:1;transform:scale(1)}}
@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}}
@keyframes glowPulse{0%,100%{box-shadow:0 0 8px rgba(59,214,113,0.2)}50%{box-shadow:0 0 20px rgba(59,214,113,0.4)}}
@keyframes slideUp{from{opacity:0;transform:translateY(30px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideDown{from{opacity:0;transform:translateY(-20px)}to{opacity:1;transform:translateY(0)}}
@keyframes zoomBounce{0%{transform:scale(0.8);opacity:0}60%{transform:scale(1.05)}100%{transform:scale(1);opacity:1}}
@keyframes typewriter{from{width:0}to{width:100%}}
@keyframes blink{0%,100%{border-color:var(--accent)}50%{border-color:transparent}}
@keyframes ripple{0%{transform:scale(0);opacity:.6}100%{transform:scale(2.5);opacity:0}}
@keyframes countUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes borderSlide{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
@keyframes iconSpin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
@keyframes heroGlow{0%,100%{filter:drop-shadow(0 0 6px rgba(59,214,113,0.3))}50%{filter:drop-shadow(0 0 16px rgba(59,214,113,0.6))}}
.header{background:rgba(255,255,255,0.85);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);color:var(--text-primary);height:64px;display:flex;align-items:center;padding:0 24px;position:fixed;top:0;left:0;right:0;z-index:1000;border-bottom:1px solid rgba(0,0,0,0.06);box-shadow:var(--shadow-sm);animation:slideDown .5s ease both}
.header::after{content:'';position:absolute;bottom:0;left:0;right:0;height:2px;background:linear-gradient(90deg,#3BD671,#473FE0,#F18C2A,#DF484A,#3BD671);background-size:200% 100%;animation:borderSlide 8s linear infinite}
.header .brand{display:flex;align-items:center;gap:12px;width:var(--sidebar-width);font-size:16px;font-weight:700;letter-spacing:.3px;color:#1C2434}
.header .brand img{height:40px;width:auto;flex-shrink:0;transition:transform .3s;animation:heroGlow 3s ease-in-out infinite}
.header .brand img:hover{transform:scale(1.08) rotate(-2deg)}
.header .top-right{margin-left:auto;display:flex;align-items:center;gap:10px;font-size:12px;color:#4A5568}
.header .top-right .status-dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--green);margin-right:4px;animation:glowPulse 2s ease-in-out infinite}
.hdr-btn{background:rgba(59,214,113,0.08);border:1px solid rgba(59,214,113,0.2);color:var(--accent);cursor:pointer;padding:5px 12px;border-radius:8px;font-size:12px;font-weight:500;display:inline-flex;align-items:center;gap:6px;transition:all .2s ease;font-family:Inter,system-ui,sans-serif}
.hdr-btn:hover{background:rgba(59,214,113,0.15);border-color:rgba(59,214,113,0.35);transform:translateY(-1px);box-shadow:0 2px 8px rgba(59,214,113,0.15)}
.hdr-btn:active{transform:translateY(0)}
.hdr-btn i{font-size:13px}
.hdr-btn.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.month-opt{display:flex;align-items:center;gap:6px;padding:5px 12px;font-size:11px;cursor:pointer;color:var(--text-secondary);transition:background .15s}
.month-opt:hover{background:rgba(59,214,113,0.08)}
.month-opt input[type=checkbox]{accent-color:var(--accent);width:14px;height:14px;cursor:pointer}
.ms-wrap{display:flex;align-items:center;gap:5px;background:var(--card-bg);border:1px solid var(--card-border);border-radius:8px;padding:5px 10px;position:relative}
.ms-trigger{display:flex;align-items:center;background:transparent;border:none;color:#334155;font-size:12px;font-weight:500;cursor:pointer;padding:0;min-width:60px}
.ms-panel{position:absolute;top:100%;left:0;min-width:100%;background:var(--card-bg);border:1px solid var(--card-border);border-radius:8px;padding:6px 0;z-index:2000;box-shadow:0 4px 16px rgba(0,0,0,0.12);max-height:260px;overflow-y:auto}
body.dark-mode .ms-wrap{border-color:rgba(255,255,255,0.15)}
body.dark-mode .ms-trigger{color:#A0AEC0}
body.dark-mode .ms-panel{background:#1C2434;border-color:rgba(255,255,255,0.15)}
body.dark-mode .month-opt{color:#3BD671}
body.dark-mode .month-opt:hover{background:rgba(255,255,255,0.1)}
.header .currency-selector{display:flex;align-items:center;gap:6px;margin-left:14px;padding-left:14px;border-left:1px solid rgba(0,0,0,0.06)}
.header .currency-selector label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#2DA356}
.header .currency-selector select{background:#f8fafc;border:1px solid rgba(0,0,0,0.08);border-radius:var(--radius-sm);padding:4px 28px 4px 10px;color:#334155;font-size:12px;font-weight:500;cursor:pointer;appearance:none;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2364748b' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 8px center;transition:border-color .2s,box-shadow .2s}
.header .currency-selector select:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(59,214,113,0.1)}
.sidebar{position:fixed;top:64px;left:0;bottom:0;width:var(--sidebar-width);background:rgba(255,255,255,0.92);backdrop-filter:blur(10px);-webkit-backdrop-filter:blur(10px);color:var(--text-secondary);overflow-y:auto;z-index:999;padding-top:8px;border-right:1px solid rgba(0,0,0,0.06);transition:transform .3s cubic-bezier(.4,0,.2,1),opacity .3s ease}
.sidebar.hidden{transform:translateX(-100%);opacity:0;pointer-events:none}
.sidebar-toggle{background:none;border:none;font-size:14px;color:var(--text-secondary);cursor:pointer;padding:0 6px;border-radius:4px;transition:.2s;margin-left:auto}
.sidebar-toggle:hover{color:var(--accent);transform:scale(1.1)}
.sidebar-toggle-float{display:none;position:fixed;left:8px;top:420px;z-index:1001;font-size:16px;padding:10px 12px;background:linear-gradient(135deg,var(--accent),var(--accent2));color:#fff;border-radius:0 12px 12px 0;box-shadow:var(--shadow-lg);transition:all .35s cubic-bezier(.4,0,.2,1);animation:float 3s ease-in-out infinite}
.sidebar-toggle-float:hover{padding-left:16px;box-shadow:0 8px 30px rgba(59,214,113,0.4);transform:scale(1.08)}
body.sidebar-hidden .sidebar-toggle-float{display:block}
.sidebar .nav-item{padding:10px 20px;display:flex;align-items:center;gap:11px;cursor:pointer;font-size:13px;transition:all .25s cubic-bezier(.4,0,.2,1);border-left:3px solid transparent;border-radius:0 12px 12px 0;margin-right:8px;font-weight:500;animation:slideInLeft .4s ease both}
.sidebar .nav-item:nth-child(1){animation-delay:.05s}.sidebar .nav-item:nth-child(2){animation-delay:.1s}.sidebar .nav-item:nth-child(3){animation-delay:.15s}.sidebar .nav-item:nth-child(4){animation-delay:.2s}.sidebar .nav-item:nth-child(5){animation-delay:.25s}.sidebar .nav-item:nth-child(6){animation-delay:.3s}.sidebar .nav-item:nth-child(7){animation-delay:.35s}.sidebar .nav-item:nth-child(8){animation-delay:.4s}.sidebar .nav-item:nth-child(9){animation-delay:.45s}.sidebar .nav-item:nth-child(10){animation-delay:.5s}
.sidebar .nav-item:hover{background:rgba(59,214,113,0.1);color:#1C2434;transform:translateX(6px);box-shadow:2px 0 12px rgba(59,214,113,0.1)}
.sidebar .nav-item.active{background:linear-gradient(135deg,rgba(59,214,113,0.15),rgba(59,214,113,0.05));color:#1C2434;border-left-color:var(--accent);font-weight:600;box-shadow:2px 0 12px rgba(59,214,113,0.12)}
body.dark-mode .sidebar .nav-item:hover{color:#EDF2F7}
body.dark-mode .sidebar .nav-item.active{color:#EDF2F7}
.sidebar .nav-item i{width:18px;text-align:center;font-size:14px;color:#2DA356}
.sidebar .nav-section{font-size:10px;text-transform:uppercase;letter-spacing:1.2px;color:#2DA356;padding:16px 20px 5px;font-weight:700}
.main{margin-left:var(--sidebar-width);margin-top:64px;padding:16px 20px;min-height:calc(100vh - 64px);transition:margin-left .3s cubic-bezier(.4,0,.2,1);overflow-x:visible}
body.sidebar-hidden .main{margin-left:0}
.page-title{font-size:18px;font-weight:800;color:var(--text-primary);margin-bottom:12px;letter-spacing:-.3px;animation:slideDown .5s ease both}
.kpi-grid{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px;animation:fadeInUp .5s ease both}
@keyframes kpiStagger{from{opacity:0;transform:translateY(10px) scale(0.97)}to{opacity:1;transform:translateY(0) scale(1)}}
@media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:600px){.kpi-grid{grid-template-columns:1fr 1fr}}
.kpi-card{background:var(--card-bg);border-radius:var(--radius-md);padding:12px;border:1px solid var(--card-border);box-shadow:var(--shadow-sm);display:flex;flex-direction:row;align-items:center;gap:10px;transition:all .3s ease;animation:kpiStagger .4s ease both;position:relative;overflow:visible}
.kpi-card::before{content:'';position:absolute;inset:-1px;border-radius:calc(var(--radius-md) + 1px);background:linear-gradient(135deg,#3BD671,#473FE0,#F18C2A,#3BD671);background-size:300% 300%;opacity:0;z-index:-1;transition:opacity .4s ease;animation:borderGlow 6s ease infinite}
.kpi-card:hover{box-shadow:0 6px 20px rgba(59,214,113,0.12);transform:translateY(-2px)}
.kpi-card:hover::before{opacity:1}
.kpi-card:nth-child(1){animation-delay:.05s}.kpi-card:nth-child(2){animation-delay:.1s}.kpi-card:nth-child(3){animation-delay:.15s}.kpi-card:nth-child(4){animation-delay:.2s}.kpi-card:nth-child(5){animation-delay:.25s}.kpi-card:nth-child(6){animation-delay:.3s}
.kpi-card .kpi-icon{width:36px;height:36px;min-width:36px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:14px;color:#fff;transition:all .3s ease}
.kpi-card:hover .kpi-icon{transform:scale(1.08)}
.kpi-card .kpi-label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#2DA356;margin-bottom:2px;font-weight:600}
.kpi-card .kpi-value{font-size:16px;font-weight:800;color:var(--text-primary);line-height:1.1;letter-spacing:-.3px}
.kpi-card .kpi-sub{font-size:9px;color:#3BD671;margin-top:2px;font-weight:500}
.kpi-card.kpi-gs{border-left:4px solid var(--green)}.kpi-card.kpi-gs .kpi-icon{background:linear-gradient(135deg,#3BD671,#059669)}
.kpi-card.kpi-ns{border-left:4px solid var(--blue)}.kpi-card.kpi-ns .kpi-icon{background:linear-gradient(135deg,#473FE0,#2563eb)}
.kpi-card.kpi-cogs{border-left:4px solid var(--orange)}.kpi-card.kpi-cogs .kpi-icon{background:linear-gradient(135deg,#F18C2A,#d97706)}
.kpi-card.kpi-gp{border-left:4px solid var(--purple)}.kpi-card.kpi-gp .kpi-icon{background:linear-gradient(135deg,#473FE0,#2DA356)}
.kpi-card.kpi-exp{border-left:4px solid var(--red)}.kpi-card.kpi-exp .kpi-icon{background:linear-gradient(135deg,#DF484A,#dc2626)}
.kpi-card.kpi-ni{border-left:4px solid var(--teal)}.kpi-card.kpi-ni .kpi-icon{background:linear-gradient(135deg,#00B5D8,#0891b2)}
.kpi-card .kpi-value.green{color:var(--green)}.kpi-card .kpi-value.red{color:var(--red)}
#pnlKpiGrid{grid-template-columns:repeat(5,1fr)!important;gap:6px!important;margin-bottom:10px!important}
#pnlKpiGrid .kpi-card{padding:8px 8px!important}
#pnlKpiGrid .kpi-icon{width:24px!important;height:24px!important;min-width:24px!important;font-size:11px!important}
#pnlKpiGrid .kpi-card .kpi-value{font-size:15px!important}
#pnlKpiGrid .kpi-card .kpi-sub{font-size:10px!important}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;overflow:visible;padding:0 10px}
.chart-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:12px}
@media(max-width:1100px){.chart-grid-3{grid-template-columns:1fr 1fr}}
@media(max-width:900px){.chart-grid,.chart-grid-3{grid-template-columns:1fr}}
.mobile-nav{display:none;position:fixed;bottom:0;left:0;right:0;z-index:1001;background:rgba(255,255,255,0.95);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);border-top:1px solid rgba(0,0,0,0.08);box-shadow:0 -2px 12px rgba(0,0,0,0.06);padding:6px 4px env(safe-area-inset-bottom)}
.mobile-nav .nav-row{display:flex;justify-content:space-around;align-items:center}
.mobile-nav .mnav-item{display:flex;flex-direction:column;align-items:center;gap:2px;padding:6px 4px;border-radius:10px;cursor:pointer;font-size:9px;font-weight:600;color:var(--text-secondary);transition:all .2s;min-width:48px;text-decoration:none;-webkit-tap-highlight-color:transparent}
.mobile-nav .mnav-item i{font-size:16px;color:#2DA356;transition:transform .2s}
.mobile-nav .mnav-item.active{color:var(--text-primary);background:rgba(59,214,113,0.1)}
.mobile-nav .mnav-item.active i{transform:scale(1.15);color:var(--accent)}
.mobile-nav .mnav-item:active{transform:scale(0.92)}
.mobile-nav .mnav-dark{background:linear-gradient(135deg,var(--accent),var(--accent2));color:#fff!important;border-radius:50%;width:38px;height:38px;display:flex;align-items:center;justify-content:center;min-width:38px;padding:0;box-shadow:0 2px 10px rgba(59,214,113,0.3)}
.mobile-nav .mnav-dark i{color:#fff!important;font-size:15px}
.mobile-nav .mnav-action{background:rgba(59,214,113,0.08);border-radius:50%;width:36px;height:36px;padding:0;min-width:36px}
.mobile-nav .mnav-action i{font-size:14px;color:var(--accent)!important}
.mobile-nav .mnav-action.active{background:rgba(59,214,113,0.2)}
body.dark-mode .mobile-nav .mnav-action{background:rgba(255,255,255,0.1)}
body.dark-mode .mobile-nav .mnav-action.active{background:rgba(255,255,255,0.2)}
@media(max-width:768px){
  .mobile-nav{display:block}
  .main{margin-bottom:70px}
  body.sidebar-hidden .sidebar-toggle-float{bottom:80px;top:auto}
  .header{padding:0 8px;height:50px}
  .header .brand{width:auto;gap:6px}
  .header .brand img{height:26px}
  .header .brand span{font-size:12px}
  .header .top-right{display:none}
  .hdr-btn{padding:4px 6px;font-size:10px;gap:3px}
  .hdr-btn span{display:none}
  .hdr-btn i{font-size:12px}
  .header .currency-selector{display:none}
  .sidebar{display:none!important}
  .main{margin-left:0!important;margin-top:50px;padding:10px 10px}
  .page-title{font-size:15px;margin-bottom:8px}
  #globalFilters{flex-wrap:wrap;gap:6px;padding:6px 8px}
  #globalFilters select{font-size:11px;padding:5px 22px 5px 8px}
  .kpi-grid{grid-template-columns:1fr 1fr;gap:8px}
  .kpi-card{padding:10px;gap:8px}
  .kpi-card .kpi-value{font-size:13px!important}
  .kpi-card .kpi-label{font-size:10px}
  .kpi-card .kpi-sub{font-size:9px}
  .kpi-card .kpi-icon{width:30px;height:30px;font-size:12px}
  .chart-grid,.chart-grid-3{grid-template-columns:1fr;gap:8px}
  .chart-card{border-radius:10px}
  .chart-card .chart-header h5{font-size:11px}
  .chart-body{height:220px}
  .pnl-card{font-size:10px}
  #pnlTable td,#pnlTable th{padding:4px 6px;font-size:10px}
  .footer{font-size:9px;padding:8px}
}
@media(max-width:400px){
  .kpi-grid{grid-template-columns:1fr 1fr;gap:6px}
  .kpi-card{padding:8px;gap:6px;flex-direction:column;text-align:center}
  .kpi-card .kpi-value{font-size:12px!important}
  .mobile-nav .mnav-item{font-size:8px;min-width:38px;padding:4px 2px}
  .mobile-nav .mnav-item i{font-size:14px}
  .mobile-nav .mnav-action{width:32px;height:32px;min-width:32px}
  .mobile-nav .mnav-action i{font-size:12px}
}
@keyframes borderGlow{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
@keyframes cardIn{from{opacity:0;transform:translateY(16px) scale(0.98)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes colorShift{0%{filter:hue-rotate(0deg)}50%{filter:hue-rotate(15deg)}100%{filter:hue-rotate(0deg)}}
.chart-card{background:var(--card-bg);border-radius:var(--radius-md);border:1px solid var(--card-border);box-shadow:var(--shadow-sm);overflow:visible;position:relative;transition:all .3s ease;animation:slideUp .5s ease both;z-index:1}
.chart-card::before{content:'';position:absolute;inset:-1px;border-radius:calc(var(--radius-md) + 1px);background:linear-gradient(135deg,#3BD671,#473FE0,#F18C2A,#3BD671);background-size:300% 300%;opacity:0;z-index:-1;transition:opacity .4s ease;animation:borderGlow 6s ease infinite}
.chart-card:hover{box-shadow:0 6px 20px rgba(59,214,113,0.1);border-color:transparent}
.chart-card:hover::before{opacity:1}
.chart-card:nth-child(1){animation-delay:.05s}.chart-card:nth-child(2){animation-delay:.1s}.chart-card:nth-child(3){animation-delay:.15s}.chart-card:nth-child(4){animation-delay:.2s}.chart-card:nth-child(5){animation-delay:.25s}.chart-card:nth-child(6){animation-delay:.3s}
.chart-card .chart-header{display:flex;align-items:center;justify-content:space-between;padding:12px 16px 0}
.chart-card .chart-header h5{font-size:13px;font-weight:600;color:var(--text-primary);margin:0;letter-spacing:-.2px}
.chart-card .chart-body{padding:4px 6px 6px;height:var(--chart-h);overflow:visible;touch-action:manipulation;-webkit-touch-callout:none;-webkit-user-select:none;user-select:none}
.chart-card .chart-body .js-plotly-plot,.chart-card .chart-body .plotly{overflow:visible!important}
.chart-card .chart-body .js-plotly-plot .main-svg,.chart-card .chart-body .plotly .main-svg{overflow:visible!important}
.chart-card .pnl-body{padding:0;height:var(--chart-h);overflow:visible}
.chart-card.full{grid-column:1/-1}
.chart-resize-handle{display:none;position:absolute;bottom:0;left:0;right:0;height:8px;cursor:ns-resize;background:linear-gradient(180deg,transparent,rgba(59,214,113,0.15));border-radius:0 0 var(--radius-md) var(--radius-md);z-index:10;transition:background .2s}
.chart-resize-handle:hover{background:linear-gradient(180deg,transparent,rgba(59,214,113,0.4))}
body.edit-mode .chart-resize-handle{display:block}
.pnl-card{background:var(--card-bg);border-radius:var(--radius-md);border:1px solid var(--card-border);box-shadow:var(--shadow-sm);padding:0;height:100%;display:flex;flex-direction:column;width:100%;transition:all .35s cubic-bezier(.4,0,.2,1)}
.pnl-card:hover{box-shadow:0 8px 28px rgba(0,0,0,0.08);transform:translateY(-2px)}
.pnl-card .pnl-header{padding:6px 10px 0;display:flex;align-items:center;justify-content:space-between}
.pnl-card .pnl-header h5{font-size:12px;font-weight:700;color:var(--text-primary);margin:0}
.pnl-card .pnl-header .pnl-period{font-size:10px;color:var(--text-secondary)}
.pnl-card .pnl-inner{padding:0;flex:1;overflow-x:auto}
.pnl-card table{width:100%;border-collapse:collapse;font-size:12px}
.pnl-card td{padding:5px 14px;border:0}
.pnl-card td:last-child{text-align:right;font-weight:700;font-variant-numeric:tabular-nums;letter-spacing:.3px;width:170px;padding-right:24px}
.pnl-card .pnl-label{color:#4A5568;padding-left:22px}
.pnl-card .pnl-label.sub{padding-left:40px;font-size:12px;color:#2DA356}
.pnl-card .pnl-label.sub2{padding-left:52px;font-size:11.5px;color:#2DA356}
.pnl-card .pnl-label.section-title{color:var(--text-primary);font-weight:700;font-size:12px;padding-left:22px;text-transform:uppercase;letter-spacing:.5px}
.pnl-card tr.section-row td{height:1px;padding:0;border:0}
.pnl-card tr.section-row td::before{content:"";display:block;width:100%;height:1px;background:rgba(0,0,0,0.06)}
.pnl-card tr.section-divider td{height:4px;padding:0;border:0}
.pnl-card tr.subtotal td{padding-top:8px}
.pnl-card tr.subtotal .pnl-label{color:#2D3748;font-weight:600}
.pnl-card tr.subtotal td:last-child{border-top:1px solid rgba(59,214,113,0.15);padding-top:8px;font-weight:700;color:var(--text-primary)}
.pnl-card tr.grand-total td{padding-top:12px}
.pnl-card tr.grand-total .pnl-label{color:#1a4a1a;font-size:12px;font-weight:800}
.pnl-card tr.grand-total td:last-child{border-top:3px double #2DA356;padding-top:8px;font-size:13px;font-weight:800;color:#1a4a1a}
.pnl-card .neg{color:var(--red)}.pnl-card .pos{color:var(--green)}
.pnl-card tr.deduction td:last-child{color:var(--red)}
.pnl-card tr.revenue-header td{padding-top:4px}
.pnl-card tr{animation:fadeIn .3s ease both}
.pnl-card tr:nth-child(1){animation-delay:.05s}.pnl-card tr:nth-child(2){animation-delay:.08s}.pnl-card tr:nth-child(3){animation-delay:.11s}.pnl-card tr:nth-child(4){animation-delay:.14s}.pnl-card tr:nth-child(5){animation-delay:.17s}.pnl-card tr:nth-child(6){animation-delay:.2s}.pnl-card tr:nth-child(7){animation-delay:.23s}.pnl-card tr:nth-child(8){animation-delay:.26s}.pnl-card tr:nth-child(9){animation-delay:.29s}.pnl-card tr:nth-child(10){animation-delay:.32s}
.chart-grid+.chart-grid{margin-top:14px}
.tab-content{display:none;animation:fadeInUp .4s cubic-bezier(.4,0,.2,1) both}.tab-content.active{display:block}
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
.footer{text-align:center;padding:16px 0;font-size:11px;color:#3BD671;animation:fadeIn .6s ease both}
.loading-shimmer{background:linear-gradient(90deg,rgba(59,214,113,0.05) 25%,rgba(59,214,113,0.12) 50%,rgba(59,214,113,0.05) 75%);background-size:200% 100%;animation:shimmer 1.5s infinite;border-radius:var(--radius-md)}
.hoverlayer .hovertext rect{fill:#ffffff!important;stroke:#3BD671!important;filter:drop-shadow(0 4px 12px rgba(59,214,113,0.15))!important}
.hoverlayer .hovertext text{fill:#1C2434!important;font-family:'Inter',system-ui,sans-serif!important;font-size:12px!important}
.hoverlayer .hovertext .legendtext{fill:#4A5568!important;font-size:11px!important}
body.edit-mode .editable{cursor:pointer;border-radius:var(--radius-sm);transition:all .2s}
body.edit-mode .editable:hover{background:rgba(59,214,113,0.08);outline:2px dashed rgba(59,214,113,0.3)}
body.edit-mode .editable.editing{background:rgba(59,214,113,0.06);outline:2px solid var(--accent)}
body.edit-mode .editable-note{position:relative}
body.edit-mode .editable-note.has-note::after{content:attr(data-note);position:absolute;top:-6px;right:-6px;background:var(--orange);color:#fff;font-size:8px;width:14px;height:14px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:700;cursor:pointer}
.edit-save-bar{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:linear-gradient(135deg,var(--accent),var(--accent2));color:#fff;padding:12px 28px;border-radius:14px;box-shadow:0 8px 30px rgba(59,214,113,0.4);display:none;z-index:9999;gap:12px;align-items:center;font-size:13px;animation:zoomBounce .4s cubic-bezier(.4,0,.2,1)}
body.edit-mode .edit-save-bar{display:flex}
.edit-save-bar button{border:none;border-radius:8px;padding:8px 18px;font-size:12px;font-weight:600;cursor:pointer;transition:all .2s}
.edit-save-bar button:hover{transform:translateY(-1px)}
.edit-save-bar .btn-save{background:#fff;color:var(--accent)}
.edit-save-bar .btn-cancel{background:rgba(255,255,255,0.2);color:#fff}
.edit-save-bar .btn-clear{background:#DF484A;color:#fff}
.edit-save-bar .btn-undo{background:rgba(255,255,255,0.2);color:#fff}
.edit-overlay{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.4);backdrop-filter:blur(4px);z-index:10000;align-items:center;justify-content:center;animation:fadeIn .2s ease}
.edit-overlay.active{display:flex}
.edit-modal{background:#fff;border-radius:var(--radius-lg);padding:28px;min-width:320px;max-width:500px;box-shadow:0 24px 60px rgba(0,0,0,0.3);animation:zoomBounce .35s cubic-bezier(.4,0,.2,1)}
.edit-modal h4{margin:0 0 16px;font-size:16px;color:var(--text-primary);font-weight:700}
.edit-modal input,.edit-modal textarea{width:100%;border:1px solid rgba(0,0,0,0.1);border-radius:var(--radius-sm);padding:10px 14px;font-size:13px;color:var(--text-primary);box-sizing:border-box;margin-bottom:12px;transition:border-color .2s,box-shadow .2s;font-family:inherit}
.edit-modal input:focus,.edit-modal textarea:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(59,214,113,0.1)}
.edit-modal textarea{height:80px;resize:vertical}
.edit-modal .modal-actions{display:flex;gap:8px;justify-content:flex-end}
.edit-modal .modal-actions button{border:none;border-radius:8px;padding:8px 18px;font-size:12px;font-weight:600;cursor:pointer;transition:all .2s}
.edit-modal .modal-actions button:hover{transform:translateY(-1px)}
.modal-btn-primary{background:linear-gradient(135deg,var(--accent),var(--accent2));color:#fff}
.modal-btn-cancel{background:#f1f5f9;color:#72839E}
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
        <option value="EGP">EGP — Egyptian Pound (1.00)</option>
      </select>
    </div>
    <button id="refreshBtn" onclick="refreshData()" class="hdr-btn"><i class="fas fa-sync-alt"></i> Refresh</button>
    <button id="editToggle" onclick="toggleEditMode()" class="hdr-btn" title="Toggle edit mode"><i class="fas fa-pen"></i> Edit</button>
    <button id="blurToggle" onclick="toggleBlur()" class="hdr-btn" title="Toggle number visibility"><i class="fas fa-eye"></i></button>
    <button id="darkToggle" onclick="toggleDarkMode()" class="hdr-btn" title="Toggle dark/light mode"><i class="fas fa-moon"></i></button>
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
  <div class="nav-section" style="display:flex;align-items:center">Data <button class="sidebar-toggle" onclick="toggleSidebar()" title="Toggle sidebar"><i class="fas fa-bars"></i></button></div>
  <div class="nav-item"><i class="fas fa-table"></i> Raw Data</div>
  <div class="nav-item" style="cursor:pointer" onclick="document.getElementById('fileInput').click()"><i class="fas fa-upload"></i> Upload Excel</div>
  <input type="file" id="fileInput" accept=".xlsx" style="display:none" onchange="uploadExcel(this)">
</div>
<button class="sidebar-toggle sidebar-toggle-float" onclick="toggleSidebar()" title="Toggle sidebar"><i class="fas fa-bars"></i></button>

<!-- Main -->
<div class="main">

<div id="globalFilters" style="display:flex;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap">
  <div class="ms-wrap" id="periodWrap">
    <i class="fas fa-calendar" style="font-size:11px;color:var(--accent)"></i>
    <label style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">Period</label>
    <div class="ms-trigger" onclick="togglePanel('periodPanel')"><span id="periodLabel">YTD</span><i class="fas fa-chevron-down" style="font-size:8px;margin-left:4px;color:var(--text-secondary)"></i></div>
    <div class="ms-panel" id="periodPanel" style="display:none">
      <div class="month-opt" onclick="msClick(this,'period')"><input type="checkbox" value="ytd" checked> YTD</div>
      <div class="month-opt" onclick="msClick(this,'period')"><input type="checkbox" value="q1"> Q1</div>
      <div class="month-opt" onclick="msClick(this,'period')"><input type="checkbox" value="q2"> Q2</div>
      <div class="month-opt" onclick="msClick(this,'period')"><input type="checkbox" value="q3"> Q3</div>
    </div>
  </div>
  <div class="ms-wrap" id="buWrap">
    <i class="fas fa-building" style="font-size:11px;color:var(--accent)"></i>
    <label style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">B.U</label>
    <div class="ms-trigger" onclick="togglePanel('buPanel')"><span id="buLabel">All</span><i class="fas fa-chevron-down" style="font-size:8px;margin-left:4px;color:var(--text-secondary)"></i></div>
    <div class="ms-panel" id="buPanel" style="display:none">
      <div class="month-opt" onclick="msClick(this,'bu')"><input type="checkbox" value="all" checked> All</div>
      <div class="month-opt" onclick="msClick(this,'bu')"><input type="checkbox" value="Export"> Export</div>
      <div class="month-opt" onclick="msClick(this,'bu')"><input type="checkbox" value="B2B"> B2B</div>
      <div class="month-opt" onclick="msClick(this,'bu')"><input type="checkbox" value="B2C"> B2C</div>
      <div class="month-opt" onclick="msClick(this,'bu')"><input type="checkbox" value="CM"> CM</div>
    </div>
  </div>
  <div class="ms-wrap" id="monthFilterWrap">
    <i class="fas fa-calendar-day" style="font-size:11px;color:var(--accent)"></i>
    <label style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">Month</label>
    <div class="ms-trigger" onclick="togglePanel('monthPanel')"><span id="monthLabel">All</span><i class="fas fa-chevron-down" style="font-size:8px;margin-left:4px;color:var(--text-secondary)"></i></div>
    <div class="ms-panel" id="monthPanel" style="display:none">
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="all" checked> All</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="1"> Jan</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="2"> Feb</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="3"> Mar</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="4"> Apr</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="5"> May</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="6"> Jun</div>
      <div class="month-opt" onclick="msClick(this,'month')"><input type="checkbox" value="7"> Jul</div>
    </div>
  </div>
</div>

<div id="tab-dashboard" class="tab-content active">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
    <div class="page-title" style="margin-bottom:0"><i class="fas fa-tachometer-alt" style="color:var(--accent);margin-right:8px"></i>Dashboard</div>
  </div>
  <div class="kpi-grid" id="kpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-arrows-alt-h" style="color:var(--accent);margin-right:6px"></i>Monthly Net Sales vs Expenses</h5></div><div class="chart-body" id="monthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:var(--accent);margin-right:6px"></i>Monthly Gross Sales</h5></div><div class="chart-body" id="salesChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-receipt" style="color:#DF484A;margin-right:6px"></i>Monthly Expenses</h5></div><div class="chart-body" id="expChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#63B3ED;margin-right:6px"></i>Expenses by Department</h5></div><div class="chart-body" id="deptChart"></div></div>
    <div class="chart-card"><div class="pnl-card"><div class="pnl-header"><h5><i class="fas fa-file-invoice-dollar" style="color:var(--accent);margin-right:6px"></i>Profit &amp; Loss Statement</h5><span class="pnl-period">Jan - Jun 2026</span></div><div class="pnl-inner" id="pnlCard"></div></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-trophy" style="color:var(--accent);margin-right:6px"></i>Top 5 <span id="topSlicerLabel">Customers</span> by Sales</h5><div class="slicer-selector" style="display:flex;align-items:center;gap:6px"><label style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-secondary)">View</label><select id="topSlicer" onchange="changeTopSlicer(this.value)" style="background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:4px 28px 4px 10px;color:#334155;font-size:12px;font-weight:500;cursor:pointer;appearance:none;background-image:url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%237c3aed' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E\");background-repeat:no-repeat;background-position:right 8px center"><option value="customers">Customers</option><option value="types">Types</option><option value="fabrics">Fabrics</option></select></div></div><div class="chart-body" id="topCustomersChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-cubes" style="color:#4A5568;margin-right:6px"></i>Monthly Quantity Sold</h5></div><div class="chart-body" id="qtyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-balance-scale" style="color:var(--accent);margin-right:6px"></i>Monthly Net Sales vs COGS</h5></div><div class="chart-body" id="cogsChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:var(--accent);margin-right:6px"></i>BU Sales %</h5></div><div class="chart-body" id="buSalesChart"></div></div>
  </div>
</div>

<div id="tab-pnl" class="tab-content">
  <div class="page-title"><i class="fas fa-file-invoice-dollar" style="color:var(--accent);margin-right:8px"></i>P&amp;L Actuals vs Forecast</div>
  <div class="kpi-grid" id="pnlKpiGrid" style="grid-template-columns:repeat(5,1fr)"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:var(--accent);margin-right:6px"></i>Monthly Net Sales — Actual vs Forecast</h5></div><div class="chart-body" id="pnlSalesChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#4A5568;margin-right:6px"></i>Monthly Gross Profit — Actual vs Forecast</h5></div><div class="chart-body" id="pnlGpChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#3BD671;margin-right:6px"></i>Monthly Net Income — Actual vs Forecast</h5></div><div class="chart-body" id="pnlNiChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#3BD671;margin-right:6px"></i>Gross Profit Margin % Trend</h5></div><div class="chart-body" id="pnlMarginChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#2DA356;margin-right:6px"></i>Department Expenses — Actual vs Forecast</h5></div><div class="chart-body" id="pnlDeptChart"></div></div>
  </div>
    <div class="chart-grid">
    <div class="chart-card full" style="overflow:visible"><div class="chart-header"><h5><i class="fas fa-arrows-alt" style="color:#2DA356;margin-right:6px"></i>P&amp;L Waterfall — YTD Actual vs Forecast</h5></div><div class="chart-body" id="pnlWaterfallChart" style="min-height:var(--chart-h-lg);overflow:visible"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#334155;margin-right:6px"></i>P&amp;L Summary — Actual vs Forecast</h5></div><div class="pnl-body" id="pnlSummaryTable"></div></div>
  </div>
</div>

<div id="tab-sales" class="tab-content">
  <div class="page-title"><i class="fas fa-chart-bar" style="color:var(--accent);margin-right:8px"></i>Sales Analytics</div>
  <div class="kpi-grid" id="salesKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:var(--accent);margin-right:6px"></i>Monthly Gross Sales — Actual vs Budget</h5></div><div class="chart-body" id="salesGSChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-area" style="color:#4A5568;margin-right:6px"></i>Monthly Returns &amp; Discounts</h5></div><div class="chart-body" id="salesDedChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-cubes" style="color:var(--accent);margin-right:6px"></i>Monthly Quantity Sold</h5></div><div class="chart-body" id="salesQtyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#63B3ED;margin-right:6px"></i>Sales by Business Unit</h5></div><div class="chart-body" id="salesBUChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-trophy" style="color:#3BD671;margin-right:6px"></i>Top 5 Customers</h5></div><div class="chart-body" id="salesTopChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-list" style="color:#2DA356;margin-right:6px"></i>Top 5 Types</h5></div><div class="chart-body" id="salesTypeChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-tshirt" style="color:#473FE0;margin-right:6px"></i>Top 5 Fabrics</h5></div><div class="chart-body" id="salesFabricChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-percent" style="color:#3BD671;margin-right:6px"></i>Net Sales Conversion Rate %</h5></div><div class="chart-body" id="salesConversionChart"></div></div>
  </div>
</div>

<div id="tab-expenses" class="tab-content">
  <div class="page-title"><i class="fas fa-receipt" style="color:var(--accent);margin-right:8px"></i>Expense Analytics</div>
  <div class="kpi-grid" id="expKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#DF484A;margin-right:6px"></i>Monthly Expenses — Actual vs Budget</h5></div><div class="chart-body" id="expMonthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#2DA356;margin-right:6px"></i>Department Expenses — Actual vs Forecast</h5></div><div class="chart-body" id="expDeptChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#3BD671;margin-right:6px"></i>Monthly Expense Variance</h5></div><div class="chart-body" id="expVarChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#2DA356;margin-right:6px"></i>Expense Distribution</h5></div><div class="chart-body" id="expPieChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-area" style="color:#473FE0;margin-right:6px"></i>Top Departments Monthly Trend</h5></div><div class="chart-body" id="expTrendChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-table" style="color:#334155;margin-right:6px"></i>Department Expense Detail — Variance Analysis</h5></div><div class="pnl-body" id="expTable"></div></div>
  </div>
</div>

<div id="tab-style" class="tab-content">
  <div class="page-title"><i class="fas fa-tshirt" style="color:var(--accent);margin-right:8px"></i>32 Degree Style Analysis — 168 Soft Bra</div>
  <div class="kpi-grid" id="styleKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-water" style="color:#2DA356;margin-right:6px"></i>Profit Waterfall — Planned to Actual</h5></div><div class="chart-body" id="styleWaterfall"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-hand-holding-usd" style="color:#3BD671;margin-right:6px"></i>Net Profit by Purchase Order</h5></div><div class="chart-body" id="styleNpChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-pie" style="color:#2DA356;margin-right:6px"></i>COGS % by Purchase Order</h5></div><div class="chart-body" id="styleCogsPie"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-th-list" style="color:#3BD671;margin-right:6px"></i>PO-Level Cost Breakdown</h5></div><div class="pnl-body" id="styleTable"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#3BD671;margin-right:6px"></i>Margin Trend Across POs</h5></div><div class="chart-body" id="styleMarginChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#473FE0;margin-right:6px"></i>Cost Per Item Comparison</h5></div><div class="chart-body" id="styleCostBar"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#3BD671;margin-right:6px"></i>Variance Analysis</h5></div><div class="pnl-body" id="styleVariance"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-lightbulb" style="color:#3BD671;margin-right:6px"></i>Recommendations</h5></div><div class="pnl-body" id="styleRecommendations"></div></div>
  </div>
</div>

<div id="tab-investment" class="tab-content">
  <div class="page-title"><i class="fas fa-coins" style="color:var(--accent);margin-right:8px"></i>Investment Insights — 2025 vs 2026 Analysis</div>
  <div class="kpi-grid" id="invKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#3BD671;margin-right:6px"></i>Return Rate Comparison — Gold / Silver / Swiss Frank</h5></div><div class="chart-body" id="invReturnChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#2DA356;margin-right:6px"></i>Risk-Adjusted Returns — Sharpe / Sortino / Calmar</h5></div><div class="chart-body" id="invRiskChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-shield-alt" style="color:#DF484A;margin-right:6px"></i>Risk Metrics — MDD / Downside Deviation / Volatility</h5></div><div class="chart-body" id="invRiskMetrics"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-project-diagram" style="color:#2DA356;margin-right:6px"></i>Correlation Matrix</h5></div><div class="chart-body" id="invCorrChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#4A5568;margin-right:6px"></i>Portfolio Allocation</h5></div><div class="chart-body" id="invPieChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:var(--accent);margin-right:6px"></i>Investment Performance — Actual vs Forecast</h5></div><div class="chart-body" id="invPerfChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-lightbulb" style="color:#3BD671;margin-right:6px"></i>Recommendation</h5></div><div class="pnl-body" id="invRecommendation"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-trophy" style="color:#F18C2A;margin-right:6px"></i>Top 3 Optimal Portfolio Allocations</h5><div style="display:flex;gap:6px;align-items:center"><select id="invOptYear" onchange="renderOptimal(this.value)" style="background:var(--card-bg);border:1px solid var(--card-border);border-radius:6px;padding:4px 10px;color:#334155;font-size:12px;cursor:pointer"><option value="y2025">2025</option><option value="y2026a" selected>2026 Actual</option><option value="y2026f">2026 Forecast</option></select></div></div><div class="pnl-body" id="invOptimal"></div></div>
  </div>
</div>

<div id="tab-cashflow" class="tab-content">
  <div class="page-title"><i class="fas fa-money-bill-wave" style="color:var(--accent);margin-right:8px"></i>Cash Flow Analysis</div>
  <div class="kpi-grid" id="cfKpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-exchange-alt" style="color:#2DA356;margin-right:6px"></i>Monthly Cash In vs Out vs Net</h5></div><div class="chart-body" id="cfMonthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-building" style="color:#2DA356;margin-right:6px"></i>Collections by Business Unit</h5></div><div class="chart-body" id="cfBuChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#4A5568;margin-right:6px"></i>Spending by Department</h5></div><div class="chart-body" id="cfDeptChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-tags" style="color:#DF484A;margin-right:6px"></i>Top Spending Categories</h5></div><div class="chart-body" id="cfCatChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-chart-line" style="color:#2DA356;margin-right:6px"></i>Collection vs Spending by Month (Stacked)</h5></div><div class="chart-body" id="cfStackChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-clock" style="color:var(--accent);margin-right:6px"></i>Aging Analysis — Receivables</h5></div><div class="chart-body" id="cfAgingChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-users" style="color:#3BD671;margin-right:6px"></i>Top 10 Customers</h5></div><div class="chart-body" id="cfCustChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-list-ol" style="color:#2DA356;margin-right:6px"></i>Top 10 Spending Categories</h5></div><div class="chart-body" id="cfCatBarChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:var(--accent);margin-right:6px"></i>Collections vs Forecast (Actuals)</h5></div><div class="chart-body" id="cfCollForecastChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#DF484A;margin-right:6px"></i>Spending vs Forecast (Actuals)</h5></div><div class="chart-body" id="cfSpendForecastChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-lightbulb" style="color:#3BD671;margin-right:6px"></i>Cash Flow Insights</h5></div><div class="pnl-body" id="cfInsights"></div></div>
  </div>
</div>

<div class="mobile-nav" id="mobileNav">
  <div class="nav-row">
    <div class="mnav-item active" onclick="switchTab('dashboard')"><i class="fas fa-tachometer-alt"></i>Home</div>
    <div class="mnav-item" onclick="switchTab('pnl')"><i class="fas fa-file-invoice-dollar"></i>P&L</div>
    <div class="mnav-item" onclick="switchTab('sales')"><i class="fas fa-chart-bar"></i>Sales</div>
    <div class="mnav-item" onclick="switchTab('expenses')"><i class="fas fa-receipt"></i>Expenses</div>
    <div class="mnav-item" onclick="switchTab('style')"><i class="fas fa-tshirt"></i>Style</div>
    <div class="mnav-item" onclick="switchTab('investment')"><i class="fas fa-coins"></i>Invest</div>
    <div class="mnav-item" onclick="switchTab('cashflow')"><i class="fas fa-money-bill-wave"></i>Cash</div>
    <div class="mnav-item mnav-action" onclick="toggleBlur()" title="Toggle blur"><i class="fas fa-eye" id="mobileBlurIcon"></i></div>
    <div class="mnav-item mnav-action" onclick="toggleEditMode()" title="Toggle edit mode"><i class="fas fa-pen" id="mobileEditIcon"></i></div>
    <div class="mnav-item mnav-dark" onclick="toggleDarkMode()" title="Toggle dark/light mode"><i class="fas fa-moon"></i></div>
  </div>
</div>

<div class="footer">VIVA 1960 Dashboard &mdash; Data refreshes every 30 seconds</div>
</div>

<div class="edit-save-bar">
  <span><i class="fas fa-pen"></i> Edit Mode Active</span>
  <button class="btn-undo" onclick="undoEdit()" title="Undo (Ctrl+Z)"><i class="fas fa-undo"></i> Undo</button>
  <button class="btn-undo" onclick="redoEdit()" title="Redo (Ctrl+Y)"><i class="fas fa-redo"></i> Redo</button>
  <button class="btn-save" onclick="saveAllEdits()"><i class="fas fa-check"></i> Save</button>
  <button class="btn-cancel" onclick="toggleEditMode()">Cancel</button>
  <button class="btn-clear" onclick="clearAllEdits()"><i class="fas fa-trash"></i> Clear All</button>
</div>

<div class="edit-overlay" id="editOverlay">
  <div class="edit-modal">
    <h4 id="editModalTitle">Edit Value</h4>
    <input type="text" id="editModalInput" style="display:none" />
    <textarea id="editModalTextarea" style="display:none"></textarea>
    <div class="modal-actions">
      <button class="modal-btn-cancel" onclick="closeEditModal()">Cancel</button>
      <button class="modal-btn-primary" onclick="confirmEdit()">Save</button>
    </div>
  </div>
</div>

<script>
// Currency rates (base: EGP) - CBE Official Rates (29 Jul 2026) - Average of Buy & Sell
let CURRENCIES = {
  EGP: {rate: 1, symbol: 'EGP', locale: 'en-US'},
  USD: {rate: 0.020089, symbol: '$', locale: 'en-US'},
  EUR: {rate: 0.017396, symbol: '\u20AC', locale: 'de-DE'},
  GBP: {rate: 0.014922, symbol: '\u00A3', locale: 'en-GB'},
  CHF: {rate: 0.016231, symbol: 'CHF', locale: 'de-CH'},
  KWD: {rate: 0.006169, symbol: 'KWD', locale: 'ar-KW'},
  BHD: {rate: 0.00744, symbol: 'BHD', locale: 'ar-BH'},
  OMR: {rate: 0.00767, symbol: 'OMR', locale: 'ar-OM'},
  JOD: {rate: 0.01409, symbol: 'JOD', locale: 'ar-JO'},
  SAR: {rate: 0.075445, symbol: 'SAR', locale: 'ar-SA'},
  AED: {rate: 0.073778, symbol: 'AED', locale: 'ar-AE'}
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

function animateKPIs() {
  document.querySelectorAll('.kpi-value').forEach((el, i) => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(10px)';
    setTimeout(() => {
      el.style.transition = 'all .5s cubic-bezier(.4,0,.2,1)';
      el.style.opacity = '1';
      el.style.transform = 'translateY(0)';
    }, 80 + i * 60);
  });
}

function animateCharts() {
  document.querySelectorAll('.chart-card').forEach((el, i) => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(20px)';
    setTimeout(() => {
      el.style.transition = 'all .6s cubic-bezier(.4,0,.2,1)';
      el.style.opacity = '1';
      el.style.transform = 'translateY(0)';
    }, 100 + i * 80);
  });
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
    marker:{color:['#2DA356','#473FE0','#3BD671','#3BD671','#81E6D9']},
    text:vals.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false}],
    {margin:{t:50,b:55,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:11,color:'#4A5568'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'transparent'}, height:280,
     hovermode:'x unified', showlegend:false}, {responsive:true, displayModeBar:false});
}

function changeTopSlicer(val) {
  topSlicer = val;
  if (_data) renderTopChart(_data, getCurrency());
  setTimeout(updatePlotlyBlur, 200);
}

async function loadData() {
  const f = getFilters();
  const period = f.period;
  const bu = f.bu;
  const month = f.month;
  const r = await authFetch('/api/data?period=' + encodeURIComponent(period) + '&bu=' + encodeURIComponent(bu) + '&month=' + encodeURIComponent(month));
  const d = await r.json();
  _data = d;
  const T = d.totals;
  document.getElementById('kpiGrid').innerHTML =
    buildKPI('Gross Sales', T.gs, 'chart-line', '#2ecc71', 'gs', 'Total Revenue') +
    buildKPI('Net Sales', T.ns, 'shopping-cart', '#3498db', 'ns', 'After Returns &amp; Discounts') +
    buildKPI('COGS', T.cogs, 'truck', '#f39c12', 'cogs', 'Cost of Goods Sold') +
    buildKPI('Gross Profit', T.gp, 'arrow-up', '#9b59b6', 'gp') +
    buildKPI('Expenses', T.exp, 'receipt', '#e74c3c', 'exp') +
    buildKPI('Net Income', T.ni, 'wallet', '#1abc9c', 'ni');
  animateKPIs();

  const M = d.months;
  const monthLabels = M.map((_,i) => ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][i] || ('M'+(i+1)));

  const c = getCurrency();
  const gsV = M.map(m => m.gs * c.rate);
  Plotly.newPlot('salesChart', [{type:'bar', x:monthLabels, y:gsV, marker:{color:_CALM.palette[0],line:{width:0},opacity:0.85}, text:gsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false, marker_line:{width:0}}],
    _calmLayout({height:280, hovermode:'x unified', showlegend:false, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'transparent'}}),
    {responsive:true, displayModeBar:false});

  const expV = M.map(m => m.exp * c.rate);
  Plotly.newPlot('expChart', [{type:'bar', x:monthLabels, y:expV, marker:{color:_CALM.palette[3],opacity:0.85}, text:expV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false}],
    _calmLayout({height:280, hovermode:'x unified', showlegend:false, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'transparent'}}),
    {responsive:true, displayModeBar:false});

  const nsV = M.map(m => m.ns * c.rate);
  const cogsV = M.map(m => m.cogs * c.rate);
  Plotly.newPlot('monthlyChart', [
    {type:'bar', name:'Net Sales', x:monthLabels, y:nsV, marker:{color:_CALM.palette[0],opacity:0.85}, text:nsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false},
    {type:'bar', name:'COGS', x:monthLabels, y:cogsV, marker:{color:_CALM.palette[1],opacity:0.85}, text:cogsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false, offsetgroup:1},
    {type:'bar', name:'Expenses', x:monthLabels, y:expV, marker:{color:_CALM.palette[3],opacity:0.85}, text:expV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false, offsetgroup:1, base:cogsV}
  ], _calmLayout({barmode:'group', height:300, hovermode:'x unified',
      legend:{orientation:'h',y:1.18,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}}), {responsive:true, displayModeBar:false});

  const depts = Object.keys(d.depts);
  const dVals = Object.values(d.depts).map(v => v * c.rate);
  Plotly.newPlot('deptChart', [{type:'bar', orientation:'h', x:dVals, y:depts,
    marker:{color:_CALM.palette.slice(0,depts.length),opacity:0.85}, text:dVals.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false}],
    _calmLayout({height:Math.max(300, depts.length*45), hovermode:'y unified', showlegend:false,
     margin:{t:15,b:25,l:15,r:15},
     yaxis:{automargin:true, tickfont:{size:11,color:'#4A5568'}},
     xaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent',automargin:true}}), {responsive:true, displayModeBar:false});

  document.getElementById('pnlCard').innerHTML =
    '<table>' +
    '<tr class="section-divider"><td colspan="2"></td></tr>' +
    '<tr class="revenue-header"><td class="pnl-label section-title">Revenue</td><td></td></tr>' +
    '<tr><td class="pnl-label">Gross Sales</td><td class="pos num">' + fmtShort(T.gs) + '</td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Less: Returns</td><td class="neg num">' + fmtShort(-T.ret) + '</td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Less: Discounts</td><td class="neg num">' + fmtShort(-T.disc) + '</td></tr>' +
    '<tr class="section-row"><td colspan="2"></td></tr>' +
    '<tr class="subtotal"><td class="pnl-label">Net Sales</td><td class="num">' + fmtShort(T.ns) + '</td></tr>' +
    '<tr class="section-divider"><td colspan="2"></td></tr>' +
    '<tr class="revenue-header"><td class="pnl-label section-title">Cost of Goods Sold</td><td></td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Less: COGS</td><td class="neg num">' + fmtShort(-T.cogs) + '</td></tr>' +
    '<tr class="section-row"><td colspan="2"></td></tr>' +
    '<tr class="subtotal"><td class="pnl-label">Gross Profit</td><td class="pos num">' + fmtShort(T.gp) + '</td></tr>' +
    '<tr class="section-divider"><td colspan="2"></td></tr>' +
    '<tr class="revenue-header"><td class="pnl-label section-title">Operating Expenses</td><td></td></tr>' +
    '<tr class="deduction"><td class="pnl-label sub">Operating Expenses</td><td class="neg num">' + fmtShort(-T.exp) + '</td></tr>' +
    '<tr class="section-row"><td colspan="2"></td></tr>' +
    '<tr class="grand-total"><td class="pnl-label">Net Income</td><td class="pos num">' + fmtShort(T.ni) + '</td></tr>' +
    '</table>';

  renderTopChart(d, c);

  // Business Units
  const busUnits = d.business_units;
  const buNames = busUnits.map(x => x.name);
  const buColors = ['#4A5568','#4A5568','#63B3ED','#63B3ED','#DF484A'];

  const buSalesPct = busUnits.map(x => x.sales_pct);
  Plotly.newPlot('buSalesChart', [{
    type:'pie', labels:buNames, values:buSalesPct,
    text:buSalesPct.map(v => v.toFixed(1) + '%'), textinfo:'label+percent', textfont:{size:11,color:_CALM.txtDark,family:'Inter'},
    marker:{colors:_CALM.palette.slice(0,buNames.length),line:{color:'#ffffff',width:3}},
    hovertemplate:'%{label}<br>%{value:.1f}%<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:280, showlegend:true,
     legend:{orientation:'h',y:-0.12,font:{size:11,color:'#4A5568'}}},
    {responsive:true, displayModeBar:false});

  // Monthly Quantity Sold
  const qtyVals = M.map(m => m.qty);
  Plotly.newPlot('qtyChart', [{type:'bar', x:monthLabels, y:qtyVals,
    marker:{color:_CALM.palette[2],opacity:0.85}, text:qtyVals.map(v => v > 0 ? v.toLocaleString() : ''), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false}],
    _calmLayout({height:280, hovermode:'x unified', showlegend:false, bargap:0.3,
     yaxis:{rangemode:'tozero',gridcolor:'transparent',automargin:true}}),
    {responsive:true, displayModeBar:false});

  // Monthly Net Sales vs COGS
  const ncogsV = M.map(m => m.cogs * c.rate);
  const nnsV = M.map(m => m.ns * c.rate);
  const fmtBig = v => { const a=Math.abs(v); if(a>=1e6)return (a/1e6).toFixed(1)+'M'; if(a>=1e3)return (a/1e3).toFixed(0)+'K'; return Math.round(a).toString(); };
  const nsLabels = nnsV.map(v => v > 0 ? fmtBig(v)+' '+c.symbol : '');
  const cogsLabels = ncogsV.map(v => v > 0 ? fmtBig(v)+' '+c.symbol : '');
  Plotly.newPlot('cogsChart', [
    {type:'bar', name:'Net Sales', x:monthLabels, y:nnsV, marker:{color:_CALM.palette[0],opacity:0.85}, text:nsLabels, textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false},
    {type:'bar', name:'COGS', x:monthLabels, y:ncogsV, marker:{color:_CALM.palette[1],opacity:0.85}, text:cogsLabels, textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false}
  ], _calmLayout({barmode:'group', height:280, hovermode:'x unified', bargap:0.25, bargroupgap:0.1,
      legend:{orientation:'h',y:1.12,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent',automargin:true}}), {responsive:true, displayModeBar:false});
  animateCharts();
  setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
  setTimeout(preloadTabs, 500);
}

function preloadTabs() {
  const f = getFilters();
  const qs = '?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month);
  ['/api/sales_forecast'+qs,'/api/pnl_forecast'+qs,'/api/data'+qs,'/api/cashflow'+qs,'/api/style_analysis'+qs,'/api/investment'+qs].forEach(url => {
    authFetch(url).then(r => r.json()).catch(() => {});
  });
}

function _plotBatch(charts) {
  return new Promise(resolve => {
    let i = 0;
    function next() {
      if (i >= charts.length) { resolve(); return; }
      const {id, data, layout, config} = charts[i++];
      var el = document.getElementById(id);
      if (!el) { next(); return; }
      try {
        var p = Plotly.newPlot(id, data, layout, config);
        if (p && p.then) p.then(next).catch(next);
        else { next(); }
      } catch(e) { next(); }
    }
    next();
  });
}

async function loadStyleAnalysis() {
  try {
    const f = getFilters();
    const r = await authFetch('/api/style_analysis?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month));
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
      {label:'Total Revenue', value:fmtS(totalRev*c.rate)+' '+c.symbol, color:'#3BD671'},
      {label:'Total COGS', value:fmtS(totalCogs*c.rate)+' '+c.symbol, color:'#B794F4'},
      {label:'Gross Profit', value:fmtS(totalGP*c.rate)+' '+c.symbol, color:'#B794F4'},
      {label:'Net Profit', value:fmtS(totalNP*c.rate)+' '+c.symbol, color:'#3BD671'},
      {label:'Avg GP Margin', value:fmtPct(avgGPM), color:'#3BD671'},
      {label:'Avg NP Margin', value:fmtPct(avgNPM), color:'#3BD671'}
    ];
    document.getElementById('styleKpiGrid').innerHTML = kpis.map(k =>
      '<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas fa-chart-line" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+'">'+k.value+'</div></div>'
    ).join('');
    // Waterfall
    const wfLabels = ['Planned Profit','Labor\\nAdjustment','Overtime','Packing\\nSavings','Clearance\\nSavings','Finance\\nSavings','SG&A\\nSavings','Actual Profit'];
    const wfVals = [wf.planned_profit, wf.labor_cost_adjustment, wf.overtime, wf.saving_packing, wf.saving_clearance, wf.saving_finance, wf.saving_sga, wf.actual_profit];
    // Net Profit by PO (bar chart)
    const npVals = pf.total_net_profit.values.map(v => v * c.rate);
    const npColors = npVals.map(v => v >= 0 ? '#3BD671' : '#DF484A');
    // Margin trend
    const gpMargins = cb.gross_profit_margin.values.map(v => v * 100);
    const npMargins = pf.net_profit_margin.values.map(v => v * 100);
    // Cost per item stacked bar
    const laborV = cb.labor_cost_per_item.values;
    const otV = cb.overtime_per_item.values;
    const packV = cb.packing_per_item.values;
    Plotly.newPlot('styleWaterfall', [{type:'waterfall',orientation:'v',x:wfLabels,y:wfVals,connector:{line:{color:'#A0AEC0',width:1,dash:'dot'}},decreasing:{marker:{color:'#DF484A'}},increasing:{marker:{color:'#3BD671'}},totals:{marker:{color:'#B794F4'}},text:wfVals.map(v=>fmtS(v*c.rate)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}], {margin:{t:50,b:80,l:90,r:30},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},height:280,showlegend:false,yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('styleNpChart', [{type:'bar',x:po,y:npVals,marker:{color:npColors},text:npVals.map(v=>fmtS(v)+' '+c.symbol),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}], {margin:{t:50,b:55,l:70,r:30},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},height:280,showlegend:false,xaxis:{tickangle:-30},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('styleCogsPie', [{type:'pie',labels:po,values:cb.total_cogs.values,text:cb.total_cogs.values.map(v=>fmtS(v*c.rate)),textinfo:'label+percent',textfont:{size:11,color:'#4A5568',family:'Inter'},marker:{colors:['#2DA356','#3BD671','#3BD671','#2DA356','#DF484A'],line:{color:'#CBD5E0',width:2}},hovertemplate:'%{label}<br>'+c.symbol+': %{text}<br>%{percent}<extra></extra>'}], {margin:{t:10,b:10,l:10,r:10},paper_bgcolor:'rgba(0,0,0,0)',height:280,showlegend:true,legend:{orientation:'h',y:-0.1,font:{size:11,color:'#4A5568'}}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('styleMarginChart', [{type:'scatter',mode:'lines+markers',name:'GP Margin',x:po,y:gpMargins,line:{color:'#B794F4',width:3},marker:{size:10,color:'#B794F4'},text:gpMargins.map(v=>v.toFixed(1)+'%'),textposition:'top center',textfont:{size:11,color:'#4A5568',family:'Inter'}},{type:'scatter',mode:'lines+markers',name:'NP Margin',x:po,y:npMargins,line:{color:'#3BD671',width:3},marker:{size:10,color:'#3BD671'},text:npMargins.map(v=>v.toFixed(1)+'%'),textposition:'top center',textfont:{size:11,color:'#4A5568',family:'Inter'}}], {margin:{t:50,b:50,l:60,r:25},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},height:280,showlegend:true,xaxis:{tickangle:-30},legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:'%',gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('styleCostBar', [{type:'bar',name:'Labor',x:po,y:laborV,marker:{color:'#B794F4'}},{type:'bar',name:'Overtime',x:po,y:otV,marker:{color:'#B794F4'}},{type:'bar',name:'Packing',x:po,y:packV,marker:{color:'#3BD671'}}], {margin:{t:40,b:45,l:60,r:25},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},barmode:'stack',height:280,xaxis:{tickangle:-30},legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
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
    let tbl = '<table style="width:100%;border-collapse:collapse;font-size:11px;color:#475569"><thead><tr style="border-bottom:1px solid #CBD5E0"><th style="padding:6px 10px;text-align:left;color:#72839E;font-weight:500">Metric</th>';
    po.forEach(p => { tbl += '<th style="padding:6px 10px;text-align:right;color:#72839E;font-weight:500">'+p+'</th>'; });
    tbl += '<th style="padding:6px 10px;text-align:right;color:#1e293b;font-weight:600">Total</th><th style="padding:6px 10px;text-align:right;color:#2DA356;font-weight:600">Forecast</th><th style="padding:6px 10px;text-align:right;color:#72839E;font-weight:600">Variance</th></tr></thead><tbody>';
    tableRows.forEach((row, i) => {
      const d = cb[row.key];
      const bg = i%2===0 ? 'rgba(0,0,0,0.02)' : 'transparent';
      tbl += '<tr style="border-bottom:1px solid #CBD5E0;background:'+bg+'"><td style="padding:6px 10px;color:#1e293b;font-weight:500">'+row.label+'</td>';
      d.values.forEach(v => { tbl += '<td style="padding:6px 10px;text-align:right">'+row.fmt(v)+'</td>'; });
      tbl += '<td style="padding:6px 10px;text-align:right;color:#1e293b;font-weight:600">'+row.fmt(d.total)+'</td>';
      tbl += '<td style="padding:6px 10px;text-align:right;color:#2DA356">'+row.fmt(d.forecasted||0)+'</td>';
      const vv = d.variance||0;
      tbl += '<td style="padding:6px 10px;text-align:right;color:'+(vv>0?'#3BD671':vv<0?'#DF484A':'#72839E')+'">'+row.fmt(vv)+'</td></tr>';
    });
    tbl += '</tbody></table>';
    document.getElementById('styleTable').innerHTML = tbl;
    // Variance text
    if (s.variance_detail) {
      document.getElementById('styleVariance').innerHTML = '<div style="padding:18px 22px"><div style="font-size:11px;font-weight:600;color:#3BD671;margin-bottom:12px">'+s.variance_title+'</div><div style="font-size:11px;color:#72839E;line-height:1.8;white-space:pre-line">'+s.variance_detail+'</div></div>';
    }
    // Recommendations
    if (s.recommendations) {
      document.getElementById('styleRecommendations').innerHTML = '<div style="padding:18px 22px"><div style="font-size:11px;font-weight:600;color:#3BD671;margin-bottom:12px">'+s.rec_title+'</div><div style="font-size:11px;color:#72839E;line-height:1.8;white-space:pre-line">'+s.recommendations+'</div></div>';
    }
    setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
  } catch(e) { console.error('Style analysis error:', e); }
}

async function loadInvestmentData() {
  try {
    const f = getFilters();
    const r = await authFetch('/api/investment?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month));
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
      {label:'Portfolio Return 2025', value:fmtR(p.return_rate.y2025), color:p.return_rate.y2025>=0?'#3BD671':'#DF484A'},
      {label:'Portfolio Return 2026A', value:fmtR(p.return_rate.y2026a), color:p.return_rate.y2026a>=0?'#3BD671':'#DF484A'},
      {label:'Portfolio Return 2026F', value:fmtR(p.return_rate.y2026f), color:p.return_rate.y2026f>=0?'#3BD671':'#DF484A'},
      {label:'Sharpe Ratio 2026A', value:p.sharpe_ratio.y2026a.toFixed(3), color:'#B794F4'},
      {label:'Max Drawdown 2026A', value:pct(p.mdd.y2026a), color:'#DF484A'},
      {label:'Initial Investment', value:fmtM(inv.initial)+' '+c.symbol, color:'#B794F4'}
    ];
    document.getElementById('invKpiGrid').innerHTML = kpis.map(k =>
      '<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas fa-chart-line" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+';font-size:15px">'+k.value+'</div></div>'
    ).join('');
    const assets = ['Gold','Silver','Swiss Frank'];
    const metrics = ['return_rate','avg_monthly_return','avg_monthly_real_return'];
    const metricNames = ['Annual Return','Avg Monthly Return','Avg Monthly Real Return'];
    const yrs = ['y2025','y2026a','y2026f'];
    const yrLabels = ['2025','2026 Actual','2026 Forecast'];
    const traces_return = [];
    const colors = {'Gold':'#3BD671','Silver':'#A0AEC0','Swiss Frank':'#2DA356'};
    const yrKeys = ['y2025','y2026a','y2026f'];
    assets.forEach((asset, ai) => {
      const data = asset === 'Gold' ? s.gold : asset === 'Silver' ? s.silver : s.swiss;
      traces_return.push({
        type:'bar', name:asset,
        x:yrLabels,
        y:yrKeys.map(yr => data['return_rate'][yr]*100),
        marker:{color:colors[asset],opacity:0.9},
        text:yrKeys.map(yr => pct(data['return_rate'][yr])),
        textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false
      });
    });
    const riskLabels2 = ['Max Drawdown','Downside Deviation','Monthly Volatility'];
    const riskKeys = ['mdd','downside_dev','monthly_risk'];
    const traces_risk = [];
    assets.forEach((asset) => {
      const data = asset === 'Gold' ? s.gold : asset === 'Silver' ? s.silver : s.swiss;
      traces_risk.push({
        type:'scatter', mode:'lines+markers', name:asset,
        x:yrLabels, y:yrKeys.map(yr => data['sharpe_ratio'][yr]),
        line:{color:colors[asset],width:3}, marker:{size:10}
      });
    });
    const traces_rm = [];
    assets.forEach((asset) => {
      const data = asset === 'Gold' ? s.gold : asset === 'Silver' ? s.silver : s.swiss;
      traces_rm.push({
        type:'bar', name:asset,
        x:riskLabels2,
        y:riskKeys.map(k => Math.abs(data[k].y2026a)*100),
        marker:{color:colors[asset]},
        text:riskKeys.map(k => pct(Math.abs(data[k].y2026a))),
        textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false
      });
    });
    const corrLabels = ['Gold','Silver','Swiss Frank'];
    const corrMatrix = [
      [1, s.correlations.gold_silver, s.correlations.gold_swiss],
      [s.correlations.gold_silver, 1, s.correlations.silver_swiss],
      [s.correlations.gold_swiss, s.correlations.silver_swiss, 1]
    ];
    const corrText = corrMatrix.map(row => row.map(v => v.toFixed(3)));
    const pw = s.portfolio.weights || {gold:0.3,silver:0.2,swiss:0.5};
    const perfLabels = ['Annual Return','Return After Loss','Present Value'];
    const perfActual = [inv.annual_return.actual, inv.return_after_loss.actual, inv.pv.actual];
    const perfForecast = [inv.annual_return.forecast, inv.return_after_loss.forecast, inv.pv.forecast];
    Plotly.newPlot('invReturnChart', traces_return, {margin:{t:50,b:70,l:70,r:25},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},barmode:'group',height:280,legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:'%',gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('invRiskChart', traces_risk, {margin:{t:50,b:60,l:60,r:25},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},height:280,legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('invRiskMetrics', traces_rm, {margin:{t:50,b:60,l:60,r:25},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},barmode:'group',height:280,legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:'%',gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('invCorrChart', [{type:'heatmap',z:corrMatrix,x:corrLabels,y:corrLabels,text:corrText,texttemplate:'%{text}',textfont:{size:11,color:'#4A5568',family:'Inter'},colorscale:[[0,'#DF484A'],[0.5,'#CBD5E0'],[1,'#3BD671']],zmin:-1,zmax:1,hovertemplate:'%{x} vs %{y}<br>Correlation: %{z:.3f}<extra></extra>'}], {margin:{t:30,b:50,l:80,r:30},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},height:280,xaxis:{side:'bottom'},yaxis:{autorange:'reversed'}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('invPieChart', [{type:'pie',labels:['Gold','Silver','Swiss Frank'],values:[pw.gold*100,pw.silver*100,pw.swiss*100],text:[(pw.gold*100).toFixed(0)+'%',(pw.silver*100).toFixed(0)+'%',(pw.swiss*100).toFixed(0)+'%'],textinfo:'label+percent',textfont:{size:11,color:'#4A5568',family:'Inter'},marker:{colors:['#3BD671','#A0AEC0','#2DA356'],line:{color:'#CBD5E0',width:3}},hovertemplate:'%{label}<br>%{value:.0f}%<extra></extra>'}], {margin:{t:10,b:10,l:10,r:10},paper_bgcolor:'rgba(0,0,0,0)',height:280,showlegend:true,legend:{orientation:'h',y:-0.1,font:{size:11,color:'#4A5568'}}}, {responsive:true,displayModeBar:false});
    Plotly.newPlot('invPerfChart', [{type:'bar',name:'2026 Actual',x:perfLabels,y:perfActual,marker:{color:'#00B5D8',line:{width:0}},text:perfActual.map(v=>fmtM(v*c.rate)+' '+c.symbol),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false},{type:'bar',name:'2026 Forecast',x:perfLabels,y:perfForecast,marker:{color:'#3BD671',line:{width:0}},text:perfForecast.map(v=>fmtM(v*c.rate)+' '+c.symbol),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}], {margin:{t:50,b:60,l:80,r:25},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'},barmode:'group',height:280,legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}}, {responsive:true,displayModeBar:false});
    if (s.recommendation) {
      document.getElementById('invRecommendation').innerHTML = '<div style="padding:18px 22px;font-size:11px;color:#72839E;line-height:1.8;white-space:pre-line">'+s.recommendation+'</div>';
    }
    if (s.optimal_portfolios) {
      window._optimalData = s.optimal_portfolios;
      renderOptimal(document.getElementById('invOptYear').value);
    }
    setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
  } catch(e) { console.error('Investment data error:', e); }
}

function renderOptimal(yr) {
  var data = window._optimalData;
  if (!data || !data[yr]) { document.getElementById('invOptimal').innerHTML = '<div style="padding:22px;color:#72839E;font-size:12px">No data available for this period.</div>'; return; }
  var list = data[yr];
  var medals = ['1st','2nd','3rd'];
  var labels = ['Highest Sharpe (Best Risk-Adjusted)','Second Best','Third Best'];
  var colors = {'gold':'#3BD671','silver':'#A0AEC0','swiss':'#2DA356'};
  document.getElementById('invOptimal').innerHTML = '<div style="padding:16px 22px;display:grid;grid-template-columns:repeat(3,1fr);gap:14px">' +
    list.map(function(p, i) {
      var total = p.gold + p.silver + p.swiss;
      return '<div style="background:#f8fafc;border:1px solid #CBD5E0;border-radius:10px;padding:16px;position:relative;border-top:3px solid '+(i===0?'#F18C2A':i===1?'#A0AEC0':'#cd7f32')+'">' +
        '<div style="font-size:11px;font-weight:700;color:'+(i===0?'#F18C2A':i===1?'#A0AEC0':'#cd7f32')+';margin-bottom:6px;text-transform:uppercase;letter-spacing:0.5px">'+medals[i]+'</div>' +
        '<div style="font-size:13px;font-weight:700;color:#1e293b;margin-bottom:10px">'+labels[i]+'</div>' +
        '<div style="display:flex;gap:2px;height:18px;border-radius:9px;overflow:hidden;margin-bottom:10px">' +
          '<div style="width:'+p.gold+'%;background:#3BD671" title="Gold '+p.gold+'%"></div>' +
          '<div style="width:'+p.silver+'%;background:#A0AEC0" title="Silver '+p.silver+'%"></div>' +
          '<div style="width:'+p.swiss+'%;background:#2DA356" title="Swiss Frank '+p.swiss+'%"></div>' +
        '</div>' +
        '<div style="font-size:11px;color:#72839E;line-height:2">' +
          '<div><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#3BD671;margin-right:4px"></span>Gold <b>'+p.gold+'%</b></div>' +
          '<div><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#A0AEC0;margin-right:4px"></span>Silver <b>'+p.silver+'%</b></div>' +
          '<div><span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#2DA356;margin-right:4px"></span>Swiss Frank <b>'+p.swiss+'%</b></div>' +
        '</div>' +
        '<div style="border-top:1px solid #CBD5E0;margin-top:10px;padding-top:10px;display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:11px">' +
          '<div><span style="color:#72839E">Return</span><br><b style="color:#3BD671">'+p.return+'%</b></div>' +
          '<div><span style="color:#72839E">Risk</span><br><b style="color:#DF484A">'+p.risk+'%</b></div>' +
          '<div style="grid-column:1/-1"><span style="color:#72839E">Sharpe Ratio</span><br><b style="color:#2DA356">'+p.sharpe+'</b></div>' +
        '</div>' +
      '</div>';
    }).join('') + '</div>';
}

async function loadCashflowData() {
  try {
    const f = getFilters();
    const r = await authFetch('/api/cashflow?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month));
    if (!r.ok) return;
    const s = await r.json();
    if (s.error) { console.error(s.error); return; }
    const c = getCurrency();
    const fmtFull = v => { const n = Math.abs(v*c.rate); if (n >= 1e6) return (v*c.rate/1e6).toFixed(2) + 'M ' + c.symbol; if (n >= 1e3) return (v*c.rate/1e3).toFixed(1) + 'K ' + c.symbol; return (v*c.rate).toFixed(2) + ' ' + c.symbol; };
    const monthNames = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const months = Object.keys(s.monthly_collections).sort((a,b)=>Number(a)-Number(b));
    const netCF = s.total_collections - s.total_spending;
    document.getElementById('cfKpiGrid').innerHTML = [
      {label:'Total Collections',value:fmtFull(s.total_collections),color:'#3BD671',icon:'fa-arrow-down'},
      {label:'Total Spending',value:fmtFull(s.total_spending),color:'#DF484A',icon:'fa-arrow-up'},
      {label:'Net Cash Flow',value:fmtFull(netCF),color:netCF>=0?'#2DA356':'#3BD671',icon:'fa-balance-scale'},
      {label:'Days Receivable',value:s.days_receivable.toFixed(1)+' days',color:'#B794F4',icon:'fa-clock'},
      {label:'Days Payable',value:s.days_payable.toFixed(1)+' days',color:'#3BD671',icon:'fa-clock'}
    ].map(k=>'<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas '+k.icon+'" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+';font-size:15px">'+k.value+'</div></div>').join('');
    animateKPIs();

    const inVals = months.map(m=>s.monthly_collections[m]*c.rate);
    const outVals = months.map(m=>-s.monthly_spending[m]*c.rate);
    const netVals = months.map(m=>s.monthly_net[m]*c.rate);
    const mxLabels = months.map(m=>monthNames[m]);
    const buKeys = Object.keys(s.bu_collections_actual).filter(k=>k!=='TOTAL');
    const buPalette = ['#2DA356','#3BD671','#3BD671','#DF484A','#473FE0'];
    const deptKeys = Object.keys(s.dept_spending_actual).filter(k=>k!=='TOTAL');
    const deptColors = {'Production':'#3BD671','G&A':'#2DA356','Assets':'#2DA356','Financing Expenses':'#DF484A','S&M':'#3BD671','R&D':'#63B3ED'};
    const catKeys = Object.keys(s.category_spending).slice().reverse();
    const buList = [...new Set(Object.values(s.monthly_bu).flatMap(v=>Object.keys(v)))];
    const deptList = [...new Set(Object.values(s.monthly_dept).flatMap(v=>Object.keys(v)))];
    const stackTraces = [];
    const buColors2 = {'CM':'#3BD671','Export':'#2DA356','B2B':'#2DA356','B2C':'#3BD671'};
    buList.forEach(bu => { stackTraces.push({type:'bar',name:'In: '+bu,x:mxLabels,y:months.map(m=>(s.monthly_bu[m]&&s.monthly_bu[m][bu]||0)*c.rate),marker:{color:buColors2[bu]||'#666'}}); });
    const deptColors2 = {'Production':'#DF484A','G&A':'#63B3ED','Assets':'#A0AEC0'};
    deptList.forEach(dept => { stackTraces.push({type:'bar',name:'Out: '+dept,x:mxLabels,y:months.map(m=>-(s.monthly_dept[m]&&s.monthly_dept[m][dept]||0)*c.rate),marker:{color:deptColors2[dept]||'#666'}}); });
    const agingKeys = Object.keys(s.aging);
    const custKeys = Object.keys(s.customer_collections).slice().reverse();
    const catKeysBar = Object.keys(s.category_spending).sort((a,b)=>s.category_spending[b]-s.category_spending[a]).slice(0,10);
    const cfBUs = Object.keys(s.bu_collections_forecast).filter(k=>k!=='TOTAL');
    const cfDepts = Object.keys(s.dept_spending_forecast).filter(k=>k!=='TOTAL');
    const baseL = {paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{size:11,color:'#4A5568'}};
    const baseCfg = {responsive:true,displayModeBar:false};

    Plotly.newPlot('cfMonthlyChart',[{type:'bar',name:'Collections',x:mxLabels,y:inVals,marker:{color:'#3BD671'},text:inVals.map(v=>fmtFull(v/c.rate)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false},{type:'bar',name:'Spending',x:mxLabels,y:outVals,marker:{color:'#DF484A'},text:outVals.map(v=>'(-'+fmtFull(-v/c.rate)+')'),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false},{type:'scatter',mode:'lines+markers',name:'Net Cash Flow',x:mxLabels,y:netVals,line:{color:'#B794F4',width:3},marker:{size:10}}],Object.assign({barmode:'relative',height:280,margin:{t:50,b:60,l:80,r:25},legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);
    Plotly.newPlot('cfBuChart',[{type:'pie',labels:buKeys,values:buKeys.map(k=>s.bu_collections_actual[k]*c.rate),textinfo:'label+percent',textfont:{size:11,color:'#4A5568',family:'Inter'},marker:{colors:buKeys.map((_,i)=>buPalette[i%buPalette.length]),line:{color:'#CBD5E0',width:3}},hovertemplate:'%{label}<br>'+c.symbol+' %{value:,.0f}<br>%{percent}<extra></extra>'}],Object.assign({height:280,margin:{t:10,b:10,l:10,r:10},showlegend:true,legend:{orientation:'h',y:-0.1,font:{size:11,color:'#4A5568'}}},baseL),baseCfg);
    Plotly.newPlot('cfDeptChart',[{type:'bar',x:deptKeys,y:deptKeys.map(k=>s.dept_spending_actual[k]*c.rate),marker:{color:deptKeys.map(k=>deptColors[k]||'#3BD671')},text:deptKeys.map(k=>fmtFull(s.dept_spending_actual[k])),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({height:280,margin:{t:30,b:60,l:80,r:25},xaxis:{tickangle:-30},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);
    Plotly.newPlot('cfCatChart',[{type:'bar',y:catKeys,x:catKeys.map(k=>s.category_spending[k]*c.rate),orientation:'h',marker:{color:'#DF484A'},text:catKeys.map(k=>fmtFull(s.category_spending[k])),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({height:280,margin:{t:15,b:45,l:15,r:15},yaxis:{automargin:true,tickfont:{size:11,color:'#4A5568'}},xaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent',automargin:true}},baseL),baseCfg);
    Plotly.newPlot('cfStackChart',stackTraces,Object.assign({barmode:'relative',height:280,margin:{t:50,b:60,l:80,r:25},legend:{orientation:'h',y:1.2,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);
    if(agingKeys.length){Plotly.newPlot('cfAgingChart',[{type:'bar',x:agingKeys,y:agingKeys.map(k=>s.aging[k].amount*c.rate),marker:{color:agingKeys.map((_,i)=>['#3BD671','#3BD671','#DF484A'][i]||'#666')},text:agingKeys.map(k=>s.aging[k].aging+'<br>'+fmtFull(s.aging[k].amount)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({height:280,margin:{t:30,b:60,l:80,r:25},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);}
    Plotly.newPlot('cfCustChart',[{type:'bar',y:custKeys,x:custKeys.map(k=>s.customer_collections[k]*c.rate),orientation:'h',marker:{color:'#3BD671'},text:custKeys.map(k=>fmtFull(s.customer_collections[k])),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({height:280,margin:{t:15,b:45,l:15,r:15},yaxis:{automargin:true,tickfont:{size:11,color:'#4A5568'}},xaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent',automargin:true}},baseL),baseCfg);
    Plotly.newPlot('cfCatBarChart',[{type:'bar',x:catKeysBar,y:catKeysBar.map(k=>s.category_spending[k]*c.rate),marker:{color:'#B794F4'},text:catKeysBar.map(k=>fmtFull(s.category_spending[k])),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({height:280,margin:{t:30,b:80,l:80,r:25},xaxis:{tickangle:-40},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);
    Plotly.newPlot('cfCollForecastChart',[{type:'bar',name:'Actual',x:cfBUs,y:cfBUs.map(bu=>(s.bu_collections_actual[bu]||0)*c.rate),marker:{color:'#B794F4'},text:cfBUs.map(bu=>fmtFull(s.bu_collections_actual[bu]||0)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false},{type:'bar',name:'Forecast',x:cfBUs,y:cfBUs.map(bu=>s.bu_collections_forecast[bu]*c.rate),marker:{color:'#B794F4'},text:cfBUs.map(bu=>fmtFull(s.bu_collections_forecast[bu])),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({barmode:'group',height:280,margin:{t:50,b:60,l:80,r:25},legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);
    Plotly.newPlot('cfSpendForecastChart',[{type:'bar',name:'Actual',x:cfDepts,y:cfDepts.map(d=>(s.dept_spending_actual[d]||0)*c.rate),marker:{color:'#DF484A'},text:cfDepts.map(d=>fmtFull(s.dept_spending_actual[d]||0)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false},{type:'bar',name:'Forecast',x:cfDepts,y:cfDepts.map(d=>s.dept_spending_forecast[d]*c.rate),marker:{color:'#B794F4'},text:cfDepts.map(d=>fmtFull(s.dept_spending_forecast[d])),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],Object.assign({barmode:'group',height:280,margin:{t:50,b:60,l:80,r:25},legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent'}},baseL),baseCfg);

    const collVar = s.total_forecast_coll ? ((s.total_collections - s.total_forecast_coll)/s.total_forecast_coll*100) : 0;
    const spendVar = s.total_forecast_spend ? ((s.total_spending - s.total_forecast_spend)/s.total_forecast_spend*100) : 0;
    document.getElementById('cfInsights').innerHTML = '<div style="padding:18px 22px;font-size:11px;color:#72839E;line-height:1.8">' + [
      '<div style="margin-bottom:14px"><b style="color:#2DA356;font-size:12px">Cash Flow Summary</b></div>',
      '<div style="margin-bottom:10px"><b style="color:#3BD671">Days Receivable:</b> ' + s.days_receivable.toFixed(1) + ' days — ' + (s.days_receivable <= 30 ? '<span style="color:#3BD671">Healthy</span>' : s.days_receivable <= 60 ? '<span style="color:#3BD671">Moderate — consider tighter credit terms</span>' : '<span style="color:#DF484A">High — collections are slow, review receivables process</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:var(--accent)">Days Payable:</b> ' + s.days_payable.toFixed(1) + ' days — ' + (s.days_payable >= 30 ? '<span style="color:#3BD671">Good leverage of payment terms</span>' : '<span style="color:#3BD671">Paying too fast — consider optimizing cash retention</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#2DA356">Collection vs Forecast:</b> ' + (collVar>=0?'+':'') + collVar.toFixed(1) + '% variance — ' + (collVar>=0?'<span style="color:#3BD671">Above target</span>':'<span style="color:#DF484A">Below target — review sales pipeline</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#DF484A">Spending vs Forecast:</b> ' + (spendVar>=0?'+':'') + spendVar.toFixed(1) + '% variance — ' + (spendVar<=0?'<span style="color:#3BD671">Under budget</span>':'<span style="color:#DF484A">Over budget — review departmental spend</span>') + '</div>',
      '<div style="margin-bottom:10px"><b style="color:#2DA356">Net Cash Position:</b> ' + fmtFull(netCF) + ' — ' + (netCF>=0?'<span style="color:#3BD671">Positive cash generation</span>':'<span style="color:#DF484A">Cash burn — monitor closely</span>') + '</div>'
    ].join('') + '</div>';
    setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
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
  var btn = document.getElementById('refreshBtn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<i class="fas fa-sync-alt fa-spin"></i>'; }
  try { await authFetch('/api/refresh?_=' + Date.now(), {cache:'no-store'}); } catch(e) {}
  _refreshCurrentTab();
  if (btn) setTimeout(function(){ btn.disabled = false; btn.innerHTML = '<i class="fas fa-sync-alt"></i> Refresh'; }, 1500);
}

// ---- Blur Toggle ----
let _blurred = localStorage.getItem('dashBlur') === 'true';
let _sidebarHidden = localStorage.getItem('dashSidebarHidden') === 'true';
if (_blurred) document.body.classList.add('blur-mode');
if (_sidebarHidden) { document.body.classList.add('sidebar-hidden'); document.querySelector('.sidebar').classList.add('hidden'); }
function toggleSidebar() {
  _sidebarHidden = !_sidebarHidden;
  document.body.classList.toggle('sidebar-hidden', _sidebarHidden);
  document.querySelector('.sidebar').classList.toggle('hidden', _sidebarHidden);
  localStorage.setItem('dashSidebarHidden', _sidebarHidden);
  const mainEl = document.querySelector('.main');
  function _doResize() {
    document.querySelectorAll('.js-plotly-plot').forEach(el => {
      Plotly.relayout(el, {autosize:true}).then(()=>Plotly.Plots.resize(el));
    });
  }
  _doResize();
  setTimeout(_doResize, 100);
  setTimeout(_doResize, 350);
}
function toggleBlur() {
  _blurred = !_blurred;
  document.body.classList.toggle('blur-mode', _blurred);
  localStorage.setItem('dashBlur', _blurred);
  document.querySelector('#blurToggle i').className = _blurred ? 'fas fa-eye-slash' : 'fas fa-eye';
  document.getElementById('blurToggle').classList.toggle('active', _blurred);
  const mb = document.getElementById('mobileBlurIcon');
  if (mb) { mb.className = _blurred ? 'fas fa-eye-slash' : 'fas fa-eye'; mb.closest('.mnav-item').classList.toggle('active', _blurred); }
  updatePlotlyBlur();
}
let _darkMode = localStorage.getItem('dashDarkMode') === 'true';
if (_darkMode) document.body.classList.add('dark-mode');
function toggleDarkMode() {
  _darkMode = !_darkMode;
  document.body.classList.toggle('dark-mode', _darkMode);
  localStorage.setItem('dashDarkMode', _darkMode);
  document.querySelector('#darkToggle i').className = _darkMode ? 'fas fa-sun' : 'fas fa-moon';
  document.getElementById('darkToggle').classList.toggle('active', _darkMode);
  updatePlotlyDarkMode();
}
function updatePlotlyDarkMode() {
  const dark = document.body.classList.contains('dark-mode');
  const bgColor = dark ? '#1C2434' : 'rgba(0,0,0,0)';
  const fontColor = dark ? '#EDF2F7' : '#4A5568';
  const tickColor = dark ? '#A0AEC0' : '#4A5568';
  document.querySelectorAll('.js-plotly-plot').forEach(el => {
    Plotly.relayout(el, {
      'paper_bgcolor': bgColor,
      'plot_bgcolor': bgColor,
      'font.color': fontColor,
      'xaxis.tickfont.color': tickColor,
      'yaxis.tickfont.color': tickColor,
      'hoverlabel.bordercolor': dark ? '#3BD671' : '#3BD671',
      'hoverlabel.bgcolor': dark ? '#1C2434' : '#ffffff',
      'hoverlabel.font.color': dark ? '#EDF2F7' : '#1e293b'
    });
  });
}
setTimeout(() => { document.querySelector('#darkToggle i').className = _darkMode ? 'fas fa-sun' : 'fas fa-moon'; document.getElementById('darkToggle').classList.toggle('active', _darkMode); }, 0);
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

// ---- Edit Mode ----
let _editMode = false;
let _overrides = {values:{}, labels:{}, notes:{}, layout:{}};
let _editTarget = null;
let _undoStack = [];
let _redoStack = [];
const MAX_UNDO = 50;

function _snapshotOverrides() {
  return JSON.parse(JSON.stringify(_overrides));
}

function _pushUndo() {
  _undoStack.push(_snapshotOverrides());
  if (_undoStack.length > MAX_UNDO) _undoStack.shift();
  _redoStack = [];
}

function loadOverrides() {
  try {
    const saved = localStorage.getItem('viva_overrides');
    if (saved) { _overrides = JSON.parse(saved); }
  } catch(e) {}
  authFetch('/api/overrides').then(r => r.ok ? r.json() : null).then(server => {
    if (server && server.values && Object.keys(server.values).length > Object.keys(_overrides.values||{}).length) {
      _overrides = server;
      localStorage.setItem('viva_overrides', JSON.stringify(_overrides));
    }
  }).catch(() => {});
}

function toggleEditMode() {
  _editMode = !_editMode;
  document.body.classList.toggle('edit-mode', _editMode);
  document.getElementById('editToggle').classList.toggle('active', _editMode);
  const me = document.getElementById('mobileEditIcon');
  if (me) { me.closest('.mnav-item').classList.toggle('active', _editMode); }
  if (_editMode) {
    markEditables();
    initChartResize();
  } else {
    document.querySelectorAll('.editable').forEach(el => {
      el.contentEditable = 'false';
      el.classList.remove('editable','editing');
    });
    document.querySelectorAll('.chart-resize-handle').forEach(h => h.remove());
  }
}

function getActiveTabId() {
  var el = document.querySelector('.tab-content.active');
  return el ? el.id.replace('tab-','') : 'dashboard';
}

function generateEditId(el) {
  const tab = el.closest('.tab-content');
  const tabId = tab ? tab.id.replace('tab-','') : 'global';
  const parent = el.closest('[id]');
  const parentId = parent ? parent.id : 'none';
  const field = el.classList.contains('kpi-value') ? 'v' : el.classList.contains('kpi-label') ? 'l' : el.classList.contains('pnl-label') ? 'pl' : el.classList.contains('num') ? 'n' : 't';
  const sibIdx = Array.from(el.parentElement.children).indexOf(el);
  return tabId + '/' + parentId + '/' + field + '.' + sibIdx;
}

function _scopeQuery(selector) {
  var active = document.querySelector('.tab-content.active');
  if (!active) return document.querySelectorAll(selector);
  return active.querySelectorAll(selector);
}

function markEditables() {
  var activeTab = document.querySelector('.tab-content.active');
  if (!activeTab) return;
  activeTab.querySelectorAll('.kpi-value, .kpi-label, .chart-header h5, .pnl-label, td.num').forEach(el => {
    if (el.closest('.edit-save-bar') || el.closest('.edit-overlay')) return;
    el.classList.add('editable');
    el.contentEditable = _editMode ? 'false' : 'false';
    const editId = generateEditId(el);
    el.setAttribute('data-edit-id', editId);
    const type = el.classList.contains('kpi-value') || el.classList.contains('num') ? 'value' : 'label';
    el.setAttribute('data-type', type);
    el.onclick = function(e) {
      if (!_editMode) return;
      e.stopPropagation();
      document.querySelectorAll('.editable.editing').forEach(x => x.classList.remove('editing'));
      el.classList.add('editing');
      _editTarget = {el, type, editId, original: el.textContent.trim()};
      openEditModal(type, el.textContent.trim(), el);
    };
  });
  applyAllOverrides();
  initChartResize();
}

function applyAllOverrides() {
  if (!_overrides) return;
  var activeTab = document.querySelector('.tab-content.active');
  if (!activeTab) return;
  activeTab.querySelectorAll('.editable').forEach(el => {
    const editId = el.getAttribute('data-edit-id');
    if (!editId) return;
    if (_overrides.values && _overrides.values[editId] !== undefined) { el.textContent = _overrides.values[editId]; el.style.borderBottom = '2px solid #F18C2A'; }
    if (_overrides.labels && _overrides.labels[editId] !== undefined) { el.textContent = _overrides.labels[editId]; el.style.borderBottom = '2px solid #F18C2A'; }
  });
  highlightDirty();
  applyLayoutOverrides();
}

function openEditModal(type, currentVal, el) {
  const overlay = document.getElementById('editOverlay');
  const title = document.getElementById('editModalTitle');
  const input = document.getElementById('editModalInput');
  const textarea = document.getElementById('editModalTextarea');
  title.textContent = type === 'label' ? 'Edit Label' : 'Edit Value';
  input.style.display = 'block';
  textarea.style.display = 'none';
  input.value = currentVal;
  input.focus();
  overlay.classList.add('active');
  input.onkeydown = function(e) { if (e.key === 'Enter') confirmEdit(); if (e.key === 'Escape') closeEditModal(); };
}

function closeEditModal() {
  document.getElementById('editOverlay').classList.remove('active');
  if (_editTarget && _editTarget.el) _editTarget.el.classList.remove('editing');
  _editTarget = null;
}

function confirmEdit() {
  if (!_editTarget) return;
  const val = document.getElementById('editModalInput').value.trim();
  if (!val && val !== '0') { closeEditModal(); return; }
  _pushUndo();
  const editId = _editTarget.editId;
  const edType = _editTarget.type;

  if (_editTarget._dept) {
    const numVal = parseFloat(val.replace(/[^0-9.\-]/g, ''));
    if (isNaN(numVal)) { closeEditModal(); return; }
    const prefix = _editTarget._prefix || 'pnl';
    const storeKey = prefix + '_dept_' + _editTarget._dept + '_' + (_editTarget._isActual ? 'act' : 'fct');
    if (!_overrides.values) _overrides.values = {};
    _overrides.values[storeKey] = numVal;
    closeEditModal();
    _persistOverrides();
    _refreshCurrentTab();
    return;
  }

  if (_editTarget.el) _editTarget.el.textContent = val;
  if (edType === 'value') {
    if (!_overrides.values) _overrides.values = {};
    _overrides.values[editId] = val;
  } else {
    if (!_overrides.labels) _overrides.labels = {};
    _overrides.labels[editId] = val;
  }
  if (_editTarget.el) _editTarget.el.classList.remove('editing');
  closeEditModal();
  highlightDirty();
  _persistOverrides();
}

function _persistOverrides() {
  localStorage.setItem('viva_overrides', JSON.stringify(_overrides));
  authFetch('/api/overrides', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(_overrides)}).catch(()=>{});
}

function undoEdit() {
  if (_undoStack.length === 0) return;
  _redoStack.push(_snapshotOverrides());
  _overrides = _undoStack.pop();
  localStorage.setItem('viva_overrides', JSON.stringify(_overrides));
  authFetch('/api/overrides', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(_overrides)}).catch(()=>{});
  _refreshCurrentTab();
}

function redoEdit() {
  if (_redoStack.length === 0) return;
  _undoStack.push(_snapshotOverrides());
  _overrides = _redoStack.pop();
  localStorage.setItem('viva_overrides', JSON.stringify(_overrides));
  authFetch('/api/overrides', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(_overrides)}).catch(()=>{});
  _refreshCurrentTab();
}

function _refreshCurrentTab() {
  var tab = getActiveTabId();
  if (tab === 'dashboard') { loadData(); }
  else if (tab === 'pnl') { window._pnlLoaded = false; loadPnlData(); }
  else if (tab === 'sales') { window._salesLoaded = false; loadSalesData(); }
  else if (tab === 'expenses') { window._expLoaded = false; loadExpensesData(); }
  else if (tab === 'style') { window._styleLoaded = false; loadStyleAnalysis(); }
  else if (tab === 'investment') { window._invLoaded = false; loadInvestmentData(); }
  else if (tab === 'cashflow') { window._cfLoaded = false; loadCashflowData(); }
}

function highlightDirty() {
  document.querySelectorAll('.editable').forEach(el => {
    const k = el.getAttribute('data-edit-id');
    const isDirty = (_overrides.values && _overrides.values[k]) || (_overrides.labels && _overrides.labels[k]);
    el.style.borderBottom = isDirty ? '2px solid #F18C2A' : '';
  });
}

async function saveAllEdits() {
  try {
    const r = await authFetch('/api/overrides', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify(_overrides)
    });
    if (r.ok) { alert('Changes saved!'); }
  } catch(e) { alert('Error saving: ' + e.message); }
}

async function clearAllEdits() {
  if (!confirm('Clear all edits? This cannot be undone.')) return;
  _overrides = {values:{}, labels:{}, notes:{}, layout:{}};
  _undoStack = [];
  _redoStack = [];
  localStorage.removeItem('viva_overrides');
  try {
    await authFetch('/api/overrides/clear', {method:'POST'}).catch(()=>{});
  } catch(e) {}
  document.querySelectorAll('.editable').forEach(el => { el.style.borderBottom = ''; });
  location.reload();
}

document.addEventListener('keydown', function(e) {
  if (!_editMode) return;
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) { e.preventDefault(); undoEdit(); }
  if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) { e.preventDefault(); redoEdit(); }
});

// ---- Chart Resize ----
function initChartResize() {
  document.querySelectorAll('.chart-card').forEach(card => {
    let handle = card.querySelector('.chart-resize-handle');
    if (_editMode && !handle) {
      handle = document.createElement('div');
      handle.className = 'chart-resize-handle';
      card.appendChild(handle);
    } else if (!_editMode && handle) {
      handle.remove();
      return;
    } else if (!_editMode) return;
    handle.addEventListener('mousedown', function(e) {
      e.preventDefault(); e.stopPropagation();
      const startY = e.clientY;
      const body = card.querySelector('.chart-body, .pnl-body');
      const startH = body ? body.offsetHeight : 220;
      document.body.style.cursor = 'ns-resize';
      document.body.style.userSelect = 'none';
      function onMove(ev) {
        const newH = Math.max(100, startH + ev.clientY - startY);
        if (body) body.style.minHeight = newH + 'px';
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        const body = card.querySelector('.chart-body, .pnl-body');
        const chartEl = body ? (body.querySelector('.js-plotly-plot') || body.querySelector('[id]')) : null;
        const key = chartEl ? chartEl.id : card.querySelector('h5').textContent.trim().replace(/\s+/g,'_').substring(0,30);
        if (body && key) {
          if (!_overrides.layout) _overrides.layout = {};
          _overrides.layout[key] = {h: body.offsetHeight};
          _persistOverrides();
        }
        if (chartEl && chartEl.classList && chartEl.classList.contains('js-plotly-plot')) {
          Plotly.Plots.resize(chartEl);
        }
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  });
}

function applyLayoutOverrides() {
  if (!_overrides || !_overrides.layout) return;
  document.querySelectorAll('.chart-card').forEach(card => {
    const chartEl = card.querySelector('.chart-body [id], .pnl-body [id]');
    const key = chartEl ? chartEl.id : (card.querySelector('h5') ? card.querySelector('h5').textContent.trim().replace(/\s+/g,'_').substring(0,30) : null);
    if (key && _overrides.layout[key]) {
      const body = card.querySelector('.chart-body, .pnl-body');
      if (body && _overrides.layout[key].h) {
        body.style.minHeight = _overrides.layout[key].h + 'px';
        body.style.height = _overrides.layout[key].h + 'px';
        setTimeout(() => {
          const plot = body.querySelector('.js-plotly-plot');
          if (plot) Plotly.Plots.resize(plot);
        }, 200);
      }
    }
  });
}

loadOverrides();
(function _liveClock() {
  function tick() {
    var el = document.getElementById('timestamp');
    if (el) {
      var now = new Date();
      el.textContent = now.toLocaleDateString() + ', ' + now.toLocaleTimeString();
    }
  }
  tick();
  setInterval(tick, 1000);
})();

async function fetchRates() {
  try {
    var r = await fetch('/api/rates');
    var rates = await r.json();
    if (rates && rates.USD) {
      var symbols = {EGP:'EGP',USD:'$',EUR:'\u20AC',GBP:'\u00A3',CHF:'CHF',KWD:'KWD',BHD:'BHD',OMR:'OMR',JOD:'JOD',SAR:'SAR',AED:'AED'};
      var locales = {EGP:'en-US',USD:'en-US',EUR:'de-DE',GBP:'en-GB',CHF:'de-CH',KWD:'ar-KW',BHD:'ar-BH',OMR:'ar-OM',JOD:'ar-JO',SAR:'ar-SA',AED:'ar-AE'};
      var names = {EGP:'Egyptian Pound',USD:'US Dollar',EUR:'Euro',GBP:'British Pound',CHF:'Swiss Franc',KWD:'Kuwaiti Dinar',BHD:'Bahraini Dinar',OMR:'Omani Rial',JOD:'Jordanian Dinar',SAR:'Saudi Riyal',AED:'UAE Dirham'};
      for (var k in rates) {
        if (CURRENCIES[k]) { CURRENCIES[k].rate = rates[k]; }
        else { CURRENCIES[k] = {rate: rates[k], symbol: symbols[k]||k, locale: locales[k]||'en-US'}; }
      }
      var sel = document.getElementById('currencySelect');
      if (sel) {
        var val = sel.value;
        var html = '';
        var order = ['EGP','USD','EUR','GBP','CHF','KWD','BHD','OMR','JOD','SAR','AED'];
        for (var i = 0; i < order.length; i++) {
          var k = order[i];
          var c = CURRENCIES[k];
          if (!c) continue;
          var rateStr = k === 'EGP' ? '1.00' : (1 / c.rate).toFixed(2);
          html += '<option value="' + k + '"' + (k === val ? ' selected' : '') + '>' + k + ' (' + c.symbol + ') ' + rateStr + '</option>';
        }
        sel.innerHTML = html;
      }
    }
  } catch(e) { console.warn('Live rates unavailable, using defaults'); }
}
fetchRates().then(function() { loadData(); });

// ---- Tab Switching ----
let _currentTab = 'dashboard';
function getActiveTab() { return _currentTab; }
function getFilters() {
  const monthChecked = Array.from(document.querySelectorAll('#monthPanel input:checked')).map(c=>c.value);
  const buChecked = Array.from(document.querySelectorAll('#buPanel input:checked')).map(c=>c.value);
  const periodChecked = Array.from(document.querySelectorAll('#periodPanel input:checked')).map(c=>c.value);
  const month = (monthChecked.length===0||monthChecked.includes('all'))?'all':monthChecked.join(',');
  const bu = (buChecked.length===0||buChecked.includes('all'))?'all':buChecked.join(',');
  const period = (periodChecked.length===0||periodChecked.includes('ytd'))?'ytd':periodChecked.join(',');
  return {period:period, bu:bu, month:month};
}
function togglePanel(id) {
  const p = document.getElementById(id);
  const isOpen = p.style.display !== 'none';
  document.querySelectorAll('.ms-panel').forEach(el => el.style.display = 'none');
  if (!isOpen) p.style.display = 'block';
}
function msClick(row, group) {
  var cb = row.querySelector('input[type=checkbox]');
  cb.checked = !cb.checked;
  var panel = document.getElementById(group+'Panel');
  var allVal = group==='period'?'ytd':'all';
  var allCb = panel.querySelector('input[value="'+allVal+'"]');
  if (cb.value === allVal) {
    panel.querySelectorAll('input[type=checkbox]').forEach(c => { if(c!==allCb) c.checked=false; });
  } else {
    allCb.checked = false;
    if (!panel.querySelector('input:checked')) allCb.checked = true;
  }
  updateMsLabel(group);
  debouncedFilter();
}
function updateMsLabel(group) {
  var panel = document.getElementById(group+'Panel');
  var checked = Array.from(panel.querySelectorAll('input:checked')).map(c=>c.value);
  var labels = {all:'All',ytd:'YTD',q1:'Q1',q2:'Q2',q3:'Q3',Export:'Export',B2B:'B2B',B2C:'B2C',CM:'CM',1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul'};
  var allVal = group==='period'?'ytd':'all';
  var text = (checked.length===0||checked.includes(allVal))?(allVal==='ytd'?'YTD':'All'):checked.map(v=>labels[v]||v).join('+');
  var el = document.getElementById(group==='period'?'periodLabel':group==='bu'?'buLabel':'monthLabel');
  if (el) el.textContent = text;
}
var _filterTimer = null;
function debouncedFilter() {
  if (_filterTimer) clearTimeout(_filterTimer);
  _filterTimer = setTimeout(function() { onFilterChange(); }, 300);
}
document.addEventListener('click', function(e) {
  if (!e.target.closest('.ms-wrap')) document.querySelectorAll('.ms-panel').forEach(el => el.style.display = 'none');
});
function onFilterChange() {
  const tab = _currentTab;
  if (tab === 'dashboard') loadData();
  else if (tab === 'pnl') { window._pnlLoaded = false; loadPnlData(); }
  else if (tab === 'sales') { window._salesLoaded = false; loadSalesData(); }
  else if (tab === 'expenses') { window._expLoaded = false; loadExpensesData(); }
  else if (tab === 'style') { window._styleLoaded = false; loadStyleAnalysis(); }
  else if (tab === 'investment') { window._invLoaded = false; loadInvestmentData(); }
  else if (tab === 'cashflow') { window._cfLoaded = false; loadCashflowData(); }
}
function switchTab(tab) {
  _currentTab = tab;
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  const tabEl = document.getElementById('tab-' + tab);
  tabEl.classList.add('active');
  tabEl.style.animation = 'none';
  tabEl.offsetHeight;
  tabEl.style.animation = 'fadeInUp .4s cubic-bezier(.4,0,.2,1) both';
  document.querySelectorAll('.sidebar .nav-item').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.sidebar .nav-item').forEach(el => {
    const t = el.textContent.trim().toLowerCase();
    const match = tab === 'pnl' ? t.includes('p&l') : (tab === 'cashflow' ? t.includes('cash flow') : t.includes(tab));
    if (match) el.classList.add('active');
  });
  document.querySelectorAll('.mobile-nav .mnav-item').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.mobile-nav .mnav-item').forEach(el => {
    const t = el.textContent.trim().toLowerCase();
    if (t.includes(tab === 'pnl' ? 'p&l' : tab)) el.classList.add('active');
  });
  if (tab === 'dashboard') loadData();
  if (tab === 'pnl') { window._pnlLoaded = true; loadPnlData(); }
  if (tab === 'sales') { window._salesLoaded = true; loadSalesData(); }
  if (tab === 'expenses') { window._expLoaded = true; loadExpensesData(); }
  if (tab === 'style') { window._styleLoaded = true; loadStyleAnalysis(); }
  if (tab === 'investment') { window._invLoaded = true; loadInvestmentData(); }
  if (tab === 'cashflow') { window._cfLoaded = true; loadCashflowData(); }
  setTimeout(() => { updatePlotlyBlur(); markEditables(); animateKPIs(); }, 300);
}

// ---- P&L Actual vs Forecast ----
const monthNameMap = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul'};

function pnlFmt(v) { const c=getCurrency(); const n=Math.abs(v);
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M '+c.symbol;
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K '+c.symbol;
  return n.toFixed(0)+' '+c.symbol;
}

function pnlFmtShort(v) { return pnlFmt(v); }

async function loadPnlData() {
  const f = getFilters();
  const r = await authFetch('/api/pnl_forecast?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month));
  const d = await r.json();
  const ytd = d.ytd;
  const monthly = d.monthly;
  const activeMonths = monthly.map(m => monthNameMap[m.month] || ('M'+m.month));

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

  const colAct = '#2DA356';
  const colFct = '#A0AEC0';

  // 1. Net Sales chart - use Sales Summary actuals from PNL Dashboard
  const nsAct = monthly.map(m => m.net_sales.actual * cr.rate);
  const nsFct = monthly.map(m => m.net_sales.forecast * cr.rate);
  buildGroupedBar('pnlSalesChart', activeMonths, [
    {name:'Actual', y:nsAct, color:colAct},
    {name:'Forecast', y:nsFct, color:colFct}
  ], 'Net Sales');

  // 2. Gross Profit chart
  const gpAct = monthly.map(m => m.gross_profit.actual * cr.rate);
  const gpFct = monthly.map(m => m.gross_profit.forecast * cr.rate);
  buildGroupedBar('pnlGpChart', activeMonths, [
    {name:'Actual', y:gpAct, color:'#4A5568'},
    {name:'Forecast', y:gpFct, color:colFct}
  ], 'Gross Profit');

  // 3. Net Income chart
  const niAct = monthly.map(m => m.net_income.actual * cr.rate);
  const niFct = monthly.map(m => m.net_income.forecast * cr.rate);
  buildGroupedBar('pnlNiChart', activeMonths, [
    {name:'Actual', y:niAct, color:'#3BD671'},
    {name:'Forecast', y:niFct, color:colFct}
  ], 'Net Income');

  // 4. GP Margin trend
  const gpMarginAct = monthly.map(m => m.net_sales.actual > 0 ? (m.gross_profit.actual / m.net_sales.actual * 100) : 0);
  const gpMarginFct = monthly.map(m => m.net_sales.forecast > 0 ? (m.gross_profit.forecast / m.net_sales.forecast * 100) : 0);
  Plotly.newPlot('pnlMarginChart', [
    {type:'scatter', mode:'lines+markers', name:'Actual Margin', x:activeMonths, y:gpMarginAct, line:{color:'#3BD671', width:3}, marker:{size:8, color:'#3BD671'}},
    {type:'scatter', mode:'lines+markers', name:'Forecast Margin', x:activeMonths, y:gpMarginFct, line:{color:'#A0AEC0', width:3, dash:'dot'}, marker:{size:8, color:'#A0AEC0'}}
  ], {margin:{t:15,b:35,l:50,r:15}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#A0AEC0',size:11},
      yaxis:{ticksuffix:'%', gridcolor:'transparent', rangemode:'tozero'}, height:280,
      hovermode:'x unified', legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}}},
    {responsive:true, displayModeBar:false});

  // 5. Department Expenses
  const depts = Object.keys(d.dept_expenses);
  const deptAct = depts.map(k => {
    const ovKey = 'pnl_dept_'+k+'_act';
    const raw = _overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? _overrides.values[ovKey] : d.dept_expenses[k].actual;
    return (typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/[^0-9.\-]/g,'')) || 0) * (_overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? 1 : cr.rate);
  });
  const deptFct = depts.map(k => {
    const ovKey = 'pnl_dept_'+k+'_fct';
    const raw = _overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? _overrides.values[ovKey] : d.dept_expenses[k].forecast;
    return (typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/[^0-9.\-]/g,'')) || 0) * (_overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? 1 : cr.rate);
  });
  Plotly.newPlot('pnlDeptChart', [
    {type:'bar', orientation:'h', name:'Actual', x:deptAct, y:depts, marker:{color:'#B794F4'}, text:deptAct.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false},
    {type:'bar', orientation:'h', name:'Forecast', x:deptFct, y:depts, marker:{color:'#3BD671'}, text:deptFct.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false}
  ], {margin:{t:15,b:25,l:10,r:10}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#4A5568',size:11},
      barmode:'group', bargap:0.25, bargroupgap:0.1, height:Math.max(280, depts.length*45), hovermode:'y unified',
      legend:{orientation:'h',y:1.05,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},
      yaxis:{automargin:true, tickfont:{size:11,color:'#4A5568'}},
      xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true}}, {responsive:true, displayModeBar:false});
  document.getElementById('pnlDeptChart').on('plotly_click', function(data) {
    if (!_editMode) return;
    const pt = data.points[0];
    const dept = pt.y;
    const traceIdx = pt.curveNumber;
    const isActual = traceIdx === 0;
    const currentVal = isActual ? d.dept_expenses[dept].actual * cr.rate : d.dept_expenses[dept].forecast * cr.rate;
    _editTarget = {el:null, type:'value', editId:'pnl_dept_'+dept+'_'+(isActual?'act':'fct'), original: pnlFmtShort(currentVal), _dept: dept, _isActual: isActual, _rate: cr.rate, _raw: isActual ? d.dept_expenses[dept].actual : d.dept_expenses[dept].forecast};
    openEditModal('value', pnlFmtShort(currentVal), null);
  });

  // 6. P&L Waterfall
  const pnlLabels = ['Net Sales', 'COGS', 'Gross Profit', 'Expenses', 'Net Income'];
  const pnlActVals = [ytd.net_sales.actual, -ytd.cogs.actual, 0, -ytd.expenses.actual, 0].map(v=>v*cr.rate);
  const pnlFctVals = [ytd.net_sales.forecast, -ytd.cogs.forecast, 0, -ytd.expenses.forecast, 0].map(v=>v*cr.rate);
  const pnlActText = [ytd.net_sales.actual, -ytd.cogs.actual, ytd.gross_profit.actual, -ytd.expenses.actual, ytd.net_income.actual].map(v=>pnlFmtShort(v*cr.rate));
  const pnlFctText = [ytd.net_sales.forecast, -ytd.cogs.forecast, ytd.gross_profit.forecast, -ytd.expenses.forecast, ytd.net_income.forecast].map(v=>pnlFmtShort(v*cr.rate));
  const waterfallMeasures = ['relative', 'relative', 'total', 'relative', 'total'];
  Plotly.newPlot('pnlWaterfallChart', [
    {type:'waterfall', name:'Actual', x:pnlLabels, y:pnlActVals, measure:waterfallMeasures,
     decreasing:{marker:{color:'#DF484A'}}, increasing:{marker:{color:'#3BD671'}}, totals:{marker:{color:'#B794F4'}},
     connector:{line:{color:'rgba(255,255,255,0.2)', width:2}},
     cliponaxis:false,
     text:pnlActText, textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'}},
    {type:'waterfall', name:'Forecast', x:pnlLabels, y:pnlFctVals, measure:waterfallMeasures,
     decreasing:{marker:{color:'#DF484A'}}, increasing:{marker:{color:'#3BD671'}}, totals:{marker:{color:'#3BD671'}},
     connector:{line:{color:'rgba(255,255,255,0.1)', width:1}},
     cliponaxis:false,
     text:pnlFctText, textposition:'outside', textfont:{size:11,color:'#4A5568',family:'Inter'},
     opacity:0.7}
  ], {margin:{t:55,b:70,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#A0AEC0',size:11},
      autosize:true, height:380, hovermode:'x unified',
      legend:{orientation:'h',y:1.08,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},
      yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true}}, {responsive:true, displayModeBar:false, cliponaxis:false});

  // 7. P&L Summary Table
  const pnlLines = [
    {label:'Net Sales', a:ytd.net_sales.actual*cr.rate, f:ytd.net_sales.forecast*cr.rate},
    {label:'COGS', a:-ytd.cogs.actual*cr.rate, f:ytd.cogs.forecast*cr.rate, neg:true},
    {label:'Gross Profit', a:ytd.gross_profit.actual*cr.rate, f:ytd.gross_profit.forecast*cr.rate, bold:true},
    {label:'Expenses', a:-ytd.expenses.actual*cr.rate, f:ytd.expenses.forecast*cr.rate, neg:true},
    {label:'Net Income', a:ytd.net_income.actual*cr.rate, f:ytd.net_income.forecast*cr.rate, bold:true}
  ];
  document.getElementById('pnlSummaryTable').innerHTML =
    '<div style="padding:12px 18px;width:100%"><table style="width:100%;border-collapse:collapse;font-size:12px">' +
    '<tr style="color:var(--text-secondary);font-size:11px;text-transform:uppercase;letter-spacing:.5px">' +
    '<td style="padding:8px 12px">Line Item</td><td style="padding:8px 12px;text-align:right">Actual</td>' +
    '<td style="padding:8px 12px;text-align:right">Forecast</td><td style="padding:8px 12px;text-align:right">Variance</td>' +
    '<td style="padding:8px 12px;text-align:right">Var %</td></tr>' +
    pnlLines.map(l => {
      const varAmt = l.a - (l.neg ? -l.f : l.f);
      const varPct = l.f !== 0 ? ((l.a - (l.neg ? -l.f : l.f)) / (l.neg ? -l.f : l.f) * 100).toFixed(1) : 'N/A';
      const isGood = l.neg ? (varAmt >= 0) : (varAmt >= 0);
      return '<tr' + (l.bold ? ' style="border-top:2px solid #475569;font-weight:700"' : '') + '>' +
        '<td style="padding:8px 12px;color:' + (l.bold ? '#1e293b' : 'var(--text-secondary)') + ';font-weight:' + (l.bold ? '700' : '500') + '">' + l.label + '</td>' +
        '<td class="num" style="padding:8px 12px;text-align:right;font-weight:600;font-size:13px">' + pnlFmt(l.a) + '</td>' +
        '<td class="num" style="padding:8px 12px;text-align:right;color:#72839E;font-size:13px">' + pnlFmt(l.neg ? -l.f : l.f) + '</td>' +
        '<td class="num" style="padding:8px 12px;text-align:right;color:' + (isGood ? 'var(--green)' : 'var(--red)') + ';font-weight:600;font-size:13px">' +
        (varAmt >= 0 ? '+' : '') + pnlFmt(varAmt) + '</td>' +
        '<td class="num" style="padding:8px 12px;text-align:right;color:' + (isGood ? 'var(--green)' : 'var(--red)') + ';font-size:12px">' +
        (varPct !== 'N/A' ? (varAmt >= 0 ? '+' : '') + varPct + '%' : 'N/A') + '</td></tr>';
    }).join('') + '</table></div>';
  setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
}

function buildGroupedBar(divId, labels, seriesArr, extraLayout) {
  const c = getCurrency();
  const traces = seriesArr.map(s => ({
    type:'bar', name:s.name, x:labels, y:s.y,
    marker:{color:s.color, opacity:0.85},
    text:s.y.map(v => fmtShort(v)), textposition:'outside',
    textfont:{size:11,color:'#4A5568',family:'Inter'}, cliponaxis:false
  }));
  Plotly.newPlot(divId, traces, {
    margin:{t:65,b:55,l:70,r:30}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
    font:{color:'#A0AEC0',size:11}, barmode:'group', height:320, hovermode:'x unified',
    legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},
    yaxis:{ticksuffix:' '+c.symbol,gridcolor:'transparent',automargin:true},
    xaxis:{automargin:true},
    cliponaxis:false,
    ...(extraLayout||{})
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
  const f = getFilters();
  const qs = '?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month);
  const [sf, d] = await Promise.all([
    authFetch('/api/sales_forecast'+qs).then(r=>r.json()),
    authFetch('/api/data'+qs).then(r=>r.json())
  ]);
  const T_gs = sf.ytd.gross_sales.actual;
  const T_ret = sf.ytd.returns.actual;
  const T_disc = sf.ytd.discounts.actual;
  const T_ns = sf.ytd.net_sales.actual;
  const cr = getCurrency();
  const months = sf.monthly.map(m => monthNameMap[m.month] || ('M'+m.month));
  const dMonths = d.months.slice(0, sf.monthly.length);

  const nsPct = T_gs > 0 ? ((T_ns / T_gs) * 100).toFixed(1) : '0';
  document.getElementById('salesKpiGrid').innerHTML =
    '<div class="kpi-card kpi-gs"><div class="kpi-icon"><i class="fas fa-chart-line"></i></div><div class="kpi-label">Gross Sales</div><div class="kpi-value">'+salesFmt(T_gs*cr.rate)+'</div><div class="kpi-sub">Budget: '+salesFmt(sf.ytd.gross_sales.forecast*cr.rate)+'</div></div>' +
    '<div class="kpi-card kpi-ns"><div class="kpi-icon"><i class="fas fa-shopping-cart"></i></div><div class="kpi-label">Net Sales</div><div class="kpi-value">'+salesFmt(T_ns*cr.rate)+'</div><div class="kpi-sub">'+nsPct+'% of Gross</div></div>' +
    '<div class="kpi-card kpi-cogs"><div class="kpi-icon"><i class="fas fa-undo"></i></div><div class="kpi-label">Returns</div><div class="kpi-value" style="color:var(--red)">'+salesFmt(T_ret*cr.rate)+'</div><div class="kpi-sub">'+(T_gs>0?((T_ret/T_gs)*100).toFixed(2):'0')+'% of Sales</div></div>' +
    '<div class="kpi-card kpi-gp"><div class="kpi-icon"><i class="fas fa-percent"></i></div><div class="kpi-label">Discounts</div><div class="kpi-value" style="color:var(--red)">'+salesFmt(T_disc*cr.rate)+'</div><div class="kpi-sub">'+(T_gs>0?((T_disc/T_gs)*100).toFixed(2):'0')+'% of Sales</div></div>' +
    '<div class="kpi-card kpi-ni"><div class="kpi-icon"><i class="fas fa-box"></i></div><div class="kpi-label">Qty Sold</div><div class="kpi-value">'+dMonths.reduce((s,m)=>s+m.qty,0).toLocaleString()+'</div><div class="kpi-sub">Jan-Jul total</div></div>';
  animateKPIs();

  const fmtBig = v => { const a=Math.abs(v); if(a>=1e6)return (a/1e6).toFixed(1)+'M'; if(a>=1e3)return (a/1e3).toFixed(0)+'K'; return Math.round(a).toString(); };
  const gsAct = sf.monthly.map(m=>m.gross_sales.actual*cr.rate);
  const gsFct = sf.monthly.map(m=>m.gross_sales.forecast*cr.rate);
  buildGroupedBar('salesGSChart', months, [{name:'Actual',y:gsAct,color:'#4A5568'},{name:'Budget',y:gsFct,color:'#A0AEC0'}]);

  const retAct = sf.monthly.map(m=>m.returns.actual*cr.rate);
  const discAct = sf.monthly.map(m=>m.discounts.actual*cr.rate);
  Plotly.newPlot('salesDedChart', [
    {type:'bar',name:'Returns Actual',x:months,y:retAct,marker:{color:'#DF484A'},text:retAct.map(v=>fmtBig(v)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'}},
    {type:'bar',name:'Discounts Actual',x:months,y:discAct,marker:{color:'#B794F4'},text:discAct.map(v=>fmtBig(v)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'}}
  ], {margin:{t:65,b:55,l:70,r:30},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#A0AEC0',size:11},barmode:'group',height:320,hovermode:'x unified',legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true},xaxis:{automargin:true},cliponaxis:false},{responsive:true,displayModeBar:false});

  const qtyV = dMonths.map(m => m.qty);
  Plotly.newPlot('salesQtyChart',[{type:'bar',x:months,y:qtyV,marker:{color:'#3BD671',opacity:0.85},text:qtyV.map(v=>v.toLocaleString()),textposition:'outside',textfont:{size:10,color:'#4A5568',family:'Inter'},cliponaxis:false}],{margin:{t:80,b:65,l:90,r:50},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#A0AEC0',size:11},yaxis:{rangemode:'tozero',gridcolor:'transparent',automargin:true},xaxis:{automargin:true,fixedrange:true},height:340,hovermode:'x unified',showlegend:false,cliponaxis:false},{responsive:true,displayModeBar:false});

  const bu = d.business_units;
  Plotly.newPlot('salesBUChart',[{type:'pie',labels:bu.map(x=>x.name),values:bu.map(x=>x.sales_pct),text:bu.map(x=>x.name+'<br>'+x.sales_pct.toFixed(1)+'%'),textinfo:'text',textfont:{size:12,color:'#475569',family:'Inter'},marker:{colors:['#2DA356','#4A5568','#3BD671','#3BD671'].slice(0,bu.length),line:{color:'#fff',width:2}},hole:0.4,hovertemplate:'%{label}<br>%{value:.1f}%<extra></extra>'}],{margin:{t:20,b:30,l:20,r:20},paper_bgcolor:'rgba(0,0,0,0)',height:300,showlegend:true,legend:{orientation:'h',y:-0.05,font:{size:11,color:'#72839E'}}},{responsive:true,displayModeBar:false});

  const tc = d.top_customers;
  Plotly.newPlot('salesTopChart',[{type:'bar',orientation:'h',x:tc.map(x=>x.sales*cr.rate),y:tc.map(x=>x.name),marker:{color:['#2DA356','#473FE0','#3BD671','#3BD671','#81E6D9'],opacity:0.9},text:tc.map(x=>fmtBig(x.sales*cr.rate)),textposition:'outside',textfont:{size:10,color:'#4A5568',family:'Inter'},cliponaxis:false}],{margin:{t:15,b:30,l:15,r:80},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#4A5568',size:11},height:Math.max(260,tc.length*48),hovermode:'y unified',showlegend:false,yaxis:{automargin:true,tickfont:{size:11,color:'#4A5568'}},xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true,fixedrange:true}},{responsive:true,displayModeBar:false});

  const tt = d.top_types;
  Plotly.newPlot('salesTypeChart',[{type:'bar',orientation:'h',x:tt.map(x=>x.sales*cr.rate),y:tt.map(x=>x.name),marker:{color:'#3BD671',opacity:0.9},text:tt.map(x=>fmtBig(x.sales*cr.rate)),textposition:'outside',textfont:{size:10,color:'#4A5568',family:'Inter'},cliponaxis:false}],{margin:{t:15,b:30,l:15,r:80},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#4A5568',size:11},height:Math.max(260,tt.length*48),hovermode:'y unified',showlegend:false,yaxis:{automargin:true,tickfont:{size:11,color:'#4A5568'}},xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true,fixedrange:true}},{responsive:true,displayModeBar:false});

  const tf = d.top_fabrics;
  Plotly.newPlot('salesFabricChart',[{type:'bar',orientation:'h',x:tf.map(x=>x.sales*cr.rate),y:tf.map(x=>x.name),marker:{color:'#00B5D8',opacity:0.9},text:tf.map(x=>fmtBig(x.sales*cr.rate)),textposition:'outside',textfont:{size:10,color:'#4A5568',family:'Inter'},cliponaxis:false}],{margin:{t:15,b:30,l:15,r:80},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#4A5568',size:11},height:Math.max(260,tf.length*48),hovermode:'y unified',showlegend:false,yaxis:{automargin:true,tickfont:{size:11,color:'#4A5568'}},xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true,fixedrange:true}},{responsive:true,displayModeBar:false});

  const convAct = sf.monthly.map(m=>m.gross_sales.actual > 0 ? ((m.net_sales.actual / m.gross_sales.actual) * 100) : 0);
  const convFct = sf.monthly.map(m=>m.gross_sales.forecast > 0 ? ((m.net_sales.forecast / m.gross_sales.forecast) * 100) : 0);
  Plotly.newPlot('salesConversionChart',[
    {type:'scatter',mode:'lines+markers',name:'Actual',x:months,y:convAct,line:{color:'#2DA356',width:3},marker:{size:10,color:'#2DA356',symbol:'circle'},fill:'tozeroy',fillcolor:'rgba(59,214,113,0.08)'},
    {type:'scatter',mode:'lines+markers',name:'Budget',x:months,y:convFct,line:{color:'#A0AEC0',width:3,dash:'dot'},marker:{size:10,color:'#A0AEC0',symbol:'diamond'}}
  ],{margin:{t:20,b:40,l:55,r:20},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#A0AEC0',size:11},yaxis:{ticksuffix:'%',gridcolor:'transparent',rangemode:'tozero'},height:280,hovermode:'x unified',legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:11,color:'#72839E'}}},{responsive:true,displayModeBar:false});
  setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
}

// ---- Expenses Tab ----
function expFmt(v) { const c=getCurrency(); const n=Math.abs(v);
  if (n >= 1e6) return (n/1e6).toFixed(1)+'M '+c.symbol;
  if (n >= 1e3) return (n/1e3).toFixed(1)+'K '+c.symbol;
  return n.toFixed(0)+' '+c.symbol;
}
function expFmtShort(v) { return expFmt(v); }

async function loadExpensesData() {
  const f = getFilters();
  const qs = '?period='+encodeURIComponent(f.period)+'&bu='+encodeURIComponent(f.bu)+'&month='+encodeURIComponent(f.month);
  const [p, d] = await Promise.all([
    authFetch('/api/pnl_forecast'+qs).then(r=>r.json()),
    authFetch('/api/data'+qs).then(r=>r.json())
  ]);
  const ytd = p.ytd.expenses;
  const depts = Object.keys(p.dept_expenses);
  const monthly = p.monthly;
  const months = monthly.map(m => monthNameMap[m.month] || ('M'+m.month));
  const cr = getCurrency();

  // KPI cards
  const topDept = depts.reduce((a,b)=>p.dept_expenses[a].actual > p.dept_expenses[b].actual ? a : b);
  const avgMonthly = ytd.actual * cr.rate / 7;
  const budgetUtil = ytd.forecast > 0 ? ((ytd.actual / ytd.forecast) * 100).toFixed(1) : 'N/A';
  document.getElementById('expKpiGrid').innerHTML =
    '<div class="kpi-card kpi-exp"><div class="kpi-icon"><i class="fas fa-receipt"></i></div><div class="kpi-label">Total Expenses</div><div class="kpi-value">'+expFmt(ytd.actual*cr.rate)+'</div><div class="kpi-sub">Budget: '+expFmt(ytd.forecast*cr.rate)+'</div></div>' +
    '<div class="kpi-card kpi-ni"><div class="kpi-icon"><i class="fas fa-building"></i></div><div class="kpi-label">Largest Dept</div><div class="kpi-value" style="font-size:13px">'+topDept+'</div><div class="kpi-sub">'+expFmt(p.dept_expenses[topDept].actual*cr.rate)+'</div></div>' +
    '<div class="kpi-card kpi-gs"><div class="kpi-icon"><i class="fas fa-calendar"></i></div><div class="kpi-label">Monthly Avg</div><div class="kpi-value">'+expFmt(avgMonthly)+'</div><div class="kpi-sub">Jan-Jul average</div></div>' +
    '<div class="kpi-card kpi-gp"><div class="kpi-icon"><i class="fas fa-percent"></i></div><div class="kpi-label">Budget Util.</div><div class="kpi-value" style="color:'+(ytd.actual<=ytd.forecast?'var(--green)':'var(--red)')+'">'+budgetUtil+'%</div><div class="kpi-sub">of budget used</div></div>';

  // 1. Monthly Expenses Actual vs Budget
  const expAct = monthly.map(m=>m.expenses.actual*cr.rate);
  const expFct = monthly.map(m=>m.expenses.forecast*cr.rate);

  // 2. Department Expenses
  const deptAct = depts.map(k=>{
    const ovKey = 'exp_dept_'+k+'_act';
    const raw = _overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? _overrides.values[ovKey] : p.dept_expenses[k].actual;
    return (typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/[^0-9.\-]/g,'')) || 0) * (_overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? 1 : cr.rate);
  });
  const deptFct = depts.map(k=>{
    const ovKey = 'exp_dept_'+k+'_fct';
    const raw = _overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? _overrides.values[ovKey] : p.dept_expenses[k].forecast;
    return (typeof raw === 'number' ? raw : parseFloat(String(raw).replace(/[^0-9.\-]/g,'')) || 0) * (_overrides && _overrides.values && _overrides.values[ovKey] !== undefined ? 1 : cr.rate);
  });
  const varVals = monthly.map(m=>(m.expenses.actual - m.expenses.forecast)*cr.rate);
  const varColors = varVals.map(v=>v <= 0 ? '#3BD671' : '#DF484A');
  const pieColors = ['#DF484A','#3BD671','#2DA356','#2DA356','#473FE0','#3BD671','#DF484A'];
  const top3 = depts.sort((a,b)=>p.dept_expenses[b].actual - p.dept_expenses[a].actual).slice(0,3);
  const trendColors = ['#DF484A','#3BD671','#2DA356'];
  const trendTraces = top3.map((d,i)=>({
    type:'scatter', mode:'lines+markers', name:d,
    x:months, y:p.dept_expenses[d].monthly.map(m=>m.actual*cr.rate),
    line:{color:trendColors[i], width:3}, marker:{size:7, color:trendColors[i]}
  }));
  buildGroupedBar('expMonthlyChart', months, [{name:'Actual',y:expAct,color:'#DF484A'},{name:'Budget',y:expFct,color:'#A0AEC0'}]);

  Plotly.newPlot('expDeptChart', [
    {type:'bar',orientation:'h',name:'Actual',x:deptAct,y:depts,marker:{color:'#B794F4'},text:deptAct.map(v=>expFmtShort(v)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false},
    {type:'bar',orientation:'h',name:'Forecast',x:deptFct,y:depts,marker:{color:'#3BD671'},text:deptFct.map(v=>expFmtShort(v)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}
  ], {margin:{t:25,b:40,l:15,r:20},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#4A5568',size:11},barmode:'group',bargap:0.25,bargroupgap:0.1,height:Math.max(280,depts.length*45),hovermode:'y unified',legend:{orientation:'h',y:1.08,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}},yaxis:{automargin:true,tickfont:{size:11,color:'#4A5568'}},xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true}},{responsive:true,displayModeBar:false});
  document.getElementById('expDeptChart').on('plotly_click', function(data) {
    if (!_editMode) return;
    const pt = data.points[0]; const dept = pt.y; const traceIdx = pt.curveNumber; const isActual = traceIdx === 0;
    const currentVal = isActual ? p.dept_expenses[dept].actual*cr.rate : p.dept_expenses[dept].forecast*cr.rate;
    _editTarget = {el:null,type:'value',editId:'exp_dept_'+dept+'_' +(isActual?'act':'fct'),original:expFmtShort(currentVal),_dept:dept,_isActual:isActual,_rate:cr.rate,_raw:isActual?p.dept_expenses[dept].actual:p.dept_expenses[dept].forecast,_prefix:'exp'};
    openEditModal('value',expFmtShort(currentVal),null);
  });

  Plotly.newPlot('expVarChart',[{type:'bar',x:months,y:varVals,marker:{color:varColors},text:varVals.map(v=>(v<=0?'':'+')+expFmtShort(v)),textposition:'outside',textfont:{size:11,color:'#4A5568',family:'Inter'},cliponaxis:false}],{margin:{t:70,b:60,l:80,r:40},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#A0AEC0',size:11},yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true},xaxis:{automargin:true,fixedrange:true},height:320,hovermode:'x unified',showlegend:false,cliponaxis:false},{responsive:true,displayModeBar:false});

  Plotly.newPlot('expPieChart',[{type:'pie',labels:depts,values:deptAct,text:deptAct.map(v=>expFmtShort(v)),textinfo:'label+text+percent',textposition:'inside',textfont:{size:11,color:'#ffffff',family:'Inter'},marker:{colors:pieColors.slice(0,depts.length),line:{color:'#ffffff',width:2}},hovertemplate:'%{label}<br>'+cr.symbol+': %{value:,.0f}<br>%{percent}<extra></extra>'}],{margin:{t:5,b:5,l:5,r:5},paper_bgcolor:'rgba(0,0,0,0)',height:320,showlegend:true,legend:{orientation:'h',y:-0.1,font:{size:11,color:'#4A5568'}}},{responsive:true,displayModeBar:false});

  Plotly.newPlot('expTrendChart',trendTraces,{margin:{t:70,b:60,l:80,r:40},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#A0AEC0',size:11},yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'transparent',automargin:true},xaxis:{automargin:true,fixedrange:true},height:320,hovermode:'x unified',legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:11,color:'#4A5568'}}},{responsive:true,displayModeBar:false});

  // 6. Department Detail Table
  const sortedDepts = depts.sort((a,b)=>p.dept_expenses[b].actual - p.dept_expenses[a].actual);
  document.getElementById('expTable').innerHTML =
    '<div style="padding:10px 14px;width:100%"><table style="width:100%;border-collapse:collapse;font-size:11px">' +
    '<tr style="color:var(--text-secondary);font-size:10px;text-transform:uppercase;letter-spacing:.5px">' +
    '<td style="padding:6px 14px">Department</td><td style="padding:6px 14px;text-align:right">Actual</td>' +
    '<td style="padding:6px 14px;text-align:right">Budget</td><td style="padding:6px 14px;text-align:right">Variance</td>' +
    '<td style="padding:6px 14px;text-align:right">Var %</td><td style="padding:6px 14px;text-align:right">% of Total</td></tr>' +
    sortedDepts.map(k => {
      const a = p.dept_expenses[k].actual*cr.rate, f = p.dept_expenses[k].forecast*cr.rate, varAmt = a - f, varPct = f>0?((a-f)/f*100).toFixed(1):'N/A', pct = ytd.actual>0?(a/(ytd.actual*cr.rate)*100).toFixed(1):'0';
      const good = varAmt <= 0;
      return '<tr><td style="padding:6px 10px;color:var(--text-secondary)">' + k + '</td>' +
        '<td class="num" style="padding:6px 10px;text-align:right;font-weight:600">' + expFmt(a) + '</td>' +
        '<td class="num" style="padding:6px 10px;text-align:right;color:#72839E">' + expFmt(f) + '</td>' +
        '<td class="num" style="padding:6px 10px;text-align:right;font-weight:600;color:' + (good?'var(--green)':'var(--red)') + '">' + (varAmt>0?'+':'') + expFmt(varAmt) + '</td>' +
        '<td class="num" style="padding:6px 10px;text-align:right;color:' + (good?'var(--green)':'var(--red)') + '">' + (varPct!=='N/A'?(varAmt>0?'+':'')+varPct+'%':'N/A') + '</td>' +
        '<td class="num" style="padding:6px 10px;text-align:right;color:#72839E">' + pct + '%</td></tr>';
    }).join('') +
    '<tr style="border-top:2px solid #CBD5E0"><td style="padding:6px 10px;font-weight:700;color:#1e293b">Total</td>' +
    '<td class="num" style="padding:6px 10px;text-align:right;font-weight:700">' + expFmt(ytd.actual*cr.rate) + '</td>' +
    '<td class="num" style="padding:6px 10px;text-align:right;color:#72839E">' + expFmt(ytd.forecast*cr.rate) + '</td>' +
    '<td class="num" style="padding:6px 10px;text-align:right;font-weight:700;color:' + (ytd.actual<=ytd.forecast?'var(--green)':'var(--red)') + '">' + (ytd.actual>ytd.forecast?'+':'') + expFmt((ytd.actual-ytd.forecast)*cr.rate) + '</td>' +
    '<td class="num" style="padding:6px 10px;text-align:right;color:' + (ytd.actual<=ytd.forecast?'var(--green)':'var(--red)') + '">' + (ytd.forecast>0?((ytd.actual-ytd.forecast)/ytd.forecast*100).toFixed(1):'N/A') + '%</td>' +
    '<td class="num" style="padding:6px 10px;text-align:right;color:#72839E">100%</td></tr></table></div>';
  setTimeout(() => { updatePlotlyBlur(); markEditables(); }, 100);
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
        bu_list = [x.strip().title() for x in bu.split(",") if x.strip()]
        sales_raw = sales_raw[sales_raw["Business Unit"].str.strip().str.title().isin(bu_list)]
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
    if month and month != "all":
        month_list = [int(x.strip()) for x in month.split(",") if x.strip().isdigit()]
        sales_raw = sales_raw[sales_raw["Month"].isin(month_list)]
    elif period and period != "ytd":
        period_months_map = {"q1": [1,2,3], "q2": [4,5,6], "q3": [7,8,9]}
        p_list = [x.strip() for x in period.split(",") if x.strip()]
        valid_months = []
        for p in p_list:
            valid_months.extend(period_months_map.get(p, []))
        if valid_months:
            sales_raw = sales_raw[sales_raw["Month"].isin(valid_months)]
    exp_raw["amount"] = pd.to_numeric(exp_raw["amount"], errors="coerce").fillna(0)
    exp_raw["Department"] = exp_raw["Department"].str.strip().str.title()
    exp_raw["Month"] = pd.to_numeric(exp_raw["Month"], errors="coerce")
    exp_raw = exp_raw[exp_raw["Month"].notna() & exp_raw["Month"].between(1, 12)]
    exp_raw["Month"] = exp_raw["Month"].astype(int)
    # Apply period or month filter to expenses too
    if month and month != "all":
        month_list = [int(x.strip()) for x in month.split(",") if x.strip().isdigit()]
        exp_raw = exp_raw[exp_raw["Month"].isin(month_list)]
    elif period and period != "ytd":
        period_months_map = {"q1": [1,2,3], "q2": [4,5,6], "q3": [7,8,9]}
        p_list = [x.strip() for x in period.split(",") if x.strip()]
        valid_months = []
        for p in p_list:
            valid_months.extend(period_months_map.get(p, []))
        if valid_months:
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

def _read_rev_forecast():
    df = pd.read_excel(_EXCEL_PATH, sheet_name="Revenue", header=None)
    bu_rows = [13, 28, 89, 100]
    monthly = []
    for m in range(12):
        total = 0
        for r in bu_rows:
            val = float(df.iloc[r, 5+m]) if 5+m < df.shape[1] and pd.notna(df.iloc[r, 5+m]) else 0
            total += val
        monthly.append(total)
    grand_total = float(df.iloc[101, 17]) if pd.notna(df.iloc[101, 17]) else sum(monthly)
    return {"yearly_rev": grand_total, "monthly_rev": monthly,
            "monthly": [{"month": i+1, "gross_sales": {"actual": 0, "forecast": v},
                         "returns": 0, "discounts": 0,
                         "net_sales": v} for i, v in enumerate(monthly)]}

def _read_exp_forecast():
    df = pd.read_excel(_EXCEL_PATH, sheet_name="Expenses by Department", header=None)
    dept_rows = {"G&A": 18, "Marketing": 21, "Financing Expenses": 28,
                 "Sales Export": 32, "Sales Stores Expenses": 35, "Sales B2B": 42, "R&D": 48}
    monthly = []
    ytd_totals = {}
    for d, r in dept_rows.items():
        total = 0
        for m in range(12):
            val = float(df.iloc[r, 2+m]) if 2+m < df.shape[1] and pd.notna(df.iloc[r, 2+m]) else 0
            total += val
        ytd_totals[d] = total
    for m in range(7):
        month_data = {}
        for d, r in dept_rows.items():
            val = float(df.iloc[r, 2+m]) if 2+m < df.shape[1] and pd.notna(df.iloc[r, 2+m]) else 0
            month_data[d] = val
        monthly.append(month_data)
    return {"ytd_totals": ytd_totals, "monthly": monthly}

def get_pnl_forecast():
    df = pd.read_excel(_EXCEL_PATH, sheet_name="PNL Dashboard ", header=None)
    sections = {"ytd": 0, "q1": 9, "jan": 18, "feb": 27, "mar": 36, "q2": 45, "apr": 54, "may": 63, "jun": 72, "jul": 81}
    months_order = ["jan", "feb", "mar", "apr", "may", "jun", "jul"]
    a_val = lambda r, c: float(df.iloc[r, c+1]) if c+1 < df.shape[1] and pd.notna(df.iloc[r, c+1]) else 0
    rev_data = _read_rev_forecast()
    exp_data = _read_exp_forecast()
    ytd_col = sections["ytd"]
    gs_val = lambda r, c: float(df.iloc[r, c+1]) if c+1 < df.shape[1] and pd.notna(df.iloc[r, c+1]) else 0
    f_val = lambda r, c: float(df.iloc[r, c+2]) if c+2 < df.shape[1] and pd.notna(df.iloc[r, c+2]) else 0
    total_exp_ytd = sum(exp_data["ytd_totals"].values())
    num_months = len(months_order)
    rev_ytd = sum(rev_data["monthly"][i]["gross_sales"]["forecast"] for i in range(min(num_months, len(rev_data["monthly"]))))
    ytd = {
        "net_sales": {"actual": gs_val(6, ytd_col), "forecast": rev_ytd},
        "cogs": {"actual": abs(a_val(46, ytd_col)), "forecast": abs(f_val(46, ytd_col))},
        "expenses": {"actual": a_val(40, ytd_col), "forecast": f_val(40, ytd_col)},
    }
    ytd["gross_profit"] = {"actual": ytd["net_sales"]["actual"] - ytd["cogs"]["actual"], "forecast": ytd["net_sales"]["forecast"] - ytd["cogs"]["forecast"]}
    ytd["net_income"] = {"actual": ytd["gross_profit"]["actual"] - ytd["expenses"]["actual"], "forecast": ytd["gross_profit"]["forecast"] - ytd["expenses"]["forecast"]}
    monthly = []
    for i, m in enumerate(months_order):
        col = sections[m]
        ns_a = gs_val(6, col); ns_f = rev_data["monthly"][i]["net_sales"]
        cogs_a = abs(a_val(46, col)); cogs_f = abs(f_val(46, col))
        exp_a = a_val(40, col); exp_f = sum(exp_data["monthly"][i].values()) if i < len(exp_data["monthly"]) else 0
        gp_a = ns_a - cogs_a; gp_f = ns_f - cogs_f
        ni_a = gp_a - exp_a; ni_f = gp_f - exp_f
        monthly.append({"month": i+1, "net_sales": {"actual": ns_a, "forecast": ns_f}, "cogs": {"actual": cogs_a, "forecast": cogs_f}, "gross_profit": {"actual": gp_a, "forecast": gp_f}, "expenses": {"actual": exp_a, "forecast": exp_f}, "net_income": {"actual": ni_a, "forecast": ni_f}})
    dept_names = ["G&A", "Marketing", "Financing Expenses", "Sales Export", "Sales Stores Expenses", "Sales B2B", "R&D"]
    dept_expenses = {}
    for i, d in enumerate(dept_names):
        r = 33 + i
        actual = a_val(r, ytd_col)
        forecast = exp_data["ytd_totals"].get(d, 0)
        monthly_dept = []
        for m_idx in range(len(months_order)):
            ma = a_val(r, sections[months_order[m_idx]])
            mf = exp_data["monthly"][m_idx].get(d, 0) if m_idx < len(exp_data["monthly"]) else 0
            monthly_dept.append({"actual": ma, "forecast": mf})
        dept_expenses[d] = {"actual": actual, "forecast": forecast, "monthly": monthly_dept}
    return {"ytd": ytd, "monthly": monthly, "dept_expenses": dept_expenses}

_sales_cache = {"data": None, "ts": 0}

def get_sales_forecast():
    df = pd.read_excel(_EXCEL_PATH, sheet_name="PNL Dashboard ", header=None)
    sections = {"ytd": 0, "jan": 18, "feb": 27, "mar": 36, "apr": 54, "may": 63, "jun": 72, "jul": 81}
    months_order = ["jan", "feb", "mar", "apr", "may", "jun", "jul"]
    a_val = lambda r, c: float(df.iloc[r, c+1]) if c+1 < df.shape[1] and pd.notna(df.iloc[r, c+1]) else 0
    f_val = lambda r, c: float(df.iloc[r, c+2]) if c+2 < df.shape[1] and pd.notna(df.iloc[r, c+2]) else 0
    ret_val = lambda r, c: float(df.iloc[r, c+4]) if c+4 < df.shape[1] and pd.notna(df.iloc[r, c+4]) else 0
    disc_val = lambda r, c: float(df.iloc[r, c+5]) if c+5 < df.shape[1] and pd.notna(df.iloc[r, c+5]) else 0
    ns_val = lambda r, c: float(df.iloc[r, c+6]) if c+6 < df.shape[1] and pd.notna(df.iloc[r, c+6]) else 0
    rev_data = _read_rev_forecast()
    ytd_col = sections["ytd"]
    num_months = len(months_order)
    rev_ytd = sum(rev_data["monthly"][i]["gross_sales"]["forecast"] for i in range(min(num_months, len(rev_data["monthly"]))))
    ytd = {
        "gross_sales": {"actual": a_val(6, ytd_col), "forecast": rev_ytd},
        "returns": {"actual": ret_val(6, ytd_col), "forecast": 0},
        "discounts": {"actual": disc_val(6, ytd_col), "forecast": 0},
        "net_sales": {"actual": ns_val(6, ytd_col), "forecast": 0},
    }
    monthly = []
    for i, m in enumerate(months_order):
        col = sections[m]
        gs_a = a_val(6, col)
        gs_f = rev_data["monthly"][i]["gross_sales"]["forecast"]
        ret_a = ret_val(6, col)
        disc_a = disc_val(6, col)
        ns_a = ns_val(6, col)
        ns_f = rev_data["monthly"][i]["net_sales"]
        monthly.append({"month": months_order.index(m)+1, "gross_sales": {"actual": gs_a, "forecast": gs_f}, "returns": {"actual": ret_a, "forecast": 0}, "discounts": {"actual": disc_a, "forecast": 0}, "net_sales": {"actual": ns_a, "forecast": ns_f}})
    return {"ytd": ytd, "monthly": monthly}

def get_sales_cached(period="ytd", bu="all", month="all"):
    if _sales_cache["data"] is None:
        _sales_cache["data"] = get_sales_forecast()
    data = _sales_cache["data"]
    return _apply_filters(data, period, bu, month)

def get_pnl_cached(period="ytd", bu="all", month="all"):
    if _pnl_cache["data"] is None:
        _pnl_cache["data"] = get_pnl_forecast()
    data = _pnl_cache["data"]
    return _apply_filters(data, period, bu, month)

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

@app.route("/ping")
def ping():
    return "ok"

@app.route("/")
@auth.login_required
def index():
    return render_template_string(HTML_TEMPLATE)

_rates_cache = {"data": None, "ts": 0}

@app.route("/api/rates")
def api_rates():
    import urllib.request as _urllib
    import re as _re
    now = time.time()
    if _rates_cache["data"] and (now - _rates_cache["ts"]) < 86400:
        return jsonify(_rates_cache["data"])
    try:
        req = _urllib.Request("https://www.fexant.com/bank/CBE", headers={"User-Agent": "Mozilla/5.0"})
        resp = _urllib.urlopen(req, timeout=15)
        html = resp.read().decode("utf-8", errors="ignore")
        rows = _re.findall(
            r'<a[^>]*title="View exchange rate history"[^>]*>.*?</a>\s*</td>\s*'
            r'<td><strong>([A-Z]{3})</strong></td>\s*'
            r'<td[^>]*>\s*([\d.]+)\s*EGP.*?</td>\s*'
            r'<td[^>]*>\s*([\d.]+)\s*EGP',
            html, _re.DOTALL
        )
        result = {"EGP": 1}
        for code, buy, sell in rows:
            if code == "CBE":
                continue
            avg = (float(buy) + float(sell)) / 2.0
            if avg > 0:
                result[code] = round(1.0 / avg, 6)
        if len(result) > 2:
            _rates_cache["data"] = result
            _rates_cache["ts"] = now
        return jsonify(result)
    except:
        return jsonify({"EGP": 1, "USD": 0.020089, "EUR": 0.017396, "GBP": 0.014922, "CHF": 0.016231, "KWD": 0.006169, "BHD": 0.00744, "OMR": 0.00767, "JOD": 0.01409, "SAR": 0.075445, "AED": 0.073778})

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
    period = request.args.get("period", "ytd")
    bu = request.args.get("bu", "all")
    month = request.args.get("month", "all")
    return jsonify(get_pnl_cached(period=period, bu=bu, month=month))

@app.route("/api/sales_forecast")
@auth.login_required
def api_sales_forecast():
    period = request.args.get("period", "ytd")
    bu = request.args.get("bu", "all")
    month = request.args.get("month", "all")
    return jsonify(get_sales_cached(period=period, bu=bu, month=month))

_style_cache = {"data": None, "ts": 0}
_inv_cache = {"data": None, "ts": 0}
_cf_cache = {"data": None, "ts": 0}

def _apply_filters(data, period="ytd", bu="all", month="all"):
    import copy
    d = copy.deepcopy(data)
    monthly = d.get("monthly", [])
    def _build_agg():
        agg = {}
        if monthly:
            for key in monthly[0]:
                if key == "month":
                    continue
                if isinstance(monthly[0][key], dict):
                    agg[key] = {"actual": 0, "forecast": 0}
        return agg
    def _sum_months(indices):
        agg = _build_agg()
        for i in indices:
            if i < len(monthly):
                m_data = monthly[i]
                for key in agg:
                    for sub in ["actual", "forecast"]:
                        agg[key][sub] += m_data.get(key, {}).get(sub, 0)
        return agg
    if month != "all":
        m_indices = []
        for mp in month.split(","):
            mp = mp.strip()
            if mp.isdigit():
                idx = int(mp) - 1
                if 0 <= idx < len(monthly):
                    m_indices.append(idx)
        if m_indices:
            d["ytd"] = _sum_months(m_indices)
            d["monthly"] = [monthly[i] for i in m_indices]
    elif period != "ytd":
        period_months = {"q1": [0,1,2], "q2": [3,4,5], "q3": [6]}
        p_list = [x.strip() for x in period.split(",") if x.strip()]
        indices = []
        for p in p_list:
            indices.extend(period_months.get(p, []))
        d["ytd"] = _sum_months(indices)
        d["monthly"] = [monthly[i] for i in indices if i < len(monthly)]
    dept = d.get("dept_expenses", {})
    if dept and d.get("monthly") and (month != "all" or period != "ytd"):
        month_indices = []
        if month != "all":
            for mp in month.split(","):
                mp = mp.strip()
                if mp.isdigit():
                    month_indices.append(int(mp) - 1)
        elif period != "ytd":
            period_months = {"q1": [0,1,2], "q2": [3,4,5], "q3": [6]}
            p_list = [x.strip() for x in period.split(",") if x.strip()]
            for p in p_list:
                month_indices.extend(period_months.get(p, []))
        for dk in dept:
            orig = dept[dk].get("monthly", [])
            dept[dk]["monthly"] = [orig[i] if i < len(orig) else {"actual":0,"forecast":0} for i in month_indices]
    return d

@app.route("/api/style_analysis")
@auth.login_required
def api_style_analysis():
    now = time.time()
    period = request.args.get("period", "ytd")
    bu = request.args.get("bu", "all")
    month = request.args.get("month", "all")
    if _style_cache["data"] and (now - _style_cache["ts"]) < 300:
        return jsonify(_style_cache["data"])
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
        _style_cache["data"] = result
        _style_cache["ts"] = time.time()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/investment")
@auth.login_required
def api_investment():
    now = time.time()
    period = request.args.get("period", "ytd")
    bu = request.args.get("bu", "all")
    month = request.args.get("month", "all")
    if _inv_cache["data"] and (now - _inv_cache["ts"]) < 300:
        return jsonify(_inv_cache["data"])
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
        # Gold metrics (rows 4-20)
        gold = {}
        gold_labels = {4:"return_rate",6:"monthly_risk",7:"risk_free_rate",8:"avg_monthly_return",9:"avg_monthly_real_return",10:"sharpe_ratio",11:"real_sharpe",12:"inflation_impact",13:"inflation_rate",14:"monthly_inflation",15:"mdd",16:"downside_dev",17:"sortino",18:"calmar",19:"kurtosis",20:"skewness"}
        for row_idx, key in gold_labels.items():
            gold[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["gold"] = gold
        # Silver metrics (rows 24-39)
        silver = {}
        silver_labels = {24:"return_rate",25:"monthly_risk",26:"risk_free_rate",27:"avg_monthly_return",28:"avg_monthly_real_return",29:"sharpe_ratio",30:"real_sharpe",31:"inflation_impact",32:"inflation_rate",33:"monthly_inflation",34:"mdd",35:"downside_dev",36:"sortino",37:"calmar",38:"kurtosis",39:"skewness"}
        for row_idx, key in silver_labels.items():
            silver[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["silver"] = silver
        # Swiss Frank metrics (rows 43-58)
        swiss = {}
        swiss_labels = {43:"return_rate",44:"monthly_risk",45:"risk_free_rate",46:"avg_monthly_return",47:"avg_monthly_real_return",48:"sharpe_ratio",49:"real_sharpe",50:"inflation_impact",51:"inflation_rate",52:"monthly_inflation",53:"mdd",54:"downside_dev",55:"sortino",56:"calmar",57:"kurtosis",58:"skewness"}
        for row_idx, key in swiss_labels.items():
            swiss[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["swiss"] = swiss
        # Correlations (rows 60-62)
        result["correlations"] = {
            "gold_silver": safe_float(60, 1),
            "gold_swiss": safe_float(61, 1),
            "silver_swiss": safe_float(62, 1)
        }
        result["betas"] = {
            "gold_silver": safe_float(64, 1),
            "gold_swiss": safe_float(65, 1),
            "silver_swiss": safe_float(66, 1)
        }
        # Portfolio weights (row 69)
        w_g = safe_float(69, 1); w_s = safe_float(69, 2); w_sw = safe_float(69, 3)
        w_total = w_g + w_s + w_sw
        # Portfolio metrics (rows 72-87)
        portfolio = {}
        portfolio_labels = {72:"return_rate",73:"monthly_risk",74:"risk_free_rate",75:"avg_monthly_return",76:"avg_monthly_real_return",77:"sharpe_ratio",78:"real_sharpe",79:"inflation_impact",80:"inflation_rate",81:"monthly_inflation",82:"mdd",83:"downside_dev",84:"sortino",85:"calmar",86:"kurtosis",87:"skewness"}
        for row_idx, key in portfolio_labels.items():
            portfolio[key] = {"y2025": safe_float(row_idx, 1), "y2026a": safe_float(row_idx, 2), "y2026f": safe_float(row_idx, 3)}
        result["portfolio"] = portfolio
        if w_total > 0:
            result["portfolio"]["weights"] = {"gold": w_g/w_total, "silver": w_s/w_total, "swiss": w_sw/w_total}
        else:
            result["portfolio"]["weights"] = {"gold": 0.33, "silver": 0.33, "swiss": 0.34}
        result["portfolio"]["prices"] = {"current": safe_float(70, 6), "prev": safe_float(70, 7)}
        result["portfolio"]["growth"] = safe_float(71, 7)
        # Investment results (rows 90-97)
        result["investment"] = {
            "initial": safe_float(90, 1),
            "annual_return": {"actual": safe_float(92, 1), "forecast": safe_float(92, 2)},
            "mdd": {"actual": safe_float(93, 1), "forecast": safe_float(93, 2)},
            "return_after_loss": {"actual": safe_float(94, 1), "forecast": safe_float(94, 2)},
            "discount_rate": safe_float(95, 1),
            "pv": {"actual": safe_float(96, 1), "forecast": safe_float(96, 2)},
            "hedge_pct": {"actual": safe_float(97, 1), "forecast": safe_float(97, 2)}
        }
        result["recommendation"] = safe_str(99, 0)
        # Optimal portfolio allocations (top 3 by Sharpe ratio for each year)
        try:
            rfr = portfolio.get("risk_free_rate", {}).get("y2026a", 0)
            optimal_all = {}
            for yr_label, yr_key in [("2025","y2025"),("2026 Actual","y2026a"),("2026 Forecast","y2026f")]:
                ar = [gold["return_rate"][yr_key], silver["return_rate"][yr_key], swiss["return_rate"][yr_key]]
                arisk = [gold["monthly_risk"][yr_key], silver["monthly_risk"][yr_key], swiss["monthly_risk"][yr_key]]
                rfr_yr = portfolio.get("risk_free_rate", {}).get(yr_key, 0)
                rho_gs = result["correlations"]["gold_silver"]
                rho_gsw = result["correlations"]["gold_swiss"]
                rho_ssw = result["correlations"]["silver_swiss"]
                best = []
                step = 5
                for w1 in range(0, 101, step):
                    for w2 in range(0, 101 - w1, step):
                        w3 = 100 - w1 - w2
                        wg, ws, wsw = w1/100, w2/100, w3/100
                        pr = wg*ar[0] + ws*ar[1] + wsw*ar[2]
                        pv = (wg*arisk[0])**2 + (ws*arisk[1])**2 + (wsw*arisk[2])**2 + \
                             2*wg*ws*rho_gs*arisk[0]*arisk[1] + \
                             2*wg*wsw*rho_gsw*arisk[0]*arisk[2] + \
                             2*ws*wsw*rho_ssw*arisk[1]*arisk[2]
                        prisk = pv**0.5
                        sharpe = (pr - rfr_yr) / prisk if prisk > 0 else 0
                        best.append({"gold": round(wg*100), "silver": round(ws*100), "swiss": round(wsw*100),
                                     "return": round(pr*100, 2), "risk": round(prisk*100, 2), "sharpe": round(sharpe, 3)})
                best.sort(key=lambda x: x["sharpe"], reverse=True)
                optimal_all[yr_key] = best[:3]
            result["optimal_portfolios"] = optimal_all
        except Exception:
            result["optimal_portfolios"] = {}
        _inv_cache["data"] = result
        _inv_cache["ts"] = time.time()
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/cashflow")
@auth.login_required
def api_cashflow():
    now = time.time()
    period = request.args.get("period", "ytd")
    bu = request.args.get("bu", "all")
    month = request.args.get("month", "all")
    if _cf_cache["data"] and (now - _cf_cache["ts"]) < 300:
        return jsonify(_cf_cache["data"])
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
        # Top Spending Categories (rows 53-62, col 1=Amount)
        cat_spend = {}
        for i in range(53, min(63, len(dash))):
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
        cf_result = {
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
        }
        _cf_cache["data"] = cf_result
        _cf_cache["ts"] = time.time()
        return jsonify(cf_result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/refresh")
@auth.login_required
def api_refresh():
    _cache["data"] = None
    _style_cache["data"] = None
    _inv_cache["data"] = None
    _cf_cache["data"] = None
    _pnl_cache["data"] = None
    _sales_cache["data"] = None
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

@app.route("/api/overrides", methods=["GET"])
@auth.login_required
def get_overrides():
    return jsonify(_load_overrides())

@app.route("/api/overrides", methods=["POST"])
@auth.login_required
def save_overrides():
    data = request.get_json(force=True)
    _save_overrides(data)
    _cache["data"] = None
    _pnl_cache["data"] = None
    _sales_cache["data"] = None
    _style_cache["data"] = None
    _inv_cache["data"] = None
    _cf_cache["data"] = None
    return jsonify({"ok": True})

@app.route("/api/overrides/clear", methods=["POST"])
@auth.login_required
def clear_overrides():
    _save_overrides({"values": {}, "labels": {}, "notes": {}, "layout": {}})
    _cache["data"] = None
    _pnl_cache["data"] = None
    _sales_cache["data"] = None
    return jsonify({"ok": True})

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
                req5 = urllib.request.Request("http://localhost:" + str(PORT) + "/api/cashflow")
                req5.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req5, timeout=30)
                req6 = urllib.request.Request("http://localhost:" + str(PORT) + "/api/style_analysis")
                req6.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req6, timeout=30)
                req7 = urllib.request.Request("http://localhost:" + str(PORT) + "/api/investment")
                req7.add_header("Authorization", "Basic VklWQSAxOTYwOmlNbFdvSnYxSHBlRDZmR0NVSDBVY2xNNkp2bz0zKEpL")
                urllib.request.urlopen(req7, timeout=30)
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
