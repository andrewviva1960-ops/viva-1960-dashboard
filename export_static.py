import pandas as pd, json, os, warnings
warnings.filterwarnings("ignore")

DIR = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(DIR, "Viva Financial model 2026 (4).xlsx")

# --- Data logic (same as app.py) ---
def compute_data():
    sales_raw = pd.read_excel(PATH, sheet_name="Sales Raw Data 2026", header=1)
    exp_raw = pd.read_excel(PATH, sheet_name="Expenses Raw Data 2026")
    for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
        sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
    sales_raw["Discount Total"] = sales_raw["Sales Amount"] * sales_raw["Discount"]
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
    exp_raw["amount"] = pd.to_numeric(exp_raw["amount"], errors="coerce").fillna(0)
    exp_raw["Department"] = exp_raw["Department"].str.strip().str.title()
    exp_raw["Month"] = pd.to_numeric(exp_raw["Month"], errors="coerce")
    exp_raw = exp_raw[exp_raw["Month"].notna() & exp_raw["Month"].between(1, 12)]
    exp_raw["Month"] = exp_raw["Month"].astype(int)
    TOTAL_COGS = 31223509.0; TOTAL_EXPENSES = 8226344.0; SALES_STORES_MONTHLY = 10000.0
    gs = sales_raw.groupby("Month")["Sales Amount"].sum()
    ret_raw = sales_raw["Return"].abs().groupby(sales_raw["Month"]).sum()
    disc_raw = sales_raw["Discount Total"].groupby(sales_raw["Month"]).sum()
    ss_mask = exp_raw["Department"] == "Sales Stores"
    exp_no_ss = exp_raw[~ss_mask].copy()
    exp = exp_no_ss.groupby("Month")["amount"].sum()
    for m in range(1, 7): exp[m] = exp.get(m, 0) + SALES_STORES_MONTHLY
    dept_exp = exp_no_ss.groupby("Department")["amount"].sum()
    dept_exp["Sales Stores"] = SALES_STORES_MONTHLY * 6
    total_gs = float(gs.sum()); total_disc = 2250925.0; total_ret = float(sales_raw["Return"].abs().sum())
    total_exp = TOTAL_EXPENSES; total_cogs = TOTAL_COGS
    total_ns = total_gs - total_ret - total_disc; gp = total_ns - total_cogs; ni = gp - total_exp
    raw_disc_total = float(disc_raw.sum())
    months_list = []
    for m in range(1, 13):
        if m <= 6:
            m_gs = float(gs.get(m,0)); m_ret = float(ret_raw.get(m,0))
            m_disc_raw = float(disc_raw.get(m,0)); m_disc = total_disc * (m_disc_raw / raw_disc_total) if raw_disc_total > 0 else 0
            m_exp = float(exp.get(m,0))
        else: m_gs=0; m_ret=0; m_disc=0; m_exp=0
        months_list.append({"gs": m_gs, "ns": m_gs - m_ret - m_disc, "exp": m_exp})
    dept_data = {k: float(v) for k, v in sorted(dept_exp.items(), key=lambda x: x[1], reverse=True)}
    top5 = sales_raw.groupby("Client")["Sales Amount"].sum().sort_values(ascending=False).head(5)
    top_customers = [{"name": str(k), "sales": float(v)} for k, v in top5.items()]
    bu_all = sales_raw.groupby("Business Unit")
    bu_sales = bu_all["Sales Amount"].sum(); bu_returns = bu_all["Return"].sum().abs(); bu_discounts = bu_all["Discount Value"].sum()
    bu_list = []
    for bu_name in bu_sales.index:
        s = float(bu_sales[bu_name]); r = float(bu_returns.get(bu_name,0)); d = float(bu_discounts.get(bu_name,0))
        bu_list.append({"name": str(bu_name), "sales_pct": round(s/total_gs*100,2) if total_gs else 0,
                        "return_pct": round(r/total_gs*100,2) if total_gs else 0, "discount_pct": round(d/total_gs*100,2) if total_gs else 0})
    bu_list.sort(key=lambda x: x["sales_pct"], reverse=True)
    return {
        "totals": {"gs": round(total_gs), "ret": round(total_ret), "disc": round(total_disc),
                   "ns": round(total_ns), "cogs": round(total_cogs), "gp": round(gp),
                   "exp": round(total_exp), "ni": round(ni)},
        "months": months_list, "depts": dept_data, "top_customers": top_customers, "business_units": bu_list
    }

DATA = compute_data()
DATA_JSON = json.dumps(DATA)

HTML = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Viva 2026 — Financial Dashboard</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.7.2/css/all.min.css">
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
:root{{--sidebar-bg:#0a0e17;--sidebar-width:220px;--header-bg:#000;--accent:#4f8cff;--card-bg:#131825;--card-border:#1e293b;--text-primary:#e2e8f0;--text-secondary:#8899aa;--green:#22c55e;--red:#ef4444}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:#080c14;color:var(--text-primary);overflow-x:hidden}}
.header{{background:var(--header-bg);color:#fff;height:56px;display:flex;align-items:center;padding:0 20px;position:fixed;top:0;left:0;right:0;z-index:1000;border-bottom:1px solid #1e293b}}
.header .brand{{display:flex;align-items:center;gap:10px;font-size:17px;font-weight:600;letter-spacing:.3px;width:var(--sidebar-width)}}
.header .brand i{{color:var(--accent);font-size:22px}}
.header .top-right{{margin-left:auto;display:flex;align-items:center;gap:12px;font-size:13px;color:var(--text-secondary)}}
.header .top-right .status-dot{{display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--green);margin-right:5px}}
.sidebar{{position:fixed;top:56px;left:0;bottom:0;width:var(--sidebar-width);background:var(--sidebar-bg);color:var(--text-secondary);overflow-y:auto;z-index:999;padding-top:10px;border-right:1px solid #1e293b}}
.sidebar .nav-item{{padding:11px 20px;display:flex;align-items:center;gap:12px;cursor:pointer;font-size:13.5px;transition:.15s;border-left:3px solid transparent;color:var(--text-secondary)}}
.sidebar .nav-item:hover{{background:rgba(255,255,255,.04);color:#fff}}
.sidebar .nav-item.active{{background:rgba(79,140,255,.1);color:#fff;border-left-color:var(--accent);font-weight:600}}
.sidebar .nav-item i{{width:18px;text-align:center;font-size:15px;color:var(--accent)}}
.sidebar .nav-section{{font-size:10px;text-transform:uppercase;letter-spacing:1.2px;color:#475569;padding:18px 20px 6px;font-weight:600}}
.main{{margin-left:var(--sidebar-width);margin-top:56px;padding:20px 24px;min-height:calc(100vh - 56px)}}
.page-title{{font-size:20px;font-weight:600;color:#f1f5f9;margin-bottom:20px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin-bottom:24px}}
@media(max-width:1200px){{.kpi-grid{{grid-template-columns:repeat(3,1fr)}}}}
@media(max-width:600px){{.kpi-grid{{grid-template-columns:1fr 1fr}}}}
.kpi-card{{background:var(--card-bg);border-radius:10px;padding:16px;border:1px solid var(--card-border);border-top:3px solid #334155}}
.kpi-card .kpi-icon{{float:right;width:36px;height:36px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;color:#fff}}
.kpi-card .kpi-label{{font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:var(--text-secondary);margin-bottom:2px}}
.kpi-card .kpi-value{{font-size:20px;font-weight:700;color:#f1f5f9}}
.kpi-card .kpi-sub{{font-size:11px;color:var(--text-secondary);margin-top:2px}}
.kpi-card.kpi-gs{{border-top-color:#22c55e}}.kpi-card.kpi-gs .kpi-icon{{background:#22c55e}}
.kpi-card.kpi-ns{{border-top-color:#3b82f6}}.kpi-card.kpi-ns .kpi-icon{{background:#3b82f6}}
.kpi-card.kpi-cogs{{border-top-color:#f59e0b}}.kpi-card.kpi-cogs .kpi-icon{{background:#f59e0b}}
.kpi-card.kpi-gp{{border-top-color:#8b5cf6}}.kpi-card.kpi-gp .kpi-icon{{background:#8b5cf6}}
.kpi-card.kpi-exp{{border-top-color:#ef4444}}.kpi-card.kpi-exp .kpi-icon{{background:#ef4444}}
.kpi-card.kpi-ni{{border-top-color:#14b8a6}}.kpi-card.kpi-ni .kpi-icon{{background:#14b8a6}}
.kpi-card .kpi-value.green{{color:var(--green)}}.kpi-card .kpi-value.red{{color:var(--red)}}
.chart-grid{{display:grid;grid-template-columns:1fr 1fr;gap:22px;margin-bottom:22px}}
.chart-grid-3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:22px;margin-bottom:22px}}
@media(max-width:1100px){{.chart-grid-3{{grid-template-columns:1fr 1fr}}}}
@media(max-width:900px){{.chart-grid,.chart-grid-3{{grid-template-columns:1fr}}}}
.chart-card{{background:var(--card-bg);border-radius:10px;border:1px solid var(--card-border);overflow:hidden}}
.chart-card .chart-header{{display:flex;align-items:center;justify-content:space-between;padding:14px 18px 0}}
.chart-card .chart-header h5{{font-size:14px;font-weight:600;color:#f1f5f9;margin:0}}
.chart-card .chart-body{{padding:2px 4px 6px;min-height:240px}}
.chart-card .pnl-body{{padding:8px 12px 12px;min-height:240px;display:flex;align-items:center}}
.chart-card.full{{grid-column:1/-1}}
.pnl-card{{background:var(--card-bg);border-radius:10px;border:1px solid var(--card-border);padding:18px;height:100%}}
.pnl-card h5{{font-size:14px;font-weight:600;color:#f1f5f9;margin-bottom:12px}}
.pnl-card table{{width:100%;border-collapse:collapse;font-size:14px}}
.pnl-card td{{padding:10px 16px;border-bottom:1px solid #1e293b}}
.pnl-card td:last-child{{text-align:right;font-weight:600;font-variant-numeric:tabular-nums}}
.pnl-card tr.total td{{border-top:2px solid #475569;padding-top:12px;font-weight:700}}
.pnl-card .label-cell{{color:var(--text-secondary)}}.pnl-card .value-cell{{letter-spacing:.2px}}
.pnl-card .neg{{color:var(--red)}}.pnl-card .pos{{color:var(--green)}}
.chart-grid+.chart-grid{{margin-top:18px}}
.footer{{text-align:center;padding:20px 0;font-size:12px;color:var(--text-secondary)}}
</style>
</head>
<body>
  <div class="header">
    <div class="brand"><img src="download.png" alt="Logo" style="height:56px;width:auto;filter:brightness(1.15) drop-shadow(0 2px 4px rgba(0,0,0,0.2));border-radius:8px;background:#fff;padding:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);"></div>
  <div class="top-right"><span><span class="status-dot"></span>Live</span><span id="tsDisplay"><i class="far fa-clock"></i> <span id="timestamp"></span></span></div>
</div>
<div class="sidebar">
  <div class="nav-section">Overview</div>
  <div class="nav-item active"><i class="fas fa-tachometer-alt"></i> Dashboard</div>
  <div class="nav-item"><i class="fas fa-file-invoice-dollar"></i> P&L</div>
  <div class="nav-item"><i class="fas fa-chart-bar"></i> Sales</div>
  <div class="nav-item"><i class="fas fa-receipt"></i> Expenses</div>
  <div class="nav-item"><i class="fas fa-building"></i> Departments</div>
  <div class="nav-section">Data</div>
  <div class="nav-item"><i class="fas fa-table"></i> Raw Data</div>
  <div class="nav-item"><i class="fas fa-cog"></i> Settings</div>
</div>
<div class="main">
  <div class="page-title"><i class="fas fa-tachometer-alt" style="color:var(--accent);margin-right:8px"></i>Dashboard</div>
  <div class="kpi-grid" id="kpiGrid"></div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-chart-bar" style="color:#3b82f6;margin-right:6px"></i>Monthly Gross Sales</h5></div><div class="chart-body" id="salesChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-receipt" style="color:#ef4444;margin-right:6px"></i>Monthly Expenses</h5></div><div class="chart-body" id="expChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-arrows-alt-h" style="color:#3b82f6;margin-right:6px"></i>Monthly Net Sales vs Expenses</h5></div><div class="chart-body" id="monthlyChart"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-sitemap" style="color:#8b5cf6;margin-right:6px"></i>Expenses by Department</h5></div><div class="chart-body" id="deptChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-file-invoice-dollar" style="color:#f1f5f9;margin-right:6px"></i>Profit &amp; Loss Summary</h5></div><div class="pnl-body" id="pnlCard"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card full"><div class="chart-header"><h5><i class="fas fa-trophy" style="color:#d97706;margin-right:6px"></i>Top 5 Customers by Sales</h5></div><div class="chart-body" id="topCustomersChart"></div></div>
  </div>
  <div class="chart-grid-3">
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-pie-chart" style="color:#3b82f6;margin-right:6px"></i>BU Sales %</h5></div><div class="chart-body" id="buSalesChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-undo-alt" style="color:#059669;margin-right:6px"></i>BU Returns % of Sales</h5></div><div class="chart-body" id="buReturnsChart"></div></div>
    <div class="chart-card"><div class="chart-header"><h5><i class="fas fa-percent" style="color:#d97706;margin-right:6px"></i>BU Discounts % of Sales</h5></div><div class="chart-body" id="buDiscountsChart"></div></div>
  </div>
  <div class="footer">Viva 2026 Financial Dashboard &mdash; Exported on <span id="exportDate"></span></div>
</div>
<script>
const DATA = {DATA_JSON};
function fmt(v){{const s=Math.abs(v).toLocaleString('en-US');return v<0?'($'+s+')':'$'+s}}
function fmtRaw(v){{return Math.abs(v).toLocaleString('en-US')}}
function buildKPI(l,v,icon,cls){{const n=v<0;const a=Math.abs(v).toLocaleString('en-US');return '<div class="kpi-card kpi-'+cls+'"><div class="kpi-icon"><i class="fas fa-'+icon+'"></i></div><div class="kpi-label">'+l+'</div><div class="kpi-value '+(n?'red':'green')+'">'+(n?'-':'')+'$'+a+'</div></div>'}}
document.getElementById('exportDate').textContent=new Date().toLocaleString();
document.getElementById('timestamp').textContent=new Date().toLocaleString();
const T=DATA.totals;
document.getElementById('kpiGrid').innerHTML=buildKPI('Gross Sales',T.gs,'chart-line','gs')+buildKPI('Net Sales',T.ns,'shopping-cart','ns')+buildKPI('COGS',T.cogs,'truck','cogs')+buildKPI('Gross Profit',T.gp,'arrow-up','gp')+buildKPI('Expenses',T.exp,'receipt','exp')+buildKPI('Net Income',T.ni,'wallet','ni');
const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];const M=DATA.months;
const gsV=months.map((_,i)=>M[i].gs);
Plotly.newPlot('salesChart',[{{type:'bar',x:months,y:gsV,marker:{{color:'#3b82f6'}},text:gsV.map(v=>'$'+fmtRaw(v)),textposition:'outside',textfont:{{size:13,color:'#e2e8f0'}},cliponaxis:false}}],{{margin:{{t:10,b:30,l:50,r:15}},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{{color:'#cbd5e1'}},yaxis:{{rangemode:'tozero',tickprefix:'$',gridcolor:'rgba(255,255,255,0.05)'}},height:260,hovermode:'x unified',showlegend:false}},{{responsive:true,displayModeBar:false}});
const expV=months.map((_,i)=>M[i].exp);
Plotly.newPlot('expChart',[{{type:'bar',x:months,y:expV,marker:{{color:'#ef4444'}},text:expV.map(v=>'$'+fmtRaw(v)),textposition:'outside',textfont:{{size:13,color:'#e2e8f0'}},cliponaxis:false}}],{{margin:{{t:10,b:30,l:50,r:15}},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{{color:'#cbd5e1'}},yaxis:{{rangemode:'tozero',tickprefix:'$',gridcolor:'rgba(255,255,255,0.05)'}},height:260,hovermode:'x unified',showlegend:false}},{{responsive:true,displayModeBar:false}});
const nsV=months.map((_,i)=>M[i].ns);
Plotly.newPlot('monthlyChart',[{{type:'bar',name:'Net Sales',x:months,y:nsV,marker:{{color:'#3b82f6'}},text:nsV.map(v=>'$'+fmtRaw(v)),textposition:'outside',textfont:{{size:12,color:'#e2e8f0'}},cliponaxis:false}},{{type:'bar',name:'Expenses',x:months,y:expV,marker:{{color:'#ef4444'}},text:expV.map(v=>'$'+fmtRaw(v)),textposition:'outside',textfont:{{size:12,color:'#e2e8f0'}},cliponaxis:false}}],{{margin:{{t:10,b:30,l:50,r:15}},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{{color:'#cbd5e1'}},barmode:'group',height:280,hovermode:'x unified',legend:{{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{{color:'#cbd5e1'}}}},yaxis:{{tickprefix:'$',gridcolor:'rgba(255,255,255,0.05)'}}}},{{responsive:true,displayModeBar:false}});
const depts=Object.keys(DATA.depts);const dVals=Object.values(DATA.depts);
Plotly.newPlot('deptChart',[{{type:'bar',orientation:'h',x:dVals,y:depts,marker:{{color:'#8b5cf6'}},text:dVals.map(v=>'$'+fmtRaw(v)),textposition:'outside',textfont:{{size:13,color:'#e2e8f0'}},cliponaxis:false}}],{{margin:{{t:10,b:20,l:130,r:60}},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{{color:'#cbd5e1'}},height:Math.max(240,depts.length*40),hovermode:'y unified',showlegend:false,xaxis:{{tickprefix:'$',gridcolor:'rgba(255,255,255,0.05)'}}}},{{responsive:true,displayModeBar:false}});
document.getElementById('pnlCard').innerHTML='<div style="padding:4px 8px"><table><tr><td class="label-cell">Gross Sales</td><td class="value-cell">'+fmt(T.gs)+'</td></tr><tr><td class="label-cell">Less: Returns</td><td class="value-cell">'+fmt(T.ret)+'</td></tr><tr><td class="label-cell">Less: Discounts</td><td class="value-cell">'+fmt(T.disc)+'</td></tr><tr><td class="label-cell"><strong>Net Sales</strong></td><td class="value-cell"><strong>'+fmt(T.ns)+'</strong></td></tr><tr><td class="label-cell">Less: COGS</td><td class="value-cell">'+fmt(T.cogs)+'</td></tr><tr class="total"><td class="label-cell"><strong>Gross Profit</strong></td><td class="value-cell"><strong>'+fmt(T.gp)+'</strong></td></tr><tr><td class="label-cell">Expenses</td><td class="value-cell">'+fmt(T.exp)+'</td></tr><tr class="total"><td class="label-cell"><strong>Net Income</strong></td><td class="value-cell" style="font-size:16px"><strong>'+fmt(T.ni)+'</strong></td></tr></table></div>';
const tc=DATA.top_customers;const tcNames=tc.map(x=>x.name);const tcVals=tc.map(x=>x.sales);
Plotly.newPlot('topCustomersChart',[{{type:'bar',x:tcNames,y:tcVals,marker:{{color:['#3b82f6','#60a5fa','#93c5fd','#bfdbfe','#dbeafe']}},text:tcVals.map(v=>'$'+fmtRaw(v)),textposition:'outside',textfont:{{size:14,color:'#e2e8f0'}},cliponaxis:false}}],{{margin:{{t:10,b:40,l:60,r:20}},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{{color:'#cbd5e1'}},yaxis:{{rangemode:'tozero',tickprefix:'$',gridcolor:'rgba(255,255,255,0.05)'}},height:280,hovermode:'x unified',showlegend:false}},{{responsive:true,displayModeBar:false}});
const bu=DATA.business_units;const buNames=bu.map(x=>x.name);const buColors=['#3b82f6','#059669','#d97706','#7c3aed'];
const buSalesPct=bu.map(x=>x.sales_pct);
Plotly.newPlot('buSalesChart',[{{type:'pie',labels:buNames,values:buSalesPct,text:buSalesPct.map(v=>v.toFixed(1)+'%'),textinfo:'label+percent',textfont:{{size:14,color:'#fff'}},marker:{{colors:buColors.slice(0,buNames.length),line:{{color:'#fff',width:2}}}},hovertemplate:'%{{label}}<br>%{{value:.1f}}%<extra></extra>'}}],{{margin:{{t:5,b:5,l:5,r:5}},paper_bgcolor:'rgba(0,0,0,0)',height:290,showlegend:true,legend:{{orientation:'h',y:-0.08,font:{{size:11,color:'#cbd5e1'}}}}}},{{responsive:true,displayModeBar:false}});
const buRetPct=bu.map(x=>x.return_pct);
Plotly.newPlot('buReturnsChart',[{{type:'pie',labels:buNames,values:buRetPct,text:buRetPct.map(v=>v.toFixed(2)+'%'),textinfo:'label+percent',textfont:{{size:13,color:'#fff'}},marker:{{colors:['#059669','#34d399','#6ee7b7','#a7f3d0'].slice(0,buNames.length),line:{{color:'#fff',width:2}}}},hovertemplate:'%{{label}}<br>%{{value:.2f}}%<extra></extra>'}}],{{margin:{{t:5,b:5,l:5,r:5}},paper_bgcolor:'rgba(0,0,0,0)',height:290,showlegend:true,legend:{{orientation:'h',y:-0.08,font:{{size:11,color:'#cbd5e1'}}}}}},{{responsive:true,displayModeBar:false}});
const buDiscPct=bu.map(x=>x.discount_pct);
Plotly.newPlot('buDiscountsChart',[{{type:'pie',labels:buNames,values:buDiscPct,text:buDiscPct.map(v=>v.toFixed(2)+'%'),textinfo:'label+percent',textfont:{{size:13,color:'#fff'}},marker:{{colors:['#d97706','#f59e0b','#fbbf24','#fcd34d'].slice(0,buNames.length),line:{{color:'#fff',width:2}}}},hovertemplate:'%{{label}}<br>%{{value:.2f}}%<extra></extra>'}}],{{margin:{{t:5,b:5,l:5,r:5}},paper_bgcolor:'rgba(0,0,0,0)',height:290,showlegend:true,legend:{{orientation:'h',y:-0.08,font:{{size:11,color:'#cbd5e1'}}}}}},{{responsive:true,displayModeBar:false}});
</script>
</body>
</html>'''

OUT = os.path.join(DIR, "dashboard_export.html")
with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML)
size_kb = os.path.getsize(OUT) / 1024
print(f"Exported: {OUT}")
print(f"Size: {size_kb:.0f} KB")
print(f"\nTo share with your manager:")
print(f"1. Go to https://app.netlify.com/drop")
print(f"2. Drag and drop this file")
print(f"3. Share the URL you get")
