import pandas as pd, json, os, warnings
warnings.filterwarnings("ignore")

DIR = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(DIR, "Viva Financial model 2026 (4).xlsx")

sales_raw = pd.read_excel(path, sheet_name="Sales Raw Data 2026", header=1)
exp_raw = pd.read_excel(path, sheet_name="Expenses Raw Data 2026")

for c in ["Sales Amount", "Return", "Discount", "Discount Value", "QTY"]:
    sales_raw[c] = pd.to_numeric(sales_raw[c], errors="coerce").fillna(0)
d_pct = sales_raw["Discount"] > 0
sales_raw["Discount Total"] = sales_raw["Sales Amount"] * sales_raw["Discount"]
sales_raw["Net Sales"] = sales_raw["Sales Amount"] - sales_raw["Return"].abs() - sales_raw["Discount Total"]
# Use Excel Month column where available, date-parse the rest
sales_raw["Month_excel"] = pd.to_numeric(sales_raw["Month"], errors="coerce") if "Month" in sales_raw.columns else float("nan")
sales_raw["Date_num"] = pd.to_numeric(sales_raw["Date"], errors="coerce")
sales_raw["Date_dt"] = pd.to_datetime(sales_raw["Date"], format="%d/%m/%Y", errors="coerce")
mask1 = sales_raw["Date_dt"].isna() & sales_raw["Date_num"].notna()
sales_raw.loc[mask1, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask1, "Date_num"], origin="1899-12-30", unit="D")
mask2 = sales_raw["Date_dt"].isna()
sales_raw.loc[mask2, "Date_dt"] = pd.to_datetime(sales_raw.loc[mask2, "Date"], format="%Y/%d/%m", errors="coerce")
sales_raw["Month_parsed"] = sales_raw["Date_dt"].dt.month
sales_raw["Month"] = sales_raw["Month_excel"].fillna(sales_raw["Month_parsed"])
sales_raw = sales_raw[sales_raw["Month"].notna() & sales_raw["Month"].between(1, 12)]
sales_raw["Month"] = sales_raw["Month"].astype(int)

exp_raw["amount"] = pd.to_numeric(exp_raw["amount"], errors="coerce").fillna(0)
exp_raw["Department"] = exp_raw["Department"].str.strip().str.title()
exp_raw["Month"] = pd.to_numeric(exp_raw["Month"], errors="coerce")
exp_raw = exp_raw[exp_raw["Month"].notna() & exp_raw["Month"].between(1, 12)]
exp_raw["Month"] = exp_raw["Month"].astype(int)

TOTAL_COGS = 31223509.0
TOTAL_EXPENSES = 8226344.0
SALES_STORES_MONTHLY = 10000.0
gs = sales_raw.groupby("Month")["Sales Amount"].sum()
ns_raw = sales_raw.groupby("Month")["Net Sales"].sum()
ret_raw = sales_raw["Return"].abs().groupby(sales_raw["Month"]).sum()
# Remove existing Sales Stores from expense data, replace with $10k/month Jan-Jun
ss_mask = exp_raw["Department"] == "Sales Stores"
exp_no_ss = exp_raw[~ss_mask].copy()
exp = exp_no_ss.groupby("Month")["amount"].sum()
for m in range(1, 7):
    exp[m] = exp.get(m, 0) + SALES_STORES_MONTHLY
dept_exp = exp_no_ss.groupby("Department")["amount"].sum()
dept_exp["Sales Stores"] = SALES_STORES_MONTHLY * 6

disc_raw = sales_raw["Discount Total"].groupby(sales_raw["Month"]).sum()
total_gs = float(gs.sum())
total_disc = 2250925.0
total_ret = float(sales_raw["Return"].abs().sum())
total_exp = TOTAL_EXPENSES
total_cogs = TOTAL_COGS
total_ns = total_gs - total_ret - total_disc
gp = total_ns - total_cogs
ni = gp - total_exp

raw_disc_total = float(disc_raw.sum())
raw_exp_total = float(exp.sum())
months_list = []
for m in range(1, 13):
    if m <= 6:
        m_gs = float(gs.get(m, 0))
        m_ret = float(ret_raw.get(m, 0))
        m_disc_raw = float(disc_raw.get(m, 0))
        m_disc = total_disc * (m_disc_raw / raw_disc_total) if raw_disc_total > 0 else 0
        m_exp = float(exp.get(m, 0))
    else:
        m_gs = 0
        m_ret = 0
        m_disc = 0
        m_exp = 0
    months_list.append({
        "gs": m_gs,
        "ns": m_gs - m_ret - m_disc,
        "exp": m_exp
    })

dept_data = {k: float(v) for k, v in sorted(dept_exp.items(), key=lambda x: x[1], reverse=True)}

data = {
    "totals": {
        "gs": round(total_gs), "ret": round(total_ret),
        "disc": round(total_disc),
        "ns": round(total_ns), "cogs": round(total_cogs),
        "gp": round(gp), "exp": round(total_exp), "ni": round(ni)
    },
    "months": months_list,
    "depts": dept_data
}

json_data = json.dumps(data)

html = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Viva 2026 Financial Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f2f5;color:#333;padding:20px}
h1{font-size:22px;margin-bottom:20px;color:#1a1a2e}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:24px}
.kpi{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}
.kpi .label{font-size:11px;text-transform:uppercase;color:#888;letter-spacing:.5px}
.kpi .value{font-size:20px;font-weight:700;margin:4px 0 2px}
.kpi .sub{font-size:11px;color:#888}
.kpi.green .value{color:#0d6e2d}
.kpi.red .value{color:#b71c1c}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.chart{background:#fff;border-radius:10px;padding:12px;box-shadow:0 1px 3px rgba(0,0,0,0.08);min-height:320px}
.chart.full{grid-column:1/-1}
.chart h3{font-size:13px;color:#555;margin-bottom:8px}
.pnl-box{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}
.pnl-box table{width:100%;border-collapse:collapse;font-size:13px}
.pnl-box td{padding:6px 12px;border-bottom:1px solid #eee}
.pnl-box td:last-child{text-align:right;font-weight:600}
.pnl-box tr.total td{border-top:2px solid #333;padding-top:8px}
@media(max-width:768px){.charts{grid-template-columns:1fr}}
</style>
</head>
<body>
<h1>Viva 2026 Financial Dashboard</h1>
<div class="kpi-grid" id="kpis"></div>
<div class="charts">
  <div class="chart" id="sales_chart"><h3>Monthly Gross Sales</h3></div>
  <div class="chart" id="exp_chart"><h3>Monthly Expenses</h3></div>
  <div class="chart" id="dept_chart"><h3>Expenses by Department</h3></div>
  <div class="chart full" id="monthly_chart"><h3>Monthly Net Sales &amp; Expenses</h3></div>
  <div class="chart full" id="pnl"><h3>Profit &amp; Loss Summary</h3></div>
</div>
<script>
const DATA = ''' + json_data + ''';

// KPIs
const kpiList = [
  {label:'Gross Sales',val:DATA.totals.gs,cls:'green',sub:'Total Revenue'},
  {label:'Net Sales',val:DATA.totals.ns,cls:'green',sub:'After Returns & Discounts'},
  {label:'COGS',val:DATA.totals.cogs,cls:'',sub:'Cost of Goods Sold'},
  {label:'Gross Profit',val:DATA.totals.gp,cls:DATA.totals.gp>=0?'green':'red'},
  {label:'Expenses',val:DATA.totals.exp,cls:''},
  {label:'Net Income',val:DATA.totals.ni,cls:DATA.totals.ni>=0?'green':'red'}
];
document.getElementById('kpis').innerHTML = kpiList.map(k =>
  '<div class="kpi '+k.cls+'"><div class="label">'+k.label+'</div><div class="value">$'+k.val.toLocaleString()+'</div>'+(k.sub?'<div class="sub">'+k.sub+'</div>':'')+'</div>'
).join('');

const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const M = DATA.months;

// Sales chart
const gsVals = months.map((_,i)=>M[i].gs);
Plotly.newPlot('sales_chart',[{type:'bar',x:months,y:gsVals,marker:{color:'#1f77b4'},text:gsVals.map(v=>'$'+v.toLocaleString()),textposition:'outside'}],
  {title:{text:'Monthly Gross Sales'},margin:{t:40,b:30,l:50,r:10},yaxis:{rangemode:'tozero'},height:300,hovermode:'x unified'},{responsive:true});

// Expenses chart
const expVals = months.map((_,i)=>M[i].exp);
Plotly.newPlot('exp_chart',[{type:'bar',x:months,y:expVals,marker:{color:'#d62728'},text:expVals.map(v=>'$'+v.toLocaleString()),textposition:'outside'}],
  {title:{text:'Monthly Expenses'},margin:{t:40,b:30,l:50,r:10},yaxis:{rangemode:'tozero'},height:300,hovermode:'x unified'},{responsive:true});

// By dept
const depts = Object.keys(DATA.depts);
const dVals = Object.values(DATA.depts);
Plotly.newPlot('dept_chart',[{type:'bar',orientation:'h',x:dVals,y:depts,marker:{color:'#9467bd'},text:dVals.map(v=>'$'+v.toLocaleString()),textposition:'outside'}],
  {title:{text:'Expenses by Department'},margin:{t:40,b:30,l:130,r:80},height:Math.max(250,depts.length*35),hovermode:'y unified'},{responsive:true});

// Monthly net sales vs expenses
const nsVals = months.map((_,i)=>M[i].ns);
Plotly.newPlot('monthly_chart',[
  {type:'bar',name:'Net Sales',x:months,y:nsVals,marker:{color:'#1f77b4'}},
  {type:'bar',name:'Expenses',x:months,y:expVals,marker:{color:'#d62728'}}
],{title:{text:'Monthly Net Sales vs Expenses'},margin:{t:40,b:30,l:50,r:10},barmode:'group',height:320,hovermode:'x unified',legend:{orientation:'h',y:1.1}},{responsive:true});

// PNL Table
const T = DATA.totals;
document.getElementById('pnl').innerHTML = '<h3>Profit & Loss Summary</h3>'+
'<div class="pnl-box"><table>'+
'<tr><td>Gross Sales</td><td>$'+T.gs.toLocaleString()+'</td></tr>'+
'<tr><td>Less: Returns</td><td>($'+T.ret.toLocaleString()+')</td></tr>'+
'<tr><td>Less: Discounts</td><td>($'+T.disc.toLocaleString()+')</td></tr>'+
'<tr><td><strong>Net Sales</strong></td><td><strong>$'+T.ns.toLocaleString()+'</strong></td></tr>'+
'<tr><td>Less: COGS</td><td>($'+T.cogs.toLocaleString()+')</td></tr>'+
'<tr class="total"><td><strong>Gross Profit</strong></td><td><strong>$'+T.gp.toLocaleString()+'</strong></td></tr>'+
'<tr><td>Expenses</td><td>($'+T.exp.toLocaleString()+')</td></tr>'+
'<tr class="total"><td><strong>Net Income</strong></td><td><strong>$'+T.ni.toLocaleString()+'</strong></td></tr>'+
'</table></div>';
</script>
</body>
</html>'''

out = os.path.join(DIR, "dashboard.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)
print("Dashboard generated: " + out)
print("Size: " + str(os.path.getsize(out)) + " bytes")
print("Open dashboard.html in your browser")
