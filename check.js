
// Currency rates (base: EGP)
const CURRENCIES = {
  EGP: {rate: 1, symbol: 'EGP', locale: 'en-US'},
  USD: {rate: 0.020051, symbol: '$', locale: 'en-US'},
  EUR: {rate: 0.017482, symbol: '€', locale: 'de-DE'},
  GBP: {rate: 0.015151, symbol: '£', locale: 'en-GB'},
  CHF: {rate: 0.016182, symbol: 'CHF', locale: 'de-CH'},
  KWD: {rate: 0.006167, symbol: 'KWD', locale: 'ar-KW'},
  BHD: {rate: 0.007559, symbol: 'BHD', locale: 'ar-BH'},
  OMR: {rate: 0.007709, symbol: 'OMR', locale: 'ar-OM'},
  JOD: {rate: 0.014216, symbol: 'JOD', locale: 'ar-JO'},
  SAR: {rate: 0.075262, symbol: 'SAR', locale: 'ar-SA'}
};
let currentCurrency = 'EGP';
let _data = null;
let topSlicer = 'customers';

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
    marker:{color:['#2563eb','#3b82f6','#60a5fa','#93c5fd','#bfdbfe']},
    text:vals.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:14,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:50,b:55,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#cbd5e1'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'}, height:340,
     hovermode:'x unified', showlegend:false}, {responsive:true, displayModeBar:false});
}

function changeTopSlicer(val) {
  topSlicer = val;
  if (_data) renderTopChart(_data, getCurrency());
  setTimeout(updatePlotlyBlur, 200);
}

async function loadData() {
  const r = await fetch('/api/data');
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
  Plotly.newPlot('salesChart', [{type:'bar', x:months, y:gsV, marker:{color:'#3b82f6'}, text:gsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:14,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:40,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#cbd5e1'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'},
     height:320, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  const expV = months.map((_,i) => M[i].exp * c.rate);
  Plotly.newPlot('expChart', [{type:'bar', x:months, y:expV, marker:{color:'#dc2626'},     text:expV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:14,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:40,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#cbd5e1'}, yaxis:{rangemode:'tozero',ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'},
     height:320, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  const nsV = months.map((_,i) => M[i].ns * c.rate);
  Plotly.newPlot('monthlyChart', [
    {type:'bar', name:'Net Sales', x:months, y:nsV, marker:{color:'#3b82f6'}, text:nsV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false},
    {type:'bar', name:'Expenses', x:months, y:expV, marker:{color:'#ef4444'}, text:expV.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false}
  ], {margin:{t:50,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:14,color:'#cbd5e1'}, barmode:'group', height:330, hovermode:'x unified',
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:14,color:'#cbd5e1'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  const depts = Object.keys(d.depts);
  const dVals = Object.values(d.depts).map(v => v * c.rate);
  Plotly.newPlot('deptChart', [{type:'bar', orientation:'h', x:dVals, y:depts,
    marker:{color:'#8b5cf6'}, text:dVals.map(v=>fmtShort(v)), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:15,b:25,l:160,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#cbd5e1'}, height:Math.max(300, depts.length*45), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

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
  const bu = d.business_units;
  const buNames = bu.map(x => x.name);
  const buColors = ['#2563eb','#059669','#d97706','#7c3aed','#dc2626'];

  const buSalesPct = bu.map(x => x.sales_pct);
  Plotly.newPlot('buSalesChart', [{
    type:'pie', labels:buNames, values:buSalesPct,
    text:buSalesPct.map(v => v.toFixed(1) + '%'), textinfo:'label+percent', textfont:{size:20,color:'#ffffff'},
    marker:{colors:buColors.slice(0,buNames.length),line:{color:'#fff',width:3}},
    hovertemplate:'%{label}<br>%{value:.1f}%<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:320, showlegend:true,
     legend:{orientation:'h',y:-0.12,font:{size:14,color:'#cbd5e1'}}},
    {responsive:true, displayModeBar:false});

  // Monthly Quantity Sold
  const qtyVals = months.map((_,i) => M[i].qty);
  Plotly.newPlot('qtyChart', [{type:'bar', x:months, y:qtyVals,
    marker:{color:'#059669'}, text:qtyVals.map(v => v.toLocaleString()), textposition:'outside', textfont:{size:20,color:'#ffffff'}, cliponaxis:false}],
    {margin:{t:30,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
     font:{size:14,color:'#cbd5e1'}, yaxis:{rangemode:'tozero',gridcolor:'rgba(255,255,255,0.05)'},
     height:320, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  // Monthly Net Sales vs COGS
  const ncogsV = months.map((_,i) => M[i].cogs * c.rate);
  const nnsV = months.map((_,i) => M[i].ns * c.rate);
  Plotly.newPlot('cogsChart', [
    {type:'bar', name:'Net Sales', x:months, y:nnsV, marker:{color:'#3b82f6'}, text:nnsV.map(v=>fmtNoSymbol(v)+' '+c.symbol), textposition:'outside', textfont:{size:16,color:'#ffffff'}, cliponaxis:false},
    {type:'bar', name:'COGS', x:months, y:ncogsV, marker:{color:'#d97706'}, text:ncogsV.map(v=>fmtNoSymbol(v)+' '+c.symbol), textposition:'outside', textfont:{size:16,color:'#ffffff'}, cliponaxis:false}
  ], {margin:{t:30,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:14,color:'#cbd5e1'}, barmode:'group', height:330, hovermode:'x unified',
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:14,color:'#cbd5e1'}},
      yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});
  setTimeout(updatePlotlyBlur, 100);
}

async function loadStyleAnalysis() {
  try {
    const r = await fetch('/api/style_analysis');
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
      {label:'Total Revenue', value:fmtS(totalRev*c.rate)+' '+c.symbol, color:'#22c55e'},
      {label:'Total COGS', value:fmtS(totalCogs*c.rate)+' '+c.symbol, color:'#f59e0b'},
      {label:'Gross Profit', value:fmtS(totalGP*c.rate)+' '+c.symbol, color:'#8b5cf6'},
      {label:'Net Profit', value:fmtS(totalNP*c.rate)+' '+c.symbol, color:'#14b8a6'},
      {label:'Avg GP Margin', value:fmtPct(avgGPM), color:'#22c55e'},
      {label:'Avg NP Margin', value:fmtPct(avgNPM), color:'#14b8a6'}
    ];
    document.getElementById('styleKpiGrid').innerHTML = kpis.map(k =>
      '<div class="kpi-card"><div class="kpi-icon" style="background:'+k.color+'"><i class="fas fa-chart-line" style="color:#fff"></i></div><div class="kpi-label">'+k.label+'</div><div class="kpi-value" style="color:'+k.color+'">'+k.value+'</div></div>'
    ).join('');
    // Waterfall
    const wfLabels = ['Planned Profit','Labor\nAdjustment','Overtime','Packing\nSavings','Clearance\nSavings','Finance\nSavings','SG&A\nSavings','Actual Profit'];
    const wfVals = [wf.planned_profit, wf.labor_cost_adjustment, wf.overtime, wf.saving_packing, wf.saving_clearance, wf.saving_finance, wf.saving_sga, wf.actual_profit];
    Plotly.newPlot('styleWaterfall', [{
      type:'waterfall', orientation:'v',
      x: wfLabels, y: wfVals,
      connector:{line:{color:'#475569',width:1,dash:'dot'}},
      decreasing:{marker:{color:'#ef4444'}},
      increasing:{marker:{color:'#22c55e'}},
      totals:{marker:{color:'#3b82f6'}},
      text: wfVals.map(v => fmtS(v*c.rate)),
      textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:40,b:70,l:80,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#cbd5e1'}, height:400, showlegend:false,
      yaxis:{ticksuffix:' '+c.symbol, gridcolor:'rgba(255,255,255,0.05)'}
    }, {responsive:true, displayModeBar:false});
    // Net Profit by PO (bar chart)
    const npVals = pf.total_net_profit.values.map(v => v * c.rate);
    const npColors = npVals.map(v => v >= 0 ? '#22c55e' : '#ef4444');
    Plotly.newPlot('styleNpChart', [{
      type:'bar', x:po, y:npVals, marker:{color:npColors},
      text:npVals.map(v=>fmtS(v)+' '+c.symbol), textposition:'outside', textfont:{size:13,color:'#ffffff'}, cliponaxis:false
    }], {
      margin:{t:40,b:45,l:65,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#cbd5e1'}, height:350, showlegend:false, xaxis:{tickangle:-30},
      yaxis:{ticksuffix:' '+c.symbol, gridcolor:'rgba(255,255,255,0.05)'}
    }, {responsive:true, displayModeBar:false});
    // COGS % pie
    Plotly.newPlot('styleCogsPie', [{
      type:'pie', labels:po, values:cb.total_cogs.values,
      text:cb.total_cogs.values.map(v=>fmtS(v*c.rate)), textinfo:'label+percent',
      textfont:{size:14,color:'#ffffff'},
      marker:{colors:['#3b82f6','#f59e0b','#22c55e','#8b5cf6','#ef4444'],line:{color:'#fff',width:2}},
      hovertemplate:'%{label}<br>'+c.symbol+': %{text}<br>%{percent}<extra></extra>'
    }], {
      margin:{t:10,b:10,l:10,r:10}, paper_bgcolor:'rgba(0,0,0,0)', height:350,
      showlegend:true, legend:{orientation:'h',y:-0.1,font:{size:12,color:'#cbd5e1'}}
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
    tbl += '<th style="padding:8px 12px;text-align:right;color:#f1f5f9;font-weight:600">Total</th><th style="padding:8px 12px;text-align:right;color:#3b82f6;font-weight:600">Forecast</th><th style="padding:8px 12px;text-align:right;color:#94a3b8;font-weight:600">Variance</th></tr></thead><tbody>';
    tableRows.forEach((row, i) => {
      const d = cb[row.key];
      const bg = i%2===0 ? 'rgba(255,255,255,0.02)' : 'transparent';
      tbl += '<tr style="border-bottom:1px solid #1e293b;background:'+bg+'"><td style="padding:8px 12px;color:#f1f5f9;font-weight:500">'+row.label+'</td>';
      d.values.forEach(v => { tbl += '<td style="padding:8px 12px;text-align:right">'+row.fmt(v)+'</td>'; });
      tbl += '<td style="padding:8px 12px;text-align:right;color:#f1f5f9;font-weight:600">'+row.fmt(d.total)+'</td>';
      tbl += '<td style="padding:8px 12px;text-align:right;color:#3b82f6">'+row.fmt(d.forecasted||0)+'</td>';
      const vv = d.variance||0;
      tbl += '<td style="padding:8px 12px;text-align:right;color:'+(vv>0?'#22c55e':vv<0?'#ef4444':'#94a3b8')+'">'+row.fmt(vv)+'</td></tr>';
    });
    tbl += '</tbody></table>';
    document.getElementById('styleTable').innerHTML = tbl;
    // Margin trend
    const gpMargins = cb.gross_profit_margin.values.map(v => v * 100);
    const npMargins = pf.net_profit_margin.values.map(v => v * 100);
    Plotly.newPlot('styleMarginChart', [
      {type:'scatter', mode:'lines+markers', name:'GP Margin', x:po, y:gpMargins, line:{color:'#8b5cf6',width:3}, marker:{size:10,color:'#8b5cf6'}, text:gpMargins.map(v=>v.toFixed(1)+'%'), textposition:'top center', textfont:{size:13,color:'#8b5cf6'}},
      {type:'scatter', mode:'lines+markers', name:'NP Margin', x:po, y:npMargins, line:{color:'#14b8a6',width:3}, marker:{size:10,color:'#14b8a6'}, text:npMargins.map(v=>v.toFixed(1)+'%'), textposition:'top center', textfont:{size:13,color:'#14b8a6'}}
    ], {
      margin:{t:50,b:50,l:60,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#cbd5e1'}, height:350, showlegend:true, xaxis:{tickangle:-30},
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#cbd5e1'}},
      yaxis:{ticksuffix:'%', gridcolor:'rgba(255,255,255,0.05)'}
    }, {responsive:true, displayModeBar:false});
    // Cost per item stacked bar
    const laborV = cb.labor_cost_per_item.values;
    const otV = cb.overtime_per_item.values;
    const packV = cb.packing_per_item.values;
    Plotly.newPlot('styleCostBar', [
      {type:'bar', name:'Labor', x:po, y:laborV, marker:{color:'#3b82f6'}},
      {type:'bar', name:'Overtime', x:po, y:otV, marker:{color:'#f59e0b'}},
      {type:'bar', name:'Packing', x:po, y:packV, marker:{color:'#22c55e'}}
    ], {
      margin:{t:40,b:45,l:60,r:25}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
      font:{size:12,color:'#cbd5e1'}, barmode:'stack', height:350, xaxis:{tickangle:-30},
      legend:{orientation:'h',y:1.15,x:.5,xanchor:'center',font:{size:13,color:'#cbd5e1'}},
      yaxis:{gridcolor:'rgba(255,255,255,0.05)'}
    }, {responsive:true, displayModeBar:false});
    // Variance text
    if (s.variance_detail) {
      document.getElementById('styleVariance').innerHTML = '<div style="padding:18px 22px"><div style="font-size:13px;font-weight:600;color:#f59e0b;margin-bottom:12px">'+s.variance_title+'</div><div style="font-size:13px;color:#94a3b8;line-height:1.8;white-space:pre-line">'+s.variance_detail+'</div></div>';
    }
    // Recommendations
    if (s.recommendations) {
      document.getElementById('styleRecommendations').innerHTML = '<div style="padding:18px 22px"><div style="font-size:13px;font-weight:600;color:#22c55e;margin-bottom:12px">'+s.rec_title+'</div><div style="font-size:13px;color:#94a3b8;line-height:1.8;white-space:pre-line">'+s.recommendations+'</div></div>';
    }
    setTimeout(updatePlotlyBlur, 100);
  } catch(e) { console.error('Style analysis error:', e); }
}

async function uploadExcel(input) {
  const file = input.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('file', file);
  const r = await fetch('/upload', {method:'POST', body:fd});
  const j = await r.json();
  if (j.ok) { loadData(); alert('Excel updated successfully!'); }
  else { alert('Error: ' + j.error); }
  input.value = '';
}

async function refreshData() {
  await fetch('/api/refresh');
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
  const r = await fetch('/api/pnl_forecast');
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
      '<div class="kpi-sub">Forecast: ' + pnlFmt(f) + '</div>' +
      '<div class="kpi-sub" style="font-weight:600;color:' + (good ? 'var(--green)' : 'var(--red)') + '">' +
      (isPos ? '+' : '') + pnlFmtShort(varAmt) + ' (' + (varPct !== 'N/A' ? (isPos ? '+' : '') + varPct + '%' : 'N/A') + ')</div></div>';
  }).join('');

  const colAct = '#3b82f6';
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
    {name:'Actual', y:gpAct, color:'#059669'},
    {name:'Forecast', y:gpFct, color:colFct}
  ], 'Gross Profit');

  // 3. Net Income chart
  const niAct = monthly.map(m => m.net_income.actual * cr.rate);
  const niFct = monthly.map(m => m.net_income.forecast * cr.rate);
  buildGroupedBar('pnlNiChart', pnlMonths, [
    {name:'Actual', y:niAct, color:'#14b8a6'},
    {name:'Forecast', y:niFct, color:colFct}
  ], 'Net Income');

  // 4. GP Margin trend
  const gpMarginAct = monthly.map(m => m.net_sales.actual > 0 ? (m.gross_profit.actual / m.net_sales.actual * 100) : 0);
  const gpMarginFct = monthly.map(m => m.net_sales.forecast > 0 ? (m.gross_profit.forecast / m.net_sales.forecast * 100) : 0);
  Plotly.newPlot('pnlMarginChart', [
    {type:'scatter', mode:'lines+markers', name:'Actual Margin', x:pnlMonths, y:gpMarginAct, line:{color:'#22c55e', width:3}, marker:{size:8, color:'#22c55e'}},
    {type:'scatter', mode:'lines+markers', name:'Forecast Margin', x:pnlMonths, y:gpMarginFct, line:{color:'#94a3b8', width:3, dash:'dot'}, marker:{size:8, color:'#94a3b8'}}
  ], {margin:{t:20,b:40,l:55,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
      yaxis:{ticksuffix:'%', gridcolor:'rgba(255,255,255,0.05)', rangemode:'tozero'}, height:300,
      hovermode:'x unified', legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}}},
    {responsive:true, displayModeBar:false});

  // 5. Department Expenses
  const depts = Object.keys(d.dept_expenses);
  const deptAct = depts.map(k => d.dept_expenses[k].actual * cr.rate);
  const deptFct = depts.map(k => d.dept_expenses[k].forecast * cr.rate);
  Plotly.newPlot('pnlDeptChart', [
    {type:'bar', orientation:'h', name:'Actual', x:deptAct, y:depts, marker:{color:'#8b5cf6'}, text:deptAct.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}},
    {type:'bar', orientation:'h', name:'Forecast', x:deptFct, y:depts, marker:{color:'#c4b5fd'}, text:deptFct.map(v=>pnlFmtShort(v)), textposition:'outside', textfont:{size:13,color:'#94a3b8'}}
  ], {margin:{t:15,b:25,l:160,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
      barmode:'group', height:Math.max(250, depts.length*50), hovermode:'y unified',
      legend:{orientation:'h',y:1.05,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}},
      xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  // 6. P&L Waterfall
  const pnlLabels = ['Net Sales', 'COGS', 'Gross Profit', 'Expenses', 'Net Income'];
  const pnlActVals = [ytd.net_sales.actual, -ytd.cogs.actual, ytd.gross_profit.actual, -ytd.expenses.actual, ytd.net_income.actual].map(v=>v*cr.rate);
  const pnlFctVals = [ytd.net_sales.forecast, -ytd.cogs.forecast, ytd.gross_profit.forecast, -ytd.expenses.forecast, ytd.net_income.forecast].map(v=>v*cr.rate);
  const waterfallMeasures = ['relative', 'relative', 'total', 'relative', 'total'];
  Plotly.newPlot('pnlWaterfallChart', [
    {type:'waterfall', name:'Actual', x:pnlLabels, y:pnlActVals, measure:waterfallMeasures,
     decreasing:{marker:{color:'#ef4444'}}, increasing:{marker:{color:'#22c55e'}}, totals:{marker:{color:'#3b82f6'}},
     connector:{line:{color:'rgba(255,255,255,0.2)', width:2}},
     text:pnlActVals.map(v => pnlFmtShort(v)), textposition:'outside', textfont:{size:11,color:'#e2e8f0'}},
    {type:'waterfall', name:'Forecast', x:pnlLabels, y:pnlFctVals, measure:waterfallMeasures,
     decreasing:{marker:{color:'#fca5a5'}}, increasing:{marker:{color:'#86efac'}}, totals:{marker:{color:'#93c5fd'}},
     connector:{line:{color:'rgba(255,255,255,0.1)', width:1}},
     text:pnlFctVals.map(v => pnlFmtShort(v)), textposition:'outside', textfont:{size:11,color:'#94a3b8'},
     opacity:0.7}
  ], {margin:{t:80,b:50,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
      height:350, hovermode:'x unified',
      legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}},
      yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

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
    margin:{t:50,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
    font:{color:'#cbd5e1',size:13}, barmode:'group', height:300, hovermode:'x unified',
    legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}},
    yaxis:{ticksuffix:' '+c.symbol,gridcolor:'rgba(255,255,255,0.05)'}
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
    fetch('/api/data').then(r=>r.json()),
    fetch('/api/sales_forecast').then(r=>r.json())
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
    {name:'Actual', y:gsAct, color:'#2563eb'},
    {name:'Budget', y:gsFct, color:'#94a3b8'}
  ]);

  // 2. Returns & Discounts
  const retVals = sf.monthly.map(m=>m.returns * cr.rate);
  const discVals = sf.monthly.map(m=>m.discounts * cr.rate);
  Plotly.newPlot('salesDedChart', [
    {type:'bar', name:'Returns', x:actMonths, y:retVals, marker:{color:'#ef4444'}, text:retVals.map(v=>salesFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#e2e8f0'}},
    {type:'bar', name:'Discounts', x:actMonths, y:discVals, marker:{color:'#f59e0b'}, text:discVals.map(v=>salesFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#e2e8f0'}}
  ], {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
      barmode:'group', height:300, hovermode:'x unified',
      legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}},
      yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  // 3. Quantity Sold
  const qtyV = actMonths.map((_,i)=>M[i].qty);
  Plotly.newPlot('salesQtyChart', [{type:'bar', x:actMonths, y:qtyV, marker:{color:'#d97706'}, text:qtyV.map(v=>v.toLocaleString()), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}, cliponaxis:false}],
    {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
     yaxis:{rangemode:'tozero',gridcolor:'rgba(255,255,255,0.05)'}, height:300, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  // 4. Sales by BU (pie)
  const bu = d.business_units;
  const buColors = ['#2563eb','#059669','#d97706','#7c3aed'];
  Plotly.newPlot('salesBUChart', [{
    type:'pie', labels:bu.map(x=>x.name), values:bu.map(x=>x.sales_pct),
    text:bu.map(x=>x.sales_pct.toFixed(1)+'%'), textinfo:'label+percent', textfont:{size:15,color:'#ffffff'},
    marker:{colors:buColors.slice(0,bu.length),line:{color:'#fff',width:2}},
    hovertemplate:'%{label}<br>%{value:.1f}%<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:290, showlegend:true,
     legend:{orientation:'h',y:-0.08,font:{size:12,color:'#cbd5e1'}}},
    {responsive:true, displayModeBar:false});

  // 5. Top Customers
  const tc = d.top_customers;
  Plotly.newPlot('salesTopChart', [{type:'bar', orientation:'h', x:tc.map(x=>x.sales*cr.rate), y:tc.map(x=>x.name),
    marker:{color:['#3b82f6','#60a5fa','#93c5fd','#bfdbfe','#dbeafe']},
    text:tc.map(x=>salesFmtShort(x.sales*cr.rate)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}}],
    {margin:{t:10,b:20,l:100,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
     height:Math.max(200, tc.length*40), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  // 6. Top Types
  const tt = d.top_types;
  Plotly.newPlot('salesTypeChart', [{type:'bar', orientation:'h', x:tt.map(x=>x.sales*cr.rate), y:tt.map(x=>x.name),
    marker:{color:'#3b82f6'}, text:tt.map(x=>salesFmtShort(x.sales*cr.rate)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}}],
    {margin:{t:10,b:20,l:120,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
     height:Math.max(200, tt.length*40), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  // 7. Top Fabrics
  const tf = d.top_fabrics;
  Plotly.newPlot('salesFabricChart', [{type:'bar', orientation:'h', x:tf.map(x=>x.sales*cr.rate), y:tf.map(x=>x.name),
    marker:{color:'#06b6d4'}, text:tf.map(x=>salesFmtShort(x.sales*cr.rate)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}}],
    {margin:{t:10,b:20,l:120,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
     height:Math.max(200, tf.length*40), hovermode:'y unified', showlegend:false,
     xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  // 8. Net Sales Conversion Rate %
  const convAct = actMonths.map((_,i)=>M[i].gs > 0 ? ((M[i].ns / M[i].gs) * 100) : 0);
  const convFct = sf.monthly.map(m=>m.gross_sales.forecast > 0 ? ((m.net_sales / m.gross_sales.forecast) * 100) : 0);
  Plotly.newPlot('salesConversionChart', [
    {type:'scatter', mode:'lines+markers', name:'Actual', x:actMonths, y:convAct, line:{color:'#22c55e', width:3}, marker:{size:8, color:'#22c55e'}},
    {type:'scatter', mode:'lines+markers', name:'Budget', x:actMonths, y:convFct, line:{color:'#94a3b8', width:3, dash:'dot'}, marker:{size:8, color:'#94a3b8'}}
  ], {margin:{t:15,b:40,l:55,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
      yaxis:{ticksuffix:'%', gridcolor:'rgba(255,255,255,0.05)', rangemode:'tozero'}, height:300,
      hovermode:'x unified', legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}}},
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
  const p = await fetch('/api/pnl_forecast').then(r=>r.json());
  const d = await fetch('/api/data').then(r=>r.json());
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
    {name:'Actual', y:expAct, color:'#ef4444'},
    {name:'Budget', y:expFct, color:'#94a3b8'}
  ]);

  // 2. Department Expenses
  const deptAct = depts.map(k=>p.dept_expenses[k].actual*cr.rate);
  const deptFct = depts.map(k=>p.dept_expenses[k].forecast*cr.rate);
  Plotly.newPlot('expDeptChart', [
    {type:'bar', orientation:'h', name:'Actual', x:deptAct, y:depts, marker:{color:'#8b5cf6'}, text:deptAct.map(v=>expFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#e2e8f0'}},
    {type:'bar', orientation:'h', name:'Forecast', x:deptFct, y:depts, marker:{color:'#c4b5fd'}, text:deptFct.map(v=>expFmtShort(v)), textposition:'outside', textfont:{size:12,color:'#94a3b8'}}
  ], {margin:{t:15,b:25,l:165,r:150}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
      barmode:'group', height:Math.max(250, depts.length*50), hovermode:'y unified',
      legend:{orientation:'h',y:1.05,x:.5,xanchor:'center',font:{size:12,color:'#cbd5e1'}},
      xaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}}, {responsive:true, displayModeBar:false});

  // 3. Monthly Variance (Actual - Forecast)
  const varVals = monthly.map(m=>(m.expenses.actual - m.expenses.forecast)*cr.rate);
  const varColors = varVals.map(v=>v <= 0 ? '#22c55e' : '#ef4444');
  Plotly.newPlot('expVarChart', [{type:'bar', x:months, y:varVals, marker:{color:varColors},
    text:varVals.map(v=>(v<=0?'':'+')+expFmtShort(v)), textposition:'outside', textfont:{size:13,color:'#e2e8f0'}, cliponaxis:false}],
    {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
     yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}, height:300, hovermode:'x unified', showlegend:false},
    {responsive:true, displayModeBar:false});

  // 4. Expense Distribution Pie
  const pieColors = ['#ef4444','#f59e0b','#8b5cf6','#3b82f6','#06b6d4','#22c55e','#ec4899'];
  Plotly.newPlot('expPieChart', [{
    type:'pie', labels:depts, values:deptAct,
    text:deptAct.map(v=>expFmtShort(v)), textinfo:'label+percent', textfont:{size:13,color:'#ffffff'},
    marker:{colors:pieColors.slice(0,depts.length),line:{color:'#fff',width:2}},
    hovertemplate:'%{label}<br>%{value:,.0f} '+cr.symbol+'<extra></extra>'}],
    {margin:{t:5,b:5,l:5,r:5}, paper_bgcolor:'rgba(0,0,0,0)', height:300, showlegend:true,
     legend:{orientation:'h',y:-0.1,font:{size:11,color:'#cbd5e1'}}},
    {responsive:true, displayModeBar:false});

  // 5. Top 3 Departments Monthly Trend
  const top3 = depts.sort((a,b)=>p.dept_expenses[b].actual - p.dept_expenses[a].actual).slice(0,3);
  const trendColors = ['#ef4444','#f59e0b','#8b5cf6'];
  const trendTraces = top3.map((d,i)=>({
    type:'scatter', mode:'lines+markers', name:d,
    x:months, y:p.dept_expenses[d].monthly.map(m=>m.actual*cr.rate),
    line:{color:trendColors[i], width:3}, marker:{size:7, color:trendColors[i]}
  }));
  Plotly.newPlot('expTrendChart', trendTraces,
    {margin:{t:45,b:40,l:60,r:20}, paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'#cbd5e1',size:13},
     yaxis:{ticksuffix:' '+cr.symbol,gridcolor:'rgba(255,255,255,0.05)'}, height:300, hovermode:'x unified',
     legend:{orientation:'h',y:1.1,x:.5,xanchor:'center',font:{size:11,color:'#cbd5e1'}}},
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
