// Reference to buttons
const downloadBtn = document.getElementById('download-template');
const loadBtn = document.getElementById('load-template');
const refreshBtn = document.getElementById('refresh-template');

let currentData = null;

// Function to create and download XLSX template
function downloadTemplate() {
  const ws_data = [
    [
      "KPI Name",
      "Current",
      "Last"
    ],
    // Starting data rows
    ["Current Assets", 0, 0],
    ["Average Inventory", 0, 0],
    ["Inventories", 0, 0],
    ["Average Net Fixed Assets", 0, 0],
    ["Average Total Assets", 0, 0],
    ["Investment", 0, 0],
    ["Total Assets", 0, 0],
    ["Current Liabilities", 0, 0],
    ["Average Accounts Payable", 0, 0],
    ["Long Term Debt", 0, 0],
    ["Total Liabilities", 0, 0],
    ["Number of Outstanding Shares", 0, 0],
    ["Shareholders Equity", 0, 0],
    ["Net Sales", 0, 0],
    ["COGS", 0, 0],
    ["Purchases of COGS", 0, 0],
    ["EBITDA", 0, 0],
    ["Net Income", 0, 0],
    ["Net Revenue", 0, 0],
    ["Net Profit", 0, 0],
    ["EBIT", 0, 0],
    ["Interest", 0, 0],
    ["Tax", 0, 0],
    ["Average Working Capital", 0, 0],
    ["Cash Flow From Operating Activities", 0, 0],
    ["Capital Expenditures", 0, 0],
    ["Principal", 0, 0],
    ["Dividends", 0, 0],
    ["Market Value per Share", 0, 0],
    ["Earnings Per Share (EPS)", 0, 0],
    ["Non-Operating Cash", 0, 0],
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, "Template");

  XLSX.writeFile(wb, "KPI_Template.xlsx");
}

// Function to load XLSX template
function loadTemplate() {
  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = '.xlsx';

  fileInput.onchange = e => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, {type: 'array'});
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      currentData = XLSX.utils.sheet_to_json(worksheet, {header:1});
      processData();
    };
    reader.readAsArrayBuffer(file);
  };
  fileInput.click();
}

// Function to refresh the template (clear data)
function refreshTemplate() {
  currentData = null;
  clearKPIs();
}

// Function to clear KPI display
function clearKPIs() {
  document.querySelectorAll('.value').forEach(span => span.textContent = '');
}

// Function to process loaded data and calculate KPIs
function processData() {
  if (!currentData || currentData.length === 0) return;

  // Map KPI name to data row
  const dataMap = {};
  currentData.forEach(row => {
    if (row.length >=3 && row[0]) {
      dataMap[row[0]] = { current: parseFloat(row[1]) || 0, last: parseFloat(row[2]) || 0 };
    }
  });

  // Helper to get value
  function getVal(kpi) {
    return dataMap[kpi] ? dataMap[kpi].current : 0;
  }
  function getValLast(kpi) {
    return dataMap[kpi] ? dataMap[kpi].last : 0;
  }

  // Calculate and set KPIs for each group
  // For example, General KPIs
  setKPI('CR', getVal('Current Assets'), getValLast('Current Assets'), (c,l) => c / getVal('Current Liabilities'));
  setKPI('QR', getVal('Current Assets'), getValLast('Current Assets'), (c,l) => (c - getVal('Inventories')) / getVal('Current Liabilities'));
  setKPI('DER', getVal('Total Liabilities'), getVal('Shareholders Equity'), (t, s) => t / s);
  setKPI('DR', getVal('Total Liabilities'), getVal('Total Assets'), (t, a) => t / a);
  setKPI('ER', getVal('Shareholders Equity'), getVal('Total Assets'), (s, a) => s / a);
  setKPI('BVPS', getVal('Shareholders Equity'), getVal('Number of Outstanding Shares'), (s, n) => s / n);
  setKPI('EPS', getVal('Net Income'), getVal('Number of Outstanding Shares'), (n, ) => n);
  // P/E ratio
  const eps = getVal('Earnings Per Share (EPS)');
  const marketPrice = getVal('Market Value per Share');
  document.getElementById('PE').textContent = eps ? (marketPrice / eps).toFixed(2) : '0.00';

  // Operative KPIs
  setKPI('WC', getVal('Current Assets'), getValLast('Current Assets'), (c, l) => c - getVal('Current Liabilities'));
  setKPI('IT', getVal('Net Sales'), getVal('Average Inventory'), (s, i) => s / i);
  setKPI('APT', getVal('Purchases of COGS'), getVal('Average Accounts Payable'), (p, a) => p / a);
  setKPI('ATR', getVal('Net Sales'), getVal('Average Total Assets'), (s, a) => s / a);
  setKPI('FAT', getVal('Net Sales'), getVal('Average Net Fixed Assets'), (s, a) => s / a);
  setKPI('WCT', getVal('Net Sales'), getVal('Average Working Capital'), (s, wc) => s / wc);
  setKPI('TR', getVal('Net Revenue'), getVal('Total Assets'), (s, a) => s / a);
  setKPI('DPR', getVal('Dividends'), getVal('Net Income'), (d, n) => d / n);

  // Cash Flow KPIs
  setKPI('OCFR', getVal('Cash Flow From Operating Activities'), getVal('Current Liabilities'), (c, l) => c / l);
  setKPI('FCF', getVal('Cash Flow From Operating Activities'), getVal('Capital Expenditures'), (c, capex) => c - capex);
  setKPI('CFDR', getVal('Cash Flow From Operating Activities'), getVal('Total Liabilities'), (c, t) => t ? c / t : 0);
  setKPI('CFM', getVal('Cash Flow From Operating Activities'), getVal('Net Sales'), (c, s) => s ? c / s : 0);
  setKPI('CRA', getVal('Cash Flow From Operating Activities'), getVal('Total Assets'), (c, a) => c / a);
  setKPI('DCR', getVal('EBITDA'), getVal('Principal') + getVal('Interest'), (e, d) => e / d);
  setKPI('CCR', getVal('Cash Flow From Operating Activities'), getVal('Net Income'), (c, n) => n ? c / n : 0);

  // Returns KPIs
  setKPI('ROA', getVal('Net Income'), getVal('Total Assets'), (n, a) => n / a);
  setKPI('ROE', getVal('Net Income'), getVal('Shareholders Equity'), (n, s) => n / s);
  setKPI('ROI', getVal('Net Profit'), getVal('Investment'), (np, inv) => np / inv);
  setKPI('ROIC', getVal('EBIT') * (1 - (getVal('Tax')/100)), getVal('Long Term Debt') + getVal('Shareholders Equity') - getVal('Non-Operating Cash'), (e, d) => e / d);
  setKPI('ROE_alt', getVal('EBIT') - getVal('Interest') - getVal('Tax'), getVal('Shareholders Equity'), (e, s) => e / s);
  setKPI('ROCE', getVal('EBIT'), getVal('Long Term Debt') + getVal('Shareholders Equity'), (e, d) => e / d);
  setKPI('ROA_alt', getVal('EBIT') - getVal('Interest') - getVal('Tax'), getVal('Total Assets'), (e, a) => e / a);
}

// Helper to set KPI values
function setKPI(id, currentVal, lastVal, calcFn) {
  const currentValResult = calcFn(currentVal, lastVal);
  document.getElementById(id).textContent = currentValResult !== undefined && !isNaN(currentValResult)
    ? currentValResult.toFixed(2)
    : '0.00';
}

// Button event listeners
downloadBtn.onclick = downloadTemplate;
loadBtn.onclick = loadTemplate;
refreshBtn.onclick = refreshTemplate;

// Toggle sections
document.querySelectorAll('.toggle-btn').forEach(btn => {
  btn.onclick = () => {
    const sectionId = 'group-' + btn.dataset.section + '-values';
    const section = document.getElementById(sectionId);
    if (section.style.display === 'none') {
      section.style.display = 'block';
    } else {
      section.style.display = 'none';
    }
  };
});
