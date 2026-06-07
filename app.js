// State Management
let state = {
  fileName: '',
  workbook: null,
  activeSheet: '',
  headers: [],
  rawData: [],
  columnMapping: {
    customer: '',
    family: '',
    product: '',
    date: '',
    location: '',
    teeth: ''
  },
  filteredData: [],
  excludedData: [], // { row: Object, reason: String }
  reportData: [],   // Grouped summary
  chartInstance: null,
  sortField: 'RequiredDate',
  sortAsc: true
};

// Exclude lists (in uppercase, trimmed)
const EXCLUDED_LOCATIONS = ["PENDING", "WAITING", "REGISTRATION IMPLANT COMPONENTS", "QA", "FINALATTACHEMENTPACK"];
const EXCLUDED_CUSTOMERS = ["SMILE DENTAL CARE", "EMPIRE DENTAL CARE"];

// Initialize UI Elements
document.addEventListener('DOMContentLoaded', () => {
  initIcons();
  setupEventListeners();
  setupDefaultDate();
});

function initIcons() {
  if (typeof lucide !== 'undefined') {
    lucide.createIcons();
  }
}

function setupDefaultDate() {
  // Set date filter to today
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  const formattedToday = `${yyyy}-${mm}-${dd}`;
  document.getElementById('filter-date-input').value = formattedToday;
}

// Setup Event Listeners
function setupEventListeners() {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  
  // Drag and drop events
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    }, false);
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
    }, false);
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
      handleFile(files[0]);
    }
  }, false);

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0]);
    }
  });

  // Reset button
  document.getElementById('btn-reset').addEventListener('click', resetApp);

  // Sheet selection change
  document.getElementById('sheet-selector').addEventListener('change', (e) => {
    state.activeSheet = e.target.value;
    loadSheetData();
  });

  // Sidebar Filter toggles and inputs
  const filterToggles = [
    'filter-toggle-date',
    'filter-toggle-family',
    'filter-toggle-trial',
    'filter-toggle-location',
    'filter-toggle-customer'
  ];
  filterToggles.forEach(id => {
    document.getElementById(id).addEventListener('change', (e) => {
      const item = e.target.closest('.filter-item');
      if (e.target.checked) {
        item.classList.remove('inactive');
      } else {
        item.classList.add('inactive');
      }
      processFiltersAndRender();
    });
  });

  document.getElementById('filter-date-input').addEventListener('change', processFiltersAndRender);

  // Tabs navigation
  const tabs = document.querySelectorAll('.tab-btn');
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      
      const tabId = tab.dataset.tab;
      document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
      });
      document.getElementById(tabId).classList.add('active');
      
      // Update icons inside tabs
      initIcons();
    });
  });

  // Search input listeners
  document.getElementById('report-search').addEventListener('input', renderReportTable);
  document.getElementById('raw-search').addEventListener('input', renderRawTable);
  document.getElementById('excluded-search').addEventListener('input', renderExcludedTable);

  // Exporters
  document.getElementById('btn-export-csv').addEventListener('click', exportCSV);
  document.getElementById('btn-export-excel').addEventListener('click', exportExcel);

  // Sort report headers
  document.querySelectorAll('#summary-table th.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const field = th.dataset.sort;
      if (state.sortField === field) {
        state.sortAsc = !state.sortAsc;
      } else {
        state.sortField = field;
        state.sortAsc = true;
      }
      
      // Update header visual classes
      document.querySelectorAll('#summary-table th.sortable').forEach(header => {
        header.classList.remove('sort-asc', 'sort-desc');
      });
      th.classList.add(state.sortAsc ? 'sort-asc' : 'sort-desc');
      
      sortAndRenderReportTable();
    });
  });

  // Modal actions
  document.getElementById('btn-close-mapping').addEventListener('click', closeModal);
  document.getElementById('btn-cancel-mapping').addEventListener('click', closeModal);
  document.getElementById('btn-confirm-mapping').addEventListener('click', saveManualMapping);
}

// Handle Uploaded File
function handleFile(file) {
  state.fileName = file.name;
  document.getElementById('meta-filename').textContent = file.name;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      state.workbook = XLSX.read(data, { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
      
      // Populate sheets drop-down
      const sheetSelector = document.getElementById('sheet-selector');
      sheetSelector.innerHTML = '';
      
      state.workbook.SheetNames.forEach(sheetName => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelector.appendChild(option);
      });

      // Show reset button & dashboard, hide upload
      document.getElementById('upload-stage').style.display = 'none';
      document.getElementById('dashboard-stage').style.display = 'block';
      document.getElementById('reset-container').style.display = 'block';

      // Load first sheet by default
      state.activeSheet = state.workbook.SheetNames[0];
      loadSheetData();
      
      initIcons();
    } catch (error) {
      alert('Error parsing Excel file. Please ensure it is a valid spreadsheet. ' + error.message);
      console.error(error);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Load spreadsheet sheet and detect headers
function loadSheetData() {
  const sheet = state.workbook.Sheets[state.activeSheet];
  
  // Read raw JSON data including headers
  const rawJson = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  if (rawJson.length === 0) {
    alert('The selected sheet is empty.');
    return;
  }

  // Detect Headers
  state.headers = Object.keys(rawJson[0]);
  state.rawData = rawJson;
  
  // Auto-map headers
  const success = autoMapHeaders();
  
  if (success) {
    processFiltersAndRender();
  } else {
    // Open mapping modal
    openMappingModal();
  }
}

// Auto map headers using keywords
function isToothNumberHeader(header) {
  const norm = header.toLowerCase().replace(/[\s_-]/g, '');
  return norm.includes('no') || norm.includes('num') || norm.includes('id') || norm === 'unit' || norm === 'units';
}

function autoMapHeaders() {
  const mappingRules = {
    customer: ['customername', 'customernam', 'customer', 'clinicname', 'clinic', 'account'],
    family: ['familyname', 'family_name', 'family name', 'family'],
    product: ['productname', 'product_name', 'product name', 'product', 'item'],
    date: ['requireddate', 'required date', 'required_date', 'reqdate', 'req date', 'date', 'orderdate', 'regdate'],
    location: ['lastscanninglocation', 'scanninglocation', 'last scanning location', 'scanning location', 'scanlocation', 'location', 'status'],
    teeth: ['teethcount', 'teeth count', 'teeth qty', 'teethqty', 'no of teeth', 'no. of teeth', 'teeth_count', 'qty', 'quantity', 'teeth']
  };

  // Reset mapping
  for (let key in state.columnMapping) {
    state.columnMapping[key] = '';
  }

  // Pass 1: Try exact matches (normalized)
  for (let key in mappingRules) {
    const rules = mappingRules[key];
    const match = state.headers.find(header => {
      const normHeader = header.toLowerCase().replace(/[\s_-]/g, '');
      
      // For teeth count, explicitly ignore headers that represent tooth numbers/IDs
      if (key === 'teeth' && isToothNumberHeader(header)) {
        return false;
      }
      
      return rules.some(rule => {
        const normRule = rule.toLowerCase().replace(/[\s_-]/g, '');
        return normHeader === normRule;
      });
    });
    if (match) {
      state.columnMapping[key] = match;
    }
  }

  // Pass 2: Fallback to substring matching for any unmapped fields
  for (let key in mappingRules) {
    if (state.columnMapping[key]) continue; // Already mapped in Pass 1

    const rules = mappingRules[key];
    const match = state.headers.find(header => {
      const normHeader = header.toLowerCase().replace(/[\s_-]/g, '');
      
      // For teeth count, explicitly ignore headers that represent tooth numbers/IDs
      if (key === 'teeth' && isToothNumberHeader(header)) {
        return false;
      }
      
      return rules.some(rule => {
        const normRule = rule.toLowerCase().replace(/[\s_-]/g, '');
        return normHeader.includes(normRule);
      });
    });
    if (match) {
      state.columnMapping[key] = match;
    }
  }

  // Check if mapping succeeded (all required keys are mapped)
  const missing = Object.keys(state.columnMapping).filter(k => !state.columnMapping[k]);
  
  return missing.length === 0;
}

// Open Mapping Modal UI
function openMappingModal() {
  const modal = document.getElementById('mapping-modal');
  modal.classList.add('open');

  const fields = ['customer', 'family', 'product', 'date', 'location', 'teeth'];
  fields.forEach(field => {
    const select = document.getElementById(`map-${field}`);
    select.innerHTML = '';
    
    // Add empty option
    const emptyOpt = document.createElement('option');
    emptyOpt.value = '';
    emptyOpt.textContent = '-- Select Excel Column --';
    select.appendChild(emptyOpt);

    state.headers.forEach(header => {
      const option = document.createElement('option');
      option.value = header;
      option.textContent = header;
      if (state.columnMapping[field] === header) {
        option.selected = true;
      }
      select.appendChild(option);
    });
  });
}

function closeModal() {
  document.getElementById('mapping-modal').classList.remove('open');
}

function saveManualMapping() {
  const fields = ['customer', 'family', 'product', 'date', 'location', 'teeth'];
  let allMapped = true;
  
  fields.forEach(field => {
    const val = document.getElementById(`map-${field}`).value;
    state.columnMapping[field] = val;
    if (!val) allMapped = false;
  });

  closeModal();
  
  if (!allMapped) {
    alert('Warning: Some columns are not mapped. Reports might calculate incomplete totals.');
  }
  
  processFiltersAndRender();
}

// Helper to safely parse Dates from Excel format
function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return val;
  }
  
  // If Excel number serial (e.g. 44561)
  if (typeof val === 'number') {
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }

  // Parse string
  const str = String(val).trim();
  
  // If string contains only digits (Excel date serial stored as string)
  if (/^\d+(\.\d+)?$/.test(str)) {
    const num = parseFloat(str);
    if (num > 10000 && num < 100000) {
      return new Date(Math.round((num - 25569) * 86400 * 1000));
    }
  }

  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    return new Date(parsed);
  }

  // Custom DD/MM/YYYY or similar parsing if standard fails
  const parts = str.split(/[-/]/);
  if (parts.length === 3) {
    // Try DD-MM-YYYY
    if (parts[0].length <= 2 && parts[2].length === 4) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const year = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }
    // Try YYYY-MM-DD
    if (parts[0].length === 4 && parts[2].length <= 2) {
      const year = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      const day = parseInt(parts[2], 10);
      return new Date(year, month, day);
    }
  }

  return null;
}

// Format Date object to YYYY-MM-DD string
function formatDateToString(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return 'N/A';
  const yyyy = dateObj.getFullYear();
  const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// Safe parsing of teeth count to integer
function parseTeethCount(val) {
  if (val === undefined || val === null || val === '') return 1; // Default to 1 if empty
  const parsed = parseInt(val, 10);
  return isNaN(parsed) ? 1 : parsed;
}

// Core filter and processing engine
function processFiltersAndRender() {
  state.filteredData = [];
  state.excludedData = [];

  // Active filter toggle checkmarks
  const checkDate = document.getElementById('filter-toggle-date').checked;
  const checkFamily = document.getElementById('filter-toggle-family').checked;
  const checkTrial = document.getElementById('filter-toggle-trial').checked;
  const checkLocation = document.getElementById('filter-toggle-location').checked;
  const checkCustomer = document.getElementById('filter-toggle-customer').checked;

  const dateCutoffStr = document.getElementById('filter-date-input').value;
  const cutoffDate = dateCutoffStr ? new Date(dateCutoffStr + 'T23:59:59') : null;

  const mapping = state.columnMapping;

  // Process rows
  state.rawData.forEach((row, idx) => {
    // Read raw values based on column mapping
    const custVal = String(row[mapping.customer] || '').trim();
    const familyVal = String(row[mapping.family] || '').trim();
    const prodVal = String(row[mapping.product] || '').trim();
    const dateRawVal = row[mapping.date];
    const locVal = String(row[mapping.location] || '').trim();
    const teethVal = row[mapping.teeth];

    const parsedDate = parseExcelDate(dateRawVal);
    const parsedTeeth = parseTeethCount(teethVal);

    // Exclusion Reasons
    let isExcluded = false;
    let excludeReason = '';

    // Rule 1: Till Date cutoff (RequiredDate <= cutoffDate)
    if (checkDate && cutoffDate) {
      if (!parsedDate) {
        isExcluded = true;
        excludeReason = 'Invalid/Missing Date';
      } else if (parsedDate > cutoffDate) {
        isExcluded = true;
        excludeReason = `Future Date (Required: ${formatDateToString(parsedDate)} > Cutoff: ${dateCutoffStr})`;
      }
    }

    // Rule 2: Exclude FamilyName == "MISCELLANEOUS"
    if (!isExcluded && checkFamily && familyVal.toUpperCase() === 'MISCELLANEOUS') {
      isExcluded = true;
      excludeReason = 'Family Name matches "MISCELLANEOUS"';
    }

    // Rule 3: Exclude ProductName contains "TRIAL"
    if (!isExcluded && checkTrial && prodVal.toUpperCase().includes('TRIAL')) {
      isExcluded = true;
      excludeReason = 'Trial Product (Product Name contains "TRIAL")';
    }

    // Rule 4: Exclude LastScanningLocation
    if (!isExcluded && checkLocation) {
      const upperLoc = locVal.toUpperCase();
      if (EXCLUDED_LOCATIONS.includes(upperLoc)) {
        isExcluded = true;
        excludeReason = `Excluded Location Stage (${locVal})`;
      }
    }

    // Rule 5: Exclude Customers ("SMILE DENTAL CARE", "EMPIRE DENTAL CARE")
    if (!isExcluded && checkCustomer) {
      const upperCust = custVal.toUpperCase();
      const match = EXCLUDED_CUSTOMERS.some(blockedCust => upperCust.includes(blockedCust));
      if (match) {
        isExcluded = true;
        excludeReason = `Excluded Customer (${custVal})`;
      }
    }

    // Assemble unified record structure
    const record = {
      id: idx,
      customerName: custVal,
      familyName: familyVal,
      productName: prodVal,
      requiredDateStr: parsedDate ? formatDateToString(parsedDate) : 'N/A',
      requiredDateObj: parsedDate,
      scanningLocation: locVal || 'Empty Location',
      teethCount: parsedTeeth,
      rawRow: row
    };

    if (isExcluded) {
      state.excludedData.push({
        record: record,
        reason: excludeReason
      });
    } else {
      state.filteredData.push(record);
    }
  });

  // Calculate Aggregated Teeth Count Report
  // Group by RequiredDate and LastScanningLocation
  const groups = {};
  state.filteredData.forEach(item => {
    const key = `${item.requiredDateStr} | ${item.scanningLocation}`;
    if (!groups[key]) {
      groups[key] = {
        RequiredDate: item.requiredDateStr,
        LastScanningLocation: item.scanningLocation,
        TeethCount: 0
      };
    }
    groups[key].TeethCount += item.teethCount;
  });

  state.reportData = Object.values(groups);

  // Update Stats Counters
  document.getElementById('stat-total-rows').textContent = state.rawData.length;
  document.getElementById('stat-processed-rows').textContent = state.filteredData.length;
  document.getElementById('stat-excluded-rows').textContent = state.excludedData.length;
  
  const totalTeeth = state.filteredData.reduce((sum, item) => sum + item.teethCount, 0);
  document.getElementById('stat-teeth-sum').textContent = totalTeeth;

  // Render lists and charts
  sortAndRenderReportTable();
  renderRawTable();
  renderExcludedTable();
  renderTeethChart();
}

// Sort & Render Report Table
function sortAndRenderReportTable() {
  const field = state.sortField;
  const asc = state.sortAsc;

  state.reportData.sort((a, b) => {
    let valA = a[field];
    let valB = b[field];

    if (field === 'TeethCount') {
      return asc ? valA - valB : valB - valA;
    }

    // String comparison (date string works lexicographically as YYYY-MM-DD)
    valA = String(valA).toLowerCase();
    valB = String(valB).toLowerCase();
    
    if (valA < valB) return asc ? -1 : 1;
    if (valA > valB) return asc ? 1 : -1;
    return 0;
  });

  renderReportTable();
}

// Populate main grouped report table
function renderReportTable() {
  const searchQuery = document.getElementById('report-search').value.toLowerCase();
  const tbody = document.getElementById('summary-table-body');
  const emptyState = document.getElementById('report-empty');
  const table = document.getElementById('summary-table');
  
  tbody.innerHTML = '';
  
  const filteredReport = state.reportData.filter(row => {
    return row.RequiredDate.toLowerCase().includes(searchQuery) || 
           row.LastScanningLocation.toLowerCase().includes(searchQuery);
  });

  if (filteredReport.length === 0) {
    table.style.display = 'none';
    emptyState.style.display = 'flex';
  } else {
    table.style.display = 'table';
    emptyState.style.display = 'none';

    filteredReport.forEach(row => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${row.RequiredDate}</strong></td>
        <td>${row.LastScanningLocation}</td>
        <td><span class="badge badge-primary" style="font-size: 0.85rem; padding: 0.25rem 0.65rem;">${row.TeethCount}</span></td>
      `;
      tbody.appendChild(tr);
    });
  }
}

// Populate filtered raw rows tab
function renderRawTable() {
  const searchQuery = document.getElementById('raw-search').value.toLowerCase();
  const tbody = document.getElementById('raw-table-body');
  
  tbody.innerHTML = '';

  const matched = state.filteredData.filter(item => {
    return item.customerName.toLowerCase().includes(searchQuery) ||
           item.familyName.toLowerCase().includes(searchQuery) ||
           item.productName.toLowerCase().includes(searchQuery) ||
           item.requiredDateStr.toLowerCase().includes(searchQuery) ||
           item.scanningLocation.toLowerCase().includes(searchQuery);
  });

  document.getElementById('raw-row-count').textContent = `Showing ${matched.length} of ${state.filteredData.length} records`;

  matched.slice(0, 300).forEach(item => { // Limit to 300 rows for DOM performance
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${item.customerName || '<span class="text-muted">N/A</span>'}</td>
      <td>${item.familyName}</td>
      <td title="${item.productName}" style="max-width: 280px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${item.productName}</td>
      <td>${item.requiredDateStr}</td>
      <td>${item.scanningLocation}</td>
      <td>${item.teethCount}</td>
    `;
    tbody.appendChild(tr);
  });

  if (matched.length > 300) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="6" style="text-align: center; color: var(--text-muted); font-style: italic;">Showing top 300 rows. Use Excel/CSV export to download all data.</td>`;
    tbody.appendChild(tr);
  }
}

// Populate excluded logs tab
function renderExcludedTable() {
  const searchQuery = document.getElementById('excluded-search').value.toLowerCase();
  const tbody = document.getElementById('excluded-table-body');
  
  tbody.innerHTML = '';

  const matched = state.excludedData.filter(item => {
    return item.reason.toLowerCase().includes(searchQuery) ||
           item.record.customerName.toLowerCase().includes(searchQuery) ||
           item.record.familyName.toLowerCase().includes(searchQuery) ||
           item.record.productName.toLowerCase().includes(searchQuery) ||
           item.record.scanningLocation.toLowerCase().includes(searchQuery);
  });

  document.getElementById('excluded-row-count').textContent = `${matched.length} rows matching query (Total Excluded: ${state.excludedData.length})`;

  matched.slice(0, 300).forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="color: var(--danger); font-weight: 500;"><span style="display: flex; align-items: center; gap: 0.25rem;"><i data-lucide="x-circle" style="width: 14px; height: 14px;"></i> ${item.reason}</span></td>
      <td>${item.record.customerName || '<span class="text-muted">N/A</span>'}</td>
      <td>${item.record.familyName}</td>
      <td title="${item.record.productName}" style="max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${item.record.productName}</td>
      <td>${item.record.requiredDateStr}</td>
      <td>${item.record.scanningLocation}</td>
      <td>${item.record.teethCount}</td>
    `;
    tbody.appendChild(tr);
  });

  if (matched.length > 300) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="7" style="text-align: center; color: var(--text-muted); font-style: italic;">Showing top 300 rows.</td>`;
    tbody.appendChild(tr);
  }
  
  initIcons();
}

// Generate & Render Visual Charts using ChartJS
function renderTeethChart() {
  const ctx = document.getElementById('teethChart').getContext('2d');
  
  // Aggregate sum of teeth by scanning location
  const locTotals = {};
  state.filteredData.forEach(item => {
    const loc = item.scanningLocation;
    locTotals[loc] = (locTotals[loc] || 0) + item.teethCount;
  });

  const labels = Object.keys(locTotals);
  const data = Object.values(locTotals);

  if (state.chartInstance) {
    state.chartInstance.destroy();
  }

  // Render beautiful horizontal bar chart
  state.chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Teeth Count Sum',
        data: data,
        backgroundColor: [
          'rgba(139, 92, 246, 0.65)',
          'rgba(6, 182, 212, 0.65)',
          'rgba(16, 185, 129, 0.65)',
          'rgba(245, 158, 11, 0.65)',
          'rgba(236, 72, 153, 0.65)',
          'rgba(59, 130, 246, 0.65)'
        ],
        borderColor: [
          '#8b5cf6',
          '#06b6d4',
          '#10b981',
          '#f59e0b',
          '#ec4899',
          '#3b82f6'
        ],
        borderWidth: 1,
        borderRadius: 6,
        borderSkipped: false
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false
        },
        tooltip: {
          backgroundColor: '#121826',
          titleFont: { family: 'Outfit', size: 14 },
          bodyFont: { family: 'Inter', size: 12 },
          borderColor: 'rgba(255, 255, 255, 0.1)',
          borderWidth: 1,
          padding: 10
        }
      },
      scales: {
        x: {
          grid: {
            color: 'rgba(255, 255, 255, 0.05)',
          },
          ticks: {
            color: '#9ca3af',
            font: { family: 'Inter', size: 11 }
          }
        },
        y: {
          grid: {
            color: 'rgba(255, 255, 255, 0.05)',
          },
          ticks: {
            color: '#9ca3af',
            font: { family: 'Inter', size: 11 }
          }
        }
      }
    }
  });
}

// Reset App to upload new spreadsheet
function resetApp() {
  state.fileName = '';
  state.workbook = null;
  state.activeSheet = '';
  state.headers = [];
  state.rawData = [];
  state.filteredData = [];
  state.excludedData = [];
  state.reportData = [];
  
  if (state.chartInstance) {
    state.chartInstance.destroy();
    state.chartInstance = null;
  }

  document.getElementById('file-input').value = '';
  document.getElementById('upload-stage').style.display = 'block';
  document.getElementById('dashboard-stage').style.display = 'none';
  document.getElementById('reset-container').style.display = 'none';
}

// Export Summary Report to CSV
function exportCSV() {
  if (state.reportData.length === 0) return;

  let csvContent = 'data:text/csv;charset=utf-8,';
  csvContent += 'Required Date,Last Scanning Location,Sum of Teeth Count\n';

  state.reportData.forEach(row => {
    // Quote scanning location to escape commas
    const escapedLoc = `"${row.LastScanningLocation.replace(/"/g, '""')}"`;
    csvContent += `${row.RequiredDate},${escapedLoc},${row.TeethCount}\n`;
  });

  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `Teeth_Count_Summary_Report_${state.fileName.split('.')[0]}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// Export Summary Report to fully structured Excel workbook
function exportExcel() {
  if (state.reportData.length === 0) return;

  // Create worksheet for summary
  const summarySheetData = state.reportData.map(row => ({
    'Required Date': row.RequiredDate,
    'Last Scanning Location': row.LastScanningLocation,
    'Sum of Teeth Count': row.TeethCount
  }));
  const wsSummary = XLSX.utils.json_to_sheet(summarySheetData);

  // Set column widths
  const wscols = [
    { wch: 15 }, // Required Date
    { wch: 35 }, // Last Scanning Location
    { wch: 20 }  // Sum of Teeth Count
  ];
  wsSummary['!cols'] = wscols;

  // Create workbook
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary Report');

  // Also append filtered raw data as secondary audit sheet
  const rawSheetData = state.filteredData.map(item => ({
    'Customer Name': item.customerName,
    'Family Name': item.familyName,
    'Product Name': item.productName,
    'Required Date': item.requiredDateStr,
    'Scanning Location': item.scanningLocation,
    'Teeth Count': item.teethCount
  }));
  const wsRaw = XLSX.utils.json_to_sheet(rawSheetData);
  XLSX.utils.book_append_sheet(wb, wsRaw, 'Filtered Raw Data');

  // Download Excel Workbook
  const filename = `Teeth_Count_Summary_Report_${state.fileName.split('.')[0]}.xlsx`;
  XLSX.writeFile(wb, filename);
}
