const IGNORED_SHEETS = ['Export Summary'];

function populateTable(rows) {
  const tbody = document.getElementById('data-table').querySelector('tbody');
  tbody.innerHTML = '';
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.DBirthDate.v}</td>
      <td>${row.DDeathDate.v}</td>
      <td>${row.Y.v}</td>
      <td>${row.M.v}</td>
      <td>${row.D.v}</td>
    `;
    tbody.appendChild(tr);
  });

  // Apply fancy styles
  const table = document.getElementById('data-table');
  table.style.borderCollapse = 'collapse';
  table.style.width = '100%';
  table.querySelectorAll('th, td').forEach((cell) => {
    cell.style.border = '1px solid #ccc';
    cell.style.padding = '10px';
  });
  table.querySelectorAll('tr:nth-child(even)').forEach((row) => {
    row.style.backgroundColor = '#e6f7ff';
  });
  table.querySelectorAll('tr:hover').forEach((row) => {
    row.style.backgroundColor = '#cceeff';
  });
  table.querySelectorAll('th').forEach((header) => {
    header.style.paddingTop = '14px';
    header.style.paddingBottom = '14px';
    header.style.textAlign = 'left';
    header.style.backgroundColor = '#0073e6';
    header.style.color = 'white';
  });
}

function createCharts(rows) {
  const years = rows.map((row) => row.Y.v);
  const ctx = document.getElementById('ageChart').getContext('2d');
  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: years,
      datasets: [
        {
          label: 'Age in Years',
          data: years,
          backgroundColor: 'rgba(75, 192, 192, 0.2)',
          borderColor: 'rgba(75, 192, 192, 1)',
          borderWidth: 1,
        },
      ],
    },
    options: {
      scales: {
        y: {
          beginAtZero: true,
        },
      },
    },
  });
}

function generateExcel(data) {
  const newSheet = XLSX.utils.json_to_sheet(data, { skipHeader: false });
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Processed Data');

  XLSX.writeFile(newWorkbook, 'Processed_Age_Data.xlsx');
  alert('Excel file with calculated ages has been generated!');
}

function readExcelFile(state) {
  return (event) => {
    const file = event.target.files[0];

    if (!file) {
      alert('Please select an Excel file.');
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      state.set('workbook', XLSX.read(data, { type: 'array' }));
      console.log('Excel file loaded successfully.');
    };
    reader.readAsArrayBuffer(file);
  };
}

function processExcel(state) {
  return (event) => {
    const workbook = state.get('workbook');
    if (!workbook) {
      console.error('No file has been loaded yet.');
      alert('Please upload an Excel file first.');
      return;
    }

    const [firstSheetName] = workbook.SheetNames.filter(
      (sheetName) => !IGNORED_SHEETS.includes(sheetName)
    );
    const sheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { raw: true });

    if (jsonData.length === 0) {
      console.warn('The Excel file is empty.');
      alert('No data found in the Excel file.');
      return;
    }

    const result = jsonData.map(injectAge);
    generateExcel(result);
    populateTable(result);
    createCharts(result);
  };
}
