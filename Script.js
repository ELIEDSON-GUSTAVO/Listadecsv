const uploadInput = document.getElementById('upload');
const downloadExcelBtn = document.getElementById('download-excel');
const downloadCsvBtn = document.getElementById('download-csv');
const importTableBtn = document.getElementById('import-table');
const tableContainer = document.getElementById('table-container');

let workbook;

uploadInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        renderTable(jsonData);

        downloadExcelBtn.disabled = false;
        downloadCsvBtn.disabled = false;
        importTableBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
});

function renderTable(data) {
    let tableHtml = '<table>';
    tableHtml += '<thead><tr>';
    data[0].forEach((header) => {
        tableHtml += `<th>${header}</th>`;
    });
    tableHtml += '</tr></thead>';
    tableHtml += '<tbody>';
    data.slice(1).forEach((row) => {
        tableHtml += '<tr>';
        row.forEach((cell) => {
            tableHtml += `<td>${cell || ''}</td>`;
        });
        tableHtml += '</tr>';
    });
    tableHtml += '</tbody></table>';
    tableContainer.innerHTML = tableHtml;
}

downloadExcelBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.table_to_sheet(document.querySelector('table'));
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, sheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, 'planilha-editada.xlsx');
});

downloadCsvBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.table_to_sheet(document.querySelector('table'));
    const csv = XLSX.utils.sheet_to_csv(sheet);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'planilha-editada.csv';
    link.click();
});

importTableBtn.addEventListener('click', () => {
    alert('Função de importação de tabela será implementada em breve.');
});