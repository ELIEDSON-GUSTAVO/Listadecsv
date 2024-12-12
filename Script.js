const uploadInput = document.getElementById('upload');
const downloadExcelBtn = document.getElementById('download-excel');
const downloadCsvBtn = document.getElementById('download-csv');
const tableContainer = document.getElementById('table-container');

let workbook; // Variável para armazenar o arquivo Excel
let jsonData = []; // Armazena os dados processados

uploadInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Processa os dados
        const processedData = processData(jsonData);
        renderTableWithHeaders(processedData);

        // Habilita os botões
        downloadExcelBtn.disabled = false;
        downloadCsvBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
});

function processData(data) {
    const headers = ['ITENS', 'CODIGO', 'QND', 'DESCRIÇÃO', 'MASS', 'MATERIAL', 'LINK'];
    const processedData = [headers];

    data.slice(1).forEach((row, rowIndex) => {
        const newRow = headers.map((header, colIndex) => {
            let value = row[colIndex] || '';

            // Validações específicas por coluna
            if (header === 'CODIGO' && !/^\d{2}\.\d{2}\.\d{2}\.\d{10}$/.test(value)) {
                alert(`Erro no código na linha ${rowIndex + 2}: "${value}" não está no formato correto.`);
            }
            if (header === 'DESCRIÇÃO' && value.trim() === '') {
                alert(`Descrição ausente na linha ${rowIndex + 2}.`);
            }
            if (header === 'MASS') {
                value = '0,1'; // Define "MASS" como 0,1
            }
            if (['MATERIAL', 'LINK'].includes(header)) {
                value = ''; // Garante que essas colunas fiquem vazias
            }

            return value;
        });

        processedData.push(newRow);
    });

    return processedData;
}

function renderTableWithHeaders(data) {
    let tableHtml = '<table>';
    tableHtml += '<thead><tr>';
    data[0].forEach((header) => {
        tableHtml += `<th>${header}</th>`;
    });
    tableHtml += '</tr></thead>';
    tableHtml += '<tbody>';
    data.slice(1).forEach((row, rowIndex) => {
        tableHtml += '<tr>';
        row.forEach((cell, colIndex) => {
            tableHtml += `<td class="editable" data-row="${rowIndex}" data-col="${colIndex}" contenteditable="true">${cell}</td>`;
        });
        tableHtml += '</tr>';
    });
    tableHtml += '</tbody></table>';
    tableContainer.innerHTML = tableHtml;

    addEditListeners();
}

function addEditListeners() {
    const editableCells = document.querySelectorAll('td.editable');
    editableCells.forEach(cell => {
        cell.addEventListener('input', (e) => {
            const row = e.target.getAttribute('data-row');
            const col = e.target.getAttribute('data-col');
            jsonData[parseInt(row) + 1][parseInt(col)] = e.target.textContent; // Atualiza jsonData
        });
    });
}

downloadExcelBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.aoa_to_sheet(jsonData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, sheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, 'editado.xlsx');
});

downloadCsvBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.aoa_to_sheet(jsonData);
    const csv = XLSX.utils.sheet_to_csv(sheet);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'editado.csv';
    link.click();
});
