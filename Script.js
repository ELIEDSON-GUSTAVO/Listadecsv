const uploadInput = document.getElementById('upload');
const downloadExcelBtn = document.getElementById('download-excel');
const downloadCsvBtn = document.getElementById('download-csv');
const tableContainer = document.getElementById('table-container');

let workbook; // Variável para armazenar o arquivo Excel
let errorLog = []; // Armazena erros encontrados

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

        // Processa os dados
        const processedData = processData(jsonData);
        renderTableWithHeaders(processedData);

        // Exibe log de erros, se houver
        if (errorLog.length > 0) {
            alert("Foram encontrados os seguintes erros:\n" + errorLog.join('\n'));
        }

        downloadExcelBtn.disabled = false;
        downloadCsvBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
});

function processData(data) {
    const headers = ['ITENS', 'CODIGO', 'QND', 'DESCRIÇÃO', 'MASS', 'MATERIAL', 'LINK'];
    const headerIndexMap = {};
    errorLog = []; // Limpa erros anteriores

    data[0].forEach((header, index) => {
        if (headers.includes(header)) {
            headerIndexMap[header] = index;
        }
    });

    headers.forEach((header) => {
        if (!(header in headerIndexMap)) {
            headerIndexMap[header] = -1; // Coluna ausente
        }
    });

    const processedData = [headers];

    data.slice(1).forEach((row, rowIndex) => {
        const newRow = headers.map((header) => {
            const colIndex = headerIndexMap[header];
            let value = colIndex !== -1 ? row[colIndex] || '' : '';

            // Validações específicas por coluna
            if (header === 'CODIGO') {
                if (!/^\d{2}\.\d{2}\.\d{2}\.\d{10}$/.test(value)) {
                    errorLog.push(`Erro no código na linha ${rowIndex + 2}: "${value}" não está no formato correto.`);
                }
            } else if (header === 'DESCRIÇÃO' && value.trim() === '') {
                errorLog.push(`Descrição ausente na linha ${rowIndex + 2}.`);
            } else if (header === 'MASS') {
                value = '0,1'; // Define "MASS" como 0,1
            } else if (['MATERIAL', 'LINK'].includes(header)) {
                value = ''; // Garante que essas colunas fiquem vazias
            }

            return value;
        });

        processedData.push(newRow);
    });

    if (headerIndexMap['ITENS'] === -1) {
        processedData.forEach((row, index) => {
            if (index === 0) {
                row[0] = 'ITENS';
            } else {
                row[0] = index;
            }
        });
    }

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
    data.slice(1).forEach((row) => {
        tableHtml += '<tr>';
        row.forEach((cell) => {
            tableHtml += `<td>${cell}</td>`;
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
    XLSX.writeFile(newWorkbook, 'editado.xlsx');
});

downloadCsvBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.table_to_sheet(document.querySelector('table'));
    const csv = XLSX.utils.sheet_to_csv(sheet);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'editado.csv';
    link.click();
});
