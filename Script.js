// Seleção de elementos no DOM
const fileInput = document.getElementById('fileInput');
const uploadInput = document.getElementById('upload');
const tableContainer = document.getElementById('table-container');
const downloadExcelButton = document.getElementById('download-excel');
const downloadCsvButton = document.getElementById('download-csv');
const importTableButton = document.getElementById('import-table');

let workbookData = null;

// Função para carregar e exibir o arquivo Excel
fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Carrega a primeira planilha
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Converte os dados da planilha para JSON
            workbookData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            displayTable(workbookData);

            // Ativa os botões de download e importar
            downloadExcelButton.disabled = false;
            downloadCsvButton.disabled = false;
            importTableButton.disabled = false;
        };
        reader.readAsArrayBuffer(file);
    }
});

// Função para exibir a tabela no DOM
function displayTable(data) {
    tableContainer.innerHTML = ''; // Limpa o container

    const table = document.createElement('table');
    table.classList.add('table');

    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        row.forEach((cell) => {
            const cellElement = rowIndex === 0 ? document.createElement('th') : document.createElement('td');
            cellElement.textContent = cell || ''; // Evita células vazias
            tr.appendChild(cellElement);
        });
        table.appendChild(tr);
    });

    tableContainer.appendChild(table);
}

// Função para baixar o arquivo no formato Excel
downloadExcelButton.addEventListener('click', () => {
    if (workbookData) {
        const ws = XLSX.utils.aoa_to_sheet(workbookData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        XLSX.writeFile(wb, 'planilha_editada.xlsx');
    }
});

// Função para baixar o arquivo no formato CSV
downloadCsvButton.addEventListener('click', () => {
    if (workbookData) {
        const ws = XLSX.utils.aoa_to_sheet(workbookData);
        const csv = XLSX.utils.sheet_to_csv(ws);

        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = 'planilha_editada.csv';
        a.click();

        URL.revokeObjectURL(url);
    }
});

// Função para importar a tabela de um novo arquivo Excel
importTableButton.addEventListener('click', () => {
    uploadInput.click();
});

// Função para lidar com o upload do novo arquivo Excel
uploadInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Carrega a primeira planilha
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Converte os dados da planilha para JSON
            workbookData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            displayTable(workbookData);
        };
        reader.readAsArrayBuffer(file);
    }
});
