// Função para renderizar a tabela HTML com os dados processados
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

// Função de visualização antes da exportação
function previewData() {
    // Exibe a tabela renderizada
    if (workbook) {
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Processa e renderiza os dados
        const processedData = processData(jsonData);
        renderTableWithHeaders(processedData);
    }
}

// Carregar o arquivo Excel
uploadInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });

        // Mostrar visualização antes da exportação
        previewData();

        // Exibe log de erros, se houver
        if (errorLog.length > 0) {
            alert("Foram encontrados os seguintes erros:\n" + errorLog.join('\n'));
        }

        downloadExcelBtn.disabled = false;
        downloadCsvBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
});

// Exportar para Excel
downloadExcelBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.table_to_sheet(document.querySelector('table'));
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, sheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, 'editado.xlsx');
});

// Exportar para CSV
downloadCsvBtn.addEventListener('click', () => {
    const sheet = XLSX.utils.table_to_sheet(document.querySelector('table'));
    const csv = XLSX.utils.sheet_to_csv(sheet);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'editado.csv';
    link.click();
});
