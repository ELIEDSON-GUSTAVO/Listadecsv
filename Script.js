function mapHeader(header) {
    const headerMap = {
        itens: ['itens'],
        codigo: ['codigo', 'código'],
        quantidade: ['qt', 'quantidade'],
        descricao: ['descrição', 'descricao'],
        un_medida: ['un. medida', 'unidade', 'unidade de medida'],
        local: ['local']
    };

    const normalizedHeader = header.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    for (const [key, aliases] of Object.entries(headerMap)) {
        if (aliases.includes(normalizedHeader)) {
            return key;
        }
    }
    return null; // Cabeçalho não identificado
}

function processData(data) {
    const headers = ['ITENS', 'CODIGO', 'QT', 'UN. MEDIDA', 'DESCRIÇÃO', 'LOCAL'];
    const headerIndexMap = {};
    let errorLog = []; // Limpa erros anteriores

    // Mapeia os índices das colunas do arquivo original
    data[0].forEach((header, index) => {
        const mappedHeader = mapHeader(header);
        if (mappedHeader) {
            headerIndexMap[mappedHeader] = index;
        } else {
            errorLog.push(`Cabeçalho desconhecido: "${header}"`);
        }
    });

    // Garante que todas as colunas estejam representadas
    headers.forEach((header) => {
        if (!(header.toLowerCase() in headerIndexMap)) {
            headerIndexMap[header.toLowerCase()] = -1; // Coluna ausente
        }
    });

    const processedData = [headers];

    data.slice(1).forEach((row, rowIndex) => {
        const newRow = headers.map((header) => {
            const colIndex = headerIndexMap[header.toLowerCase()];
            return colIndex !== -1 ? row[colIndex] || '' : '';
        });

        // Validações específicas para as colunas
        if (!/^\d{2}\.\d{2}\.\d{2}\.\d{10}$/.test(newRow[1])) {
            errorLog.push(`Erro no código na linha ${rowIndex + 2}: "${newRow[1]}" não está no formato correto.`);
        }

        processedData.push(newRow);
    });

    // Preenche a coluna "ITENS" com ordem crescente
    processedData.forEach((row, index) => {
        if (index === 0) {
            row[0] = 'ITENS';
        } else {
            row[0] = index; // Preenche com valores crescentes
        }
    });

    // Exibe erros no console, se houver
    if (errorLog.length > 0) {
        console.error('Erros encontrados durante o processamento:');
        errorLog.forEach((error) => console.error(error));
    }

    return processedData;
}

function exportToCSV(data, filename) {
    const csvContent = data.map((row) => row.map((cell) => `"${cell}"`).join(';')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const fileContent = e.target.result;
        const rows = fileContent.split('\n').map((row) => row.split(';'));
        const processedData = processData(rows);
        exportToCSV(processedData, 'dados_processados.csv');
    };
    reader.readAsText(file);
});
