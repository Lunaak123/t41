let excelData = []; // Placeholder for Excel data
let currentSheetName = ''; // Placeholder for the current sheet name

// Load the Excel file when the page loads
document.addEventListener('DOMContentLoaded', async () => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');

    if (fileUrl) {
        await loadExcelData(fileUrl);
    }
});

// Function to load Excel data
async function loadExcelData(url) {
    const response = await fetch(url);
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    displayData(excelData);
}

// Function to display data in the table
function displayData(data) {
    const sheetContent = document.getElementById('sheet-content');
    sheetContent.innerHTML = '';

    const table = document.createElement('table');
    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        row.forEach((cell, cellIndex) => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });
    sheetContent.appendChild(table);
}

// Function to apply the selected operation
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim().toUpperCase();
    const operationColumns = document.getElementById('operation-columns').value.trim().toUpperCase().split(',');
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;
    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value, 10) - 1;
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value, 10) - 1;

    if (!primaryColumn || !operationColumns.length) {
        alert('Please enter both primary and operation columns.');
        return;
    }

    const table = document.querySelector('table');
    if (!table) return;

    const rows = Array.from(table.querySelectorAll('tr')).slice(1); // Ignore header row
    rows.forEach((row, rowIndex) => {
        if (rowIndex >= rowRangeFrom && rowIndex <= rowRangeTo) {
            const primaryCell = row.cells[primaryColumn.charCodeAt(0) - 65];
            const shouldHighlight = checkOperation(rowIndex, primaryCell, operationColumns, operation, operationType);

            row.style.backgroundColor = shouldHighlight ? '#d1e7dd' : ''; // Apply highlight color
        } else {
            row.style.backgroundColor = ''; // Reset color
        }
    });
}

// Check operation for each row
function checkOperation(rowIndex, primaryCell, operationColumns, operation, operationType) {
    if (!primaryCell) return false;
    const primaryValue = primaryCell.textContent.trim();

    if (operationType === 'and') {
        return operationColumns.every(col => {
            const colCell = primaryCell.parentNode.cells[col.charCodeAt(0) - 65];
            const colValue = colCell ? colCell.textContent.trim() : '';
            return operation === 'null' ? !colValue : colValue !== '';
        });
    } else if (operationType === 'or') {
        return operationColumns.some(col => {
            const colCell = primaryCell.parentNode.cells[col.charCodeAt(0) - 65];
            const colValue = colCell ? colCell.textContent.trim() : '';
            return operation === 'null' ? !colValue : colValue !== '';
        });
    }
    return false;
}

// Download functionality
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value || 'downloaded_file';
    const format = document.getElementById('file-format').value;

    if (format === 'xlsx') {
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, currentSheetName);
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else if (format === 'csv') {
        const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(excelData));
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.setAttribute('href', url);
        link.setAttribute('download', `${filename}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    document.getElementById('download-modal').style.display = 'none';
});

// Close modal
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});
