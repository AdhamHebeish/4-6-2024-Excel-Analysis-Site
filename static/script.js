const fileUpload = document.getElementById('fileUpload');
const submitButton = document.getElementById('submitButton');
const dataTable = document.getElementById('data-table');

submitButton.addEventListener('click', async () => {
    const file = fileUpload.files[0];

    if (!file) {
        alert('Please select a file to upload.');
        return;
    }

    try {
        const data = await processExcelFile(file);
        if (data.error) {
            alert(data.error);
            return;
        }

        populateTable(data);
    } catch (error) {
        console.error(error);
        alert('An error occurred during upload.');
    }
});

async function processExcelFile(file) {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('/analyze_excel', {
        method: 'POST',
        body: formData
    });

    const data = await response.json();
    if (data.error) {
        alert(data.error);
        return;
    }

    // Update the table on successful processing
    populateTable(data);
}

function populateTable(data) {
    const tableBody = dataTable.getElementsByTagName('tbody')[0];

    // Clear existing rows
    tableBody.innerHTML = '';

    // Create table header row (adjust as needed)
    const headerRow = tableBody.insertRow();
    for (const header of Object.keys(data[0])) {
        const headerCell = headerRow.insertCell();
        headerCell.textContent = header;
    }

    // Create data rows in original order
    for (const rowIndex in data) {
        const rowData = data[rowIndex];
        const dataRow = tableBody.insertRow();
        for (const value of Object.values(rowData)) {
            const dataCell = dataRow.insertCell();
            dataCell.textContent = value;
        }
    }
}
