function loadFile() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];
    
    if (file) {
        const reader = new FileReader();
        reader.onload = function(event) {
            const arrayBuffer = event.target.result;
            mammoth.extractRawText({arrayBuffer: arrayBuffer})
                .then(function(result) {
                    document.getElementById("inputText").value = result.value; // текст из файла
                })
                .catch(function(err) {
                    console.log(err);
                });
        };
        reader.readAsArrayBuffer(file);
    }
}

function convertToTable() {
    const inputText = document.getElementById("inputText").value.trim();
    const isMarkdown = inputText.startsWith('|');
    const rows = isMarkdown ? inputText.split('\n') : inputText.split('\n').filter(row => row.trim() !== '');
    const tableBody = document.querySelector("#countryTable tbody");
    const headersRow = document.getElementById("tableHeaders");
    tableBody.innerHTML = ""; // Очистка предыдущих данных
    headersRow.innerHTML = ""; // Очистка заголовков

    if (rows.length > 0) {
        if (isMarkdown) {
            const headers = rows[0].split('|').map(col => col.trim());
            headers.forEach(header => {
                const th = document.createElement("th");
                th.textContent = header;
                headersRow.appendChild(th);
            });

            for (let i = 2; i < rows.length; i++) {
                const columns = rows[i].split('|').map(col => col.trim());
                const newRow = document.createElement("tr");
                columns.forEach(col => {
                    const newCell = document.createElement("td");
                    newCell.textContent = col;
                    newRow.appendChild(newCell);
                });
                tableBody.appendChild(newRow);
            }
        } else {
            const headers = rows[0].split('|').map(col => col.trim());
            headers.forEach(header => {
                const th = document.createElement("th");
                th.textContent = header;
                headersRow.appendChild(th);
            });

            for (let i = 1; i < rows.length; i++) {
                const columns = rows[i].split('|').map(col => col.trim());
                const newRow = document.createElement("tr");
                columns.forEach(col => {
                    const newCell = document.createElement("td");
                    newCell.textContent = col;
                    newRow.appendChild(newCell);
                });
                tableBody.appendChild(newRow);
            }
        }
    }
}

function exportToExcel() {
    const table = document.getElementById("countryTable");
    const wb = XLSX.utils.table_to_book(table);
    XLSX.writeFile(wb, 'countries.xlsx');
}

function exportToWord() {
    try {
        const table = document.getElementById("countryTable").outerHTML;
        const blob = new Blob([table], {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "countries.docx";
        a.click();
        URL.revokeObjectURL(url);
    } catch (error) {
        alert("Error Word:", error);

    }
}

