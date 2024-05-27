// Función para formatear números con puntos cada 3 dígitos
function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

// Función para calcular los pagos del préstamo
function calculateLoan() {
    const fullName = document.getElementById("fullName").value;
    const idNumber = document.getElementById("idNumber").value;
    const phone = document.getElementById("phone").value;

    const loanAmount = parseFloat(document.getElementById("loanAmount").value);
    const initialPayment = parseFloat(document.getElementById("initialPayment").value) || 0;
    const finalPayment = parseFloat(document.getElementById("finalPayment").value) || 0;
    const capitalPayment = parseFloat(document.getElementById("capitalPayment").value) || 0;
    const loanPeriod = parseInt(document.getElementById("loanPeriod").value);
    const loanStartDate = document.getElementById("loanStartDate").value;

    const remainingAmount = loanAmount - initialPayment - finalPayment - capitalPayment;
    const monthlyPayment = remainingAmount / loanPeriod;
    const initialPaymentPercentage = (initialPayment / loanAmount) * 100 || 0;
    const totalLoan = loanAmount + (monthlyPayment * loanPeriod);

    // Actualizar el resumen de compra
    document.getElementById("summaryFullName").textContent = fullName;
    document.getElementById("summaryIdNumber").textContent = idNumber;
    document.getElementById("summaryMonthlyPayment").textContent = formatNumber(monthlyPayment.toFixed(2)) + ' COP';
    document.getElementById("summaryInstallments").textContent = loanPeriod;
    document.getElementById("summaryInitialPaymentPercentage").textContent = initialPaymentPercentage.toFixed(2) + '%';
    document.getElementById("summaryTotalLoan").textContent = formatNumber(totalLoan.toFixed(2)) + ' COP';

    // Actualizar la tabla de pagos
    const tableBody = document.getElementById("paymentTableBody");
    tableBody.innerHTML = '';
    let currentDate = new Date(loanStartDate);
    for (let i = 1; i <= loanPeriod; i++) {
        currentDate.setMonth(currentDate.getMonth() + 1);
        const newRow = tableBody.insertRow();
        newRow.insertCell(0).textContent = i;
        newRow.insertCell(1).textContent = formatNumber(monthlyPayment.toFixed(2)) + ' COP';
        newRow.insertCell(2).textContent = currentDate.toISOString().split('T')[0];
        const paymentCell = newRow.insertCell(3);
        const paymentCheckbox = document.createElement('input');
        paymentCheckbox.type = 'checkbox';
        paymentCheckbox.addEventListener('change', () => {
            if (paymentCheckbox.checked) {
                newRow.classList.add('paid');
            } else {
                newRow.classList.remove('paid');
            }
        });
        paymentCell.appendChild(paymentCheckbox);
    }
}

// Función para exportar la tabla a Excel
function exportToExcel() {
    const tableBody = document.getElementById("paymentTableBody");
    const rows = Array.from(tableBody.rows);
    const data = rows.map(row => {
        return [
            row.cells[0].textContent,
            row.cells[1].textContent,
            row.cells[2].textContent,
            row.cells[3].querySelector('input').checked ? 1 : 0
        ];
    });

    const fullName = document.getElementById("fullName").value;
    const idNumber = document.getElementById("idNumber").value;

    const worksheetData = [
        ["Nombre del Cliente", "Número de Identificación", "Número de Cuotas", "Valor de la Cuota", "Fecha Límite de Pago", "Estado"],
        ...data
    ];

    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Préstamos");
    XLSX.writeFile(wb, `${fullName}_${idNumber}.xlsx`);
}

// Función para importar datos desde Excel
function importFromExcel(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const [header, ...rows] = jsonData;

        if (header[0] !== "Nombre del Cliente" || header[1] !== "Número de Identificación") {
            alert("El formato del archivo no es correcto.");
            return;
        }

        const tableBody = document.getElementById("paymentTableBody");
        tableBody.innerHTML = '';

        rows.forEach(row => {
            const newRow = tableBody.insertRow();
            newRow.insertCell(0).textContent = row[0];
            newRow.insertCell(1).textContent = row[1];
            newRow.insertCell(2).textContent = row[2];
            const paymentCell = newRow.insertCell(3);
            const paymentCheckbox = document.createElement('input');
            paymentCheckbox.type = 'checkbox';
            paymentCheckbox.checked = row[3] === 1;
            if (paymentCheckbox.checked) {
                newRow.classList.add('paid');
            }
            paymentCheckbox.addEventListener('change', () => {
                if (paymentCheckbox.checked) {
                    newRow.classList.add('paid');
                } else {
                    newRow.classList.remove('paid');
                }
            });
            paymentCell.appendChild(paymentCheckbox);
        });
    };
    reader.readAsArrayBuffer(file);
}
