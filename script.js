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
    const totalLoan = loanAmount;

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
    const fullName = document.getElementById("fullName").value;
    const idNumber = document.getElementById("idNumber").value;
    const summaryMonthlyPayment = document.getElementById("summaryMonthlyPayment").textContent;
    const summaryInstallments = document.getElementById("summaryInstallments").textContent;
    const summaryInitialPaymentPercentage = document.getElementById("summaryInitialPaymentPercentage").textContent;
    const summaryTotalLoan = document.getElementById("summaryTotalLoan").textContent;

    const loanAmount = parseFloat(document.getElementById("loanAmount").value);
    const initialPayment = parseFloat(document.getElementById("initialPayment").value) || 0;
    const finalPayment = parseFloat(document.getElementById("finalPayment").value) || 0;
    const capitalPayment = parseFloat(document.getElementById("capitalPayment").value) || 0;

    const worksheetData = [
        ["Nombre del Cliente", fullName],
        ["Número de Identificación", idNumber],
        ["Pago Mensual", summaryMonthlyPayment],
        ["Número de Cuotas", summaryInstallments],
        ["Porcentaje de Cuota Inicial", summaryInitialPaymentPercentage],
        ["Valor Total del Préstamo", summaryTotalLoan],
        ["Cuota Inicial", formatNumber(initialPayment.toFixed(2)) + ' COP'],
        ["Cuota Final", formatNumber(finalPayment.toFixed(2)) + ' COP'],
        ["Valor a Capital", formatNumber(capitalPayment.toFixed(2)) + ' COP'],
        [],
        ["Número de Cuota", "Valor de la Cuota", "Fecha Límite de Pago", "Estado"]
    ];

    rows.forEach(row => {
        worksheetData.push([
            row.cells[0].textContent,
            row.cells[1].textContent,
            row.cells[2].textContent,
            row.cells[3].querySelector('input').checked ? 1 : 0
        ]);
    });

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

        const fullName = rows[0][1];
        const idNumber = rows[1][1];
        const monthlyPayment = rows[2][1];
        const installments = rows[3][1];
        const initialPaymentPercentage = rows[4][1];
        const totalLoan = rows[5][1];
        const initialPayment = parseFloat(rows[6][1].replace(/[^0-9.-]+/g,""));
        const finalPayment = parseFloat(rows[7][1].replace(/[^0-9.-]+/g,""));
        const capitalPayment = parseFloat(rows[8][1].replace(/[^0-9.-]+/g,""));

        document.getElementById("fullName").value = fullName;
        document.getElementById("idNumber").value = idNumber;

        document.getElementById("summaryFullName").textContent = fullName;
        document.getElementById("summaryIdNumber").textContent = idNumber;
        document.getElementById("summaryMonthlyPayment").textContent = monthlyPayment;
        document.getElementById("summaryInstallments").textContent = installments;
        document.getElementById("summaryInitialPaymentPercentage").textContent = initialPaymentPercentage;
        document.getElementById("summaryTotalLoan").textContent = totalLoan;

        document.getElementById("initialPayment").value = initialPayment;
        document.getElementById("finalPayment").value = finalPayment;
        document.getElementById("capitalPayment").value = capitalPayment;

        const tableBody = document.getElementById("paymentTableBody");
        tableBody.innerHTML = '';

        rows.slice(10).forEach(row => {
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
