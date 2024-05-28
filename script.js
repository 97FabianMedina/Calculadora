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

    // Calcular la suma de las cuotas pagadas
    let paidInstallmentsSum = 0;
    const tableBody = document.getElementById("paymentTableBody");
    const rows = Array.from(tableBody.rows);
    rows.forEach(row => {
        const checkbox = row.cells[3].querySelector('input');
        if (checkbox && checkbox.checked) {
            paidInstallmentsSum += parseFloat(row.cells[1].textContent.replace(/[^0-9.-]+/g,""));
        }
    });

    const remainingAmount = loanAmount - initialPayment - finalPayment - capitalPayment - paidInstallmentsSum;
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
        paymentCheckbox.addEventListener('change', function() {
            // Recalcular al cambiar el estado del pago
            if (this.checked) {
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

    const tableBody = document.getElementById("paymentTableBody");
    const rows = Array.from(tableBody.rows);

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

        // Asignar los datos importados a los campos correspondientes
        document.getElementById("fullName").value = jsonData[0][1];
        document.getElementById("idNumber").value = jsonData[1][1];
        document.getElementById("summaryMonthlyPayment").textContent = jsonData[2][1];
        document.getElementById("summaryInstallments").textContent = jsonData[3][1];
        document.getElementById("summaryInitialPaymentPercentage").textContent = jsonData[4][1];
        document.getElementById("summaryTotalLoan").textContent = jsonData[5][1];
        document.getElementById("loanAmount").value = parseFloat(jsonData[5][1].replace(/[^0-9.-]+/g,""));
        document.getElementById("initialPayment").value = parseFloat(jsonData[6][1].replace(/[^0-9.-]+/g,""));
        document.getElementById("finalPayment").value = parseFloat(jsonData[7][1].replace(/[^0-9.-]+/g,""));
        document.getElementById("capitalPayment").value = parseFloat(jsonData[8][1].replace(/[^0-9.-]+/g,""));

        const tableBody = document.getElementById("paymentTableBody");
        const existingRowsCount = tableBody.rows.length - 1;

        jsonData.slice(10).forEach((row, index) => {
            const newRow = tableBody.insertRow(existingRowsCount + index);
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
            paymentCheckbox.addEventListener('change', function() {
                // Recalcular al cambiar el estado del pago
                if (this.checked) {
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

// Función para agregar saldo al capital
function addCapital() {
    const capitalInput = parseFloat(document.getElementById("capitalInput").value);
    const tableBody = document.getElementById("paymentTableBody");
    const rows = Array.from(tableBody.rows);
    let paidInstallmentsSum = 0;
    rows.forEach(row => {
        const checkbox = row.cells[3].querySelector('input');
        if (checkbox && checkbox.checked) {
            paidInstallmentsSum += parseFloat(row.cells[1].textContent.replace(/[^0-9.-]+/g,""));
        }
    });
    const remainingAmount = loanAmount - initialPayment - finalPayment - paidInstallmentsSum;
    const newRemainingAmount = remainingAmount + capitalInput;
    const newMonthlyPayment = newRemainingAmount / loanPeriod;
    const currentDate = new Date(document.getElementById("loanStartDate").value);
    for (let i = 1; i <= loanPeriod; i++) {
        currentDate.setMonth(currentDate.getMonth() + 1);
        const row = tableBody.rows[i - 1];
        const monthlyPaymentCell = row.cells[1];
        monthlyPaymentCell.textContent = formatNumber(newMonthlyPayment.toFixed(2)) + ' COP';
    }
}
// Función para limpiar todos los campos del formulario y la tabla de pagos
function refreshForm() {
    // Limpiar campos del formulario
    document.getElementById("fullName").value = "";
    document.getElementById("idNumber").value = "";
    document.getElementById("phone").value = "";
    document.getElementById("loanAmount").value = "";
    document.getElementById("initialPayment").value = "";
    document.getElementById("finalPayment").value = "";
    document.getElementById("capitalPayment").value = "";
    document.getElementById("loanPeriod").value = "";
    document.getElementById("loanStartDate").value = "";

    // Limpiar resumen de compra
    document.getElementById("summaryFullName").textContent = "";
    document.getElementById("summaryIdNumber").textContent = "";
    document.getElementById("summaryMonthlyPayment").textContent = "";
    document.getElementById("summaryInstallments").textContent = "";
    document.getElementById("summaryInitialPaymentPercentage").textContent = "";
    document.getElementById("summaryTotalLoan").textContent = "";

    // Limpiar tabla de pagos
    const tableBody = document.getElementById("paymentTableBody");
    tableBody.innerHTML = "";
}

// Asignar evento al botón de refrescar
document.getElementById("refreshButton").addEventListener("click", refreshForm);


// Asignar eventos a los botones
document.getElementById("importButton").addEventListener("change", importFromExcel);
document.getElementById("exportButton").addEventListener("click", exportToExcel);
document.getElementById("addCapitalButton").addEventListener("click", addCapital);
