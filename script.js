let rowCount = 1;

function addRow() {
    const table = document.getElementById("expenseTable").getElementsByTagName('tbody')[0];
    const newRow = table.insertRow();
    newRow.innerHTML = `
        <td>${rowCount++}</td>
        <td><input type="date" name="date"></td>
        <td><input type="text" name="areaOfWork"></td>
        <td>
            <select name="modeOfTravel">
                <option value="Bike">Bike</option>
                <option value="Bus">Bus</option>
                <option value="Auto">Auto</option>
                <option value="Taxi">Taxi</option>
                <option value="Train">Train</option>
            </select>
        </td>
        <td><input type="number" name="kmTravelled" oninput="calculateRow(this)"></td>
        <td><input type="number" name="travelBill" readonly></td>
        <td><input type="number" name="allowance" value="100"></td>
        <td><input type="number" name="petrolBill" readonly></td>
        <td><input type="number" name="hotelBill"></td>
        <td><input type="number" name="otherExpenses"></td>
        <td><input type="number" name="totalExpense" readonly></td>
        <td><input type="text" name="remark"></td>
    `;
}

function calculateRow(input) {
    const row = input.closest('tr');
    const kmTravelled = parseFloat(row.querySelector('input[name="kmTravelled"]').value) || 0;
    const travelBill = kmTravelled * 2;  // Example calculation, adjust as needed
    const petrolBill = kmTravelled * 105.88 / 45;  // Fuel price / mileage
    const allowance = parseFloat(row.querySelector('input[name="allowance"]').value) || 0;
    const hotelBill = parseFloat(row.querySelector('input[name="hotelBill"]').value) || 0;
    const otherExpenses = parseFloat(row.querySelector('input[name="otherExpenses"]').value) || 0;

    row.querySelector('input[name="travelBill"]').value = travelBill.toFixed(2);
    row.querySelector('input[name="petrolBill"]').value = petrolBill.toFixed(2);
    row.querySelector('input[name="totalExpense"]').value = (travelBill + petrolBill + hotelBill + otherExpenses + allowance).toFixed(2);
}

function exportToExcel() {
    const table = document.getElementById("expenseTable");
    const mergedRows = {};

    // Collect data and merge rows with the same date
    for (let i = 1, row; row = table.rows[i]; i++) {
        const dateInput = row.querySelector('input[type="date"]');
        if (!dateInput) continue;

        const date = dateInput.value;
        const areaOfWork = row.querySelector('input[name^="areaOfWork"]').value || '';
        const modeOfTravel = row.querySelector('select[name^="modeOfTravel"]').value || '';
        const kmTravelled = parseFloat(row.querySelector('input[name^="kmTravelled"]').value) || 0;
        const travelBill = parseFloat(row.querySelector('input[name^="travelBill"]').value) || 0;
        const allowance = parseFloat(row.querySelector('input[name^="allowance"]').value) || 0;
        const petrolBill = parseFloat(row.querySelector('input[name^="petrolBill"]').value) || 0;
        const hotelBill = parseFloat(row.querySelector('input[name^="hotelBill"]').value) || 0;
        const otherExpenses = parseFloat(row.querySelector('input[name^="otherExpenses"]').value) || 0;
        const totalExpense = parseFloat(row.querySelector('input[name^="totalExpense"]').value) || 0;
        const remark = row.querySelector('input[name^="remark"]').value || '';

        const data = {
            date,
            areaOfWork,
            modeOfTravel,
            kmTravelled,
            travelBill,
            allowance,
            petrolBill,
            hotelBill,
            otherExpenses,
            totalExpense,
            remark
        };

        if (mergedRows[date]) {
            mergedRows[date].kmTravelled += data.kmTravelled;
            mergedRows[date].travelBill += data.travelBill;
            mergedRows[date].allowance = Math.max(mergedRows[date].allowance, data.allowance);
            mergedRows[date].petrolBill += data.petrolBill;
            mergedRows[date].hotelBill += data.hotelBill;
            mergedRows[date].otherExpenses += data.otherExpenses;
            mergedRows[date].totalExpense += data.totalExpense;
            mergedRows[date].remark += ' | ' + data.remark;
        } else {
            mergedRows[date] = data;
        }
    }

    const worksheetData = [
        ["Date", "Area of Work", "Mode of Travel", "Kilometers Traveled", "Travel Bill", "Allowance", "Petrol Bill", "Hotel Bill", "Other Expenses", "Total Expense", "Remark"]
    ];

    for (const date in mergedRows) {
        const row = mergedRows[date];
        worksheetData.push([
            row.date,
            row.areaOfWork,
            row.modeOfTravel,
            row.kmTravelled.toFixed(2),
            row.travelBill.toFixed(2),
            row.allowance.toFixed(2),
            row.petrolBill.toFixed(2),
            row.hotelBill.toFixed(2),
            row.otherExpenses.toFixed(2),
            row.totalExpense.toFixed(2),
            row.remark
        ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(worksheetData);

    // Apply formatting: red borders and auto width for columns
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell_ref = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cell_ref]) continue;

            ws[cell_ref].s = {
                border: {
                    top: { style: "thin", color: { rgb: "FF0000" } },
                    bottom: { style: "thin", color: { rgb: "FF0000" } },
                    left: { style: "thin", color: { rgb: "FF0000" } },
                    right: { style: "thin", color: { rgb: "FF0000" } }
                }
            };
        }
    }

    // Auto-width for columns
    const wscols = worksheetData[0].map((col, i) => {
        return { wpx: 150 };
    });
    ws['!cols'] = wscols;

    // Create a workbook and export to Excel
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Expenses");
    XLSX.writeFile(wb, "Monthly_Expense_Report.xlsx");
}
