document.getElementById('dataForm').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = {
        name: document.getElementById('name').value,
        age: document.getElementById('age').value,
    };
    console.log('Submitting Data:', formData);

    const response = await fetch('/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData),
    });
    alert(await response.text());
});

const ExcelJS = require('exceljs');
const fs = require('fs');

(async () => {
    const filePath = './data.xlsx';

    const workbook = new ExcelJS.Workbook();
    let sheet;

    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
        sheet = workbook.getWorksheet('Sheet1');
        console.log('File exists, loaded sheet.');
    } else {
        sheet = workbook.addWorksheet('Sheet1');
        sheet.columns = [
            { header: 'Name', key: 'name' },
            { header: 'Age', key: 'age' },
        ];
        console.log('File does not exist, created new sheet.');
    }

    sheet.addRow({ name: 'Test User', age: '25' });

    await workbook.xlsx.writeFile(filePath);
    console.log('Row added and file updated.');
})();
