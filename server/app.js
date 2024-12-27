const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.static('public'));

app.post('/submit', async (req, res) => {
    const { name, age } = req.body;
    const filePath = './data.xlsx';

    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
    } else {
        const sheet = workbook.addWorksheet('Sheet1');
        sheet.columns = [
            { header: 'Name', key: 'name' },
            { header: 'Age', key: 'age' },
        ];
    }

    const sheet = workbook.getWorksheet('Sheet1');
    sheet.addRow({ name, age });

    await workbook.xlsx.writeFile(filePath);
    res.send('Data added successfully!');
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
