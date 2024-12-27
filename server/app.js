const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
const app = express();
const PORT = 3000;

// Middleware
app.use(express.json());
app.use(express.static('public'));

// Route to add data to the Excel file
app.post('/submit', async (req, res) => {
    const { name, age } = req.body;
    console.log('Received Data:', { name, age });

    if (!name || !age) {
        return res.status(400).send('Name and age are required.');
    }

    const filePath = './data.xlsx';

    try {
        const workbook = new ExcelJS.Workbook();
        let sheet;

        if (fs.existsSync(filePath)) {
            console.log('File exists. Loading...');
            await workbook.xlsx.readFile(filePath);
            sheet = workbook.getWorksheet('Sheet1');

            // Check if sheet exists
            if (!sheet) {
                console.log('Sheet1 not found. Creating a new sheet...');
                sheet = workbook.addWorksheet('Sheet1');
                sheet.columns = [
                    { header: 'Name', key: 'name' },
                    { header: 'Age', key: 'age' },
                ];
            }
        } else {
            console.log('File does not exist. Creating...');
            sheet = workbook.addWorksheet('Sheet1');
            sheet.columns = [
                { header: 'Name', key: 'name' },
                { header: 'Age', key: 'age' },
            ];
        }

        // Append the new row
        console.log('Adding row:', { name, age });
        sheet.addRow({ name, age });

        // Save the updated workbook
        await workbook.xlsx.writeFile(filePath);
        console.log('Data successfully written to Excel file.');
        res.send('Data added successfully!');
    } catch (error) {
        console.error('Error during file operation:', error);
        res.status(500).send('Failed to update Excel file.');
    }
});




// Route to download the Excel file
app.get('/download', (req, res) => {
    const filePath = './data.xlsx';
    if (fs.existsSync(filePath)) {
        res.download(filePath, 'data.xlsx');
    } else {
        res.status(404).send('File not found.');
    }
});

// Route to fetch Excel data as JSON
app.get('/data', async (req, res) => {
    const filePath = './data.xlsx';

    try {
        const workbook = new ExcelJS.Workbook();
        if (!fs.existsSync(filePath)) {
            console.log('Excel file not found.');
            return res.status(404).send('Excel file not found.');
        }

        console.log('Reading Excel file...');
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.getWorksheet('Sheet1');
        const data = [];

        sheet.eachRow((row, rowIndex) => {
            if (rowIndex === 1) return; // Skip the header row
            const [_, name, age] = row.values; // Skip the first empty item
            data.push({ name, age });
        });

        console.log('Data from Excel:', data);
        res.json(data);
    } catch (error) {
        console.error('Error reading Excel file:', error);
        res.status(500).send('Failed to read Excel file.');
    }
});



// Start the server
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
