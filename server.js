const express = require('express');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json());
app.use(fileUpload());

function extractAmount(row) {
    const key = Object.keys(row).find(k =>
        k.toLowerCase().includes('amount') || k.toLowerCase().includes('deposit')
    );
    if (!key || !row[key]) return '0.00';
    return parseFloat(row[key].toString().replace(/,/g, '')).toFixed(2);
}

function extractUTR(row) {
    const key = Object.keys(row).find(k =>
        k.toLowerCase().includes('utr') || k.toLowerCase().includes('narration') || k.toLowerCase().includes('description')
    );
    if (!key || !row[key]) return null;
    const match = row[key].toString().match(/\b[A-Z0-9]{10,}\b/);
    return match ? match[0] : null;
}

function compareStatements(inputPath, bankPath) {
    const input = xlsx.utils.sheet_to_json(xlsx.readFile(inputPath).Sheets['Sheet1']);
    const bank = xlsx.utils.sheet_to_json(xlsx.readFile(bankPath).Sheets['Sheet1']);

    const inputMap = new Map();
    const bankMap = new Map();

    input.forEach(row => {
        const utr = row.Utr || extractUTR(row);
        if (utr) inputMap.set(utr, row);
    });

    bank.forEach(row => {
        const utr = extractUTR(row);
        if (utr) bankMap.set(utr, row);
    });

    const result = [];

    inputMap.forEach((inputRow, utr) => {
        const bankRow = bankMap.get(utr);
        const userId = inputRow['User Id'] || '';
        const updatedAmount = parseFloat(inputRow['Updated Amount'] || 0).toFixed(2);

        if (!bankRow) {
            result.push({ 'User Id': userId, 'UTR': "'" + utr, 'Status': 'Missing in Bank', 'Amount': '', 'Mismatched Amount': updatedAmount });
        } else {
            const bankAmount = extractAmount(bankRow);
            if (updatedAmount !== bankAmount) {
                result.push({ 'User Id': userId, 'UTR': "'" + utr, 'Status': 'Amount Mismatch', 'Amount': bankAmount, 'Mismatched Amount': updatedAmount });
            }
        }
    });

    bankMap.forEach((bankRow, utr) => {
        if (!inputMap.has(utr)) {
            const bankAmount = extractAmount(bankRow);
            result.push({ 'User Id': '', 'UTR': "'" + utr, 'Status': 'Excess in Bank', 'Amount': bankAmount, 'Mismatched Amount': '' });
        }
    });

    const outputSheet = xlsx.utils.json_to_sheet(result);
    const outputWb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(outputWb, outputSheet, 'Comparison');

    const outputPath = path.join(__dirname, 'comparison_output.xlsx');
    xlsx.writeFile(outputWb, outputPath);
    return outputPath;
}

app.post('/upload', (req, res) => {
    if (!req.files || !req.files.input || !req.files.bank) {
        return res.status(400).send('❌ Files missing');
    }

    const inputPath = path.join(__dirname, req.files.input.name);
    const bankPath = path.join(__dirname, req.files.bank.name);

    req.files.input.mv(inputPath, err => {
        if (err) return res.status(500).send('❌ Error saving input file');

        req.files.bank.mv(bankPath, err2 => {
            if (err2) return res.status(500).send('❌ Error saving bank file');

            const outputPath = compareStatements(inputPath, bankPath);

            res.download(outputPath, 'comparison_output.xlsx', () => {
                fs.unlinkSync(inputPath);
                fs.unlinkSync(bankPath);
                fs.unlinkSync(outputPath);
            });
        });
    });
});

app.get('/', (req, res) => {
    res.send('✅ Backend running!');
});

app.listen(PORT, () => {
    console.log(`✅ Server running at http://localhost:${PORT}`);
});
