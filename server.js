const express = require('express');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 5000;

app.use(cors());
app.use(express.json());
app.use(fileUpload());

// Improved function to extract amount
function extractAmount(row) {
    const amountKey = Object.keys(row).find(k =>
        (k.toLowerCase().includes('amount') && k.toLowerCase().includes('inr')) ||
        k.toLowerCase().includes('deposits') ||
        k.toLowerCase().includes('amount')
    );
    if (!amountKey || !row[amountKey]) return '0.00';
    const raw = row[amountKey].toString().replace(/,/g, '');
    return parseFloat(raw).toFixed(2);
}

// Improved function to extract UTR
function extractUTR(row) {
    const utrKey = Object.keys(row).find(k =>
        k.toLowerCase().includes('description') ||
        k.toLowerCase().includes('tracking') ||
        k.toLowerCase().includes('narration') ||
        k.toLowerCase().includes('utr')
    );
    if (!utrKey || !row[utrKey]) return null;

    const str = row[utrKey].toString();
    const match = str.match(/\b[A-Z0-9]{10,}\b/); // Alphanumeric UTRs with at least 10 characters
    return match ? match[0] : null;
}

// Comparison function
function compareStatements(inputPath, bankPath) {
    const inputWorkbook = xlsx.readFile(inputPath);
    const bankWorkbook = xlsx.readFile(bankPath);

    const inputData = xlsx.utils.sheet_to_json(inputWorkbook.Sheets[inputWorkbook.SheetNames[0]]);
    const bankData = xlsx.utils.sheet_to_json(bankWorkbook.Sheets[bankWorkbook.SheetNames[0]]);

    const inputMap = new Map();
    inputData.forEach(row => {
        const utr = row.Utr || extractUTR(row);
        if (utr) inputMap.set(utr, row);
    });

    const bankMap = new Map();
    bankData.forEach(row => {
        const utr = extractUTR(row);
        if (utr) bankMap.set(utr, row);
    });

    const result = [];

    inputMap.forEach((inputRow, utr) => {
        const bankRow = bankMap.get(utr);
        const userId = inputRow['User Id'] || 'N/A';
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
    const outputWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(outputWorkbook, outputSheet, 'Comparison');

    const outputPath = path.join(__dirname, 'comparison_output.xlsx');
    xlsx.writeFile(outputWorkbook, outputPath);
    return outputPath;
}

// Upload endpoint
app.post('/upload', (req, res) => {
    if (!req.files || !req.files.input || !req.files.bank) {
        return res.status(400).send('Missing files');
    }

    const inputPath = path.join(__dirname, req.files.input.name);
    const bankPath = path.join(__dirname, req.files.bank.name);

    req.files.input.mv(inputPath, err => {
        if (err) return res.status(500).send(err);

        req.files.bank.mv(bankPath, err2 => {
            if (err2) return res.status(500).send(err2);

            const outputFilePath = compareStatements(inputPath, bankPath);

            res.download(outputFilePath, 'comparison_output.xlsx', () => {
                fs.unlinkSync(inputPath);
                fs.unlinkSync(bankPath);
                fs.unlinkSync(outputFilePath);
            });
        });
    });
});

// Health check
app.get('/', (req, res) => {
    res.send('✅ Server is running!');
});

// Start server
if (require.main === module) {
    app.listen(PORT, () => {
        console.log(`🚀 Server running on http://localhost:${PORT}`);
    });
} else {
    module.exports = app;
}