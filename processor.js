const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');

// Robust Mapping Dictionary: Tally Groups -> Schedule 3 / IT Act Heads
const tallyToSchedule3Map = {
    'Capital Account': { type: 'EquityAndLiabilities', head: 'Shareholder\'s Funds', subHead: 'Share Capital' },
    'Reserves & Surplus': { type: 'EquityAndLiabilities', head: 'Shareholder\'s Funds', subHead: 'Reserves & Surplus' },
    'Long Term Borrowings': { type: 'EquityAndLiabilities', head: 'Non-Current Liabilities', subHead: 'Long-Term Borrowings' },
    'Short Term Borrowings': { type: 'EquityAndLiabilities', head: 'Current Liabilities', subHead: 'Short-Term Borrowings' },
    'Sundry Creditors': { type: 'EquityAndLiabilities', head: 'Current Liabilities', subHead: 'Trade Payables' },
    'Duties & Taxes': { type: 'EquityAndLiabilities', head: 'Current Liabilities', subHead: 'Short-Term Provisions' },
    'Provisions': { type: 'EquityAndLiabilities', head: 'Current Liabilities', subHead: 'Short-Term Provisions' },
    'Fixed Assets': { type: 'Assets', head: 'Non-Current Assets', subHead: 'Property, Plant and Equipment' },
    'Investments': { type: 'Assets', head: 'Non-Current Assets', subHead: 'Non-Current Investments' },
    'Closing Stock': { type: 'Assets', head: 'Current Assets', subHead: 'Inventories' },
    'Sundry Debtors': { type: 'Assets', head: 'Current Assets', subHead: 'Trade Receivables' },
    'Cash-in-hand': { type: 'Assets', head: 'Current Assets', subHead: 'Cash and Cash Equivalents' },
    'Bank Accounts': { type: 'Assets', head: 'Current Assets', subHead: 'Bank Balances' },
    'Loans & Advances (Asset)': { type: 'Assets', head: 'Current Assets', subHead: 'Short-Term Loans and Advances' },
    'Sales Accounts': { type: 'Income', head: 'Revenue', subHead: 'Revenue from Operations' },
    'Direct Incomes': { type: 'Income', head: 'Revenue', subHead: 'Other Income' },
    'Indirect Incomes': { type: 'Income', head: 'Revenue', subHead: 'Other Income' },
    'Purchase Accounts': { type: 'Expense', head: 'Expenses', subHead: 'Purchases of Stock-in-Trade' },
    'Direct Expenses': { type: 'Expense', head: 'Expenses', subHead: 'Direct Expenses' },
    'Indirect Expenses': { type: 'Expense', head: 'Expenses', subHead: 'Other Expenses' }
};

// PDF Extractor for Previous Year Audited Statements
async function extractPDFData(pdfPath) {
    console.log('Initiating PDF extraction for previous year data...');
    const dataBuffer = fs.readFileSync(pdfPath);
    const data = await pdfParse(dataBuffer);
    
    const pdfText = data.text;
    const extractedBalances = {};

    const lines = pdfText.split('\n');
    for (let line of lines) {
        line = line.trim();
        if (!line) continue;
        
        // Regex to find text ending with a number (handles commas and decimals)
        // e.g. "Trade Receivables 1,50,000.00" or "Share Capital 50000"
        const match = line.match(/^([a-zA-Z\s\(\),&]+)\s+([\d,]+\.?\d*)$/);
        if (match) {
            const head = match[1].trim();
            const amount = parseFloat(match[2].replace(/,/g, ''));
            if (head && !isNaN(amount)) {
                extractedBalances[head] = amount;
            }
        }
    }
    
    console.log(`Extracted ${Object.keys(extractedBalances).length} data points from PDF successfully.`);
    return extractedBalances;
}

// Advanced Tally Trial Balance Extractor
async function extractTallyData(tbFilePath) {
    console.log('Initiating Tally TB extraction engine...');
    const tbWorkbook = new ExcelJS.Workbook();
    await tbWorkbook.xlsx.readFile(tbFilePath);
    const ws = tbWorkbook.worksheets[0];

    const financialData = {
        EquityAndLiabilities: {},
        Assets: {},
        Income: {},
        Expense: {}
    };

    let currentActiveGroup = null;

    ws.eachRow((row, rowNumber) => {
        if (rowNumber < 5) return;

        const particulars = row.getCell(1).text ? row.getCell(1).text.trim() : null;
        if (!particulars) return;

        const debit = row.getCell(3).value || 0;
        const credit = row.getCell(4).value || 0;
        const closing = row.getCell(5).value || 0;

        if (tallyToSchedule3Map[particulars]) {
            currentActiveGroup = particulars;
        } else if (currentActiveGroup) {
            const mapping = tallyToSchedule3Map[currentActiveGroup];
            if (!mapping) return;

            const { type, subHead } = mapping;
            let netBalance = 0;
            
            if (typeof closing === 'number' && closing !== 0) {
                netBalance = closing;
            } else {
                netBalance = (type === 'Assets' || type === 'Expense') ? (debit - credit) : (credit - debit);
            }

            if (!financialData[type][subHead]) {
                financialData[type][subHead] = 0;
            }
            financialData[type][subHead] += netBalance;
        }
    });

    console.log('Tally extraction completed successfully.');
    return financialData;
}

// Mathematical Depreciation Engines
function calculateITDepreciation(assetValue, rate, dateOfPurchase) {
    const purchaseDate = new Date(dateOfPurchase);
    const cutoffDate = new Date(purchaseDate.getFullYear(), 9, 4);
    const applicableRate = purchaseDate > cutoffDate ? (rate / 2) : rate;
    return assetValue * (applicableRate / 100);
}

function calculateCompaniesActDepreciation(assetValue, usefulLife) {
    const salvageValue = assetValue * 0.05;
    const depreciableAmount = assetValue - salvageValue;
    return depreciableAmount / usefulLife;
}

// Main Engine
async function processFinancials(entityType, templatePath, prevYearPath, currentYearPath) {
    try {
        console.log(`Starting generation for: ${entityType.toUpperCase()}`);
        const workbook = new ExcelJS.Workbook();
        
        if (templatePath && fs.existsSync(templatePath)) {
            await workbook.xlsx.readFile(templatePath);
        } else {
            workbook.addWorksheet('Balance Sheet');
            workbook.addWorksheet('Profit & Loss');
            workbook.addWorksheet('Depreciation Schedule');
        }

        // Extract Previous Year Data (PDF Parsing integration)
        let prevYearData = {};
        if (prevYearPath && fs.existsSync(prevYearPath)) {
            if (prevYearPath.toLowerCase().endsWith('.pdf')) {
                prevYearData = await extractPDFData(prevYearPath);
            }
        }

        let tallyData = null;
        if (currentYearPath && fs.existsSync(currentYearPath)) {
            tallyData = await extractTallyData(currentYearPath);
        }

        if (tallyData) {
            const bsSheet = workbook.getWorksheet('Balance Sheet') || workbook.worksheets[0];
            bsSheet.columns = [
                { header: 'Particulars', key: 'particulars', width: 40 },
                { header: 'Note No.', key: 'note', width: 10 },
                { header: 'Current Year (₹)', key: 'cy', width: 20 },
                { header: 'Previous Year (₹)', key: 'py', width: 20 }
            ];

            bsSheet.addRow({ particulars: 'I. EQUITY AND LIABILITIES' }).font = { bold: true };
            
            for (const [subHead, amount] of Object.entries(tallyData.EquityAndLiabilities)) {
                // Automate Word-to-Word Matching from PDF for Previous Year Opening Balance
                let openingBalance = prevYearData[subHead] || '';
                bsSheet.addRow({ particulars: `   ${subHead}`, cy: amount, py: openingBalance });
            }

            bsSheet.addRow({}); 
            bsSheet.addRow({ particulars: 'II. ASSETS' }).font = { bold: true };

            for (const [subHead, amount] of Object.entries(tallyData.Assets)) {
                let openingBalance = prevYearData[subHead] || '';
                bsSheet.addRow({ particulars: `   ${subHead}`, cy: amount, py: openingBalance });
            }

            bsSheet.eachRow((row) => {
                const cyCell = row.getCell('cy');
                const pyCell = row.getCell('py');
                if (typeof cyCell.value === 'number') cyCell.numFmt = '#,##0.00';
                if (typeof pyCell.value === 'number') pyCell.numFmt = '#,##0.00';
            });
        }

        const outputPath = path.join(__dirname, `Final_Audited_Financials_${Date.now()}.xlsx`);
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Excel file successfully created at: ${outputPath}`);
        
        return { success: true, path: outputPath };
        
    } catch (error) {
        console.error('Critical Error during processing:', error);
        return { success: false, error: error.message };
    }
}

module.exports = {
    processFinancials,
    calculateITDepreciation,
    calculateCompaniesActDepreciation
};
