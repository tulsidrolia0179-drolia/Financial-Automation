const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

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

    // Parse Row by Row
    ws.eachRow((row, rowNumber) => {
        // Tally exports usually start data below row 5
        if (rowNumber < 5) return;

        // Typically: Col 1 = Particulars, Col 2 = Opening, Col 3 = Debit, Col 4 = Credit, Col 5 = Closing
        const particulars = row.getCell(1).text ? row.getCell(1).text.trim() : null;
        if (!particulars) return;

        const debit = row.getCell(3).value || 0;
        const credit = row.getCell(4).value || 0;
        const closing = row.getCell(5).value || 0;

        // Check if this row is a Group Header
        if (tallyToSchedule3Map[particulars]) {
            currentActiveGroup = particulars;
        } 
        // If it's a ledger account sitting under a group
        else if (currentActiveGroup) {
            const mapping = tallyToSchedule3Map[currentActiveGroup];
            if (!mapping) return;

            const { type, subHead } = mapping;
            
            // Determine net balance
            let netBalance = 0;
            if (typeof closing === 'number' && closing !== 0) {
                netBalance = closing; // Prefer exact closing balance if present
            } else {
                // Fallback to calculation
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
    // Income Tax Act 1961 - 180 Days rule
    const purchaseDate = new Date(dateOfPurchase);
    const cutoffDate = new Date(purchaseDate.getFullYear(), 9, 4); // Oct 4th approx for 180 days in leap/non-leap
    const applicableRate = purchaseDate > cutoffDate ? (rate / 2) : rate;
    return assetValue * (applicableRate / 100);
}

function calculateCompaniesActDepreciation(assetValue, usefulLife, yearsUsed = 0) {
    // Schedule II of Companies Act, 2013 - SLM approach with 5% Salvage
    const salvageValue = assetValue * 0.05;
    const depreciableAmount = assetValue - salvageValue;
    const annualDepreciation = depreciableAmount / usefulLife;
    return annualDepreciation;
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

        let tallyData = null;
        if (currentYearPath && fs.existsSync(currentYearPath)) {
            tallyData = await extractTallyData(currentYearPath);
        }

        if (tallyData) {
            // Populate Balance Sheet intelligently
            const bsSheet = workbook.getWorksheet('Balance Sheet') || workbook.worksheets[0];
            bsSheet.columns = [
                { header: 'Particulars', key: 'particulars', width: 40 },
                { header: 'Note No.', key: 'note', width: 10 },
                { header: 'Current Year (₹)', key: 'cy', width: 20 },
                { header: 'Previous Year (₹)', key: 'py', width: 20 }
            ];

            // Write Equity & Liabilities
            let rowCounter = 2;
            bsSheet.addRow({ particulars: 'I. EQUITY AND LIABILITIES' }).font = { bold: true };
            
            for (const [subHead, amount] of Object.entries(tallyData.EquityAndLiabilities)) {
                bsSheet.addRow({ particulars: `   ${subHead}`, cy: amount });
            }

            // Write Assets
            bsSheet.addRow({}); // Empty row
            bsSheet.addRow({ particulars: 'II. ASSETS' }).font = { bold: true };

            for (const [subHead, amount] of Object.entries(tallyData.Assets)) {
                bsSheet.addRow({ particulars: `   ${subHead}`, cy: amount });
            }

            // Format numbers
            bsSheet.eachRow((row) => {
                const cyCell = row.getCell('cy');
                if (typeof cyCell.value === 'number') {
                    cyCell.numFmt = '#,##0.00';
                }
            });
        }

        // Export logic
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
