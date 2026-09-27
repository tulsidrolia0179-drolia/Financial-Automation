const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');

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

async function extractPDFData(pdfPath) {
    console.log('Initiating PDF extraction...');
    const dataBuffer = fs.readFileSync(pdfPath);
    const data = await pdfParse(dataBuffer);
    const extractedBalances = {};
    const lines = data.text.split('\n');
    for (let line of lines) {
        line = line.trim();
        const match = line.match(/^([a-zA-Z\s\(\),&]+)\s+([\d,]+\.?\d*)$/);
        if (match) {
            const head = match[1].trim();
            const amount = parseFloat(match[2].replace(/,/g, ''));
            if (head && !isNaN(amount)) extractedBalances[head] = amount;
        }
    }
    return extractedBalances;
}

async function extractTallyData(tbFilePath) {
    const tbWorkbook = new ExcelJS.Workbook();
    await tbWorkbook.xlsx.readFile(tbFilePath);
    const ws = tbWorkbook.worksheets[0];
    const financialData = { EquityAndLiabilities: {}, Assets: {}, Income: {}, Expense: {} };
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
            if (typeof closing === 'number' && closing !== 0) netBalance = closing;
            else netBalance = (type === 'Assets' || type === 'Expense') ? (debit - credit) : (credit - debit);

            if (!financialData[type][subHead]) financialData[type][subHead] = 0;
            financialData[type][subHead] += netBalance;
        }
    });
    return financialData;
}

function calculateITDepreciation(assetValue, rate, dateOfPurchase) {
    const purchaseDate = new Date(dateOfPurchase);
    const cutoffDate = new Date(purchaseDate.getFullYear(), 9, 4);
    return assetValue * ((purchaseDate > cutoffDate ? (rate / 2) : rate) / 100);
}

function calculateCompaniesActDepreciation(assetValue, usefulLife) {
    return (assetValue - (assetValue * 0.05)) / usefulLife;
}

// NEW FORMAT PRESERVATION ENGINE
async function processFinancials(entityType, templatePath, prevYearPath, currentYearPath) {
    try {
        const workbook = new ExcelJS.Workbook();
        const hasTemplate = templatePath && fs.existsSync(templatePath);
        
        if (hasTemplate) await workbook.xlsx.readFile(templatePath);
        else {
            workbook.addWorksheet('Balance Sheet');
            workbook.addWorksheet('Profit & Loss');
        }

        let prevYearData = {};
        if (prevYearPath && fs.existsSync(prevYearPath) && prevYearPath.toLowerCase().endsWith('.pdf')) {
            prevYearData = await extractPDFData(prevYearPath);
        }

        let tallyData = null;
        if (currentYearPath && fs.existsSync(currentYearPath)) {
            tallyData = await extractTallyData(currentYearPath);
        }

        if (tallyData) {
            const bsSheet = workbook.getWorksheet('Balance Sheet') || workbook.worksheets[0];

            if (hasTemplate) {
                // --- NON-DESTRUCTIVE INJECTION ENGINE ---
                console.log('Template detected. Executing Non-Destructive Format Preservation...');
                
                // Flatten data to lowercase for robust matching
                const flatData = {};
                const normalize = (str) => str.toLowerCase().replace(/\s+/g, ' ').trim();

                for (const [subHead, amount] of Object.entries(tallyData.EquityAndLiabilities)) flatData[normalize(subHead)] = { cy: amount };
                for (const [subHead, amount] of Object.entries(tallyData.Assets)) flatData[normalize(subHead)] = { cy: amount };
                for (const [subHead, amount] of Object.entries(prevYearData)) {
                    const key = normalize(subHead);
                    if (!flatData[key]) flatData[key] = { py: amount };
                    else flatData[key].py = amount;
                }

                // Scan user's format and inject directly without overwriting formulas
                bsSheet.eachRow((row, rowNumber) => {
                    let textCol = -1;
                    let foundKey = null;

                    // Find matching label in the row
                    row.eachCell((cell, colNumber) => {
                        if (cell.type === ExcelJS.ValueType.String && cell.text) {
                            const cleanText = normalize(cell.text);
                            if (flatData[cleanText]) {
                                foundKey = cleanText;
                                textCol = colNumber;
                            }
                        }
                    });

                    // If found, safely inject into subsequent columns (Assuming Col + 2 = CY, Col + 3 = PY)
                    if (foundKey) {
                        const data = flatData[foundKey];
                        const cyCell = row.getCell(textCol + 2);
                        const pyCell = row.getCell(textCol + 3);

                        // Only overwrite if it's NOT a formula (preserves client's =SUM functions)
                        if (data.cy !== undefined && !cyCell.formula) {
                            cyCell.value = data.cy;
                            cyCell.numFmt = '#,##0.00';
                        }
                        if (data.py !== undefined && !pyCell.formula) {
                            pyCell.value = data.py;
                            pyCell.numFmt = '#,##0.00';
                        }
                    }
                });

            } else {
                // --- FALLBACK: BUILD FROM SCRATCH ---
                bsSheet.columns = [
                    { header: 'Particulars', key: 'particulars', width: 40 },
                    { header: 'Note No.', key: 'note', width: 10 },
                    { header: 'Current Year (₹)', key: 'cy', width: 20 },
                    { header: 'Previous Year (₹)', key: 'py', width: 20 }
                ];

                bsSheet.addRow({ particulars: 'I. EQUITY AND LIABILITIES' }).font = { bold: true };
                for (const [subHead, amount] of Object.entries(tallyData.EquityAndLiabilities)) {
                    bsSheet.addRow({ particulars: `   ${subHead}`, cy: amount, py: prevYearData[subHead] || '' });
                }

                bsSheet.addRow({}); 
                bsSheet.addRow({ particulars: 'II. ASSETS' }).font = { bold: true };
                for (const [subHead, amount] of Object.entries(tallyData.Assets)) {
                    bsSheet.addRow({ particulars: `   ${subHead}`, cy: amount, py: prevYearData[subHead] || '' });
                }
            }
        }

        const outputPath = path.join(__dirname, `Final_Audited_Financials_${Date.now()}.xlsx`);
        await workbook.xlsx.writeFile(outputPath);
        return { success: true, path: outputPath };
        
    } catch (error) {
        return { success: false, error: error.message };
    }
}

module.exports = { processFinancials, calculateITDepreciation, calculateCompaniesActDepreciation };
