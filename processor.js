const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const xml2js = require('xml2js');

const MEMORY_FILE = path.join(__dirname, 'custom_mapping.json');

const baseTallyToSchedule3Map = {
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

// SMART MEMORY ENGINE
function getActiveMapping() {
    let customMap = {};
    if (fs.existsSync(MEMORY_FILE)) {
        try {
            customMap = JSON.parse(fs.readFileSync(MEMORY_FILE, 'utf-8'));
        } catch (e) {
            console.warn('Warning: custom_mapping.json is invalid, using defaults.');
        }
    } else {
        // Auto-create the memory database on first run
        fs.writeFileSync(MEMORY_FILE, JSON.stringify({
            "_EXAMPLE_CUSTOM_LEDGER_": { "type": "Assets", "head": "Current Assets", "subHead": "Custom Asset Name" }
        }, null, 4));
    }
    return { ...baseTallyToSchedule3Map, ...customMap };
}

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
    const activeMapping = getActiveMapping();
    const tbWorkbook = new ExcelJS.Workbook();
    await tbWorkbook.xlsx.readFile(tbFilePath);
    const ws = tbWorkbook.worksheets[0];
    const financialData = { EquityAndLiabilities: {}, Assets: {}, Income: {}, Expense: {}, AssetAdditions: [] };
    let currentActiveGroup = null;

    ws.eachRow((row, rowNumber) => {
        if (rowNumber < 5) return;
        const particulars = row.getCell(1).text ? row.getCell(1).text.trim() : null;
        if (!particulars) return;
        const debit = row.getCell(3).value || 0;
        const credit = row.getCell(4).value || 0;
        const closing = row.getCell(5).value || 0;

        if (activeMapping[particulars]) {
            currentActiveGroup = particulars;
        } else if (currentActiveGroup) {
            const mapping = activeMapping[currentActiveGroup];
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

async function extractTallyXMLData(xmlPath) {
    console.log('Initiating Transaction-Level Tally XML extraction...');
    const activeMapping = getActiveMapping();
    const xmlString = fs.readFileSync(xmlPath, 'utf-8');
    const parser = new xml2js.Parser({ explicitArray: false, ignoreAttrs: true });
    const result = await parser.parseStringPromise(xmlString);

    const financialData = { EquityAndLiabilities: {}, Assets: {}, Income: {}, Expense: {}, AssetAdditions: [] };

    try {
        const messages = result.ENVELOPE?.BODY?.DATA?.TALLYMESSAGE || [];
        const msgArray = Array.isArray(messages) ? messages : [messages];
        
        for (const msg of msgArray) {
            if (msg.LEDGER && msg.LEDGER.PARENT) {
                const parentGroup = msg.LEDGER.PARENT;
                const closingBal = parseFloat(msg.LEDGER.CLOSINGBALANCE) || 0;
                if (activeMapping[parentGroup]) {
                    const { type, subHead } = activeMapping[parentGroup];
                    if (!financialData[type][subHead]) financialData[type][subHead] = 0;
                    financialData[type][subHead] += Math.abs(closingBal);
                }
            }

            if (msg.VOUCHER && msg.VOUCHER.DATE) {
                const voucherDate = msg.VOUCHER.DATE; // YYYYMMDD
                const entries = msg.VOUCHER['ALLLEDGERENTRIES.LIST'] || [];
                const entryArray = Array.isArray(entries) ? entries : [entries];
                
                for (const entry of entryArray) {
                    const ledgerName = entry.LEDGERNAME || '';
                    const amount = parseFloat(entry.AMOUNT) || 0;
                    if (amount < 0 && (ledgerName.toLowerCase().includes('asset') || ledgerName.toLowerCase().includes('machinery') || ledgerName.toLowerCase().includes('computer'))) {
                        const year = voucherDate.substring(0, 4);
                        const month = voucherDate.substring(4, 6);
                        const day = voucherDate.substring(6, 8);
                        financialData.AssetAdditions.push({
                            asset: ledgerName,
                            amount: Math.abs(amount),
                            date: new Date(`${year}-${month}-${day}`)
                        });
                    }
                }
            }
        }
    } catch (err) {
        console.warn("XML Structure warning:", err.message);
    }

    console.log(`XML Extraction complete. Found ${financialData.AssetAdditions.length} fixed asset additions with precise dates.`);
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
            if (currentYearPath.toLowerCase().endsWith('.xml')) {
                tallyData = await extractTallyXMLData(currentYearPath);
            } else {
                tallyData = await extractTallyData(currentYearPath);
            }
        }

        if (tallyData) {
            const bsSheet = workbook.getWorksheet('Balance Sheet') || workbook.worksheets[0];

            if (hasTemplate) {
                console.log('Template detected. Executing Non-Destructive Format Preservation...');
                
                const flatData = {};
                const normalize = (str) => str.toLowerCase().replace(/\s+/g, ' ').trim();

                for (const [subHead, amount] of Object.entries(tallyData.EquityAndLiabilities)) flatData[normalize(subHead)] = { cy: amount };
                for (const [subHead, amount] of Object.entries(tallyData.Assets)) flatData[normalize(subHead)] = { cy: amount };
                for (const [subHead, amount] of Object.entries(prevYearData)) {
                    const key = normalize(subHead);
                    if (!flatData[key]) flatData[key] = { py: amount };
                    else flatData[key].py = amount;
                }

                bsSheet.eachRow((row) => {
                    let textCol = -1;
                    let foundKey = null;

                    row.eachCell((cell, colNumber) => {
                        if (cell.type === ExcelJS.ValueType.String && cell.text) {
                            const cleanText = normalize(cell.text);
                            if (flatData[cleanText]) {
                                foundKey = cleanText;
                                textCol = colNumber;
                            }
                        }
                    });

                    if (foundKey) {
                        const data = flatData[foundKey];
                        const cyCell = row.getCell(textCol + 2);
                        const pyCell = row.getCell(textCol + 3);

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

            if (tallyData.AssetAdditions && tallyData.AssetAdditions.length > 0) {
                console.log('Generating automated depreciation schedule from XML transaction dates...');
                let depSheet = workbook.getWorksheet('Depreciation Schedule');
                if (!depSheet) depSheet = workbook.addWorksheet('Depreciation Schedule');
                
                depSheet.columns = [
                    { header: 'Asset Name', key: 'asset', width: 30 },
                    { header: 'Date of Addition', key: 'date', width: 20 },
                    { header: 'Amount Added (₹)', key: 'amount', width: 20 },
                    { header: 'Days Used (>180?)', key: 'days', width: 20 },
                    { header: 'IT Act Dep (%)', key: 'dep', width: 15 }
                ];
                
                tallyData.AssetAdditions.forEach(asset => {
                    const cutoffDate = new Date(asset.date.getFullYear(), 9, 4); // Oct 4
                    const isMoreThan180 = asset.date <= cutoffDate ? 'Yes (>180)' : 'No (<180)';
                    depSheet.addRow({
                        asset: asset.asset,
                        date: asset.date.toISOString().split('T')[0],
                        amount: asset.amount,
                        days: isMoreThan180,
                        dep: 'Auto-Calculated'
                    });
                });
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
