const ExcelJS = require('exceljs');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const xml2js = require('xml2js');
const { app } = require('electron');

const userDataPath = app ? app.getPath('userData') : (process.env.APPDATA || (process.platform == 'darwin' ? process.env.HOME + '/Library/Preferences' : process.env.HOME + '/.local/share'));
const MEMORY_FILE = path.join(userDataPath, 'custom_mapping.json');

if (!fs.existsSync(MEMORY_FILE)) {
    try {
        fs.writeFileSync(MEMORY_FILE, JSON.stringify({}));
    } catch (e) {
        console.error("Failed to create custom_mapping.json:", e);
    }
}

const baseTallyToSchedule3Map = {
    'Capital Account': { type: 'EquityAndLiabilities', head: 'Shareholder\'s Funds', subHead: 'Share Capital' },
    'Reserves & Surplus': { type: 'EquityAndLiabilities', head: 'Shareholder\'s Funds', subHead: 'Reserves and Surplus' },
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
    'Sales Accounts': { type: 'PL', head: 'Revenue From Operations', subHead: 'Sale of Products/Services' },
    'Other Income': { type: 'PL', head: 'Other Income', subHead: 'Other Non-Operating Income' },
    'Purchase Accounts': { type: 'PL', head: 'Expenses', subHead: 'Cost of Materials Consumed' },
    'Direct Expenses': { type: 'PL', head: 'Expenses', subHead: 'Other Expenses' },
    'Indirect Expenses': { type: 'PL', head: 'Expenses', subHead: 'Other Expenses' },
    'Depreciation': { type: 'PL', head: 'Expenses', subHead: 'Depreciation and Amortization Expense' },
    'Interest & Finance Charges': { type: 'PL', head: 'Expenses', subHead: 'Finance Costs' }
};

function loadCustomMappings() {
    if (fs.existsSync(MEMORY_FILE)) {
        try {
            return JSON.parse(fs.readFileSync(MEMORY_FILE, 'utf8'));
        } catch (e) {
            console.error("Error reading memory file", e);
        }
    }
    return {};
}

function saveCustomMapping(userLedger, selectedSchedule3Head) {
    const current = loadCustomMappings();
    current[userLedger] = selectedSchedule3Head;
    try {
        fs.writeFileSync(MEMORY_FILE, JSON.stringify(current, null, 2));
    } catch (e) {
        console.error("Error saving memory file", e);
    }
}

async function parseTallyXML(filePath) {
    const xmlData = fs.readFileSync(filePath, 'utf8');
    const parser = new xml2js.Parser({ explicitArray: false });
    const result = await parser.parseStringPromise(xmlData);
    const data = [];
    
    let ledgers = [];
    if (result && result.ENVELOPE && result.ENVELOPE.BODY && result.ENVELOPE.BODY.DATA) {
        const tallyData = result.ENVELOPE.BODY.DATA;
        if (tallyData.TALLYMESSAGE) {
            const msgs = Array.isArray(tallyData.TALLYMESSAGE) ? tallyData.TALLYMESSAGE : [tallyData.TALLYMESSAGE];
            for (const msg of msgs) {
                if (msg.LEDGER) {
                    ledgers.push(msg.LEDGER);
                }
            }
        }
    }

    for (const leg of ledgers) {
        const name = leg.NAME || leg.$.NAME || 'Unknown';
        const parent = leg.PARENT || '';
        const opening = parseFloat(leg.OPENINGBALANCE || 0);
        const closing = parseFloat(leg.CLOSINGBALANCE || leg.BALANCETOTAL || 0);
        data.push({
            ledger: name,
            group: parent,
            opening: isNaN(opening) ? 0 : opening,
            closing: isNaN(closing) ? 0 : closing
        });
    }

    return data;
}

async function mergePreviousYearFiles(filePaths) {
    if (!filePaths) return [];
    const paths = Array.isArray(filePaths) ? filePaths : [filePaths];
    let mergedData = [];

    for (const fp of paths) {
        if (!fp) continue;
        if (fp.endsWith('.xml')) {
            const parsed = await parseTallyXML(fp);
            mergedData = mergedData.concat(parsed);
        } else if (fp.endsWith('.xlsx') || fp.endsWith('.xls')) {
            const workbook = xlsx.readFile(fp);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = xlsx.utils.sheet_to_json(sheet);
            mergedData = mergedData.concat(jsonData);
        } else if (fp.endsWith('.pdf')) {
            const dataBuffer = fs.readFileSync(fp);
            const pdfData = await pdfParse(dataBuffer);
            mergedData.push({ pdfText: pdfData.text });
        }
    }

    return mergedData;
}

async function processFinancials(entityType, templatePath, prevYearPaths, currentYearPath) {
    console.log(`Processing ${entityType} financials...`);

    const prevYearData = await mergePreviousYearFiles(prevYearPaths);

    let currentYearData = [];
    if (currentYearPath) {
        if (currentYearPath.endsWith('.xml')) {
            currentYearData = await parseTallyXML(currentYearPath);
        } else if (currentYearPath.endsWith('.xlsx') || currentYearPath.endsWith('.xls')) {
            const workbook = xlsx.readFile(currentYearPath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            currentYearData = xlsx.utils.sheet_to_json(sheet);
        }
    }

    return `Processed ${prevYearData.length} previous year records and ${currentYearData.length} current year records successfully.`;
}

module.exports = {
    processFinancials,
    parseTallyXML,
    mergePreviousYearFiles,
    loadCustomMappings,
    saveCustomMapping
};
