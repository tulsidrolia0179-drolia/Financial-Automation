const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Utility: Calculate Depreciation as per Income Tax Act 1961 (Block of Assets)
function calculateITDepreciation(assetValue, rate, daysUsed) {
    // Basic logic: if used < 180 days, apply half the rate, else full rate
    let applicableRate = daysUsed < 180 ? (rate / 2) : rate;
    return assetValue * (applicableRate / 100);
}

// Utility: Calculate Depreciation as per Schedule 3 of Companies Act (Useful Life)
function calculateCompaniesActDepreciation(assetValue, usefulLife, salvageValue = 0.05) {
    // SLM based on useful life, keeping 5% salvage value standard
    let depreciableAmount = assetValue - (assetValue * salvageValue);
    return depreciableAmount / usefulLife;
}

// Main Processing Engine
async function processFinancials(entityType, templatePath, prevYearPath, currentYearPath) {
    try {
        console.log(`Starting automated processing for entity type: ${entityType}`);
        
        const workbook = new ExcelJS.Workbook();
        
        // 1. Load User's Formatting Template
        if (templatePath && fs.existsSync(templatePath)) {
            await workbook.xlsx.readFile(templatePath);
            console.log('User formatting template loaded successfully.');
        } else {
            // Fallback base structure if no template is provided
            workbook.addWorksheet('Balance Sheet');
            workbook.addWorksheet('Profit & Loss');
            workbook.addWorksheet('Depreciation');
        }

        // 2. Map Previous Year Data (Audited pdf/excel)
        if (prevYearPath) {
            console.log(`Extracting previous year figures from: ${prevYearPath}`);
            // TODO: Extract closing balances and map to current year opening balances
        }

        // 3. Process Current Year Data (Tally Fetch / Trial Balance)
        if (currentYearPath) {
            console.log(`Parsing current year data from: ${currentYearPath}`);
            // TODO: Map Tally trial balance ledgers to respective Schedule 3 / IT Act heads
        }

        // 4. Execute Entity-Specific Depreciation & Compliance Rules
        if (['proprietorship', 'partnership', 'trust'].includes(entityType)) {
            console.log('Executing Non-Corporate Logic: Income Tax Act 1961 Depreciation...');
            // Apply block-wise WDV depreciation logic
            
        } else if (entityType === 'company') {
            console.log('Executing Corporate Logic: Schedule 3 Format & Dual Depreciation...');
            // Apply dual depreciation (IT Act + Companies Act)
            // Enforce Schedule 3 horizontal/vertical mapping
        }

        // 5. Generate Final Output
        const outputPath = path.join(__dirname, `Final_Financials_${Date.now()}.xlsx`);
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Success! Audited format ready at: ${outputPath}`);
        
        return { success: true, path: outputPath };
        
    } catch (error) {
        console.error('Error during automated processing:', error);
        return { success: false, error: error.message };
    }
}

module.exports = {
    processFinancials,
    calculateITDepreciation,
    calculateCompaniesActDepreciation
};
