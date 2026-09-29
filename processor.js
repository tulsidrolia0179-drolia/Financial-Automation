const xlsx = require('xlsx');
const fs = require('fs');
const pdfParse = require('pdf-parse');

async function extractPreviousYearData(filePath) {
    try {
        let extractedData = [];
        if (filePath.toLowerCase().endsWith('.xlsx') || filePath.toLowerCase().endsWith('.xls')) {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            extractedData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
            return { success: true, type: 'excel', data: extractedData.slice(0, 5) };
        } else if (filePath.toLowerCase().endsWith('.pdf')) {
            const dataBuffer = fs.readFileSync(filePath);
            const data = await pdfParse(dataBuffer);
            const textLines = data.text.split('\n').filter(line => line.trim() !== '');
            return { success: true, type: 'pdf', data: textLines.slice(0, 5) };
        } else {
            return { success: false, error: 'Unsupported file format.' };
        }
    } catch (error) {
        return { success: false, error: error.message };
    }
}

module.exports = { extractPreviousYearData };
