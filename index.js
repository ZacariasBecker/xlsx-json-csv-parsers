var fs = require("fs");
const xlsx = require('xlsx');

function convertExcelFileToJsonUsingXlsx() {
    // Read the file using pathname
    const file = xlsx.readFile('./Book1.xlsx');
    // Grab the sheet info from the file
    const sheetNames = file.SheetNames;
    const totalSheets = sheetNames.length;
    // Variable to store our data
    let parsedData = [];
    // Loop through sheets
    for (let i = 0; i < totalSheets; i++) {
        // Convert to json using xlsx
        const tempData = xlsx.utils.sheet_to_json(file.Sheets[sheetNames[i]]);
        // Skip header row which is the colum names
        tempData.shift();
        // Add the sheet's json to our data array
        parsedData.push(...tempData);
    }
    // call a function to save the data in a json file
    generateJSONFile(parsedData);
}

function generateJSONFile(data) {
    try {
        fs.writeFileSync('data.json', JSON.stringify(data));
    } catch (err) {
        console.error(err);
    }
}

const convertToXlsxCsv = (jsonData, outputFilePath, fileType) => {
    // Create a new workbook
    const workbook = xlsx.utils.book_new();
    // Add the JSON data to a new sheet
    const sheet = xlsx.utils.json_to_sheet(jsonData);
    // Add the sheet to the workbook
    xlsx.utils.book_append_sheet(workbook, sheet, 'Sheet 1');
    // Write the workbook to a file
    if (fileType === 'xlsx') {
        xlsx.writeFile(workbook, outputFilePath);
    } else if (fileType === 'csv') {
        const csvData = xlsx.utils.sheet_to_csv(sheet);
        fs.writeFileSync(outputFilePath, csvData);
    }
    console.log(`Conversion from JSON to ${fileType.toUpperCase()} successful!`);
};

// Example: Convert JSON to XLSX
const jsonDataToXLSX = [
    { Name: 'John Doe', Age: 30, City: 'New York' },
    { Name: 'Jane Doe', Age: 25, City: 'San Francisco' },
];

convertToXlsxCsv(jsonDataToXLSX, 'output.xlsx', 'xlsx');

// Example: Convert JSON to CSV
const jsonDataToCSV = [
    { Name: 'Alice', Age: 28, City: 'London' },
    { Name: 'Bob', Age: 32, City: 'Berlin' },
];

convertToXlsxCsv(jsonDataToCSV, 'output.csv', 'csv');
convertExcelFileToJsonUsingXlsx();