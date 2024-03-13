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



convertExcelFileToJsonUsingXlsx();
