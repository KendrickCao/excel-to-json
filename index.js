const ExcelJS = require('exceljs');
const fs = require('fs');

const workbook = new ExcelJS.Workbook();
const fileName = process.argv[2];
const excludeSame = process.argv.includes('-excludeSame',2);
workbook.xlsx.readFile(fileName)
    .then(() => {
        // Assuming the data is in the first worksheet
        const worksheet = workbook.getWorksheet(1);

        // Create an object to store the JSON data
        const jsonData = {};

        // Iterate through each row in the worksheet
        worksheet.eachRow((row, rowNumber) => {
            // Skip the first row which contains the column headers
            if (rowNumber === 1) {
                return;
            }
            // Assuming the first column contains the keys, and the find the value in the rest columns
            const key = row.getCell(1).value;
            for (let i = 2; i <= worksheet.columnCount; i++) {
                const value = row.getCell(i).value;
                if (value && value !== '') {
                    if(!excludeSame || key !== value){
                        jsonData[key] = value;                    
                        break;
                    }
                }
            }
        });

        // Convert the JSON object to a string
        const jsonString = JSON.stringify(jsonData, null, 2);

        // Save the JSON string to a file
        fs.writeFileSync(`${fileName}.json`, jsonString);

        console.log(`Conversion completed. JSON file saved as ${fileName}.json`);
    })
    .catch(error => {
        console.error('Error reading the Excel file:', error.message);
    });
