
function main(workbook: ExcelScript.Workbook) {
    // Retrieve the active worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Read all data from the worksheet.
    let range = sheet.getUsedRange();
    let values: (string | number)[][] = range.getValues(); // Specify the type of cell values

    // First row and first column as headers
    let headersRow: string[] = values[0].slice(1) as string[]; // First row without first column
    let headersCol: string[] = values.map(row => row[0]).slice(1) as string[]; // First column without first row

    // Create an object for the JSON structure
    let jsonData: { [key: string]: { [key: string]: string | number } } = {};

    // Scroll through the data and combine the row and column headers
    for (let i = 1; i < values.length; i++) { // Rows (without header)
        let rowHeader = headersCol[i - 1];
        jsonData[rowHeader] = {};

        for (let j = 1; j < values[i].length; j++) { // Columns (without first column)
            let colHeader = headersRow[j - 1];
            jsonData[rowHeader][colHeader] = values[i][j];
        }
    }

    // Convert the object into a JSON string
    let jsonString = JSON.stringify(jsonData, null, 2);

    // Output the JSON string to the console (or save it to a file)
    console.log(jsonString);
}
