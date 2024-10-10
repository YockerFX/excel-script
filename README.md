# Excel to JSON Converter

This project contains an Excel Script that converts data from an Excel file into a JSON format. The script reads the data from the first row (as column headers) and the first column (as row headers) to generate all possible combinations of values in a JSON structure.

## Features

- Converts an Excel table to a JSON format.
- Uses the first row for column headers and the first column for row headers.
- Combines row and column headers to create a nested JSON object.

## Script Usage
To use the script, follow these steps:

- Open your Excel workbook.
- Go to Automate > Code Editor.
- Paste the script into the editor.
- Run the script to convert the Excel data into JSON format.

### Example

If your Excel data looks like this:

|     | A      | B      | C      |
|-----|--------|--------|--------|
| 1   | Name   | Alter  | Stadt  |
| 2   | Max    | 25     | Berlin |
| 3   | Anna   | 30     | Hamburg |

The generated JSON will look like this:

```json
{
  "Max": {
    "Alter": 25,
    "Stadt": "Berlin"
  },
  "Anna": {
    "Alter": 30,
    "Stadt": "Hamburg"
  }
}
```
