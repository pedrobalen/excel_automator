const XLSX = require("xlsx");

const workbook = XLSX.readFile("test.xlsx");

const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const range = XLSX.utils.decode_range(sheet["!ref"]);
const startRow = range.e.r + 1;

const newData = [
    ["Robert", 35, "Sample City 2"],
];

function appendDataToSheet(sheet, startRow, newData) {
        newData.forEach((row, index) => {
                row.forEach((cellValue, colIndex) => {
                        const cellAddress = XLSX.utils.encode_cell({
                                r: startRow + index,
                                c: colIndex,
                        });
                        sheet[cellAddress] = { v: cellValue };
                });
        });
}

appendDataToSheet(sheet, startRow, newData);

range.e.r = startRow + newData.length - 1;
sheet["!ref"] = XLSX.utils.encode_range(range);

XLSX.writeFile(workbook, "test.xlsx");
