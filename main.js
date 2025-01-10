const XLSX = require("xlsx");

const workbook = XLSX.readFile("test.xlsx");

const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

const range = XLSX.utils.decode_range(sheet["!ref"]);
const startRow = range.e.r + 1;
const newData = [["pedro", 77, "Marte"]];

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

range.e.r = startRow + newData.length - 1;
sheet["!ref"] = XLSX.utils.encode_range(range);

function updateRow(sheet, rowIndex, rowData) {
  rowData.forEach((cellValue, colIndex) => {
    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    sheet[cellAddress] = { v: cellValue };
  });

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  if (rowIndex > range.e.r) {
    range.e.r = rowIndex;
  }
  if (rowData.length - 1 > range.e.c) {
    range.e.c = rowData.length - 1;
  }
  sheet["!ref"] = XLSX.utils.encode_range(range);
}

function saveWorkbook(workbook, filename) {
  XLSX.writeFile(workbook, filename);
}

//appendDataToSheet(sheet, startRow, newData);
updateRow(sheet, 3, ["Pedro", 88, "Plutao"]);
saveWorkbook(workbook, "test.xlsx");
