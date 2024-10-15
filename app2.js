const xlsx = require('xlsx');
const workbook = xlsx.readFile('sheet.xlsx');

function getReelSets(sheetName, columns, startRow, endRow){
    const worksheet = workbook.Sheets[sheetName];
    const filteredData = [];
    for (let row = startRow; row <= endRow; row++) {
        const rowData = columns.map(column => {
            const cellAddress = `${column}${row}`;
            const cell = worksheet[cellAddress];
            return cell ? cell.v : null;
        });
        filteredData.push(rowData);
    }

    return filteredData;
}
////
const sheetName ="Marketing";
const columns = ['C', 'D', 'E', 'F', 'G'];
const startRow = 84;
const endRow = 154;
const bgReelSets = getReelSets(sheetName, columns, startRow, endRow);
console.log(bgReelSets);
