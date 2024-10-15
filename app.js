const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const xlsx = require('xlsx');

const PORT = 3000;

function extractExcelData(filePath, sheetName, columns, startRow, endRow) {
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet) {
        return { error: `Sheet "${sheetName}" not found.` };
    }

    const columnArray = columns.split(',');
    const start = parseInt(startRow);
    const end = parseInt(endRow);

    const filteredData = [];
    for (let row = start; row <= end; row++) {
        const rowData = columnArray.map(column => {
            const cellAddress = `${column.trim()}${row}`;
            const cell = worksheet[cellAddress];
            return cell ? cell.v : null;
        });
        filteredData.push(rowData);
    }

    return { data: filteredData };
}

http.createServer((req, res) => {
    const parsedUrl = url.parse(req.url, true);
    const method = req.method;
    const pathname = parsedUrl.pathname;

    if (method === 'GET' && pathname === '/') {
        // Serve the HTML file
        const htmlPath = path.join(__dirname, 'index.html');
        fs.readFile(htmlPath, (err, content) => {
            if (err) {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
                res.end('Error loading page');
            } else {
                res.writeHead(200, { 'Content-Type': 'text/html' });
                res.end(content);
            }
        });
    } else if (method === 'POST' && pathname === '/upload') {
        // Handle file upload and process Excel data
        let body = '';
        req.on('data', chunk => {
            body += chunk.toString();
        });

        req.on('end', () => {
            const params = new URLSearchParams(body);
            const sheetName = params.get('sheetName');
            const columns = params.get('columns');
            const startRow = params.get('startRow');
            const endRow = params.get('endRow');

            // Using a static file path for simplicity; modify as needed.
            const filePath = path.join(__dirname, 'sheet.xlsx');
            const result = extractExcelData(filePath, sheetName, columns, startRow, endRow);

            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(result));
        });
    } else {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('Not Found');
    }
}).listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});
