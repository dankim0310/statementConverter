const fs = require('fs');
const path = require('path');
const csvParser = require('csv-parser');
const chardet = require('chardet');
const iconv = require('iconv-lite');
const ExcelJS = require('exceljs');
const { JSDOM } = require('jsdom');

// HTML 파일을 Excel로 변환하는 함수
async function convertHtmlToExcel(htmlContent, outputPath) {
    const dom = new JSDOM(htmlContent);
    const document = dom.window.document;
    const table = document.querySelector('table');

    if (!table) {
        throw new Error('HTML 파일에서 테이블을 찾을 수 없습니다.');
    }

    const newWorkbook = new ExcelJS.Workbook();
    const newSheet = newWorkbook.addWorksheet('Sheet1');

    const rows = Array.from(table.querySelectorAll('tr'));
    rows.forEach((row) => {
        const cells = Array.from(row.querySelectorAll('td, th'));
        const rowData = cells.map(cell => cell.textContent.trim());
        newSheet.addRow(rowData);
    });

    await newWorkbook.xlsx.writeFile(outputPath);
}

// 파일을 읽고 첫 10행을 출력하는 함수 (CSV)
function readCsvFile(filePath) {
    const encoding = chardet.detectFileSync(filePath);
    console.log('Detected encoding (CSV):', encoding);
    
    const results = [];
    fs.createReadStream(filePath)
        .pipe(iconv.decodeStream(encoding))
        .pipe(csvParser())
        .on('data', (data) => {
            if (results.length < 10) {
                results.push(data);
            }
        })
        .on('end', () => {
            console.log('CSV 파일 첫 10개 행:', results);
        });
}

// 파일을 읽고 첫 10행을 출력하는 함수 (HTML)
async function readHtmlFile(filePath) {
    const htmlContent = fs.readFileSync(filePath, 'utf-8');
    const tempExcelPath = path.join(__dirname, 'temp.xlsx');
    
    await convertHtmlToExcel(htmlContent, tempExcelPath);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempExcelPath);
    const sheet = workbook.getWorksheet(1);

    const data = [];
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 10) {
            data.push(row.values.filter(cell => cell !== null && cell !== ''));
        }
    });

    console.log('HTML 파일 첫 10개 행:', data);

    // 임시 파일 삭제
    fs.unlinkSync(tempExcelPath);
}

// 파일이 HTML 형식인지 확인하는 함수
function isHtmlFile(filePath) {
    const content = fs.readFileSync(filePath, 'utf-8');
    return content.includes('<html') || content.includes('<table');
}

// 메인 함수
(async () => {
    const csvFilePath = path.join('C:', 'Users', 'danki', 'Downloads', '화명센터 통장거래내역.CSV');
    const htmlFilePath = path.join('C:', 'Users', 'danki', 'Downloads', 'KB_거래내역조회(699601-04-199864_20240708072641).xls');

    try {
        // CSV 파일 처리
        console.log('CSV 파일 처리 시작');
        if (!isHtmlFile(csvFilePath)) {
            console.log('일반 CSV 파일로 감지되었습니다.');
            readCsvFile(csvFilePath);
        } else {
            console.log('CSV 파일이 HTML 형식으로 감지되었습니다. 확인이 필요합니다.');
        }

        // HTML 파일 처리
        console.log('HTML 파일 처리 시작');
        if (isHtmlFile(htmlFilePath)) {
            console.log('HTML 파일로 감지되었습니다.');
            await readHtmlFile(htmlFilePath);
        } else {
            console.log('HTML 파일이 올바르지 않습니다. 확인이 필요합니다.');
        }
    } catch (error) {
        console.error('Error processing files:', error);
    }
})();
