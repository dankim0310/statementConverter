const fs = require('fs');
const path = require('path');
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

// 파일을 읽고 첫 10행을 출력하는 함수
async function readExcelFile(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);

    const data = [];
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 10) {
            data.push(row.values.filter(cell => cell !== null && cell !== ''));
        }
    });

    console.log('파일 첫 10개 행:', data);
}

// 파일이 HTML 형식인지 확인하는 함수
function isHtmlFile(filePath) {
    const content = fs.readFileSync(filePath, 'utf-8');
    return content.includes('<html') || content.includes('<table');
}

// 메인 함수
(async () => {
    const inputFilePath = path.join('C:', 'Users', 'danki', 'Downloads', 'KB_거래내역조회(699601-04-199864_20240708072641).xls');
    const tempExcelPath = path.join('C:', 'Users', 'danki', 'Downloads', 'temp.xlsx');

    try {
        if (isHtmlFile(inputFilePath)) {
            console.log('HTML 파일로 감지되었습니다.');

            // HTML 파일을 Excel 파일로 변환
            const htmlContent = fs.readFileSync(inputFilePath, 'utf-8');
            await convertHtmlToExcel(htmlContent, tempExcelPath);
            console.log('HTML 파일을 Excel 파일로 변환 완료:', tempExcelPath);

            // 변환된 Excel 파일을 읽고 첫 10행을 출력
            await readExcelFile(tempExcelPath);

            // 임시 파일 삭제
            fs.unlinkSync(tempExcelPath);
        } else {
            console.log('일반 Excel 파일로 감지되었습니다.');

            // 일반 Excel 파일을 읽고 첫 10행을 출력
            await readExcelFile(inputFilePath);
        }
    } catch (error) {
        console.error('Error processing file:', error);
    }
})();
