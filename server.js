const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));

// 날짜 형식을 yyyy/mm/dd로 변환하는 함수
function formatDate(dateString) {
    const date = new Date(dateString);
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0'); // 월은 0부터 시작하므로 +1
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}/${mm}/${dd}`;
}

app.post('/upload', upload.single('file'), (req, res) => {
    const file = req.file;
    const workbook = xlsx.readFile(file.path);
    const sheetNames = workbook.SheetNames;
    const sheet = workbook.Sheets[sheetNames[0]];

    // 전체 데이터를 읽어옵니다.
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // 8번째 행의 헤더를 가져옵니다.
    const headers = data[7];

    // 9번째 행부터 데이터를 읽어옵니다.
    const mainData = data.slice(8);

    // 데이터 매핑 및 통일된 양식으로 변환
    const standardizedData = mainData.map(row => ({
        날짜: formatDate(row[headers.indexOf('거래일시')]), // 날짜 형식 변환
        상세내역: row[headers.indexOf('거래내용')],
        거래처: row[headers.indexOf('거래기록사항')],
        항목분류: '', // 항목 분류는 주어진 매핑에 없으므로 빈 문자열로 설정
        수입: row[headers.indexOf('입금금액')],
        지출: row[headers.indexOf('출금금액')]
    }));

    // 날짜 오름차순으로 정렬
    standardizedData.sort((a, b) => new Date(a.날짜) - new Date(b.날짜));

    // 새로운 워크북 및 시트 생성
    const newWorkbook = xlsx.utils.book_new();
    const newSheet = xlsx.utils.json_to_sheet(standardizedData, { header: ["날짜", "상세내역", "거래처", "항목분류", "수입", "지출"] });

    xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'StandardizedData');

    const outputPath = path.join(__dirname, 'uploads', 'standardized.xlsx');
    xlsx.writeFile(newWorkbook, outputPath);

    res.download(outputPath, 'standardized.xlsx');
});

app.listen(3000, () => {
    console.log('Server running on port 3000');
});
