const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const app = express();
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// 날짜 형식을 yyyy/mm/dd로 변환하는 함수
function formatDate(dateString) {
    const date = new Date(dateString);
    if (isNaN(date)) return null; // 유효하지 않은 날짜는 null로 반환합니다.
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}/${mm}/${dd}`;
}

// 텍스트가 한글 세 글자인지 확인하는 함수
function isThreeKoreanChars(text) {
    const koreanCharRegex = /^[가-힣]{3}$/;
    return koreanCharRegex.test(text);
}

// 헤더 일치 여부를 확인하는 미들웨어
function checkHeaders(req, res, next) {
    const file = req.file;
    const bankType = req.body.bankType;

    if (!file) {
        return res.status(400).json({ error: '파일이 제공되지 않았습니다.' });
    }

    try {
        const workbook = xlsx.readFile(file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        let headers;
        let expectedHeaders;

        if (bankType === '농협개인') {
            headers = data[7]; // 8번째 행이 헤더입니다.
            expectedHeaders = ['순번', '거래일시', '출금금액', '입금금액', '거래후잔액', '거래내용', '거래기록사항', '거래점', '거래메모'];
        } else if (bankType === '농협기업') {
            headers = data[9]; // 10번째 행이 헤더입니다.
            expectedHeaders = ['구분', '거래일자', '출금금액(원)', '입금금액(원)', '거래 후 잔액(원)', '거래내용', '거래기록사항', '거래점', '거래시간', '이체메모'];
        } else {
            fs.unlinkSync(file.path); // 파일을 삭제합니다.
            return res.status(400).json({ error: '지원되지 않는 은행 유형입니다.' });
        }

        // 헤더 일치 여부 확인
        const isHeaderMatching = expectedHeaders.every(header => headers.includes(header));
        if (!isHeaderMatching) {
            fs.unlinkSync(file.path); // 파일을 삭제합니다.
            return res.status(400).json({ error: '업로드한 파일이 해당 은행의 양식과 일치하지 않습니다. 다시 확인해 주십시오.' });
        }

        // 미들웨어 통과 후 다음 단계로 이동
        next();
    } catch (error) {
        console.error('Error processing file:', error);
        fs.unlinkSync(file.path); // 파일을 삭제합니다.
        return res.status(500).json({ error: '파일 처리 중 오류가 발생했습니다.' });
    }
}

app.post('/upload', upload.single('file'), checkHeaders, async (req, res) => {
    const file = req.file;
    const bankType = req.body.bankType;

    try {
        const workbook = xlsx.readFile(file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        let headers, mainData;

        if (bankType === '농협개인') {
            headers = data[7];
            mainData = data.slice(8);
        } else if (bankType === '농협기업') {
            headers = data[9];
            mainData = data.slice(10);
        }

        // 데이터를 표준 형식으로 변환합니다.
        const standardizedData = mainData.map(row => {
            const 거래처 = row[headers.indexOf('거래기록사항')];
            let 상세내역 = '';
            if (!isThreeKoreanChars(거래처)) {
                상세내역 = 거래처;
            }
            const 날짜 = formatDate(row[headers.indexOf('거래일시')] || row[headers.indexOf('거래일자')]);
            const 수입 = row[headers.indexOf('입금금액')] || row[headers.indexOf('입금금액(원)')];
            const 지출 = row[headers.indexOf('출금금액')] || row[headers.indexOf('출금금액(원)')];
            
            // 유효한 데이터만 반환합니다.
            if (날짜 && (수입 || 지출)) {
                return {
                    날짜: 날짜,
                    상세내역: 상세내역,
                    거래처: isThreeKoreanChars(거래처) ? 거래처 : '',
                    항목분류: '',
                    수입: 수입,
                    지출: 지출
                };
            }
        }).filter(row => row); // 유효하지 않은 행을 필터링합니다.

        // 데이터를 날짜 순으로 정렬합니다.
        standardizedData.sort((a, b) => new Date(a.날짜) - new Date(b.날짜));

        // 새로운 엑셀 파일을 만듭니다.
        const newWorkbook = new ExcelJS.Workbook();
        const newSheet = newWorkbook.addWorksheet('StandardizedData');

        // 열 헤더 추가
        newSheet.columns = [
            { header: '날짜', key: '날짜', width: 15 },
            { header: '상세내역', key: '상세내역', width: 30 },
            { header: '거래처', key: '거래처', width: 15 },
            { header: '항목분류', key: '항목분류', width: 15 },
            { header: '수입', key: '수입', width: 15 },
            { header: '지출', key: '지출', width: 15 }
        ];

        // 데이터 추가
        standardizedData.forEach(row => newSheet.addRow(row));

        // 스타일 설정
        newSheet.eachRow({ includeEmpty: false }, row => {
            row.font = { size: 11 };
            row.eachCell({ includeEmpty: false }, cell => {
                if (cell.address[0] === 'A') { // 날짜 열
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                }
            });
        });

        const outputPath = path.join(__dirname, 'uploads', 'standardized.xlsx');
        await newWorkbook.xlsx.writeFile(outputPath);

        res.download(outputPath, 'standardized.xlsx', () => {
            fs.unlinkSync(file.path); // 원본 업로드 파일을 삭제합니다.
            fs.unlinkSync(outputPath); // 생성된 엑셀 파일을 삭제합니다.
        });
    } catch (error) {
        console.error('Error processing file:', error);
        fs.unlinkSync(file.path); // 파일을 삭제합니다.
        res.status(500).json({ error: '파일 처리 중 오류가 발생했습니다.' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
