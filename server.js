const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const iconv = require('iconv-lite'); // 인코딩 변환을 위한 모듈 추가
const app = express();
const upload = multer({ dest: 'uploads/' });

const bankHeaders = {
    '농협은행': [
        {
            headerRow: 9,
            expectedHeaders: ['구분', '거래일자', '출금금액(원)', '입금금액(원)', '거래 후 잔액(원)', '거래내용', '거래기록사항', '거래점', '거래시간', '이체메모'],
            mappings: {
                날짜: '거래일자',
                지출: '출금금액(원)',
                수입: '입금금액(원)',
                거래처: '거래기록사항'
            }
        },
        {
            headerRow: 17,
            expectedHeaders: ['순 번', '거래일자', '시간', '찾으신금액', '맡기신금액', '남은금액', '거래내용', '기록사항', '연동은행', '연동입금계좌번호', '거래점','은행구분'],
            mappings: {
                날짜: '거래일자',
                지출: '찾으신금액',
                수입: '맡기신금액',
                거래처: '기록사항'
            }
        }
    ],
    '국민은행': [
        {
            headerRow: 6,
            expectedHeaders: ['No', '거래일시', '보낸분/받는분', '출금액(원)', '입금액(원)', '잔액(원)', '내 통장 표시', '적요', '처리점', '구분'],
            mappings: {
                날짜: '거래일시',
                지출: '출금액(원)',
                수입: '입금액(원)',
                거래처: '보낸분/받는분',
                상세내역: '내 통장 표시'
            }
        },
        {
            headerRow: 0,
            expectedHeaders: ['거래일시', '보낸분/받는분', '출금액(원)', '입금액(원)', '잔액(원)', '내 통장 표시', '적요', '처리점', '구분'],
            mappings: {
                날짜: '거래일시',
                지출: '출금액(원)',
                수입: '입금액(원)',
                거래처: '보낸분/받는분',
                상세내역: '내 통장 표시'
            }
        }
    ],
    '기업은행': {
        headerRow: 5,
        expectedHeaders: ['No', '거래일시', '출금', '입금', '거래후 잔액', '거래내용', '송금메시지', '상대계좌번호', '상대은행', '거래구분', '수표어음금액', 'CMS코드', '상대계좌예금주명'],
        mappings: {
            날짜: '거래일시',
            지출: '출금',
            수입: '입금',
            거래처: '상대계좌예금주명',
            상세내역: '거래내용'
        }
    },
    '우리은행': {
        headerRow: 3,
        expectedHeaders: ['No.', '거래일시', '기재내용', '지급(원)', '입금(원)', '거래후 잔액(원)'],
        mappings: {
            날짜: '거래일시',
            지출: '지급(원)',
            수입: '입금(원)',
            거래처: '기재내용'
        }
    },
    '부산은행': {
        headerRow: 5, //6번째 행
        expectedHeaders: ['번호', '거래일시', '적요', '기재내용', '입금금액', '출금금액','거래후잔액','취급점','메모내용','적용이율'],
        mappings: {
            날짜: '거래일시',
            지출: '출금금액',
            수입: '입금금액',
            거래처: '기재내용'
        }
    },
    '하나은행': {
        headerRow: 6, //7번째 행
        expectedHeaders: ['거래일시', '적요', '의뢰인/수취인', '입금', '출금','거래후잔액','구분','거래점','거래특이사항'],
        mappings: {
            날짜: '거래일시',
            지출: '출금',
            수입: '입금',
            거래처: '의뢰인/수취인',
            상세내역: '적요'
        }
    },
    '경남은행': {
        headerRow: 0, //1번째 행
        expectedHeaders: ['번호', '거래일시', '거래종류', '입지구분', '출금금액','입금금액','거래후잔액','적요','취급점','메모내용'],
        mappings: {
            날짜: '거래일시',
            지출: '출금금액',
            수입: '입금금액',
            거래처: '적요'
        }
    },
    '우체국은행': {
        headerRow: 9, //10번째 행
        expectedHeaders: ['거래일시', '적요','입금액','출금액','거래후잔액','내역','거래국','메모'],
        mappings: {
            날짜: '거래일시',
            지출: '출금액',
            수입: '입금액',
            거래처: '내역'
        }
    }
};

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

function formatDate(dateString) {
    const date = new Date(dateString);
    if (isNaN(date)) return null;
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}/${mm}/${dd}`;
}

function parseDate(dateString) {
    // 우체국은행 날짜 형식 처리
    const match = dateString.match(/(\d{4})\.(\d{2})\.(\d{2}) (\d{2}):(\d{2}):(\d{2})(\d{2})/);
    if (match) {
        const [_, year, month, day, hour, minute, second, millisecond] = match;
        return new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}.${millisecond}Z`);
    }
    return new Date(dateString);
}


function isThreeKoreanChars(text) {
    const koreanCharRegex = /^[가-힣]{3}$/;
    return koreanCharRegex.test(text);
}

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

        console.log('파일 첫 10개 행:', data.slice(0, 10));

        let bankInfos = bankHeaders[bankType];
        
        if (!bankInfos) {
            fs.unlinkSync(file.path);
            return res.status(400).json({ error: '지원되지 않는 은행 유형입니다.' });
        }

        // bankInfos가 배열이 아니면 배열로 만듭니다.
        if (!Array.isArray(bankInfos)) {
            bankInfos = [bankInfos];
        }

        let bankInfo = null;
        for (const info of bankInfos) {
            const headers = data[info.headerRow];
            const expectedHeaders = info.expectedHeaders;

            console.log('은행 유형:', bankType);
            console.log('읽어들인 헤더:', headers);
            console.log('기대하는 헤더:', expectedHeaders);

            const isHeaderMatching = expectedHeaders.every(header => headers.includes(header));
            if (isHeaderMatching) {
                bankInfo = info;
                break;
            }
        }

        if (!bankInfo) {
            fs.unlinkSync(file.path);
            return res.status(400).json({ error: '업로드한 파일이 해당 은행의 양식과 일치하지 않습니다. 다시 확인해 주십시오.' });
        }

        req.bankInfo = bankInfo;
        next();
    } catch (error) {
        console.error('Error processing file:', error);
        fs.unlinkSync(file.path);
        return res.status(500).json({ error: '파일 처리 중 오류가 발생했습니다.' });
    }
}

app.post('/upload', upload.single('file'), checkHeaders, async (req, res) => {
    const file = req.file;
    const bankType = req.body.bankType;
    const bankInfo = req.bankInfo;

    try {
        const buffer = fs.readFileSync(file.path);
        let workbook;
        if (bankType === '국민은행' && bankInfo.headerRow === 0) {
            const decodedBuffer = iconv.decode(buffer, 'euc-kr'); // euc-kr 인코딩으로 변환
            workbook = xlsx.read(decodedBuffer, { type: 'string' });
        } else {
            workbook = xlsx.read(buffer, { type: 'buffer' });
        }

        const sheetNames = workbook.SheetNames;
        const newWorkbook = new ExcelJS.Workbook();

        for (const sheetName of sheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

            console.log('data:', data);

            let headers = data[bankInfo.headerRow];
            let mainData = data.slice(bankInfo.headerRow + 1);

            console.log('mainData:', mainData);

            const standardizedData = mainData.map((row, index) => {
                try {
                    let 거래처 = row[headers.indexOf(bankInfo.mappings['거래처'])];
                    let 상세내역 = row[headers.indexOf(bankInfo.mappings['상세내역'])] || ''; // 상세내역에 매핑된 값
            
                    if (!isThreeKoreanChars(거래처) && !상세내역) {
                        상세내역 = 거래처;
                        거래처 = '';
                    } else if (!isThreeKoreanChars(거래처) && 상세내역) {
                        상세내역 = 상세내역;
                        거래처 = 거래처;
                    }
            
                    // 여기에서 parseDate 함수를 사용합니다
                    const 날짜 = (bankType === '우체국은행') ? parseDate(row[headers.indexOf(bankInfo.mappings['날짜'])]) : new Date(row[headers.indexOf(bankInfo.mappings['날짜'])]);
                    if (isNaN(날짜)) {
                        console.log(`Invalid date at row ${index}:`, row);
                        return null; // 날짜 값이 없는 행은 제외
                    }
                    const formattedDate = formatDate(날짜.toISOString());
                    const 수입 = row[headers.indexOf(bankInfo.mappings['수입'])];
                    const 지출 = row[headers.indexOf(bankInfo.mappings['지출'])];
            
                    if (formattedDate && (수입 || 지출)) {
                        return {
                            날짜: formattedDate,
                            상세내역: 상세내역,
                            거래처: isThreeKoreanChars(거래처) ? 거래처 : (!상세내역 ? '' : 거래처),
                            항목분류: '',
                            수입: 수입,
                            지출: 지출,
                            원본날짜: 날짜
                        };
                    } else {
                        console.log(`Invalid data at row ${index}:`, row);
                    }
                } catch (error) {
                    console.error(`Error processing row ${index}:`, row, error);
                    return null;
                }
            }).filter(row => row !== null);
            
            
            console.log('standardizedData:', standardizedData);
            

            standardizedData.sort((a, b) => a.원본날짜 - b.원본날짜);

            const newSheet = newWorkbook.addWorksheet(sheetName);
            newSheet.columns = [
                { header: '날짜', key: '날짜', width: 12 },
                { header: '상세내역', key: '상세내역', width: 24 },
                { header: '거래처', key: '거래처', width: 16 },
                { header: '항목분류', key: '항목분류', width: 12 },
                { header: '수입', key: '수입', width: 12 },
                { header: '지출', key: '지출', width: 12 }
            ];

            standardizedData.forEach(row => {
                delete row.원본날짜;
                newSheet.addRow(row);
                console.log('Added row:', row);
            });

            newSheet.eachRow({ includeEmpty: false }, row => {
                row.font = { size: 11 };
                row.eachCell({ includeEmpty: false }, cell => {
                    if (cell.address[0] === 'A') {
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                    }
                });
            });
        }

        const outputPath = path.join(__dirname, 'uploads', 'standardized.xlsx');
        await newWorkbook.xlsx.writeFile(outputPath);

        res.download(outputPath, 'standardized.xlsx', () => {
            fs.unlinkSync(file.path);
            fs.unlinkSync(outputPath);
        });
    } catch (error) {
        console.error('Error processing file:', error);
        fs.unlinkSync(file.path);
        res.status(500).json({ error: '파일 처리 중 오류가 발생했습니다.' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
