const list = [
  {
    orderNum: "A309012",
    menu: "햄버거",
    price: 12000,
    date: "2023-05-01",
  },
  {
    orderNum: "B882175",
    menu: "아메리카노(ice)",
    price: 1900,
    date: "2023-05-17",
  },
  {
    orderNum: "B677919",
    menu: "떡볶이",
    price: 6000,
    date: "2023-05-28",
  },
  {
    orderNum: "A001092",
    menu: "마라탕",
    price: 28000,
    date: "2023-06-12",
  },
  {
    orderNum: "A776511",
    menu: "후라이드치킨",
    price: 18000,
    date: "2023-06-12",
  },
  {
    orderNum: "A256512",
    menu: "고급사시미",
    price: 289900,
    date: "2023-06-12",
  },
  {
    orderNum: "C114477",
    menu: "단체도시락",
    price: 1000000,
    date: "2023-06-19",
  },
];

const headers = ["주문번호", "메뉴", "가격", "주문날짜"];
const headerWidths = [40, 16, 16, 24];

document.querySelector('#down').addEventListener('click', makeFile);
// 엑셀 파일을 생성하는 함수
function makeFile(e) {
  // ExcelJS를 이용해 새로운 Workbook(엑셀 파일)을 생성
  const workbook = new ExcelJS.Workbook();
  // Workbook에 새로운 Worksheet(엑셀 시트)를 추가. 이름은 '배달 주문 내역'
  const sheet = workbook.addWorksheet('배달 주문 내역');

  // 헤더(열 이름) 추가
  const headerRow = sheet.addRow(headers);
  // 헤더 행의 높이 설정
  headerRow.height = 30.75;
  // 각 헤더 셀에 스타일 지정 및 너비 설정
  headerRow.eachCell((cell, colNum) => {
    styleHeaderCell(cell); // 헤더 셀 스타일 지정 함수 호출
    sheet.getColumn(colNum).width = headerWidths[colNum - 1]; // 헤더 셀 너비 설정
  });

  // 데이터 추가
  list.forEach(item => {
    // 각 아이템(주문 정보)을 행으로 추가
    const row = sheet.addRow([item.orderNum, item.menu, item.price, item.date]);
    // 각 데이터 셀에 스타일 지정
    row.eachCell(styleDataCell); // 데이터 셀 스타일 지정 함수 호출
  });

  // 엑셀 파일 생성 및 다운로드
  download(workbook, '배달 주문 내역').then(r => { });
}

// 엑셀 파일을 생성하고 다운로드하는 함수
const download = async (workbook, fileName) => {
  // 작성한 Workbook을 ArrayBuffer로 변환
  const buffer = await workbook.xlsx.writeBuffer();
  // ArrayBuffer를 Blob 객체로 변환
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  // Blob 객체를 URL로 변환
  const url = window.URL.createObjectURL(blob);
  // 'a' 태그 생성 및 href에 URL 설정
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName + '.xlsx'; // 다운로드 될 파일명 설정
  anchor.click(); // 'a' 태그 클릭 이벤트 발생
  window.URL.revokeObjectURL(url); // Blob URL 해제
};

// 헤더 셀에 스타일을 지정하는 함수
const styleHeaderCell = (cell) => {
  // 셀 스타일 지정
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "ffebebeb" } };
  cell.border = { bottom: { style: "thin", color: { argb: "-100000f" } }, right: { style: "thin", color: { argb: "-100000f" } } };
  cell.font = { name: "Arial", size: 12, bold: true, color: { argb: "ff252525" } };
  cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
};

// 데이터 셀에 스타일을 지정하는 함수
const styleDataCell = (cell) => {
  // 셀 스타일 지정
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "ffffffff" } };
  cell.border = { bottom: { style: "thin", color: { argb: "-100000f" } }, right: { style: "thin", color: { argb: "-100000f" } } };
  cell.font = { name: "Arial", size: 10, color: { argb: "ff252525" } };
  cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
};
