
import fs from "fs"
import path from 'path';
import ExcelJS from 'exceljs';

const formatNumber = (number, numDefault = 0) => {
  if (!number) return '0';
  let value = number.toString();

  value = value.replace(/(\d)(?=(\d{3})+$)/g, '$1,');

  return value;
};


let toaNha = '';

export function reset() {
  toaNha = ''
  // XÓA THƯ MỤC OUTPUT CŨ
  const outputDir = path.resolve('./output');
  if (fs.existsSync(outputDir)) {
    fs.rmSync(outputDir, { recursive: true, force: true });
    console.log('🧹 Đã xóa thư mục output cũ');
  }
  fs.mkdirSync(outputDir);

}

export function convertDataFromRow(maKH, row) {
  const date = new Date();
  const month = date.getMonth() + 1 < 10 ? `0${date.getMonth() + 1}` : date.getMonth() + 1;
  const year = date.getFullYear();

  if (row.getCell(1).value) {
    toaNha = row.getCell(1).value;
  }
  // Lấy dữ liệu từ Excel
  const data = {
    TOA_NHA: toaNha,
    PHONG: row.getCell(4).value,
    MA_KH: maKH,
    DIEN_DAU: formatNumber(row.getCell(6).value, ''),
    DIEN_CUOI: formatNumber(row.getCell(7).value, ''),
    TONG_DIEN: formatNumber(row.getCell(8).value?.result),

    NUOC_DAU: formatNumber(row.getCell(10).value, ''),
    NUOC_CUOI:
      row.getCell(11).value === 1
        ? ''
        : formatNumber(row.getCell(11).value, ''),
    TONG_NUOC:
      row.getCell(12).value?.result === 1
        ? ''
        : formatNumber(row.getCell(12).value?.result),

    TIEN_PHONG: formatNumber(row.getCell(16).value),
    TIEN_DIEN: formatNumber(row.getCell(17).value?.result),
    TIEN_NUOC: formatNumber(row.getCell(18).value?.result),
    INTERNET: formatNumber(row.getCell(19).value),
    MAY_GIAT: formatNumber(row.getCell(20).value?.result),
    XE_DIEN: formatNumber(row.getCell(21).value),
    DICH_VU: formatNumber(row.getCell(24).value?.result),

    NO_CU: formatNumber(row.getCell(22).value?.result),
    TONG: formatNumber(row.getCell(25).value?.result),
    THANG: month,
    NAM: year,
    THANG_NAM: `${month}/${year}`,
  };
  return data;
}


export async function fillSoDienNuoc() {
  const date = new Date();
  const month = date.getMonth();

  let cellCuoi = month + 2;
  let cellDau = month + 1;
  // thang 1 => cần lấy số của tháng 12 và 11
  if (month === 0) {
    cellCuoi = 13;
    cellDau = 12;
  }
  if (month === 1) {
    cellDau = 13;
  }

  const cellDataDienDau = 6;
  const cellDataDienCuoi = 7;

  const fileDien = new ExcelJS.Workbook();
  await fileDien.xlsx.readFile('./dien.xlsx');
  const sheetDien = fileDien.worksheets[0];
  let toaNha1 = '';

  const fileData = new ExcelJS.Workbook();
  await fileData.xlsx.readFile('data.xlsx');
  const sheetData = fileData.worksheets[fileData.worksheets.length - 1];

  const startRow = 4;
  for (let i = startRow; i <= sheetDien.rowCount; i++) {
    const row = sheetDien.getRow(i);

    if (row.getCell(1).value) {
      toaNha1 = row.getCell(1).value;
    }
    const maToaNha = toaNha1.split('-')[0].trim();
    const soPhong = row.getCell(2).value.replace('P', '');
    const maKH = maToaNha + soPhong;

    for (let j = 0; j <= sheetData.rowCount; j++) {
      const rowData = sheetData.getRow(j);
      const maKHFromData = rowData.getCell(5).value;


      if (maKH === maKHFromData) {
        const dienDau = row.getCell(cellDau).value;
        const dienCuoi = row.getCell(cellCuoi).value;
        if (dienDau) {
          rowData.getCell(cellDataDienDau).value = dienDau;

        }
        if (dienCuoi) {
          rowData.getCell(cellDataDienCuoi).value = dienCuoi;
        }

        if (maKH.includes('WC')) {
          const soTang = maKH[maKH.length - 1];
          const listPhongUseWc = [];
          let e = j
          while (1) {
            --e;
            let rowPhongUseWc = sheetData.getRow(e);
            const soPhong = rowPhongUseWc.getCell(4).value;
            if (!soPhong || !soPhong.toString().startsWith(soTang)) break;

            listPhongUseWc.push(rowPhongUseWc);

            if (e < 3) break;
          }
          const tongSoNguoi = listPhongUseWc.reduce((total, _row) => total + (_row.getCell(14).value || 0), 0);
          const tongTienDien = (dienCuoi - dienDau) * rowData.getCell(9);
          if (!tongSoNguoi) break;
          const tienDienMoiNguoi = tongTienDien / tongSoNguoi;
          listPhongUseWc.forEach(_row => {
            const songuoi = _row.getCell(14).value;
            _row.getCell(22).value = Math.ceil(tienDienMoiNguoi * songuoi);
            _row.getCell(22).numFmt = '#,##0';
          })
        }

        break;
      }
    }

  }

  // Ghi đè file
  await fileData.xlsx.writeFile('data.xlsx');



}