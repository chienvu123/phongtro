const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const formatNumber = (number, numDefault = 0) => {
  if (!number) return '0';
  let value = number.toString();

  value = value.replace(/(\d)(?=(\d{3})+$)/g, '$1,');

  return value;
};

// === HÀM CHÍNH === //
async function generateFromTemplate() {
  // XÓA THƯ MỤC OUTPUT CŨ
  const outputDir = path.resolve(__dirname, 'output');
  if (fs.existsSync(outputDir)) {
    fs.rmSync(outputDir, {recursive: true, force: true});
    console.log('🧹 Đã xóa thư mục output cũ');
  }
  fs.mkdirSync(outputDir);

  // Đọc template Word

  const content = fs.readFileSync(
    path.resolve(__dirname, 'template.docx'),
    'binary',
  );
  //   console.log(content)

  // Đọc Excel
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('data.xlsx');
  const sheet = workbook.worksheets[2];
  const existMaKH = {};
  const startRow = 3;
  const date = new Date();
  const month = date.getMonth() + 1;
  const year = date.getFullYear();
  let toaNha = '';
  for (let i = startRow; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);

    const maKH = row.getCell(5).value;
    if (!maKH) continue;
    if (existMaKH[maKH]) continue;
    const zip = new PizZip(content);
    // Chuẩn bị template docx
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });
    existMaKH[maKH] = true;
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

    try {
      if (maKH === 'HN01101') {
        console.log('data', data);
      }
      doc.render(data);

      const buf = doc.getZip().generate({type: 'nodebuffer'});

      const outPath = path.resolve(__dirname, 'output', `${maKH}.docx`);
      fs.writeFileSync(outPath, buf);
      //   console.log(`✅ Tạo file: ${maKH}`);
    } catch (error) {
      console.error(`❌ Lỗi khi tạo file cho ${maKH}:`, error);
    }
  }
}

if (!fs.existsSync('./output')) fs.mkdirSync('./output');
generateFromTemplate();
