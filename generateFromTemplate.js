import fs from "fs"
import path from 'path';
import ExcelJS from 'exceljs';
import PizZip from 'pizzip';
import Docxtemplater from "docxtemplater"
import { convertDataFromRow, reset, fillSoDienNuoc } from './utils.js'


// === HÀM CHÍNH === //
async function generateFromTemplate() {

  reset()
  fillSoDienNuoc()
  // Đọc template Word
  const content = fs.readFileSync(
    path.resolve('./template.docx'),
    'binary',
  );
  //   console.log(content)

  // Đọc Excel
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('data.xlsx');
  const sheet = workbook.worksheets[workbook.worksheets.length - 1];
  const existMaKH = {};
  const startRow = 3;
  for (let i = startRow; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);

    const maKH = row.getCell(5).value;
    if (!maKH) continue;
    if (existMaKH[maKH]) continue;
    existMaKH[maKH] = true;
    if (maKH.includes('WC')) break;

    const data = convertDataFromRow(maKH, row);

    try {

      // Chuẩn bị template docx
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });


      doc.render(data);

      const buf = doc.getZip().generate({ type: 'nodebuffer' });

      const outPath = path.resolve('./output', `${maKH}.docx`);
      fs.writeFileSync(outPath, buf);
      //   console.log(`✅ Tạo file: ${maKH}`);
    } catch (error) {
      console.error(`❌ Lỗi khi tạo file cho ${maKH}:`, error);
    }
  }
}

if (!fs.existsSync('./output')) fs.mkdirSync('./output');
generateFromTemplate();
