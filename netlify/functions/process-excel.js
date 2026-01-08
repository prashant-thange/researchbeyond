const Busboy = require('busboy');
const ExcelJS = require('exceljs');

function parseMultipartForm(event) {
  return new Promise((resolve, reject) => {
    const busboy = Busboy({
      headers: {
        'content-type': event.headers['content-type'] || event.headers['Content-Type']
      }
    });

    const result = {
      fields: {},
      files: {}
    };

    busboy.on('file', (fieldname, file, filename, encoding, mimetype) => {
      const chunks = [];
      file.on('data', (data) => chunks.push(data));
      file.on('end', () => {
        result.files[fieldname] = {
          filename,
          encoding,
          mimetype,
          buffer: Buffer.concat(chunks)
        };
      });
    });

    busboy.on('field', (fieldname, value) => {
      result.fields[fieldname] = value;
    });

    busboy.on('error', reject);
    busboy.on('finish', () => resolve(result));

    busboy.end(Buffer.from(event.body, event.isBase64Encoded ? 'base64' : 'utf8'));
  });
}

// extract and sum all <number> चौ.मी.
function extractCarpetAreaSqMt(text) {
  if (!text) return null;

  const str = typeof text === 'string' ? text : (text.richText ? text.richText.map(r => r.text).join('') : String(text));
  const regex = /(\d+\.?\d*)\s*चौ\.मी\./g;
  let match;
  const values = [];

  while ((match = regex.exec(str)) !== null) {
    const num = parseFloat(match[1]);
    if (!isNaN(num)) values.push(num);
  }

  if (!values.length) return null;
  const total = values.reduce((s, v) => s + v, 0);
  return parseFloat(total.toFixed(2));
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: 'Method Not Allowed'
    };
  }

  try {
    const form = await parseMultipartForm(event);
    const file = form.files.file;
    if (!file || !file.buffer) {
      return { statusCode: 400, body: 'No file uploaded' };
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.buffer);

    const worksheet = workbook.getWorksheet('CRE');
    if (!worksheet) {
      return { statusCode: 400, body: 'Sheet "CRE" not found' };
    }

    // headers
    worksheet.getCell('AM1').value = 'Carpet Area sq.mt';
    worksheet.getCell('AN1').value = 'Carpet Area sq.ft';
    worksheet.getCell('AO1').value = 'Saleable area';
    worksheet.getCell('AP1').value = 'APR';

    const headerStyle = {
      font: { bold: true },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    ['AM1', 'AN1', 'AO1', 'AP1'].forEach((addr) => {
      Object.assign(worksheet.getCell(addr).style, headerStyle);
    });

    const lastRow = worksheet.rowCount;

    for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
      const row = worksheet.getRow(rowIndex);
      const marathiText = row.getCell('U').value;

      const totalSqMt = extractCarpetAreaSqMt(marathiText);
      if (totalSqMt !== null) {
        row.getCell('AM').value = totalSqMt;
        row.getCell('AM').numFmt = '0.00';

        const sqFt = parseFloat((totalSqMt * 10.764).toFixed(2));
        row.getCell('AN').value = sqFt;
        row.getCell('AN').numFmt = '0.00';
      }
    }

    worksheet.getColumn('AM').width = 18;
    worksheet.getColumn('AN').width = 18;
    worksheet.getColumn('AO').width = 18;
    worksheet.getColumn('AP').width = 10;

    const outBuffer = await workbook.xlsx.writeBuffer();

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="output.xlsx"',
        'Access-Control-Allow-Origin': '*'
      },
      body: outBuffer.toString('base64'),
      isBase64Encoded: true
    };
  } catch (err) {
    console.error(err);
    return {
      statusCode: 500,
      body: 'Internal Server Error'
    };
  }
};
