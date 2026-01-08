exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  try {
    const workbook = new (require('exceljs').Workbook());
    
    // SIMPLIFIED: Assume file in event.body (base64)
    const buffer = Buffer.from(event.body, 'base64');
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.getWorksheet('CRE');
    if (!worksheet) {
      return { statusCode: 400, body: 'No CRE sheet found' };
    }

    // Add headers
    worksheet.getCell('AM1').value = 'Carpet Area sq.mt';
    worksheet.getCell('AN1').value = 'Carpet Area sq.ft';

    // Process rows (simplified regex)
    const lastRow = worksheet.rowCount;
    for (let i = 2; i <= lastRow; i++) {
      const text = worksheet.getCell('U' + i).value?.toString() || '';
      const match = text.match(/(\d+\.?\d*)\s*चौ\.मी\./g);
      
      if (match) {
        const total = match.reduce((sum, m) => sum + parseFloat(m), 0);
        worksheet.getCell('AM' + i).value = total.toFixed(2);
        worksheet.getCell('AN' + i).value = (total * 10.764).toFixed(2);
      }
    }

    const bufferOut = await workbook.xlsx.writeBuffer();
    
    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="output.xlsx"',
        'Access-Control-Allow-Origin': '*'
      },
      body: bufferOut.toString('base64'),
      isBase64Encoded: true
    };
  } catch (error) {
    console.error('ERROR:', error);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: error.message })
    };
  }
};
