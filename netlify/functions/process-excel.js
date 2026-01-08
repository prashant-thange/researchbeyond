const Busboy = require('busboy');

// Fix: Use new ExcelJS.Workbook()
const ExcelJS = require('exceljs');

function parseMultipartForm(event) {
  return new Promise((resolve, reject) => {
    const busboy = Busboy({
      headers: {
        'content-type': event.headers['content-type'] || event.headers['Content-Type']
      }
    });

    const result
