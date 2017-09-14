const Excel = require('exceljs');
let filename = './acao_today_works_18000_cp.xlsx';
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
    .then(function() {
        let worksheet = workbook.getWorksheet(1);

        let arr = [70, 77, 78, 79];
        
        for(let i = 1; i <= worksheet.rowCount; i++) {
          let row = worksheet.getRow(i);
          let value = row.getCell(1).value.toString();
          value = value.trim();
          value = value.replace(/\s/g, '');
          value = value.replace(/[`~!@#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, '');
          row.getCell(1).value = value;
          row.commit();
          let length = row.getCell(1).value.toString().length;

          if(value == '1184116446') {
            console.log('found');
          }
          
          if(length <= 7) {
            //console.log(value);
            worksheet.spliceRows(i, 1);
          }
          if(length == 8 && arr.indexOf(Number(value.substr(0,2))) > -1) {
            //console.log('55' + '11' + value);
           row.getCell(1).value = '55' + '11' + value;
          }
          else if(length == 8 && !arr.indexOf(Number(value.substr(0,2))) > -1) {
            //console.log('55' + '11' + '9' + value);
            row.getCell(1).value = '55' + '11' + '9' + value;
          }

          if(length == 9) {
           // console.log('55' + '11' + value);
           row.getCell(1).value = '55' + '11' + value;
          }

          if(length == 10 && arr.indexOf(Number(value.substr(2,2))) > -1) {
            //console.log('55' + value);
            row.getCell(1).value = '55' + value;
          }

          else if(length == 10 && !arr.indexOf(Number(value.substr(2,2))) > -1) {
            //console.log('55' + sub);
            let sub = value.substr(0,2) + '9' + value.substr(2);
            row.getCell(1).value = '55' + sub;
          }
          if(length == 11) {
            //console.log('55' + value);
            row.getCell(1).value = '55' + value;
          }
          row.commit();
        }
        return workbook.xlsx.writeFile('acao_today_works_18000_new.xlsx');
});