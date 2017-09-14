const Excel = require('exceljs');
let filename = './acao_today_works_18000.xlsx';
var workbook = new Excel.Workbook();
workbook.xlsx.readFile(filename)
    .then(function() {
        let worksheet = workbook.getWorksheet(1);

        for(let i = 1; i <= worksheet.rowCount; i++) {
          let row = worksheet.getRow(i);
          let value = row.getCell(1).value.toString();
          value = value.trim();
          value = value.replace(/\s/g, '');
          row.getCell(1).value = value;
          row.commit();
          
        }
        
        for(let i = 1; i <= worksheet.rowCount; i++) {
          let row = worksheet.getRow(i);
          let value = row.getCell(1).value.toString();
          let length = row.getCell(1).value.toString().length;

          let arr = [70, 77, 78, 79];

           if(length <= 7) {
            //row.splice(row.actualCellCount,1, '');
            worksheet.spliceRows(i, 1);
            row.commit();
           }

           else if(length == 12) {
             let sub = '55' + value.slice(0, -1).toString();
             row.getCell(1).value = sub;
             row.commit();
           }

           else if(length == 10) {
              let sub = value.substr(0,2) + '9' + value.substr(2);
              row.getCell(1).value = '55' + sub;
              row.commit();
          }

          else if(length == 8 && arr.indexOf(Number(value.substr(0,2))) > -1) {
              row.getCell(1).value = '55' + '11' + value;
              row.commit();
          }

          else if(length == 8 && !arr.indexOf(Number(value.substr(0,2))) > -1) {
            row.getCell(1).value = '55' + '11' + '9' + value;
            row.commit();
          }

          else if(length == 11) {
            row.getCell(1).value = '55' + value;
            row.commit();
          }
          else if(length == 9) {
            row.getCell(1).value = '55' + '11' + value;
            row.commit();
          }

          else {
              console.log(value);
          }
         
        }
        return workbook.xlsx.writeFile('acao_today_works_18000_new.xlsx');
});