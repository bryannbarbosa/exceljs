const Excel = require('exceljs');
let filename = './file.xlsx';
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
          value = value.replace(/[`a-zA-Z~!@#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, '');
          row.getCell(1).value = value;
          row.commit();
          let length = row.getCell(1).value.toString().length;

          //if(value == '1184116446') {
           // console.log('found');
          //  let sub = row.getCell(1).toString;
          //  console.log('found');
         // }

          
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
            /*if(i == 15936) {
              console.log(row.getCell(1).value);
              row.getCell(1).value = '55' + sub.toString();
              console.log(row.getCell(1).value);
              //console.log('55' + sub);
            }*/
              row.getCell(1).value = '55' + sub.toString();
              //console.log('55' + sub);
           // let sub = value.substr(0,2) + '9' + value.substr(2);
            //row.getCell(1).value = '55' + sub;
          }
          if(length == 11) {
           
            row.getCell(1).value = '55' + row.getCell(1).value.toString();
          }
          if(length == 12 && !arr.indexOf(Number(value.substr(2,2))) > -1) {
            let sub = '55' + value.slice(0, -1).toString();
            row.getCell(1).value = sub;
          }

          /*if(length <= 7) {
            row.getCell(1).value = 0;
            //console.log(value);
            
          }*/
          row.commit();
        }

        for(let i = 1; i <= worksheet.rowCount; i++) {
          let row = worksheet.getRow(i);
          let value = row.getCell(1).value.toString();
          let length = row.getCell(1).value.toString().length;
          
          
          if(length <= 7) {
            worksheet.spliceRows(i, 1);
            //row.getCell(1).value = 0;
            //console.log(value);
            
          }
        }
        return workbook.xlsx.writeFile('acao_today_works_18000_new.xlsx');
});