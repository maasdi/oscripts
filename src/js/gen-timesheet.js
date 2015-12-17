var _ = require('underscore');

var Excel = require('exceljs');
var Workbook = Excel.Workbook;
var WorkbookWriter = Excel.stream.xlsx.WorkbookWriter;

var filename = __dirname + '/../res/CV-template-res.xlsx';

var fonts = {
    header: { name: "Arial", family: 2, size: 12, bold: true }
};

var wb = new Workbook();
wb.creator = "Maas";
wb.lastModifiedBy = "Maas";
wb.created = new Date(1985, 8, 30);
wb.modified = new Date();

var ws = wb.addWorksheet("01-15 of the month");

ws.getColumn(1).width = 2;
ws.getColumn(2).width = 25.29;
_.each([3,4,5,6,7,8,9,10,11,12,13,14,15,16,17], function(colno) {
   ws.getColumn(colno).width = 3.14; 
});
ws.getColumn(18).width = 10.14;

ws.getCell('B1').value = 'TIMESHEET';
ws.getCell('B1').font = fonts.header;

ws.getCell('B3').value =  'Name : Maas Dianto';
ws.getCell('B4').value =  'Period: November 2015';

ws.getCell('B6').value = 'Project Name (ID)';
ws.getCell('B6').fill = {
    type: "pattern",
    pattern:"darkTrellis",
    bgColor:{argb:"FFFFFF00"}
};

ws.getCell('C6').value = '1';
ws.getCell('D6').value = '2';
ws.getCell('E6').value = '3';
ws.getCell('F6').value = '4';
ws.getCell('G6').value = '5';
ws.getCell('H6').value = '6';
ws.getCell('I6').value = '7';
ws.getCell('J6').value = '8';
ws.getCell('K6').value = '9';
ws.getCell('L6').value = '10';
ws.getCell('M6').value = '11';
ws.getCell('N6').value = '12';
ws.getCell('O6').value = '13';
ws.getCell('P6').value = '14';
ws.getCell('Q6').value = '15';
ws.getCell('R6').value = 'Total Hours';


wb.xlsx.writeFile(filename)
    .then(function() {
        console.log("Done..!");
    })
    .catch(function(error) {
        console.log(error.message);
    });
