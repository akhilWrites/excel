
const ExcelJS = require('exceljs');

const express = require('express');
const app = express();
var bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json())

app.listen(3001, () => {
    console.log('listening on port 3001');
  });

app.post('/excel/download', (req, res)=> {

    const body = req.body;

    res.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.header("content-disposition", "attachment; filename=excel.xlsx")
    console.log(JSON.stringify(req.body))
    const options = {
        stream: res,
        filename: 'excel.xlsx',
        useStyles: true,
        useSharedStrings: true
      };

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
    const worksheet = workbook.addWorksheet('My Sheet');
    
    const A1 = worksheet.getCell('A1');
    const A2 = worksheet.getCell('A2');
    const A3 = worksheet.getCell('A3');
    const A4 = worksheet.getCell('A4');
    
    A1.value = body.A1;
    A2.value= body.A2;

    A3.value= { formula: "=SUM(E3:E5)" };
    A4.value= { formula: `SUM(A1 : A3)` };

    console.log(worksheet.getCell('A3').value);
    worksheet.commit();
    workbook.commit();


});

