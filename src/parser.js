const fs = require('fs');
const xlsx = require('node-xlsx').default;
const { genTableColumn, genTableData } = require('./table-parser');
const beautify = require('js-beautify/js').js;


function transformExcel(parseFilePath, genFileName = 'tabledata.js') {
  const dataMap = {};
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(parseFilePath));
  workSheetsFromBuffer.forEach((sheet) => {
    const sheetName = sheet.name;
    const data = sheet.data;
    const header = data[0];
    if (header) {
      const columns = genTableColumn(header);
      const dataSource = genTableData(data.slice(1), columns)
      dataMap[sheetName] = {
        columns,
        dataSource
      };
    }
  });
  fs.writeFile(
    `${process.cwd()}/${genFileName}`,
    beautify(`export default ${JSON.stringify(dataMap)}`, { indent_size: 2 }),
    (err) => {
      console.log(err ? '生成失败！' : '生成成功！');
    }
  );
}

exports.transformExcel = transformExcel;