const fs = require('fs');
const xlsx = require('node-xlsx').default;
const { genTableColumn, genTableData, getTableDataItem } = require('./table-parser');
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

// 解析新疆办事指南excel
function transformXinjiangServiceGuideExcel(parseFilePath, genFileName = 'tabledata.js') {
  const dataMap = {
    'lawBasis': [],
    'unionMatter': [],
    'applyMaterials': [],
  };
  const typeMap = {
    '法定依据': 'lawBasis',
    '联办事项': 'unionMatter',
    '申报材料': 'applyMaterials',
  };
  const types = Object.keys(typeMap);
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(parseFilePath));
  workSheetsFromBuffer.forEach((sheet) => {
    const data = sheet.data;
    if (data && data.length > 0) {
      let curType = '';
      let curTarget = null;
      let curColumns = null;
      data.forEach((row) => {
        const col_1 = row[0];
        let realRow = row;
        let isSlice = false;
        if (types.includes(col_1)) {
          curType = col_1;
          curTarget = dataMap[typeMap[col_1]] || [];
          realRow = row.slice(1);
          isSlice = true;
          curColumns = genTableColumn(realRow);
        }
        if (curType === '法定依据') {
          curTarget.push({
            name: realRow[isSlice ? 0 : 1],
            content: realRow[isSlice ? 1 : 2],
          });
        } else if (!isSlice) {
          const dataItem = getTableDataItem(curColumns);
          curColumns.forEach((col, colIndex) => {
            dataItem[col.dataIndex] = realRow[colIndex + 1] || '';
          });
          curTarget.push(dataItem);
        }
      });
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
exports.transformXinjiangServiceGuideExcel = transformXinjiangServiceGuideExcel;