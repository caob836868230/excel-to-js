const Pinyin = require("pinyin");

function getTableColumnItem({ title, attr }) {
  return {
    title,
    dataIndex: attr,
    key: attr,
  };
}

function getTableDataItem(columns) {
  return columns.reduce((res, next) => {
    res[next.dataIndex] = '';
    return res;
  }, {});
}

function genTableColumn(headerList) {
  let columns = [];
  headerList.forEach(headerItem => {
    columns.push(getTableColumnItem({
      title: headerItem,
      attr: Pinyin(headerItem, {
        style: Pinyin.STYLE_FIRST_LETTER
      }).join(''),
    }));
  });
  return columns;
}

function genTableData(dataList, columns) {
  let data = [];
  for (let row of dataList) {
    if (row.length === 0) {
      break;
    }
    const dataItem = getTableDataItem(columns);
    row.forEach((col, colIndex) => {
      dataItem[columns[colIndex].dataIndex] = col;
    });
    data.push(dataItem);
  }
  return data;
}

module.exports.genTableColumn = genTableColumn;
module.exports.genTableData = genTableData;