/** 办事指南解析 */
const fs = require('fs');
const xlsx = require('node-xlsx').default;
const beautify = require('js-beautify/js').js;
const { genTableColumn, getTableDataItem } = require('./table-parser');

// 单个主题前端配置模版
const template = {
  sceneInfo: {
    department: '', // 牵头部门
    organization: '', // 联办机构/联办部门
    supervisePhone: '12345',
    superviseWeb: '<a href="https://zwfw.xinjiang.gov.cn:9001/epoint-customerservice-rest/robot/robot?siteNo=wechat" style="color:#1677FF;">https://zwfw.xinjiang.gov.cn:9001/epoint-customerservice-rest/robot/robot?siteNo=wechat</a>',

  },
  sceneDetail: {
    decodePageConfigs: [], // 涉及的内容
    lawBasis: [],  // 法定依据
    unionMatter: [],  // 联办事项
    applyMaterials: [], // 申报材料
  },
  mattersInfo: [
    {
      projectType: '', // 办件类型
      accept_time: '', // 法定办结时限
      term: '', // 承诺办结时限
      numberOfVisits: '', // 到场次数
      onsiteReason: '', // 必须现场办理原因说明
      exerciseLevelName: '', // 行使层级
      implementingSubjectType: '', // 实施主体性质
      delegationDepartment: '', // 委托部门
      onlineHandlingDepth: '', // 网办深度（等级）
      handlingForm: '', // 办理形式
      universalRange: '', // 通办范围
      name: '', // 受理条件-事项名称
      acceptingCondition: '', // 受理条件-事项说明
    },
  ],
};

// 获取主题id
function getSceneId(dataList) {
  const row1 = dataList[0];
  const col3 = row1?.[3];
  return col3.replace('主题ID：', '');
}

const setValueMap = {
  牵头部门: (value, target) => {
    target.sceneInfo.department = value;
  },
  联办机构: (value, target) => {
    target.sceneInfo.organization = value;
  },
  办件类型: (value, target) => {
    target.mattersInfo[0].projectType = value;
  },
  法定办结时限: (value, target) => {
    target.mattersInfo[0].accept_time = value;
  },
  承诺办结时限: (value, target) => {
    target.mattersInfo[0].term = value;
  },
  到场次数: (value, target) => {
    target.mattersInfo[0].numberOfVisits = Number(value.replace('次', ''));
  },
  必须现场办理原因说明: (value, target) => {
    target.mattersInfo[0].onsiteReason = value;
  },
  行使层级: (value, target) => {
    target.mattersInfo[0].exerciseLevelName = value;
  },
  涉及的内容: (value, target) => {
    target.sceneDetail.decodePageConfigs.push({
      type: 'SJDNR',
      chart: value,
    });
  },
  实施主体性质: (value, target) => {
    target.mattersInfo[0].implementingSubjectType = value;
  },
  委托部门: (value, target) => {
    target.mattersInfo[0].delegationDepartment = value;
  },
  '网办深度（等级）': (value, target) => {
    target.mattersInfo[0].onlineHandlingDepth = value;
  },
  办理形式: (value, target) => {
    target.mattersInfo[0].handlingForm = value;
  },
  通办范围: (value, target) => {
    target.mattersInfo[0].universalRange = value;
  },
};

// 解析新疆办事指南excel
function transformXinjiangServiceGuideExcel(parseFilePath, genFileName = 'tabledata.js') {
  const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(parseFilePath));
  const data = {};
  workSheetsFromBuffer.forEach((sheet) => {
    const dataList = sheet.data;
    if (dataList && dataList.length > 0) {
      const curTarget = JSON.parse(JSON.stringify(template));
      const sceneId = getSceneId(dataList);
      // excel模版序号列索引
      let index = null;
      let curColumns;
      const shoulitiaojian = [];
      dataList.slice(2).forEach((row) => {
        const col1 = row[0];
        let isSlice = false;
        let realRow = row;
        if (typeof col1 === 'number') {
          index = col1;
          // 判断是否是申报材料、联办事项模块
          if ([3, 4, 7].includes(index)) {
            isSlice = true;
            realRow = row.slice(2);
            curColumns = genTableColumn(realRow);
          }
        }
        if (index === 1 || index === 2) {
          const fn = setValueMap[row[index].trim()];
          typeof fn === 'function' && fn(row[3], curTarget);
        } else if (!isSlice && [4, 7].includes(index)) {
          // 申报材料和联办事项
          const dataItem = getTableDataItem(curColumns);
          curColumns.forEach((col, colIndex) => {
            dataItem[col.dataIndex] = realRow[colIndex + 2] || '';
          });
          const tempTarget = index === 4 ? curTarget.sceneDetail.applyMaterials : curTarget.sceneDetail.unionMatter;
          tempTarget.push(dataItem);
        } else if (index === 5) {
          if (realRow[2] && realRow[3]) {
            // 法定依据
            curTarget.sceneDetail.lawBasis.push({
              name: realRow[2],
              content: realRow[3],
            });
          }
        } else if (!isSlice && index === 3) {
          if (realRow[2] && realRow[3]) {
            // 受理条件
            shoulitiaojian.push({
              name: realRow[2],
              acceptingCondition: realRow[3],
            });
          }
        }
      });
      shoulitiaojian.forEach((j, jdx) => {
        if (jdx === 0) {
          curTarget.mattersInfo[0].name = j.name;
          curTarget.mattersInfo[0].acceptingCondition = j.acceptingCondition;
        } else {
          curTarget.mattersInfo.push({
            name: j.name,
            acceptingCondition: j.acceptingCondition,
          });
        }
      });
      data[sceneId] = curTarget;
    }
  });
  fs.writeFile(
    `${process.cwd()}/${genFileName}`,
    beautify(`export default ${JSON.stringify(data)}`, { indent_size: 2 }),
    (err) => {
      console.log(err ? '生成失败！' : '生成成功！');
    }
  );
}

exports.transformXinjiangServiceGuideExcel = transformXinjiangServiceGuideExcel;