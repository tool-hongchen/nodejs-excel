var XLSX = require('xlsx');
var fs = require("fs")
var pathAndroid = '../file/android.xlsx';
var pathIos = '../file/ios.xlsx';
var pathAndroidJS = '../file/questionAndriod.js';
var pathIosJS = '../file/questionIOS.js';
const workbook = XLSX.readFile(pathAndroid);
const sheetNames = workbook.SheetNames;
const worksheet = workbook.Sheets[sheetNames[0]];
const data = [];
const keys = Object.keys(worksheet);
const textTypeObj = {
  'url': 'url',//跳转链接 文字 （目标页手动配置）
  'link': 'link',//可复制的链接
  'img': 'img'//img
};
let oldRow = 1, oldcol = 'A';
let title = '';
let classificationId = '';
let content = [];
let obj = {};
// console.log(JSON.stringify(worksheet))
// console.log(worksheet)
keys
// 过滤以 ! 开头的 key
  .filter(k => k[0] !== '!')
  // 遍历所有单元格
  .forEach(k => {
    // 如 A11 中的 A
    // let col = k.substring(0, 1);
    // 如 A11 中的 11
    // let row = parseInt(k.substring(1));
    //前两行是标题不需要记录

    let i = 0;
    // console.log(isNaN(k.substring(i)))
    while (isNaN(k.substring(i))) {
      i++;
    }

    // 如 A11 中的 A
    let col = k.substring(0, i);
    // 如 A11 中的 11
    let row = parseInt(k.substring(i));


    if (row <= 1) {
      return;
    }
    if (row != oldRow) {
      if (obj.id) {
        data.push(obj);
        content = [];
        obj = {};
        title = '';
        classificationId = '';
      }
    }
    // 当前单元格的值
    let value = worksheet[k].v;
    // 保存字段名
    // if (row === 1) {
    //   headers[col] = value;
    //   return;
    // }
    //根据前后差值 遍历 添加空白行
    if (k == `B${row}`) {
      //第二列代表 问题标题
      title = value;
      oldcol = col;
      return false
    }
    if (k == `A${row}`) {
      //第一列代表 所属类型
      classificationId = value.split('_')[1];
      oldcol = col;
      return false
    }

    let spaceObj = {
      type: 'text',
      content: ''
    }
    const diffNum = col.charCodeAt() - oldcol.charCodeAt();
    for (let i = 0; i < diffNum - 1; i++) {
      content.push(spaceObj);
    }

    spaceObj = {
      type: textTypeObj[value.split('_')[0]] || 'text',
      content: textTypeObj[value.split('_')[0]] ? value.slice(value.indexOf('_') + 1) : value
    }
    content.push(spaceObj);


    obj = {
      id: row - 1,
      title: title,
      classificationId: classificationId,
      content: content
    }
    oldRow = row;
    oldcol = col;
  });
data.push(obj);
let androidQuestion = `${JSON.stringify(data)}`;
fs.writeFile(pathAndroidJS, androidQuestion, function (err) {
  if (err) {
    return console.error(err);
  }
});