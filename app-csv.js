var chalk = require('chalk')

var fs = require("fs");

var tableContainer = []
var commonHeader = [
  ['上海浦东发展银行个人信用卡账户对账单'],
  ['案号：', '', '', '账户号：', '', ''],
  ['姓名：', '', '', '证件号码：', '', ''],
]
var header = []
var applyNum = 0
var fileNum = 0

fs.readFile('/Users/huangqier/Documents/测试excel.csv', function (err, data) {
    if (err) {
        console.log(err.stack);
        return;
    }
    ConvertToTable(data, function (table) {
        console.log(table);
        header = [table[0]];
        tableContainer = table.slice(1);
        handleTable(tableContainer);
    })
});

function ConvertToTable(data, callBack) {
    data = data.toString();
    let table = [];
    let rows = data.split("\r\n");
    for (var i = 0; i < rows.length; i++) {
        table.push(rows[i].split(","));
    }
    callBack(table);
}

function handleTable(table) {
  if (table.length && table[0][1]) {
    let accout = table[0][1]
    let name = table[0][2]
    let id = table[0][3]
    commonHeader[1][4] = accout
    commonHeader[2][1] = name
    commonHeader[2][4] = id
    let breakPoint = 0
    for (let i = 0; i < table.length; i++) {
      table[i][0] = i + 1
      if (!table[i] || table[i][1] !== accout) {
        breakPoint = i;
        break;
      }
    }
    let output = [...commonHeader, ...header, ...table.slice(0, breakPoint || undefined)]
    for (let i = 0; i < output.length; i++) {
      output[i] = output[i].map(x => `\t${x}`).join(',')
    }
    output = output.join('\n')
    const folderName = '账户对账单'
    const fileName = `${name}+${accout}`
    // 生成文件夹
    if (!fs.existsSync(folderName)) {
      fs.mkdir(folderName, (err) => {
        if (err) console.log(err, '--->mkdir<---')
      })
    }
    // 生成csv文件
    console.log(`提交第${++applyNum}个csv文件生成任务`)
    fs.writeFile(`./${folderName}/${fileName}.csv`, output, function(err){
      if (err) {
        console.log(err, '---->csv<---')
        return
      } else {
        console.log(chalk.green(`成功生成第${++fileNum}个csv文件`))
      }
    })
    if (breakPoint === 0) {
      handleTable([])
    } else {
      table = table.slice(breakPoint)
      handleTable(table)
    }
  } else {
    console.log(chalk.green(`提交完毕，共提交${applyNum}个任务，等待异步生成csv文件...`));
  }
}