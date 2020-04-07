

// 拆分表，修改以下这个文件绝对路径，在终端输入npm run dev即可启动程序
const sourcePath = '/Users/huangqier/Documents/原始资料/待处理表3.xlsx'

var xlsx = require('node-xlsx'); 
var chalk = require('chalk')

var fs = require('fs'); 
var path = require('path');

var commonHeader = [
  ['上海浦东发展银行个人信用卡账户对账单'],
  ['案号：', '', '', '账户号：', '', ''],
  ['姓名：', '', '', '证件号码：', '', ''],
]
var header = []
var tableContainer = []
var applyNum = 0
var fileNum = 0

// 读取文件 
// const sourcePath = '/Users/huangqier/Downloads/待处理表2.xlsx'
console.log(chalk.blue(`正在读取xlsx源文件...\n`))
const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(sourcePath));
console.log(chalk.green(`成功读取xlsx源文件，共有${workSheetsFromBuffer.length}个sheet，开始进入拆分处理流程\n`))

// 输出到控制台 
// console.log(workSheetsFromBuffer); 
// console.log(workSheetsFromBuffer[0].data);

// 只处理一个sheet
// const table = workSheetsFromBuffer[0].data
// handleTableSheet(table)

// 处理多个sheet
for (let i = 0; i < workSheetsFromBuffer.length; i++) {
  console.log(chalk.blue(`开始处理第${i + 1}个sheet...\n`))
  handleTableSheet(workSheetsFromBuffer[i].data)
}

function handleTableSheet (table) {
  // 删掉不需要的列
  const deleteList = [4, 7, 8, 9]
  const deleteCloumn = (row) => {
    for( let i = row.length - 1; i >= 0; i--) {
      if (deleteList.includes(i)) {
        row.splice(i, 1)
      }
    }
    return row
  }
  // 调换列的位置
  const changeMap = {
    '1': 3,
    '2': 4,
  }
  const changePosition = (row) => {
    for( let i = 0; i < row.length; i++) {
      if (changeMap[i] !== undefined) {
        const item = row[i]
        row[i] = row[changeMap[i]]
        row[changeMap[i]] = item
      }
    }
    return row
  }
  header = [changePosition(deleteCloumn(table[0]))];
  // console.log(header)
  tableContainer = table.slice(1);
  // console.log(tableContainer[0])
  handleTable(tableContainer);

  function handleTable(table) {
    if (table.length && table[0][9]) {
      let order = table[0][9]
      let accout = table[0][0]
      let name = table[0][7]
      let id = table[0][8]
      commonHeader[1][4] = accout
      commonHeader[2][1] = name
      commonHeader[2][4] = id
      let breakPoint = 0
      for (let i = 0; i < table.length; i++) {
        if (!table[i] || table[i][9] !== order) {
          breakPoint = i;
          break;
        }
      }
      let output = [...commonHeader, ...header, ...table.slice(0, breakPoint || undefined).map(x => changePosition(deleteCloumn(x)))]
      const folderName = '账户对账单'
      const subFolderName = `${order}${name}`
      const fileName = `${order}交易流水${name}`
      // 生成文件夹
      if (!fs.existsSync(folderName)) {
        fs.mkdirSync(folderName)
      }
      // 生成子文件夹
      if (!fs.existsSync(`${folderName}/${subFolderName}`)) {
        fs.mkdirSync(`${folderName}/${subFolderName}`)
      }
      // 生成xlsx文件
      var buffer = xlsx.build([{name: `mySheetName`, data: output}]); // Returns a buffer 
      console.log(chalk.blue(`提交第${++applyNum}个xlsx文件生成任务\n`))
      fs.writeFileSync(`${folderName}/${subFolderName}/${fileName}.xlsx`, buffer)
      console.log(chalk.green(`成功生成第${++fileNum}个xlsx文件\n`))
      if (breakPoint === 0) {
        handleTable([])
      } else {
        table = table.slice(breakPoint)
        setTimeout(() => {
          handleTable(table)
        }, 100)
      }
    } else {
      console.log(chalk.green(`程序结束，共提交${applyNum}个任务，共生成${fileNum}个xlsx文件\n`));
    }
  }
}

