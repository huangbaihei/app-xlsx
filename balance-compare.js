

// 去重表，修改以下这个文件绝对路径，在终端输入npm run com即可启动程序
const sourcePath = '/Users/huangqier/Documents/原始资料/待处理表3.xlsx'
// 输出的文件夹名
const folderName = '余额比对专用'
// 输出的文件名
const fileName = '去重表3'

var xlsx = require('node-xlsx'); 
var chalk = require('chalk')

var fs = require('fs'); 

var header = []
var tableContainer = []
var applyNum = 0
var fileNum = 0

// 读取文件 
// const sourcePath = '/Users/huangqier/Downloads/待处理表2.xlsx'
console.log(chalk.blue(`正在读取xlsx源文件...\n`))
const workSheetsFromBuffer = xlsx.parse(fs.readFileSync(sourcePath));
console.log(chalk.green(`成功读取xlsx源文件，共有${workSheetsFromBuffer.length}个sheet，开始进入去重处理流程\n`))

// 只处理一个sheet
// const table = workSheetsFromBuffer[0].data
// handleTableSheet(table)

// 处理多个sheet
for (let i = 0; i < workSheetsFromBuffer.length; i++) {
  console.log(chalk.blue(`开始处理第${i + 1}个sheet...\n`))
  handleTableSheet(workSheetsFromBuffer[i].data)
}

function handleTableSheet (table) {
  header = [table[0]];
  // console.log(header)
  tableContainer = table.slice(1);
  // console.log(tableContainer[0])
  handleTable(tableContainer);

  function handleTable(table) {
    for (let i = table.length - 2; i >= 0; i--) {
      if (table[i][0] === table[i + 1][0]) {
        table.splice(i, 1)
      }
    }
    let output = [...header, ...table]
    // 生成文件夹
    if (!fs.existsSync(folderName)) {
      fs.mkdirSync(folderName)
    }
    // 生成xlsx文件
    var buffer = xlsx.build([{name: `mySheetName`, data: output}]); // Returns a buffer 
    console.log(chalk.blue(`提交${++applyNum}个xlsx文件生成任务\n`))
    fs.writeFileSync(`${folderName}/${fileName}.xlsx`, buffer)
    console.log(chalk.green(`成功生成${++fileNum}个xlsx文件\n`))
  }
}


