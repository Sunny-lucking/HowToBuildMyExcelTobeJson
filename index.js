const Util = require('./src/util')
const downLoadExcel = require('./src/downLoadExcel')
const parseXlsx = require('./src/parser')
const generateJsonFile = require("./src/generateJsonFile")
async function excelToJson(excelPathName, outputPath) {
  if (Util.checkAddress(excelPathName) === 'google') {
    // 1.判断是谷歌excel文档，需要交给Google对象去处理，主要是下载线上的，生成本地excel文件
    const filePath = await downLoadExcel(excelPathName)

    // 2.解析本地excel成二维数组
    const data = await parseXlsx(filePath)
    
    // 3.生成json文件
    generateJsonFile(data, outputPath)
  }

}
module.exports = excelToJson

