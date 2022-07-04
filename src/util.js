const fs = require('fs')
const path = require('path')

/**
 * @param {string} url 一个用于校验的字符串，校验是否为合法http地址。返回true,false
 */
exports.isHttpUrl = function (url) {
  return /^http(s)?:\/\/([\w-_]+\.)*[\w-_]+\.[a-zA-Z]+(:\d+)?/i.test(url)
}

/**
 * @param {string}  excel路径
 * @return {string} excel文件绝对路径
 */
exports.excelAbsPath = function (str) {
  return exports.isHttpUrl(str)
    ? str
    : path.isAbsolute(str)
      ? str
      : path.join(process.cwd(), str)
}

/**
 * @param {string}  输出目录路径
 * @return {string} 输出目录绝对路径
 */
exports.outAbsPath = function (str) {
  return path.isAbsolute(str) ? str : path.join(process.cwd(), str)
}

/**
 *
 * @param {string} 解析传入excel地址
 * @return {string}
 */
exports.checkAddress = function (url) {
  if (!exports.isHttpUrl(url)) return 'local'
  return url.split(/https?:\/\/|:|\//)[1].split('.')[1]
}


/**
 *
 * @param {string} keys_arr[i] 校验是否有require写法
 */
exports.toNumber = function (str) {
  if (typeof str === 'string') {
    str = str.trim()
  }
  return /\D/.test(str) ? str : Number(str)
}


/**
 *
 * @param {array*2} matrix 清除二维数组，那些完全是空内容的项。
 */
exports.clearEmptyArrItem = function (matrix) {
  return matrix.filter(function (val) {
    return val.some(function (val1) {
      return val1.replace(/\s/g, '') !== ''
    })
  })
}


/**
 *
 * @param {array*2} matrix 一个二维数组，返回旋转后的二维数组。
 */
exports.rotateExcelDate = function (matrix) {
  if (!matrix[0]) return []
  var results = [],
    result = [],
    i,
    j,
    lens,
    len
  for (i = 0, lens = matrix[0].length; i < lens; i++) {
    result = []
    for (j = 0, len = matrix.length; j < len; j++) {
      result[j] = matrix[j][i]
    }
    results.push(result)
  }
  return results
}

exports.getEmpty2DArr = function(rows, cols) {
  let arrs = new Array(rows);
  for (var i = 0; i < arrs.length; i++) {
    arrs[i] = new Array(cols).fill(''); //每行有cols列
  }
  return arrs;
}

exports.colToInt = function (col) {
  const letters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
  col = col.trim().split('')
  let n = 0

  for (let i = 0; i < col.length; i++) {
    n *= 26
    n += letters.indexOf(col[i])
  }

  return n
}