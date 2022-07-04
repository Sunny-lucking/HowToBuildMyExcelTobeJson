const fs = require('fs')
const Stream = require('stream')
const unzip = require('unzipper')
const xpath = require('xpath')
const XMLDOM = require('xmldom')
const { getEmpty2DArr, colToInt,clearEmptyArrItem,rotateExcelDate} = require('./util')
const ns = { a: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' }
const select = xpath.useNamespaces(ns)


async function parseXlsx(path) {

  // 1. 解析本地excel文件，获取excel的sheet信息和strings信息
  const files = await extractFiles(path);

  // 2. 根据strings和sheet解析成二维数组
  const data = await extractData(files)

  // 3. 处理二维数组的内容，
  const fixData = handleData(data)
  return fixData;
}

function extractFiles(path) {

  // excel的本质是多份xml组成的压缩文件，这里我们只需要xl/sharedStrings.xml和xl/worksheets/sheet1.xml
  const files = {
    strings: {}, // strings内容
    sheet: {},
    'xl/sharedStrings.xml': 'strings',
    'xl/worksheets/sheet1.xml': 'sheet'
  }

  const stream = path instanceof Stream ? path : fs.createReadStream(path)

  return new Promise((resolve, reject) => {
    const filePromises = [] // 由于一份excel文档，会被解析成好多分xml文档，但是我们只需要两份xml文档，分别是（xl/sharedStrings.xml和xl/worksheets/sheet1.xml），所以用数组接受

    stream
      .pipe(unzip.Parse())
      .on('error', reject)
      .on('close', () => {
        Promise.all(filePromises).then(() => {
          console.log(files.strings.contents);
          console.log(files.sheet.contents);
          return resolve(files)
        })
      })
      .on('entry', entry => {
       
        // 每解析某个xml文件都会进来这里，但是我们只需要xl/sharedStrings.xml和xl/worksheets/sheet1.xml，并将内容保存在strings和sheet中
        const file = files[entry.path]
        if (file) {
          let contents = ''
          let chunks = []
          let totalLength = 0
          filePromises.push(
            new Promise(resolve => {
              entry
                .on('data', chunk => {
                  chunks.push(chunk)
                  totalLength += chunk.length
                })
                .on('end', () => {
                  contents = Buffer.concat(chunks, totalLength).toString()
                  files[file].contents = contents
                  if (/�/g.test(contents)) {
                    throw TypeError('本次转化出现乱码�')
                  } else {
                    resolve()
                  }
                })
            })
          )
        } else {
          entry.autodrain()
        }
      })
  })
}

function calculateDimensions(cells) {
  const comparator = (a, b) => a - b
  const allRows = cells.map(cell => cell.row).sort(comparator)
  const allCols = cells.map(cell => cell.column).sort(comparator)
  const minRow = allRows[0]
  const maxRow = allRows[allRows.length - 1]
  const minCol = allCols[0]
  const maxCol = allCols[allCols.length - 1]

  return [{ row: minRow, column: minCol }, { row: maxRow, column: maxCol }]
}


function extractData(files) {
  let sheet
  let values
  let data = []

  try {
    sheet = new XMLDOM.DOMParser().parseFromString(files.sheet.contents)
    const valuesDoc = new XMLDOM.DOMParser().parseFromString(
      files.strings.contents
    )

    // 把所有每个格子的内容都放进了values数组里。
    values = select('//a:si', valuesDoc).map(string =>
      select('.//a:t', string)
        .map(t => t.textContent)
        .join('')
    )

    console.log(values);
  } catch (parseError) {
    return []
  }



  const na = {
    textContent: ''
  }

  class CellCoords {
    constructor(cell) {
      cell = cell.split(/([0-9]+)/)
      this.row = parseInt(cell[1])
      this.column = colToInt(cell[0])
    }
  }

  class Cell {
    constructor(cellNode) {
      const r = cellNode.getAttribute('r')
      const type = cellNode.getAttribute('t') || ''
      const value = (select('a:v', cellNode, 1) || na).textContent
      const coords = new CellCoords(r)

      this.column = coords.column // 该格子所在列数
      this.row = coords.row // 该格子所在行数
      this.value = value // 该格子的顺序
      this.type = type // 该格子是否为空
    }
  }

  const cells = select('/a:worksheet/a:sheetData/a:row/a:c', sheet).map(
    node => new Cell(node)
  )

  // 计算该表格的最大最小列数行数
  d = calculateDimensions(cells)

  const cols = d[1].column - d[0].column + 1
  const rows = d[1].row - d[0].row + 1

  // 生成二维空数组
  data = getEmpty2DArr(rows, cols)

  // 填充二维空数组
  for (const cell of cells) {
    let value = cell.value

    // s表示该格子有内容
    if (cell.type == 's') {
      value = values[parseInt(value)]
    }

    // 填充该格子
    if (data[cell.row - d[0].row]) {
      data[cell.row - d[0].row][cell.column - d[0].column] = value
    }
  }
  return data
}



function handleData(data) {
  if (data) {
    data = clearEmptyArrItem(data)
    data = rotateExcelDate(data)
    data = clearEmptyArrItem(data)
  }
  return data
}

module.exports = parseXlsx
