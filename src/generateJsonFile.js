const fs = require('fs')
const rmobj = require('./remove')

function generateJsonFile(excelDatas, outputPath) {
  
  // 获得转化成json格式
  const jsons = convertProcess(excelDatas)

  // 生成写入文件
  writeFile(jsons, outputPath)
}

  /**
   *
   * @param {array*2} data
   * 返回处理完后的多语言数组，每一项都是一个json对象。
   */
  function convertProcess(data) {
    var keys_arr = [],
      data_arr = [],
      result_arr = [],
      i,
      j,
      data_arr_len,
      col_data_json,
      col_data_arr,
      data_arr_col_len
    // 表格合并处理，这是json属性列。
    keys_arr = data[0]
    // 第一例是json描述，后续是语言包
    data_arr = data.slice(1)

    for (i = 0, data_arr_len = data_arr.length; i < data_arr_len; i++) {
      // 取出第一个列语言包
      col_data_arr = data_arr[i]
      // 该列对应的临时对象
      col_data_json = {}
      for (
        j = 0, data_arr_col_len = col_data_arr.length;
        j < data_arr_col_len;
        j++
      ) {
        
        col_data_json[keys_arr[j]] = col_data_arr[j]
      }
      result_arr.push(col_data_json)
    }

    return result_arr
  }




  //得到的数据写入文件
  function writeFile(datas, outputPath) {
    for (let i = 0, len = datas.length; i < len; i++) {
      fs.writeFileSync(outputPath +
        (datas[i].filename || datas[i].lang) +
        '.json',
        JSON.stringify(datas[i], null, 4)
      )
    }
    rmobj.flush();
  }






module.exports = generateJsonFile
