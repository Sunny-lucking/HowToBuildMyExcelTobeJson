const fs = require('fs')
const request = require('superagent')
const rmobj = require('./remove')

/**
 * 下载google excel 文档到本地
 * @param {*} url  // https://docs.google.com/spreadsheets/d/12q3leiNxdmI_ZLWFj4LP_EA5PeJpLF18vViuyiSOuvM/edit#gid=0
 * @returns 
 */
function downLoadExcel(url) {

  // 记录当前下载文件的目录，方便删除
  rmobj.push({
    path: __dirname,
    ext: 'xlsx'
  })
  return new Promise((resolve, reject) => {
    var down1 = url.split('/')
    var down2 = down1.pop() // edit#gid=0
    var url2 = down1.join('/') // https://docs.google.com/spreadsheets/d/12q3leiNxdmI_ZLWFj4LP_EA5PeJpLF18vViuyiSOuvM
    var id = down1.pop() // 12q3leiNxdmI_ZLWFj4LP_EA5PeJpLF18vViuyiSOuvM
    var hash = down2.split('#').pop() // gid=0
    var downurl = url2 + '/export?format=xlsx&id=' + id + '&' + hash  // https://docs.google.com/spreadsheets/d/12q3leiNxdmI_ZLWFj4LP_EA5PeJpLF18vViuyiSOuvM/export?format=xlsx&id=12q3leiNxdmI_ZLWFj4LP_EA5PeJpLF18vViuyiSOuvM&gid=0
    var loadedpath = __dirname + '/' + id + '.xlsx'
    const stream = fs.createWriteStream(loadedpath)
    const req = request.get(downurl)
    req.pipe(stream).on('finish', function () {
      resolve(loadedpath)
      // 已经成功下载下来了，接下来将本地excel转化成json的工作就交给Excel对象来完成
    })
  })

}

module.exports = downLoadExcel