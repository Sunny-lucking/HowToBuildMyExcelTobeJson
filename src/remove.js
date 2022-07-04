const fs = require('fs')

module.exports = {
  rmdir: [],
  push() {
    this.rmdir.push(...arguments)
  },
  flush() {
    this.rmdir.forEach(obj => {
      fs.readdirSync(obj.path).forEach(file => {
        if (new RegExp(obj.ext).test(file)) fs.unlinkSync(obj.path + '/' + file)
      })
    })
    this.rmdir.length = 0
  }
}