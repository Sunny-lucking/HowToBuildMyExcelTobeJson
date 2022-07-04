const excelToJson = require('./index.js');
excelToJson('https://docs.google.com/spreadsheets/d/12q3leiNxdmI_ZLWFj4LP_EA5PeJpLF18vViuyiSOuvM/edit#gid=0', __dirname +
'/lang/')

// [ 'lang', 'cn','en', 'lang001','我是阳光', 'i am sunny','lang002', '前端阳光','FE Sunny', 'lang003','带带我', 'ddw']