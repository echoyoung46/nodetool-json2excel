const jexcel = require('json2excel');
const fs = require('fs');
const path = require("path");

const param = process.argv[2];
const rootPath = path.resolve();
const filePath = rootPath + '/' + param;
let listArr = [];
let actData = [
  [],
  [],
  [],
  [],
  []
];
let sheetIdx = 0;
let count = 0;
let fileLength = 0;

let JSON2EXCEL = {
  init: () => {
    JSON2EXCEL.combineData();
  },
  combineData: () => {
    fs.readdir(filePath, (err, files) => {
      if (err) {
        console.log(err);
        return;
      }

      fileLength = files.length;

      files.forEach(function (filename) {
        fs.stat(path.join(filePath, filename), function (err, stats) {
          if (err) throw err;
          if (stats.isFile()) {
            fs.readFile(filePath + '/' + filename, 'utf-8', (err, data) => {
              count++;
              if (err) {
                throw err;
              }
              try {
                listArr.push(JSON.parse(data));
              } catch (e) {
                console.log(e);
              }

              if (count == fileLength) {
                JSON2EXCEL.sortList();
              }
            })
          }
        });
      });
    })
  },
  sortList: () => {
    for (let i = 0; i < listArr.length; i++) {
      for (let j = 0; j < listArr[i].length; j++) {
        let item = listArr[i][j];
        let time = new Date(item.vUpdateDate);
        let _year = time.getFullYear();
        let _month = time.getMonth();

        if (_year == '2018') {
          actData[_month].push(item);
        }
      }
    }

    JSON2EXCEL.writeExcel();
  },
  writeFile: (data) => {
    let _data = JSON.stringify(data);
    fs.writeFile(rootPath + '/' + 'filelist.txt', _data, function (err) {
      if (err) throw err;
      JSON2EXCEL.writeExcel();
    });
  },
  writeExcel: () => {
    // console.log(actData[0].length);
    // console.log(actData[0][0]);
    let sheetData = [];

    let headerObj = {};

    for (let i in actData[0][0]) {
      headerObj[i] = i;
    }

    for (let i = 0; i < actData.length; i++) {
      sheetData.push({
        header: headerObj,
        items: actData[i],
        sheetName: 'sheet' + (i + 1)
      })
    }

    let excelData = {
      sheets: sheetData,
      filepath: 'pvpact.xlsx'
    };

    jexcel.j2e(excelData, function (err) {
      console.log('finish')
    });
  }
}

JSON2EXCEL.init();