const fs = require('fs');
const path = require('path');
const WriterManager = require("./WriterManager");

// jsファイルのディレクトリ内のExcelファイル(１件目)を取得
const dir = path.dirname(process.argv[1]);
const fileList = fs.readdirSync(dir);
const excelFilePath = fileList.find((fileName) => {
  return fileName.indexOf(".xlsx") > 0;
});

if(excelFilePath) {
  // Excelファイル読みこみ
  const manager = new WriterManager(dir + "/" + excelFilePath);
  // 議事録担当者と回数を出力
  manager.consoleWriteWriterList();
} else {
  console.log("Not found Excel file.");
}


