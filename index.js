const fs = require('fs');
const WriterManager = require("./src/WriterManager");

// jsファイルのディレクトリ内のExcelファイル(１件目)を取得
const fileList = fs.readdirSync(__dirname);
const excelFilePath = fileList.find((fileName) => {
  return fileName.indexOf(".xlsx") > 0;
});

// Excelファイル読みこみ
const manager = new WriterManager(__dirname + "/" + excelFilePath);
// 議事録担当者と回数を出力
manager.consoleWriteWriterList();
