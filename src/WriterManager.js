const xlsx = require("xlsx");
const Person = require("./Person");

/**
 * 議事録担当者の情報を管理するクラス
 */
module.exports = class WriterManager {

  /**
   * コンストラクタ
   * @param {String} fileName 議事録管理表のExcelファイルのファイル名
   */
  constructor(fileName) {
    this.book = xlsx.readFile(fileName);
    this.NAME_LIST_SHEET_NAME = "通算出席率";
    this.WRITER_SHEET_NAME = "議事確認管理";
    
    // Excelファイルから議事録担当者の情報を取得し保持する
    this.personList = this.createPersonList();
    this.setWriteCount();
  }

  /**
   * 通算出席率シートから議事録担当社員のリスト（東京事業部・一般社員）を作成する。
   * 
   * @returns {Person[]} 議事録担当社員リスト
   */
  createPersonList() {
    const json = this.getSheetJson(this.NAME_LIST_SHEET_NAME);

    let personList = [];

    json.forEach((row) => {
      if(row.__EMPTY === "一般" && row.__EMPTY_1 === "東京事業部") {
        //console.log(row);
        const name = row.議事録確認者.replace(/ |　/g, "");
        const person = new Person(name);
        personList.push(person);
      }
    });
    return personList;
  }
  
  /**
   * 議事録担当回数を担当者別に設定する
   */
  setWriteCount() {
    this.book.SheetNames.forEach((sheetName) => {
      if(sheetName.indexOf(this.WRITER_SHEET_NAME) >= 0) {
        const json = this.getSheetJson(sheetName);

        json.forEach((row) => {
          if(row.__EMPTY === "氏名") {
            this.addWriterCount(row.__EMPTY_6);
            this.addWriterCount(row.__EMPTY_8);
            this.addWriterCount(row.__EMPTY_10);
            this.addWriterCount(row.__EMPTY_12);
            this.addWriterCount(row.__EMPTY_14);
            this.addWriterCount(row.__EMPTY_16);
            this.addWriterCount(row.__EMPTY_18);
            this.addWriterCount(row.__EMPTY_20);
            this.addWriterCount(row.__EMPTY_22);
            this.addWriterCount(row.__EMPTY_24);
            this.addWriterCount(row.__EMPTY_26);
            this.addWriterCount(row.__EMPTY_28);
          }
          
        });
      }
    });
  }

  /**
   * 引数で渡されたシート名からシート情報をJSONで返却する
   * 
   * @param {String} sheetName シート名
   * @returns {JSON} シート内情報のJSONオブジェクト
   */
  getSheetJson(sheetName) {
    const sheet = this.book.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
  };

  /**
   * 議事録担当のセルデータから、議事録担当者の議事録担当回数を加算する
   * 
   * @param {String} writerData Excel内の議事録担当セルデータ
   */
  addWriterCount(writerData) {
    if(writerData.indexOf("・") > 0) {
      // 名称からかっこを削除し分割する。
      // カウント対象は２名体制後の議事録とする
      const nameList = writerData.replace(/\(|\)|（|）/g,"").split("・");
      nameList.forEach((name) => {
        const person = this.personList.find((person) => {
          return person.name.indexOf(name) >= 0;
        });
        if(person) {
          person.count = person.count + 1;
        }  
      });
    }
  }

  /**
   * 議事録担当者の情報をコンソールに出力する。
   */
  consoleWriteWriterList() {
    this.personList.forEach((person) => {
      console.log(person.name + "\t", person.count, "回");
    });
  }
}
