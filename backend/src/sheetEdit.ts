
/**
 * @description シート編集クラス
 * @author yoshikoro
 * @date 2019-06-13
 * @export
 * @class sheetEdit
 */
export class sheetEdit {
  private shName: string;
  /**
   *Creates an instance of sheetEdit.
   * @author yoshikoro
   * @date 2019-06-13
   * @param {string} sheetName
   * @memberof sheetEdit
   */
  constructor(sheetName: string) {
    this.shName = sheetName;
  }
  /**
   * @description シートをセットする
   * @author yoshikoro
   * @date 2019-06-13
   * @param {string} sheetName
   * @memberof sheetEdit
   */
  setSheetName(sheetName: string) {
    this.shName = sheetName;
  }
  getEmailData(configSheetName: string, configRow: number): string[] {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      configSheetName
    );
    const data: object[][] = sheet.getDataRange().getValues();
    let address: string[] = [];
    for (let index = 1; index < data.length; index++) {
      address.push(data[index][configRow].toString());
    }
    return address;
  }
  /**
   * @description 月データクリア用
   * @author yoshikoro
   * @date 2019-06-13
   * @memberof sheetEdit
   */
  monthDataClear(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.shName
    );
    const lastCol: number = sheet.getLastColumn();
    for (let index = 2; index < lastCol; index += 8) {
      sheet.getRange(5, index, 31, 2).clearContent();
    }
  }
  /**
   * @description allClear
   * @author yoshikoro
   * @date 2019-06-13
   * @memberof sheetEdit
   */
  allClear(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.shName
    );
    sheet.clear();
  }
  allClearBySpreadSheetId(spreadSheetId: string): void {
    const sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(
      this.shName
    );
    sheet.clear();
  }
  /**
   * @description　引数のデータをセット
   * @author yoshikoro
   * @date 2019-06-13
   * @param {(object[][] | string[][])} data
   * @memberof sheetEdit
   */
  setAll(data: object[][] | string[][]): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.shName
    );
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  setAllBySpreadSheetId(
    spreadSheetId: string,
    data: object[][] | string[][]
  ): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      spreadSheetId
    ).getSheetByName(this.shName);
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }

  /**
   * @description　指定した行数の全列全てにワークシート関数を入力
   * @author yoshikoro
   * @date 2019-06-13
   * @param {string} formulaString
   * @param {number} rowNum
   * @memberof sheetEdit
   */
  setFomulaString(formulaString: string, rowNum: number): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.shName
    );
    const colNum: number = sheet.getLastColumn();
    sheet.getRange(rowNum, 2, 1, colNum).setFormula(formulaString);
  }
  /**
   * @description
   * @author yoshitaka
   * @date 2019-08-02
   * @param {(object[][] | string[][])} data
   * @returns {number[]}
   * @memberof sheetEdit
   */
  getDataColumns(
    data: object[][] | string[][],
    kbnRowNumber: number = 2,
    searchWord: string
  ): number[] {
    var columns: number[] = [];
    for (let i = 0; i < data[kbnRowNumber].length; i++) {
      const element = data[kbnRowNumber][i];
      if (element == searchWord) {
        columns.push(i);
      }
    }
    return columns;
  }
  /**
   * @description 修正売り上げシート編集するメソッド
   * @author yoshitaka
   * @date 2019-08-02
   * @param {(object[][] | string[][])} data
   * @memberof sheetEdit
   */
  setcorrectionSalesUriage(
    data: object[][] | string[][],
    spreadSheetId?: string
  ): void {
    /**
     * @description 元データの必要列情報
     * @enum {number}
     */
    enum columnInfo {
      kdate = 0,
      uribumon = 2,
      salesman = 3, //営業担当
      //startUriage = 19, //売り上げ開始列20190802変動する為廃止
      //endUriage = 30 //売り上げ終了列20190802変動する為廃止
    }
    enum eCompleteFile {
      uribumon = 1,
      souri = 3,
      rentaluri = 4,
    }
    let results: string[][] | object[][] = data;
    //ここから本体
    const colInfo: number[] = [
      0,
      2,
      3,
      /*
      19,
      20,
      21,
      22,
      23,
      24,
      25,
      26,
      27,
      28,
      29,
      30
      */
    ];
    const uriageColummns: number[] = this.getDataColumns(data, 2, "稼動修正");
    let resultColumns = colInfo.concat(uriageColummns);
    for (let i = 0; i < results.length; i++) {
      for (let j = results[i].length - 1; j >= 0; j--) {
        if (resultColumns.indexOf(j) == -1) {
          results[i].splice(j, 1);
          //20190617へ変更
        }
      }
    }
    for (let i = 0; i < results.length; i++) {
      let val: string = results[i][eCompleteFile.uribumon].toString();
      val = val.replace(/[0-9]/g, "").trim();
      let oldsouri = results[i][eCompleteFile.souri].toString();
      let oldrentaluri = results[i][eCompleteFile.rentaluri].toString();
      let souriString = results[i][eCompleteFile.souri].toString();
      souriString = souriString.replace(/,/g, "");
      let souri: number = parseInt(souriString);

      let rentaluriString = results[i][eCompleteFile.rentaluri].toString();
      rentaluriString = rentaluriString.replace(/,/g, "");
      let rentaluri: number = parseInt(rentaluriString);
      results[i][eCompleteFile.uribumon] = val;
      if (isNaN(souri)) {
        results[i][eCompleteFile.souri] = oldsouri;
      } else {
        results[i][eCompleteFile.souri] = Math.abs(souri).toString();
      }
      if (isNaN(rentaluri)) {
        results[i][eCompleteFile.rentaluri] = oldrentaluri;
      } else {
        results[i][eCompleteFile.rentaluri] = Math.abs(rentaluri).toString();
      }
      const top = results[i].shift();
      results[i].splice(1, 0, top);
    }
    const isId = (spreadSheetIdString: string) => {
      if (spreadSheetId == undefined) {
        return SpreadsheetApp.getActiveSpreadsheet().getId();
      } else {
        return spreadSheetId;
      }
    };
    const id = isId(spreadSheetId);
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      id
    ).getSheetByName(this.shName);
    sheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  }
  setMainUriage(data: object[][] | string[][], spreadSheetId?: string): void {
    /**
     * @description 元データの必要列情報
     * @enum {number}
     */
    enum columnInfo {
      kdate = 0,
      uribumon = 2,
      salesman = 3,
      startUriage = 19,
      endUriage = 30, //売り上げ終了列
    }
    enum eCompleteFile {
      uribumon = 1,
    }
    let results: string[][] | object[][] = data;
    //ここから本体
    const colInfo: number[] = [
      0,
      2,
      3 /*,

            64,
            65,
            66,
            67,
            68,
            69,
            70,
            71,
            72,
            73,
            74,
            75
            */,
    ];
    const uriageColummns: number[] = this.getDataColumns(data, 2, "総合計");
    const resultColumns = colInfo.concat(uriageColummns);
    for (let i = 0; i < results.length; i++) {
      for (let j = results[i].length - 1; j >= 0; j--) {
        if (resultColumns.indexOf(j) == -1) {
          results[i].splice(j, 1);
          //20190617へ変更
        }
      }
    }
    for (let i = 0; i < results.length; i++) {
      let val: string = results[i][eCompleteFile.uribumon].toString();
      val = val.replace(/[0-9]/g, "").trim();
      results[i][eCompleteFile.uribumon] = val;
      const top = results[i].shift();
      results[i].splice(2, 0, top);
    }

    const isId = (spreadSheetIdString: string) => {
      if (spreadSheetId == undefined) {
        return SpreadsheetApp.getActiveSpreadsheet().getId();
      } else {
        return spreadSheetId;
      }
    };
    const id = isId(spreadSheetId);
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      id
    ).getSheetByName(this.shName);
    /*
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      "1AN4z6iFecNYA0_yayOGQZ-Ub_9166Giyf-Koy86JWc0"
    ).getSheetByName(this.shName);
    */
    sheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  }
  /**
   * @description　修理データをシートに入力
   * @description 元データの必要部分だけにフィルタする
   * @description enumに列情報記載
   * @author yoshikoro
   * @date 2019-06-13
   * @param {(object[][] | string[][])} data
   * @memberof sheetEdit
   */
  setRepairData(data: object[][] | string[][], spreadSheetId?: string): void {
    /**
     * @description 元データの必要列情報
     * @enum {number}
     */
    enum columnInfo {
      bumon = 11,
      uribumon = 16,
      salesman = 39,
      mitsudate = 40,
      daketudata = 42,
      isseiritu = 47,
      mitsumoney = 64,
      daketumoney = 66, //妥結金額
    }
    let results: string[][] | object[][] = data;
    //ここから本体
    const colInfo: number[] = [11, 16, 39, 40, 42, 47, 64, 66];
    for (let i = 0; i < results.length; i++) {
      for (let j = results[i].length - 1; j >= 0; j--) {
        if (colInfo.indexOf(j) == -1) {
          results[i].splice(j, 1);
          //20190617へ変更
        }
      }
    }

    const isId = (spreadSheetIdString: string) => {
      if (spreadSheetId == undefined) {
        return SpreadsheetApp.getActiveSpreadsheet().getId();
      } else {
        return spreadSheetId;
      }
    };
    const id = isId(spreadSheetId);
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      id
    ).getSheetByName(this.shName);
    sheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  }
  /**
   * @description 1列目データから数字を取り除きシートへ貼り付けるメソッド
   * @author sato-yoshitaka@aktio.co.jp
   * @date 2019-06-19
   * @param {(object[][] | string[][])} data
   * @memberof sheetEdit
   */
  setWarehousing(data: object[][] | string[][], spreadSheetId?: string): void {
    let results: string[][] | object[][] = data;
    //ここから本体
    const tCol = 0;
    const dayCol = 1;
    for (let i = 0; i < results.length; i++) {
      let val: string = results[i][tCol];
      //20190703 累計表示に変更の為日付データを入れない仕様に変更
      // let dayString: string = results[i][dayCol];
      //  val = val.replace(/[0-9]/g, "") + dayString.trim();
      val = val.replace(/[0-9]/g, "").trim();
      results[i][tCol] = val;
    }
    //20190617へ変更

    const isId = (spreadSheetIdString: string) => {
      if (spreadSheetId == undefined) {
        return SpreadsheetApp.getActiveSpreadsheet().getId();
      } else {
        return spreadSheetId;
      }
    };
    const id = isId(spreadSheetId);
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      id
    ).getSheetByName(this.shName);
    sheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  }

  /**
   * @description 1列目データから数字を取り除きシートへ貼り付けるメソッド
   * @author sato-yoshitaka@aktio.co.jp
   * @date 2019-06-19
   * @param {(object[][] | string[][])} data
   * @memberof sheetEdit
   */
  setAcountPay(data: object[][] | string[][], spreadSheetId?: string): void {
    let results: string[][] | object[][] = data;
    //ここから本体
    const tCol = 0;
    for (let i = 0; i < results.length; i++) {
      let val: string = results[i][tCol];
      val = val.replace(/[0-9]/g, "");
      results[i][tCol] = val.trim();
    }
    const isId = (spreadSheetIdString: string) => {
      if (spreadSheetId == undefined) {
        return SpreadsheetApp.getActiveSpreadsheet().getId();
      } else {
        return spreadSheetId;
      }
    };
    const id = isId(spreadSheetId);
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.openById(
      id
    ).getSheetByName(this.shName);
    sheet.getRange(1, 1, results.length, results[0].length).setValues(results);
  }
  /**
   * @deprecated 20190620 CSVデータ変更に伴い使用停止
   * @description 日商等を日付比較して日付が一致する行を編集するメソッド
   * @date 2019-06-08
   * @author yoshitaka
   * @param {(object[][] | string[][])} compariData
   * @param {(object[][] | string[][])} editData
   * @returns {(string[][] | object[][])} getDataRange().getValues()
   * @memberof sheetEdit
   */
  setFindData(
    compariData: object[][] | string[][],
    editData: object[][] | string[][]
  ): string[][] | object[][] {
    enum eRowColumn {
      ePlusRow = 3,
      eStartArrayIndex = 1,
      eCStartArrayIndex = 5,
      eCBlockOfficeCol = 1,
      eCShopCol = 2,
      etargetRow = 1,
      startRow = 5,
      eCReturnCol = 3,
      eCReturnRentalCol = 4,
    }
    //日付計算
    const targetData: string = String(compariData[3][0]);
    var replaceString: string = targetData.replace(/[^0-9]/g, "");
    replaceString = replaceString.slice(-2); //return yyyyMMdd
    const targetRow: number = parseInt(replaceString) + eRowColumn.ePlusRow;
    for (
      let i = eRowColumn.eStartArrayIndex;
      i < editData[targetRow].length;
      i++
    ) {
      let element: string = String(editData[eRowColumn.etargetRow][i]);
      const cnt: number = element.indexOf("ブロック");
      if (element == "") {
        continue;
      } else if (element == "横浜支店") {
        element = "総合計";
      } else if (cnt > 1) {
        element = element + "/合計";
      }
      for (let j = eRowColumn.eCStartArrayIndex; j < compariData.length; j++) {
        const value: string = String(
          compariData[j][eRowColumn.eCBlockOfficeCol]
        );
        const value2: string = String(compariData[j][eRowColumn.eCShopCol]);
        if (element == value) {
          editData[targetRow][i] = Math.abs(
            parseInt(String(compariData[j][eRowColumn.eCReturnCol]))
          );
          editData[targetRow][i + 1] = Math.abs(
            parseInt(String(compariData[j][eRowColumn.eCReturnRentalCol]))
          );
        } else if (element == value2) {
          editData[targetRow][i] = Math.abs(
            parseInt(String(compariData[j][eRowColumn.eCReturnCol]))
          );
          editData[targetRow][i + 1] = Math.abs(
            parseInt(String(compariData[j][eRowColumn.eCReturnRentalCol]))
          );
        }
      }
    }
    return editData;
    //{営業所一致} Column ＋ targetRow data[targetRow][column] = 総売り
    //{営業所一致} Column + 1 ＋ targetRow data[targetRow][column] = レンタル売り
  }
}

/**
 * @description 加工ベース
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 15/11/2022
 * @param {string} spreadSheetId
 */
function inspectionDataSpreadsheetAddFromTempFile(spreadSheetId: string) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy.MM.dd');
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const mvFolder = sp.getSheetByName('config').getRange('C2').getValue();
  //このシートをテンプレートにする
  const newSp = sp.copy(today + sp.getName());
  DriveApp.getFileById(newSp.getId()).moveTo(DriveApp.getFolderById(mvFolder));
  //configシートは必要ない為削除
  newSp.deleteSheet(newSp.getSheetByName('config'));
  const dataSp = SpreadsheetApp.openById(spreadSheetId);
  const data = dataSp.getSheets()[0].getDataRange().getValues();
  //フィルター
  const filterData = data.filter((elementArray) => {
    if (elementArray[1] !== 'フィルターする文字') {
      return false;
    } else {
      return true;
    }
  });
  //フィルターしたデータを必要列だけに加工
  for (let i = 0; i < filterData.length; i++) {
    //　必要列[ 1,5,8,9,10,12,13,14,15,23,25,29,30,31,32 ]
    // 列情報　B・F・I・J・K・M・N・O・P・X・Z・AD・AE・AF・AG
    const column = [
      0, 1, 2, 4, 6, 7, 11, 16, 17, 18, 19, 20, 21, 22, 24, 26, 27, 28,
    ];
    for (let k = 0; k < column.length; k++) {
      filterData[i].splice(column[k] - k, 1);
    }
  }

  //貼り付け処理
  newSp
    .getSheetByName('targetSheetName')
    .getRange(2, 1, data.length, data[0].length)
    .setValues(data);
