class SheetEdit {
  private sp: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private sh: GoogleAppsScript.Spreadsheet.Sheet;

  constructor(
    spwb: GoogleAppsScript.Spreadsheet.Spreadsheet,
    sheet?: GoogleAppsScript.Spreadsheet.Sheet
  ) {
    this.sp = spwb;
    this.sh = sheet ?? spwb.getSheets()[0];
  }
  /**
   * @description ２次元配列データを貼り付けする関数
   * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
   * @date 2024-01-03
   * @param {any[][]} data getDataRange() => data
   * @param {number} [stRow=1] default 1
   * @param {number} [stCol=1] default 1
   * @memberof SheetEdit
   */
  public set2DimData(data: any[][], stRow = 1, stCol = 1, sheetName?: string) {
    const sh = this.sp.getSheetByName(sheetName) ?? this.sh;

    sh.getRange(stRow, stCol, data.length, data[stRow].length).setValues(data);
  }
  /**
   * set2DimDataUseSheetAPI
   */
  public set2DimDataUseSheetAPI(
    data: any[][],
    stRange = "A1",
    sheetName?: string
  ) {
    const spId = this.sp.getId();
    const sh = this.sp.getSheetByName(sheetName) ?? this.sh;
    const shName = sh.getName();
    const opt: GoogleAppsScript.Sheets.Schema.BatchUpdateValuesRequest = {
      valueInputOption: "USER_ENTERED",
      data: [{ range: `${shName}!${stRange}`, values: data }],
    };
    Sheets.Spreadsheets.Values.batchUpdate(opt, spId);
  }
  public append2DimDataUseSheetAPI(
    data: any[][],
    stRange = "A1",
    sheetName?: string
  ) {
    const spId = this.sp.getId();
    const sh = this.sp.getSheetByName(sheetName) ?? this.sh;
    const shName = sh.getName();
    const tshName = `${sheetName}!A1`;
    const resource: GoogleAppsScript.Sheets.Schema.ValueRange = {
      values: data,
    };
    const opt = { valueInputOption: "RAW", insertDataOption: "INSERT_ROWS" };
    const sheetapiRes = Sheets.Spreadsheets.Values.append(
      resource,
      spId,
      tshName,
      opt
    );

    return sheetapiRes;
  }
}
