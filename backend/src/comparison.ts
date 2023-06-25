class Comparison {
  public sp: GoogleAppsScript.Spreadsheet.Spreadsheet;

  constructor(spreadsheetObj) {
    this.sp = spreadsheetObj;
  }
  /**
   * getFilterData
   */
  public getFilterDataBySheetName(
    targetShName: string,
    filterValue: string,
    targetArrayColumn: number
  ): string[][] | any[][] {
    const sp = this.sp;
    const sh = sp.getSheetByName(targetShName);
    const data = sh.getDataRange().getDisplayValues();
    const filterData = data.filter((element, index) => {
      const value = element[targetArrayColumn];
      if (filterValue == value) {
        return true;
      }
    });
    return filterData;
  }
  public getFilterDataBy2DimData(
    data: string[][] | any[][],
    filterValue: string,
    targetArrayColumn: number
  ): string[][] | any[][] {
    const sp = this.sp;
    const filterData = data.filter((element, index) => {
      const value = element[targetArrayColumn];
      if (filterValue == value) {
        return true;
      }
    });
    return filterData;
  }
}
