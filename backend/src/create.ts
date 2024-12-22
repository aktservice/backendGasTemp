class FileProc {
  private rootFolderId;
  private rootFolderObj;
  constructor(rootFolderId: string) {
    this.rootFolderId = rootFolderId;
    this.rootFolderObj = DriveApp.getFolderById(rootFolderId);
  }
  /**
   * createFileByList
   */
  public copyFileByList(
    targetFolderId: string,
    copyTarget: GoogleAppsScript.Drive.File,
    list: string[]
  ): object[] {
    const folder = DriveApp.getFolderById(targetFolderId);
    let result: object[] = [];
    list.forEach((element) => {
      const target = copyTarget.makeCopy(element, folder);
      result.push({ id: target.getId(), fileName: target.getName() });
    });
    return result;
  }

  /**
   * name
   */
  public getSpDataByList(spid, rng, list) {
    //listにあるIDからデータを取得するメソッド
  }
}
