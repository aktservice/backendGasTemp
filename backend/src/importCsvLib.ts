/*
 *
 * @description CSV or EXCEL をインポートするクラス
 * @author yoshitaka
 * @date 2019-07-14
 * @export
 * @class importCsvLib
 * @constructor
 * @param {string} rootFolderId ルートフォルダのＩＤ
 * @param {string}[findFolderName = "root"] ルート配下で検索するフォルダ名
 *
 */

export class importCsvLib {
  private folderId: string
  constructor(rootFolderId: string, findFolderName: string = 'root') {
    if (findFolderName == 'root') {
      this.folderId = rootFolderId
    } else {
      const id = this.findFolder(rootFolderId, findFolderName).getId()
      this.folderId = id
    }
  }
  /**
   * @description　フォルダＩＤを返す
   * @readonly
   * @type {string}
   * @memberof importCsvLib
   */
  public get getFolderId(): string {
    return this.folderId
  }
  /**
   * @description　フォルダを探すメソッド
   * @description フォルダがなければ１個目を返す
   * @author yoshitaka
   * @date 2019-11-07
   * @param {string} rootFolderId ルートフォルダＩＤ
   * @param {string} [folderName="noName"]　ルート配下で検索するフォルダ名
   * @returns {GoogleAppsScript.Drive.Folder}
   * @memberof importCsvLib
   */
  findFolder(
    rootFolderId: string,
    folderName: string = 'noName'
  ): GoogleAppsScript.Drive.Folder {
    var folder: GoogleAppsScript.Drive.Folder
    const folders = DriveApp.getFolderById(rootFolderId).getFolders()
    while (folders.hasNext()) {
      folder = folders.next()
      if (folder.getName() == folderName) {
        return folder
      } else if (folderName == 'noName') {
        return folder
      }
    }
  }

  /**
   * @description　ファイルを探すメソッド
   * @description ファイル名指定がなければ１個目を返す
   * @author yoshitaka
   * @date 2019-10-29
   * @param {string} [fileName="noName"] 探すファイル名称
   * @returns {GoogleAppsScript.Drive.File}
   * @memberof importCsvLib
   */
  findFile(
    fileName: string = 'noName'
  ): GoogleAppsScript.Drive.File | undefined {
    var file: GoogleAppsScript.Drive.File
    const files = DriveApp.getFolderById(this.folderId).getFiles()
    while (files.hasNext()) {
      file = files.next()
      if (file.getName() == fileName) {
        return file
      } else if (fileName == 'noName') {
        return file
      }
    }
  }

  findFiles(
    fileName: string = 'noName'
  ): GoogleAppsScript.Drive.FileIterator | undefined {
    var file: GoogleAppsScript.Drive.File
    const files = DriveApp.getFolderById(this.folderId).getFiles()
    while (files.hasNext()) {
      return files
    }
    return undefined
  }

  /**
   * @description　EXCELファイルをグーグルスプレッドシート形式ファイルにするメソッド
   * @author yoshitaka
   * @date 2019-07-14
   * @param {Object} file documentNode　ファイルノード（ＨＴＭＬ側からファイルでとる事を想定）
   * @param {string} fileName 新しくつけるファイル名（年月が先頭につく）
   * @returns {string} できあがったファイルのＩＤ文字列
   * @memberof importCsvLib
   */

  createDriveFileFromCSV(file: any, fileName: string): string {
    const folderID: string = this.folderId
    const yyyyMM: string = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')
    var res = Drive.Files.insert(
      {
        mimeType: 'application/vnd.google-apps.spreadsheet',
        parents: [{ id: folderID }],
        title: yyyyMM + fileName,
      },
      file.getBlob()
    )
    return res.id
  }
  createDriveFileFromExcel(
    file: GoogleAppsScript.Base.Blob,
    fileName: string
  ): string {
    const folderID: string = this.folderId
    var res = Drive.Files.insert(
      {
        mimeType: 'application/vnd.google-apps.spreadsheet',
        parents: [{ id: folderID }],
        title: fileName,
      },
      file
    )
    return res.id
  }

  /**
   * @description　ＨＴＭＬ側から入ってきたＢＬＯＢを使い
   *              グーグルドライブのファイルにするメソッド
   * @author yoshitaka
   * @date 2019-07-14
   * @param {object} form.Files
   * @returns {object}
   * @memberof importCsvLib
   */
  createDriveFileFromblob(file: any): any {
    var myfile = DriveApp.createFile(file.getBlob())
    return myfile
  }

  /**
   * @description ファイルを削除するメソッド
   * @author yoshitaka
   * @date 2019-07-14
   * @param {string} fileId
   * @memberof importCsvLib
   */
  deleteDriveFileFromId(fileId: string) {
    DriveApp.getFileById(fileId).setTrashed(true)
  }

  /**
   * @description ファイルを２次元配列データとして返すメソッド
   * @author yoshitaka
   * @date 2019-07-14
   * @param {string} fileId
   * @returns {string[][]}
   * @memberof importCsvLib
   */
  sendCsv(fileId: string): string[][] {
    return this.csvChangeJis(fileId)
  }

  /**
   * @description CSVファイルを２次元配列にするメソッド
   * @author yoshitaka
   * @date 2019-07-14
   * @param {Object} fileId csvFile(UTF8)
   * @returns {string[][]}
   * @memberof importCsvLib
   */
  csvChange(fileId: any): string[][] {
    var blob = DriveApp.getFileById(fileId).getBlob().getDataAsString()
    var data: string[][] = Utilities.parseCsv(blob)
    return data
  }

  /**
   * @description CSVファイルを２次元配列にするメソッド
   * @author yoshitaka
   * @date 2019-07-14
   * @param {Object} fileId csvFile(shift-jis)
   * @returns {string[][]}
   * @memberof importCsvLib
   */

  csvChangeJis(fileId: string): string[][] {
    var blob = DriveApp.getFileById(fileId)
      .getBlob()
      .getDataAsString('Shift_JIS')
    var data: string[][] = Utilities.parseCsv(blob)
    return data
  }
}
