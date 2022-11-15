/**
 * @description openイベント
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 2022-10-18
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('メニュー')
  menu.addItem('メール送信', 'sendEmailforEditors')

  menu.addToUi()
}

/**
 * @description メール処理
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 2022-10-15
 */
function sendEmailforEditors(
  newSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
) {
  const sp = SpreadsheetApp.getActiveSpreadsheet()
  const sh = sp.getSheetByName('config')
  const data = sh.getDataRange().getValues()

  const editors = []
  //const viewers = [];
  data.forEach((element, index) => {
    //ヘッダーはリターン
    if (index == 0) {
      return
    }
    if (element[4] !== '') {
      editors.push(element[4])
    }
    /*
    if (element[5] !== '') {
      viewers.push(element[5]);
    }
    */
  })
  // newSpreadsheet.addEditors(editors);
  //const recipientCC = viewers.join();
  const recipientTO = editors.join()

  const dt: string = Utilities.formatDate(new Date(), 'JST', 'MM月分')
  const deadlineString = sh.getRange('D2').getValue()
  const deadline = new Date(deadlineString)
  const arrayDay = ['日', '月', '火', '水', '木', '金', '土']
  const strDay = arrayDay[deadline.getDay()]
  const deadlineFormatString = Utilities.formatDate(deadline, 'JST', 'MM月dd日')
  const dl = `${deadlineFormatString}（${strDay}）`

  const subject: string = `${dt}Contents`

  const tempFilesFolderURL = sh.getRange('A2').getValue()

  const outP: GoogleAppsScript.HTML.HtmlTemplate =
    HtmlService.createTemplateFromFile('mailSend')
  //outP.TODAY = dt;
  outP.DEADLINE = dl
  outP.URL = ''
  outP.FILEFOLDERURL = tempFilesFolderURL
  const htmlObj: any = outP.evaluate().getContent()
  const options: any = {
    htmlBody: htmlObj,
    noReply: false,
    /*, cc: recipientCC*/
  }
  const body: string = '通常メール'
  GmailApp.sendEmail(recipientTO, subject, body, options)
}

/**
 * @description 毎日９時にトリガー
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 2022-10-17
 */
function triggerScript(): void {
  const functionName = ''
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create()
}
