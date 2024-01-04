/**
 * @description メール処理
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 2022-10-15
 */
function sendEmailforEditors(
  newSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
) {
  const CONFIGSHNAME = "config";
  const EMAILADRESSCOLUMN = 1;
  const sp = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sp.getSheetByName(CONFIGSHNAME);
  const data = sh.getDataRange().getValues();

  const editors = [];
  //const viewers = [];
  data.forEach((element, index) => {
    //ヘッダーはリターン
    if (index == 0) {
      return;
    }
    if (element[EMAILADRESSCOLUMN] !== "") {
      editors.push(element[EMAILADRESSCOLUMN]);
    }
    /*
    if (element[5] !== '') {
      viewers.push(element[5]);
    }
    */
  });
  //const recipientCC = viewers.join();
  const recipientTO = editors.join();

  const dt: string = Utilities.formatDate(new Date(), "JST", "MM月分");
  //日付が入っている想定
  const deadlineString = sh.getRange("D2").getValue();
  const deadline = new Date(deadlineString);
  const arrayDay = ["日", "月", "火", "水", "木", "金", "土"];
  //曜日を取得
  const strDay = arrayDay[deadline.getDay()];
  const deadlineFormatString = Utilities.formatDate(
    deadline,
    "JST",
    "MM月dd日"
  );
  const dl = `${deadlineFormatString}（${strDay}）`;

  const subject: string = `${dt}Contents`;
  //添付するファイルのフォルダURL
  const tempFilesFolderURL = sh.getRange("A2").getValue();
  //mailSend.html
  const outP: GoogleAppsScript.HTML.HtmlTemplate =
    HtmlService.createTemplateFromFile("mailSend");
  //outP.TODAY = dt;
  outP.DEADLINE = dl;
  outP.URL = "";
  outP.FILEFOLDERURL = tempFilesFolderURL;
  const htmlObj: any = outP.evaluate().getContent();
  const options: any = {
    htmlBody: htmlObj,
    noReply: false,
    /*, cc: recipientCC*/
  };
  const body: string = "通常メール";
  GmailApp.sendEmail(recipientTO, subject, body, options);
}
