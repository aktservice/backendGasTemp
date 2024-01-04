/**
 * @description openイベント
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 2022-10-18
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("メニュー");
  menu.addItem("メール送信", "sendEmailforEditors");

  menu.addToUi();
}

/**
 * @description 毎日９時にトリガー
 * @author yoshitaka <sato-yoshitaka@aktio.co.jp>
 * @date 2022-10-17
 */
function triggerScript(): void {
  const functionName = "";
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}
