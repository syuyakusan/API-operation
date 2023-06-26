/**
 * 回答状況をシートに入力するAPI関数
 * @param {SpreadSheet} spreadSheet
 */
function setAnswerStatusForApi(spreadSheet){
  const setSheet = spreadSheet.getSheetByName('ホーム');
  const reqNum = setSheet.getRange('F6').getValue();
  const memberList = setSheet.getRange(8,6,reqNum,1).getValues().flat();
  let outputArray = [];
  for (name of memberList){
    const lastDate = Utilities.formatDate(checkAnswerStatus(spreadSheet,name), "Asia/Tokyo", "MM/dd");
    outputArray.push([lastDate]);
  }
  setSheet.getRange(17,11,reqNum,1).setValues(outputArray);
}

/**
 * 個人シートがどこまで入力されているかを調べる関数
 * @param {SpreadSheet} spreadSheet
 * @param name {String} メンバー名(かつシート名)
 * preturns lastDate {Date} 入力最終日 
 */
function checkAnswerStatus(spreadSheet,name){
  const sheet = spreadSheet.getSheetByName(name);
  const range = sheet.getRange(5,2,31,74);
  const lastColumn = getLastColumnInRange(range);
  const lastDate = sheet.getRange(2,lastColumn,1,1).getValue();

  return new Date(lastDate);
}

