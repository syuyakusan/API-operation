/**
 * リンク集を共有するためのテキストを作成し、ダッシュボードに貼り付けるAPI関数
 * @param {SpreadSheet} spreadSheet
 */
function setSharedLinkTextForApi(spreadSheet) {
  const setSheet = spreadSheet.getSheetByName('ホーム');  

  const sheetUrl = spreadSheet.getUrl();
  const formAnswerUrl = setSheet.getRange('K7').getValue();

  const sharedLinks = `以下のリンクから予定を回答してください！\n【PCの方：集約さんのスプレッドシート】\n ${sheetUrl} \n 【スマホの方：集約さんのフォーム】\n ${formAnswerUrl}`;

  setSheet.getRange('F17').setValue(sharedLinks);

}

/**
 * 集約結果を共有するためのテキストにして、ダッシュボードに貼り付けるAPI関数
 * @param {SpreadSheet} spreadSheet
 */
function setSharedResultTextForApi(spreadSheet) {
  const setSheet = spreadSheet.getSheetByName('ホーム');  

  const resultTextArray = summarizedResultToTextForApi(spreadSheet);

   if(resultTextArray.length === 0){
    resultTextArray.push("全員の予定が合う日程がありません");
  }else{
    resultTextArray.unshift("全員の予定が合う日程は以下の通りです");
  }

  const resultText = resultTextArray.join("\n");

  setSheet.getRange('F18').setValue(resultText);

}
