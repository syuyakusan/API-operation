/**
 * ※※実行注意※※
 * 最終閲覧日から90日以上たっている使われていないスプレッドシートを削除する関数
 * ∵カレンダーは90日分もないので、スプレッドシートを開いて再設定する必要がある
 * トリガーで毎月1日に実行
 */
function deleteDisusedSpreadSheet() {
  const idNameDb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-Name");

  const urlList = idNameDb.getRange(1,2,1,idNameDb.getLastColumn()-1).getValues().flat();
  const today = new Date("2023/12/01");
  let deadLine = new Date();
  // 列が削除されると列番号がずれるので、カウンタで補正する
  let j=0;
  for(i=0;i<urlList.length;i++){
    const spreadSheet = SpreadsheetApp.openByUrl(urlList[i]);
    const lastViewdDate = spreadSheet.getSheetByName("ホーム").getRange("O6").getValue();
    deadLine.setDate(lastViewdDate.getDate() + 90);
    if(today > deadLine){
      // ファイルのをゴミ箱へ
      try{
        const trahFile = DriveApp.getFileById(spreadSheet.getId());
        trahFile.setTrashed(true);
      }catch(e){
        // 既に削除されていた場合のエラーを吸収
      }

      //userID-Nameから列を削除
      idNameDb.deleteColumn(i+2-j);
      j++;

      // userID-URLsから該当URLを削除
      deleteUrlFromNameDB(urlList[i]);


    }
    // 初期化
    deadLine = new Date();
  }
  
}

/**
 * userID-URLsシートから任意のURLを消す関数
 * @param {String} url
 */
function deleteUrlFromNameDB(url){
  const idUrlsDb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-URLs");
  const tableRange = idUrlsDb.getRange(2,3,idUrlsDb.getLastRow()-1,idUrlsDb.getLastColumn()-2);
  let tableMatrix = tableRange.getValues();
  let maxLength = 0;
  // 該当するURLを削除
  for(i=0;i<tableMatrix.length;i++){
    const rowArray = tableMatrix[i];
    const index = rowArray.indexOf(url);
    if(index !== -1){
      tableMatrix[i] = tableMatrix[i].splice(index,1);
    }
    if(tableMatrix[i].length > maxLength){
      maxLength = tableMatrix[i].length;
    }
  }
  // 修正した配列を書き込む
  tableRange.clearContent();
  idUrlsDb.getRange(2,3,idUrlsDb.getLastRow()-1,maxLength).setValues(tableMatrix);
  
}

/**
 * userID-Nameシートから任意のuserIDを消す関数
 * @param {String} userId
 */
function deleteUrlFromNameDB(userId){
  const idNameDb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-Name");
  try{
    const userIdList = idNameDb.getRange(3,1,idNameDb.getLastRow()-2,1).getValues().flat();
    const index = userIdList.indexOf(userId);
    if(index !== -1){
      //該当する行を削除
      idNameDb.deleteRow(index + 3);
    }
  }catch(e){
    // DBにuserIDが存在しない場合
    return;
  }
  
}

/**
 * userID-URLsシートから任意のuserIDを消す関数
 * @param {String} userId
 */
function deleteUrlFromUrlDB(userId){
  const idUrlsDb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("userID-URLs");
  try{
    const userIdList = idUrlsDb.getRange(2,1,idUrlsDb.getLastRow()-1,1).getValues().flat();
    const index = userIdList.indexOf(userId);
    if(index !== -1){
      //該当する行を削除
      idUrlsDb.deleteRow(index + 2);
    }
  }catch(e){
    // DBにuserIDが存在しない場合
    return;
  }
  
}
