/**
 * ステータスバーの表示を変える関数
 */
function statusFunc(status) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadSheet.getSheetByName('ホーム');
  setSheet.getRange('B6').setValue(status);

}

/**
 * ステータスバーの表示を変える関数
 */
function showProcess(process) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = spreadSheet.getSheetByName('ホーム');
  setSheet.getRange('B7').setValue(process);
}



// function statusClass(){
//   // TODO: すべての表示をこのクラスにまとめたい
//   const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
//   const setSheet = spreadSheet.getSheetByName('ホーム');
//   const showStatus = {
//     done : function(){
//       setSheet.getRange('B6').setValue('完了');
//       setSheet.getRange('B7').setValue('');
//     },
//     processing : function(process){
//       setSheet.getRange('B6').setValue('処理中...');
//       setSheet.getRange('B7').setValue(process);
//     },
//     error : function(errorLog){
//       setSheet.getRange('B6').setValue('エラー');
//       setSheet.getRange('B7').setValue(errorLog);
//     }
//   }
// }
