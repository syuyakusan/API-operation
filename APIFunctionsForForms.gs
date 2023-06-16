/**
 * Google Formsの回答結果をシートに反映するAPI関数
 * @param {SpreadSheet} spreadSheet
 */
function summarizeFormsForApi(spreadSheet){
  const setSheet = spreadSheet.getSheetByName('ホーム');

  const formId = setSheet.getRange('N11').getValue();
  if (formId =='') {
    setSheet.getRange('K12').setValue(0);
    setSheet.getRange('K13').setValue('フォームが存在しません'); 
    statusFunc('完了');
    showProcess('');
    return;
  }
  let form;
  try{
    form = FormApp.openById(formId);
  }catch(ERROR){
    // Formsが作られていない場合はこの処理は飛ばす
    return;
  }
  const formResponses = form.getResponses(); //全件の回答 
  const answerNum = formResponses.length;
  let latestForm = formResponses[answerNum- 1];
  if(typeof(latestForm) === "undefined") {
    setSheet.getRange('K12').setValue(0);
    setSheet.getRange('K13').setValue('回答がありません');  

    statusFunc('完了');
    showProcess('');
    return;
  }



  var counter = setSheet.getRange('K12').getValue();

  for (i=0;i < (answerNum - counter);i++) { //回答件数とカウンタが一致していない場合のみ実行
    //処理

    latestForm = formResponses[answerNum- 1 - i];
    var itemResponses = latestForm.getItemResponses(); //回答がまとまった配列

    var name = itemResponses[0].getResponse();
    var time = itemResponses[1].getResponse();

    var timeArray = time.split('\n');//行ごとに分割


    //日付,開始時間,終了時間にわける関数
    const splitFunc = function(value,index,array) {
      return value.split('-');
    }
    
    //上記関数で成形した配列にする
    var splitedTimeArray = timeArray.map(splitFunc); //[****,****,****][日付,開始時刻,終了時刻]



    //セルの位置を求める関数
    const calcFunc = function(value) {
      var month = value[0].slice(0,2);
      var date = value[0].slice(2,5);
      var startH =value[1].slice(0,2);
      var startM = value[1].slice(2,5);
      var endH =value[2].slice(0,2);
      var endM = value[2].slice(2,5);
      var startRow = ((startH-7)*2 + (startM/30) +3);
      var endRow = ((endH-7)*2 + (endM/30) +3);
      var rowLength = endRow - startRow + 1;

      const fd = setSheet.getRange('F7').getValue();

      var firstDay = new Date(fd); //カレンダーの最初の日付

      let setDay = new Date(fd); //指定された日付
          setDay.setMonth(month-1);
          setDay.setDate(date);

      var difDays = (setDay - firstDay)/86400000;

      if (difDays < 0) { //年をまたぐ入力一年後の日付に
        setDay = setDay.setFullYear(setDay.getFullYear() + 1);
        difDays = (setDay - firstDay)/86400000;
      }


      var array = [difDays+2,startRow,rowLength];

      return (array);
    }

      var calculatedTimeArray =splitedTimeArray.map(calcFunc); //[列数,開始行数,開始行から終了行までの行数]


    //配列をもとにスプレッドシートに入力
    const sheet = spreadSheet.getSheetByName(name);
    
    for (i=0;i<calculatedTimeArray.length;i++) {
      let array = calculatedTimeArray[i];
      let column = array[0];
      let row = array[1];
      let length = array[2];

      sheet.getRange(row,column,length,1).setValue('○');
    }

  }


  setSheet.getRange('K12').setValue(answerNum);
  var now = new Date();
  now = Utilities.formatDate(now, "Asia/Tokyo", "MM/dd HH:mm");
  setSheet.getRange('K13').setValue(now);
}
