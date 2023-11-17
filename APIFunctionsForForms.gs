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



  let counter = setSheet.getRange('K12').getValue();

  for (let i=0;i < (answerNum - counter);i++) { //回答件数とカウンタが一致していない場合のみ実行
    //処理

    latestForm = formResponses[answerNum- 1 - i];
    let itemResponses = latestForm.getItemResponses(); //回答がまとまった配列

    let name = itemResponses[0].getResponse();
    let time = itemResponses[1].getResponse();

    let timeArray = time.split('\n');//行ごとに分割


    //日付,開始時間,終了時間にわける関数
    const splitFunc = function(value,index,array) {
      return value.split('-');
    }
    
    //上記関数で成形した配列にする
    let splitedTimeArray = timeArray.map(splitFunc); //[****,****,****][日付,開始時刻,終了時刻]



    //セルの位置を求める関数
    const calcFunc = function(value) {
      let month = cutoffOverMonth(value[0].slice(0,2));
      let date =  cutoffOverDate(value[0].slice(2,5));
      let startH =cutoffOverHour(value[1].slice(0,2));
      let startM = cutoffOverMinute(value[1].slice(2,5));
      let endH =cutoffOverHour(value[2].slice(0,2));
      let endM = cutoffOverMinute(value[2].slice(2,5));
      let startRow = ((startH-7)*2 + (startM/30) +3);
      let endRow = ((endH-7)*2 + (endM/30) +3);
      let rowLength = endRow - startRow + 1;

      const fd = setSheet.getRange('F7').getValue();

      let firstDay = new Date(fd); //カレンダーの最初の日付

      let setDay = new Date(fd); //指定された日付
          setDay.setMonth(month-1);
          setDay.setDate(date);

      let difDays = (setDay - firstDay)/86400000;

      if (difDays < 0) { //年をまたぐ入力一年後の日付に
        setDay = setDay.setFullYear(setDay.getFullYear() + 1);
        difDays = (setDay - firstDay)/86400000;
      }


      let array = [difDays+2,startRow,rowLength];

      return (array);
    }

      let calculatedTimeArray =splitedTimeArray.map(calcFunc); //[列数,開始行数,開始行から終了行までの行数]


    //配列をもとにスプレッドシートに入力
    const sheet = spreadSheet.getSheetByName(name);
    
    for (let j=0;j<calculatedTimeArray.length;j++) {
      let array = calculatedTimeArray[j];
      let column = array[0];
      let row = array[1];
      let length = array[2];

      sheet.getRange(row,column,length,1).setValue('○');
    }

  }


  setSheet.getRange('K12').setValue(answerNum);
  let now = new Date();
  now = Utilities.formatDate(now, "Asia/Tokyo", "MM/dd HH:mm");
  setSheet.getRange('K13').setValue(now);
}