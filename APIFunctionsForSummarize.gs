/**
 * 各個人シートの値を比較して集約シートにまとめるAPI用関数
 * @param {SpreadSheet} spreadSheet
 */
function checkDiffForApi(spreadSheet){

  const sumSheet = spreadSheet.getSheetByName('集約');
  const setSheet = spreadSheet.getSheetByName('ホーム');
  //名前入りシート数を取得
  const memberNum = setSheet.getRange('F6').getValue();


  let values =[0];
  for (i=0;i<memberNum;i++) {  //各シートの値を配列に格納
    let hisSheet = spreadSheet.getSheets()[i+2];
    let hisRange = hisSheet.getRange(1, 1, hisSheet.getLastRow(), hisSheet.getLastColumn());
    values[i] = hisRange.getValues();
  }



  //表の大きさは縦5~35,横2~74
  let attendanceMatrix =[];
  let nameArray = [];
  let tmpRowArray = [];
  let tmpColumnArray = [];
  const nameList = setSheet.getRange(8,6,memberNum,1).getValues().flat();
  let sumSheetLastRow = sumSheet.getLastRow(); //行番号=行数
  let sumSheetLastColumn = sumSheet.getLastColumn();

  //表外に書き込みがある場合無視
  if (sumSheetLastRow > 35) {
    sumSheetLastRow = 35;
  }
  if (sumSheetLastColumn > 74) {
    sumSheetLastColumn = 74;
  }
  // 値を格納した配列で5Bにあたるところから判定を繰り返す
  // あるセルについて各シートを判定、終わったら次のセルへ
  // nameArray i行j列で〇だった人の配列
  // tmpRowArray[j] i行j列のnameArray
  // tmpColumnArray[i] i行のtmpRowArray 
  for (i=4;i<sumSheetLastRow;i++) {  //行の移動
    tmpRowArray =[];
    for (j=1;j<sumSheetLastColumn;j++) {  //列の移動
      nameArray = [];
      for (k=0;k < memberNum;k++) { //人の移動
        if (values[k][i][j] === "○") { //判定(漢数字ゼロではない)
          nameArray.push(nameList[k]); //〇だった人をnameArrayに追加
        }
      }
      tmpRowArray[j] = nameArray;
    }
    tmpColumnArray[i] = tmpRowArray;
    attendanceMatrix = tmpColumnArray;
    // attendanceMatrix[i][j] i行j列の出席者 [String,String,...]
  } 

  //配列を元に集約シートに書き込み
  const tableRange = sumSheet.getRange(5,2,sumSheetLastRow-4,sumSheetLastColumn-1);
  //集約シートの初期化
  tableRange.clearContent();
  tableRange.clearFormat();
  tableRange.clearNote();
  tableRange.setHorizontalAlignment("center"); 
  // 集約シートに書き込み
  let tmpList = [];
  let inputArray = sumSheet.getRange(5,2,35,74).getValues();
  for (i=4;i<sumSheetLastRow;i++) {
    for (j=1;j<sumSheetLastColumn;j++) {
      let tmpCell = sumSheet.getRange(i+1,j+1);
      tmpList = attendanceMatrix[i][j];
      if (tmpList != []){
        // i行j列 name\nname\name...
        tmpCell.setValue(tmpList.length);
        // tmpCell.setNote(tmpList.join("\n"));
        inputArray[i-4][j-1] = attendanceMatrix[i][j].join("\n")+"\n\n--既存の予定--";
      }
    }
  }
  sumSheet.getRange(5,2,35,74).setNotes(inputArray);

  //条件付き書式を再設定
  const colorRange = sumSheet.getRange(5,2,sumSheetLastRow-4,sumSheetLastColumn-1);
  const ruleGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(memberNum) 
    .setBackground("#b7e1cd") //セル背景を設定（緑）
    .setRanges([colorRange])
    .build();

  const ruleYellow = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(memberNum-1) 
    .setBackground("#fce8b2") //セル背景を設定（黄）
    .setRanges([colorRange])
    .build();
  
  const ruleRed = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(memberNum-2) 
    .setBackground("#f4c7c3") //セル背景を設定（赤）
    .setRanges([colorRange])
    .build();
  
  const calendarRange = sumSheet.getRange(2,2,sumSheetLastRow-1,sumSheetLastColumn-1);
  const ruleGray = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("OR(B$2<today(),B$2>today()+28)") 
    .setBackground("#a6a6a6")
    .setRanges([calendarRange])
    .build();

  const ruleOrange = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("B$2=today()")
    .setBackground("#fce8b2")
    .setRanges([calendarRange])
    .build();

  const rules =[ruleGreen,ruleYellow,ruleRed,ruleGray,ruleOrange];
  sumSheet.setConditionalFormatRules(rules);

}

/**
 * 集約結果をテキストにして配列に格納して返すAPI用関数
 * @param {SpreadSheet} spreadSheet
 * @returns {Array} 集約結果
 * @example [["1/1 10:00~12:00"],[]]
 */
function summarizedResultToTextForApi(spreadSheet){
  const sumSheet = spreadSheet.getSheetByName('集約');
  const setSheet = spreadSheet.getSheetByName('ホーム');
  //名前入りシート数を取得
  const memberNum = setSheet.getRange('F6').getValue();

  // 行列を転置する関数
  const transposeArray = array => array[0].map((_, colIndex) => array.map(row => row[colIndex]));

  const tableColumns = 74;
  const tableRows =31;

  const resultMatrix = sumSheet.getRange(5,2,tableRows,tableColumns).getValues();
  const timeHeaderArray = sumSheet.getRange(5,1,tableRows,1).getDisplayValues().flat();
  const dateHeaderArray = sumSheet.getRange(2,2,1,tableColumns).getValues().flat();


  // 転置する[[n行],[n+1行]]->[[n列],[n+1列]]
  const transposedResultMatrix = transposeArray(resultMatrix);

  let resultTextArray = [];
  // 列ごとに探索
  for(i=0;i<74;i++){
    let consecutiveCounter = 0;
    // 1行ずつ進める
    for(j=0;j<31;j++){
      const value = transposedResultMatrix[i][j];
      if(value === memberNum){
        consecutiveCounter++;
      }else if (value != memberNum && consecutiveCounter>0){
        let cellAmount = consecutiveCounter;
        let date = Utilities.formatDate(new Date(dateHeaderArray[i]), 'JST', 'MM/dd');
        let startTime = timeHeaderArray[j-cellAmount];
        let endTime = timeHeaderArray[j-1];
        if(startTime !== endTime){
          resultTextArray.push(`${date} ${startTime}~${endTime}`);
        }
        consecutiveCounter=0;
      }
    }
  }


  return resultTextArray;
}

/**
 * LINEの日程回答をシートに反映する関数
 * @param {SpreadSheet} spreadSheet
 * @param {String} name
 * @param {String} answerText
 * @returns {String} forUserText
 */
function summarizeLineDateAnswerToSpreadSheetForApi(spreadSheet,name,answerText){
  const setSheet = spreadSheet.getSheetByName('ホーム');

  let time = answerText;
  // LINE表示用
  let forUserText="以下の日程で反映しました！";
  const forUserArray = time.split("\n");
  for(i=0;i<forUserArray.length;i++){
    forUserText = `${forUserText}\n${convertDateFormat(forUserArray[i])}`;
  }

  // 反映用
  let timeArray = time.split('\n');//行ごとに分割


  //日付,開始時間,終了時間にわける関数
  const splitFunc = function(value,index,array) {
    return value.split('-');
  }
  
  //上記関数で成形した配列にする
  let splitedTimeArray = timeArray.map(splitFunc); //[****,****,****][日付,開始時刻,終了時刻]

  //セルの位置を求める関数
  const calcFunc = function(value) {
    let month = value[0].slice(0,2);
    let date = value[0].slice(2,5);
    let startH =value[1].slice(0,2);
    let startM = value[1].slice(2,5);
    let endH =value[2].slice(0,2);
    let endM = value[2].slice(2,5);
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
  
  for (i=0;i<calculatedTimeArray.length;i++) {
    let array = calculatedTimeArray[i];
    let column = array[0];
    let row = array[1];
    let length = array[2];

    sheet.getRange(row,column,length,1).setValue('○');
  }
  return forUserText;
}

/**LINE表示用に日時のフォーマットを変換する関数
 * @param {String} dataString
 */
function convertDateFormat(dateString) {
  let regex = /^(\d{2})(\d{2})-(\d{2})(\d{2})-(\d{2})(\d{2})$/;
  let match = dateString.match(regex);
  
  if (match) {
    let month = parseInt(match[1], 10);
    let day = parseInt(match[2], 10);
    let startHour = parseInt(match[3], 10);
    let startMinute = parseInt(match[4], 10);
    let endHour = parseInt(match[5], 10);
    let endMinute = parseInt(match[6], 10);
    
    let convertedDateString = month + "月" + day + "日 " + startHour + ":" + formatMinutes(startMinute) + "~" + endHour + ":" + formatMinutes(endMinute);
    return convertedDateString;
  } else {
    throw new Error('Invalid date format. Please use "MMdd-HHmm-HHmm" format.');
  }
}

function formatMinutes(minutes) {
  return minutes < 10 ? "0" + minutes : minutes;
}


