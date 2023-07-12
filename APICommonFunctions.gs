/**
 * 範囲内の最終列を取得する関数
 * @param range {Range} 範囲
 * @return lastColumn {Number} 最終列
 */
function getLastColumnInRange(range){
  const matrix = range.getValues();
  const firstColumn = range.getColumn();
  const rangeHeight = range.getLastRow() - range.getRow() + 1;
  let maxRowLength = 0;
  for(i=0;i<rangeHeight;i++){
    let length = 0;

    for(j=0;j<matrix[i].length;j++){
    const index = matrix[i].length -1 -j;
    if(matrix[i][index] != ""){
      length = index+1;
      break;
    }}
    if(length > maxRowLength){
      maxRowLength = length; 
    }
  }
  const lastColumn = firstColumn + maxRowLength - 1;
  return lastColumn;
  
}

function isNumber(val){
   var regex = new RegExp(/^[0-9]+$/);
   return regex.test(val);
 }

 /**
 * 24以上の数字を切り捨てる関数
 * @param {Number|String} number
 * @returns {String} properNumber
 */
function cutoffOverHour(number){
  let properNumber = number;
  if(Number(number) > 24 ){
    properNumber = 24;
  }
  return properNumber;
}

/**
 * 60以上の数字を切り捨てる関数
 * @param {Number|String} number
 * @returns {String} properNumber
 */
function cutoffOverMinute(number){
  let properNumber = number;
  if(Number(number) > 60 ){
    properNumber = 60;
  }
  return properNumber;
}

/**
 * 12以上の数字を切り捨てる関数
 * @param {Number|String} number
 * @returns {String} properNumber
 */
function cutoffOverMonth(number){
  let properNumber = number;
  if(Number(number) > 12 ){
    const today = new Date();
    properNumber = today.getMonth()+1;
  }
  return properNumber;
}

/**
 * 31以上の数字を切り捨てる関数
 * @param {Number|String} number
 * @returns {String} properNumber
 */
function cutoffOverDate(number){
  let properNumber = number;
  if(Number(number) > 31 ){
    const today = new Date();
    properNumber = today.getDate();
  }
  return properNumber;
}