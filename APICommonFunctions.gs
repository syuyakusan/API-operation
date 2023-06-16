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
    let length = matrix[i].filter(value => value).length;
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