
//先頭の値を返す
function getAttr(x){
  return sheet.getRange(1,x).getValue();
}

//keyの値を返す
function getKey(y){
  return sheet.getRange(y,1).getValue();
}

//指定された行の指定された属性の値を返す。
function getAttrValueCell(y, attr){
  var attrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var x = attrs.indexOf(attr);
  
  if (x>-1) {
    return sheet.getRange(y,x+1);
  }
}