
//先頭の値を返す
function getAttr(sheet, x){
  return sheet.getRange(1,x).getValue();
}

//keyの値を返す
function getKey(sheet, y){
  return sheet.getRange(y,1).getValue();
}

//指定された行の指定された属性の値を返す。
function getAttrValueCell(sheet, y, attr){
  var x = getAttrX(sheet, attr);
  
  if (x>0) {
    return sheet.getRange(y,x);
  }
}

//指定された属性の位置を返す。
function getAttrX(sheet, attr){
  var attrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//  Logger.log(attrs);
  var x = attrs.indexOf(attr);
  return x+1;
}


//指定された属性の値の位置を1つ見つける
//TODO: selectなどに拡張した方がいい?
function getY(sheet, attr, value){
  var x = getAttrX(sheet, attr);
  
  if (x>0){
    var values = sheet.getRange(1, x, sheet.getLastRow(), 1).getValues();
    for (var y=0; y<values.length; y++){
      if (values[y][0] == value){
        return y+1;
      }
    }
  }
  
  return -1; //not found
}