function getBackgroundColor(rangeSpecification) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  return sheet.getRange(rangeSpecification).getBackground();
}

function getForegroundColor(rangeSpecification) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  return sheet.getRange(rangeSpecification).getFontColor();
}


function sumWhereBackgroundColorIs(color, rangeSpecification) {
  var condition = function (cell) { return cell.getBackground() == color; };
  return sumByCondition(color, rangeSpecification, condition);
}

function sumWhereBackgroundColorIsNot(color, rangeSpecification) {
  var condition = function (cell) { return cell.getBackground() != color; };
  return sumByCondition(color, rangeSpecification, condition);
}

function sumWhereForegroundColorIs(color, rangeSpecification) {
  var condition = function (cell) { return cell.getFontColor() == color; };
  return sumByCondition(color, rangeSpecification, condition);
}

function sumWhereForegroundColorIsNot(color, rangeSpecification) {
  var condition = function (cell) { return cell.getFontColor() != color; };
  return sumByCondition(color, rangeSpecification, condition);
}




function sumByCondition(color, rangeSpecification, condition) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var range = sheet.getRange(rangeSpecification);
  
  var sum = 0;
  
  for (var i = 1; i <= range.getNumRows(); i++) {
    for (var j = 1; j <= range.getNumColumns(); j++) {
      
      var cell = range.getCell(i, j);
      
      if(condition(cell))
        sum += parseFloat(cell.getValue());
    }
  }
  
  return sum;
}