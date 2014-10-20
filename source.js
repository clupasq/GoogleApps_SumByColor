function _getSpreadsheet(key) {
  return key ? SpreadsheetApp.openById(key) : SpreadsheetApp.getActiveSpreadsheet();
}

function getBackgroundColor(rangeSpecification, key) {
  var sheet = _getSpreadsheet(key);
  return sheet.getRange(rangeSpecification).getBackground();
}

function getForegroundColor(rangeSpecification, key) {
  var sheet = _getSpreadsheet(key);
  return sheet.getRange(rangeSpecification).getFontColor();
}


function sumWhereBackgroundColorIs(color, rangeSpecification, key) {
  var condition = function (cell) { return cell.getBackground().toLowerCase() == color.toLowerCase(); };
  return sumByCondition(rangeSpecification, condition, key);
}

/**
 * Sums all the values of cells with a given background color.
 *
 * @param {color}
 * @param {rangeSpecification} rangeSpecification - the range to search against.
 * @param {key} key (optional) - the key of a remote spreadsheet. If provided the
 *        range lookup will be done remotely.
 * @return A sum of the cell values.
 * @customfunction
 */
function sumWhereBackgroundColorIsNot(color, rangeSpecification, key) {
  var condition = function (cell) {
    return cell.getBackground().toLowerCase() != color.toLowerCase();
  };
  return sumByCondition(rangeSpecification, condition, key);
}

function sumWhereForegroundColorIs(color, rangeSpecification, key) {
  var condition = function (cell) { return cell.getFontColor().toLowerCase() == color.toLowerCase(); };
  return sumByCondition(rangeSpecification, condition, key);
}

function sumWhereForegroundColorIsNot(color, rangeSpecification, key) {
  var condition = function (cell) { return cell.getFontColor().toLowerCase() != color.toLowerCase(); };
  return sumByCondition(rangeSpecification, condition, key);
}



/**
 * Sums all the values of cells in the given range that has a
 * specific condition.
 *
 * @param {rangeSpecification} rangeSpecification - the range to search against.
 * @param {condition} condition - a function that determines if the cell should be
 *        summed or not.
 * @param {key} key (optional) - the key of a remote spreadsheet. If provided the
 *        range lookup will be done remotely.
 * @return A sum of the cell values.
 * @customfunction
 */
function sumByCondition(rangeSpecification, condition, key) {
  var sheet = _getSpreadsheet(key);
  var range = sheet.getRange(rangeSpecification);
  
  var sum = 0;
  
  for (var i = 1; i <= range.getNumRows(); i++) {
    for (var j = 1; j <= range.getNumColumns(); j++) {
      
      var cell = range.getCell(i, j);
      
      if(condition(cell))
        sum += parseFloat(cell.getValue() || 0);
    }
  }
  
  return sum;
}
