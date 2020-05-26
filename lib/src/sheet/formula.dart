part of excel;

class Formula {
  dynamic _evaluatedValue;
  String _formulaString;

  Formula._(dynamic evaluatedValue, String formula) {
    this._evaluatedValue = evaluatedValue;
    this._formulaString = formula;
  }

  static Formula abs(Sheet sheet, CellIndex cellIndex) {
    _checkCellIndex(sheet, cellIndex);
    ///// process the values here and do the operations
    return Formula._('ERROR', '=IsSUM()');
  }

  static _checkCellIndex(Sheet sheet, CellIndex cellIndex) {
    if (cellIndex == null ||
        cellIndex.columnIndex < 0 ||
        cellIndex.rowIndex < 0) {
      _damagedExcel(text: 'dirty CellIndex');
    }
    sheet._checkMaxCol(cellIndex.columnIndex);
    sheet._checkMaxRow(cellIndex.rowIndex);
  }

  /// get evaluated String of Formula
  get value {
    return this._evaluatedValue;
  }

  /// get Formula
  get formula {
    return this._formulaString;
  }
}
