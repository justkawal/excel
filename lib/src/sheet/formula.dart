part of excel;

class Formula {
  String _formula;

  Formula._(String formula) {
    this._formula = formula;
  }

  ///
  ///```
  ///var abs = Formula.abs(-3));
  ///```
  /* static Formula abs(Sheet sheet, dynamic val) {
    dynamic formulaValue = _getParsedVal(val);
    return Formula._('ABS($formulaValue)');
  } */

  /// Helps to initiate a custom formula
  ///```
  ///var my_custom_formula = Formula.custom('SUM(1,2)');
  ///```
  static Formula custom(String formula) {
    assert(formula != null);
    return Formula._(formula);
  }

  /// returns the average.
  ///
  /// pass the cellIndexList as ["A1","B3:B6","D4:G6", CellIndex.indexByString("A2")]
  ///```
  ///var cells = ["A1","B3:B6","D4:G6", CellIndex.indexByString("A2")];
  ///var average_formula = Formula.average(sheetObject, cells);
  ///```
  /* static Formula average(Sheet sheet, List<dynamic> values) {
    List<dynamic> cellIdList = _getParsedList(sheet, values);
    return Formula._('AVERAGE(${cellIdList.join(',')})');
  } */

  /// returns the sum.
  ///
  /// pass the cellIndexList as ["A1","B3:B6","D4:G6", CellIndex.indexByString("A2")]
  ///```
  ///var cells = ["A1","B3:B6","D4:G6", CellIndex.indexByString("A2")];
  ///var sum_formula = Formula.sum(sheetObject, cells);
  ///```
  /* static Formula sum(Sheet sheet, List<dynamic> values) {
    List<dynamic> cellIdList = _getParsedList(sheet, values);
    return Formula._('SUM(${cellIdList.join(',')})');
  } */

  /***************************** Fomula Utilities *****************************/

  /* static dynamic _getParsedVal(dynamic val) {
    if (val is CellIndex) {
      return val.cellId;
    } else if (val is Formula) {
      return val._formula;
    }
    return val;
  } */

  /// returns three mapped values
  /// sum is the total sum got from the cells
  /// list is the cellIndexList in the format of CellId as A1 or B90 ... or it can be values of formulas
  /* static List<dynamic> _getParsedList(
      Sheet sheet, List<dynamic> cellIndexList) {
    List<dynamic> list = List<String>();
    for (var val in cellIndexList) {
      if (val is CellIndex) {
        list.add(val.cellId);
      } else if (val is String) {
        if (val.contains(':')) {
          var cells = val.split(':');
          if (cells.length == 2) {
            var l = _expandCellIndexList(
                    sheet,
                    CellIndex.indexByString('${cells[0]}'),
                    CellIndex.indexByString('${cells[1]}'))
                .map((e) => e.cellId)
                .toList();
            list.addAll(l);
          } else {
            list.addAll(
                cells.map((e) => CellIndex.indexByString(e.toString())));
          }
        } else {
          list.add(val);
        }
      } else if (val is Formula) {
        list.add(val._formula);
      } else {
        list.add(val);
      }
    }
    return list;
  }

  static List<CellIndex> _expandCellIndexList(
      Sheet sheet, CellIndex start, CellIndex end) {
    List<CellIndex> indexList = <CellIndex>[];
    if (start != null && end != null) {
      sheet._checkMaxCol(start.columnIndex);
      sheet._checkMaxRow(start.rowIndex);
      sheet._checkMaxCol(end.columnIndex);
      sheet._checkMaxRow(end.rowIndex);
      for (var i = start.rowIndex; i <= end.rowIndex; i++) {
        for (var j = start.columnIndex; j <= end.columnIndex; j++) {
          indexList
              .add(CellIndex.indexByColumnRow(columnIndex: j, rowIndex: i));
        }
      }
    }
    return indexList;
  } */

  /// get Formula
  get formula {
    return this._formula;
  }
}
