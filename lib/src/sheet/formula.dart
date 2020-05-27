part of excel;

class Formula {
  dynamic _value;
  String _formula;

  Formula._(dynamic evaluatedValue, String formula) {
    this._value = evaluatedValue;
    this._formula = formula;
  }

  static Formula abs(Sheet sheet, CellIndex cellIndex) {
    dynamic value = "0";
    String formulae = cellIndex.cellId;
    if (_checkCellIndex(sheet, cellIndex, checkForward: true)) {
      Data data = _data(sheet, cellIndex);
      dynamic val = _tryParse(data.value.toString());

      /// if the obtained value is number the do the processing
      value = val ?? "Error";
    }

    return Formula._(value, '=ABS($formulae)');
  }

  /// returns the average.
  static Formula average(Sheet sheet, List<CellIndex> cellIndexList) {
    List<String> cellIdList = List<String>();
    dynamic value = "#DIV/0!";
    Map<String, dynamic> map = _sumInternally(sheet, cellIndexList);
    bool evaluate = map["evaluate"] == 1;
    cellIdList = List<String>.from(map["list"]);
    int sum = map["sum"];
    if (evaluate) {
      if (cellIndexList.length > 0) {
        value = sum / cellIndexList.length;
      }
    } else {
      value = "#VALUE!";
    }

    return Formula._(value, '=AVERAGE(${cellIdList.join(',')})');
  }

  /// returns the sum.
  ///
  /// pass the cellIndexList as ["A1","B3:B6","D4:G6"]
  static Formula sum(Sheet sheet, List<CellIndex> cellIndexList) {
    List<String> cellIdList = List<String>();
    dynamic value = "#DIV/0!";
    Map<String, dynamic> map = _sumInternally(sheet, cellIndexList);
    bool evaluate = map["evaluate"] == 1;
    int sum = map["sum"];
    if (evaluate) {
      if (cellIndexList.length > 0) {
        value = sum / cellIndexList.length;
      }
    } else {
      value = "#VALUE!";
    }

    return Formula._(value, '=AVERAGE(${cellIdList.join(',')})');
  }

  /// returns three mapped values
  /// dirty is set to 1 if not able to evaluate
  /// sum is the total sum got from the cells
  /// list is the cellIndexList in the format of CellId as A1 or B90 ...
  static Map<String, dynamic> _sumInternally(
      Sheet sheet, List<CellIndex> cellIndexList) {
    List<String> cellIdList = List<String>();
    int sum = 0;
    bool evaluate = true;
    for (int i = 0; i < cellIndexList.length; i++) {
      CellIndex cellIndex = cellIndexList[i];
      cellIdList.add(cellIndex.cellId);
      if (evaluate) {
        if (_checkCellIndex(sheet, cellIndex, checkForward: true)) {
          Data data = _data(sheet, cellIndex);
          dynamic val = num.tryParse(data.value.toString());

          /// if the obtained value is number then do the processing
          if (val == null) {
            evaluate = false;
          } else {
            sum += val;
          }
        } else {
          sum += 0;
        }
      }
    }
    return {"evaluate": evaluate ? 1 : 0, "sum": sum, "list": cellIdList};
  }
/* 
  static List<CellIndex> _expandCellIndexList(List<String> cellIndexList) {
    List<CellIndex> cellList = List<CellIndex>();
    if (cellIndexList != null) {
      cellIndexList.forEach((cells) {
        if (cells != null) {
          if (cells.contains(":")) {
            List<String> cList = cells.split(':');
            if (cList.length == 2) {
              cList.forEach((element) {
                if (element != null) {
                  cellList.add(CellIndex.indexByString(element));
                }
              });
            }
          } else {
            cellList.add(CellIndex.indexByString(cells));
          }
        }
      });
    }
    return cellList;
  } */

  /// try to parse the integer or double and return null if it is not integer or double
  static num _tryParse(String input) {
    String source = input.trim();
    return int.tryParse(source) ?? double.tryParse(source);
  }

  /// return the Data object so as to extract out the value
  static Data _data(Sheet sheet, CellIndex cellIndex) {
    return sheet._sheetData[cellIndex.rowIndex][cellIndex.columnIndex];
  }

  /// check various important things
  static bool _checkCellIndex(Sheet sheet, CellIndex cellIndex,
      {bool checkForward = false}) {
    if (sheet == null) {
      _damagedExcel(text: 'null sheet reference found.');
    }
    if (cellIndex == null ||
        cellIndex.columnIndex < 0 ||
        cellIndex.rowIndex < 0) {
      _damagedExcel(text: 'dirty or null CellIndex found.');
    }
    sheet._checkMaxCol(cellIndex.columnIndex);
    sheet._checkMaxRow(cellIndex.rowIndex);
    if (checkForward) {
      return sheet._sheetData.isNotEmpty &&
          _isContain(sheet._sheetData) &&
          _isContain(sheet._sheetData[cellIndex.rowIndex]) &&
          _isContain(
              sheet._sheetData[cellIndex.rowIndex][cellIndex.columnIndex]);
    }
    return true;
  }

  /// get evaluated String of Formula
  get value {
    return this._value;
  }

  /// get Formula
  get formula {
    return this._formula;
  }
}
