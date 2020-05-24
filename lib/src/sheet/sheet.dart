part of excel;

class Sheet {
  Excel _excel;
  String _sheet;
  int _maxRows = 0;
  int _maxCols = 0;
  Map<int, Map<int, Data>> _sheetData = Map<int, Map<int, Data>>();

  Sheet(Excel excel, String sheetName, {Map<int, Map<int, Data>> sh}) {
    this._sheetData = sh == null ? Map<int, Map<int, Data>>() : sh;
    this._excel = excel;
    this._sheet = sheetName;
  }

  Data cell(CellIndex cellIndex) {
    if (cellIndex._columnIndex < 0 || cellIndex._rowIndex < 0) {
      _damagedExcel(
          text:
              '${cellIndex._columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex._columnIndex < 0 ? cellIndex._columnIndex : cellIndex._rowIndex} Negative index is not accepted.');
    }

    /// increasing the rowCount
    if (this._maxRows < (cellIndex._rowIndex + 1)) {
      this._maxRows = cellIndex._rowIndex + 1;
    }

    /// increasing the colCount
    if (this._maxCols < (cellIndex._columnIndex + 1)) {
      this._maxCols = cellIndex._columnIndex + 1;
    }

    /// checking if the map has been already initialized or not?
    /// if the user has called this class by its own
    if (!_isContain(this._sheetData)) {
      this._sheetData = Map<int, Map<int, Data>>();
    }

    /// if the sheetData contains the row then start putting the column
    if (_isContain(this._sheetData[cellIndex._rowIndex])) {
      this._sheetData[cellIndex._rowIndex][cellIndex._columnIndex] =
          Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex);
    } else {
      /// else put the column with map showing.
      this._sheetData[cellIndex._rowIndex] = {
        cellIndex._columnIndex:
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex)
      };
    }

    return this._sheetData[cellIndex._rowIndex][cellIndex._columnIndex];
  }

  List<List<Data>> get loadCells {
    List<List<Data>> _data = List<List<Data>>();

    if (!_isContain(this._sheetData) || this._sheetData.isEmpty) {
      return _data;
    }

    if (this._maxRows > 0 && this.maxCols > 0) {
      _data = List.generate(this._maxRows, (rowIndex) {
        return List.generate(this._maxCols, (colIndex) {
          if (_isContain(this._sheetData[rowIndex]) &&
              _isContain(this._sheetData[rowIndex][colIndex])) {
            return this._sheetData[rowIndex][colIndex];
          }
          return Data.newData(this, rowIndex, colIndex);
        });
      });
    }

    return _data;
  }

  /// remove Col
  removeCol(int colIndex) {
    if (colIndex < 0) {
      return;
    }
    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (colIndex <= this.maxCols - 1) {
      /// do the shifting task
      List<int> sortedKeys = this._sheetData.keys.toList()..sort();
      sortedKeys.forEach((rowKey) {
        Map<int, Data> colMap = Map<int, Data>();
        List<int> sortedColKeys = this._sheetData[rowKey].keys.toList()..sort();
        sortedColKeys.forEach((colKey) {
          if (_isContain(this._sheetData[rowKey]) &&
              _isContain(this._sheetData[rowKey][colKey])) {
            if (colKey < colIndex) {
              colMap[colKey] = this._sheetData[rowKey][colKey];
            }
            if (colIndex == colKey) {
              this._sheetData[rowKey].remove(colKey);
            }
            if (colIndex < colKey) {
              colMap[colKey - 1] = this._sheetData[rowKey][colKey];
              this._sheetData[rowKey].remove(colKey);
            }
          }
        });
        _data[rowKey] = Map<int, Data>.from(colMap);
      });
      this._sheetData = Map<int, Map<int, Data>>.from(_data);
    }
  }

  /// remove Row
  removeRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }
    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (rowIndex <= this.maxRows - 1) {
      /// do the shifting task
      List<int> sortedKeys = this._sheetData.keys.toList()..sort();
      sortedKeys.forEach((rowKey) {
        if (rowKey < rowIndex && _isContain(this._sheetData[rowKey])) {
          _data[rowKey] = Map<int, Data>.from(this._sheetData[rowKey]);
        }
        if (rowIndex == rowKey && _isContain(this._sheetData[rowKey])) {
          this._sheetData.remove(rowKey);
        }
        if (rowIndex < rowKey && _isContain(this._sheetData[rowKey])) {
          _data[rowKey - 1] = Map<int, Data>.from(this._sheetData[rowKey]);
          this._sheetData.remove(rowKey);
        }
      });
      this._sheetData = Map<int, Map<int, Data>>.from(_data);
    }
  }

  /// get sheetName
  String get sheetName {
    return this._sheet;
  }

  /// List of table's rows
  List<Data> row(int rowIndex) {
    if (rowIndex < 0) {
      return null;
    }
    if (_isContain(this._sheetData) && _isContain(this._sheetData[rowIndex])) {
      return List.generate(this.maxCols, (colIndex) {
        if (_isContain(this._sheetData[rowIndex][colIndex])) {
          return this._sheetData[rowIndex][colIndex];
        }
        return Data.newData(this, rowIndex, colIndex);
      });
    }
    return [];
  }

  /// Get max rows
  int get maxRows => this._maxRows;

  /// Get max cols
  int get maxCols => this._maxCols;
}
/* 
/// Table of a excel file
class DataTable {
  final String name;
  DataTable(this.name);

  int _maxRows = 0, _maxCols = 0;




  /// Get max cols
  int get maxCols => _maxCols;
}
 */

/* 
  /// change sheetName
  set sheetName(String newSheetName) {
    if(_isContain(this._excel._sheetMap) && _isContain(this._excel._sheetMap[])){

    }
    
    this._sheet = newSheetName;
  } */
