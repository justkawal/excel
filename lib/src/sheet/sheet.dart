part of excel;

class Sheet {
  Excel _excel;
  String _sheet;
  int _maxRows = 0;
  int _maxCols = 0;
  List<String> _spannedItems = List<String>();
  List<_Span> _spanList = List<_Span>();
  Map<int, Map<int, Data>> _sheetData = Map<int, Map<int, Data>>();

  Sheet(Excel excel, String sheetName, {Map<int, Map<int, Data>> sh}) {
    this._sheetData = sh ?? Map<int, Map<int, Data>>();
    this._spanList = List<_Span>();
    this._excel = excel;
    this._sheet = sheetName;
  }

  Data cell(CellIndex cellIndex) {
    _checkMaxCol(cellIndex.columnIndex);
    _checkMaxRow(cellIndex.rowIndex);
    if (cellIndex._columnIndex < 0 || cellIndex._rowIndex < 0) {
      _damagedExcel(
          text:
              '${cellIndex._columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex._columnIndex < 0 ? cellIndex._columnIndex : cellIndex._rowIndex} Negative index does not exist.');
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
    _countRowAndCol();

    return this._sheetData[cellIndex._rowIndex][cellIndex._columnIndex];
  }

  List<List<Data>> get cells {
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

  _countRowAndCol() {
    int maximumColIndex = -1, maximumRowIndex = -1;
    if (_isContain(this._sheetData)) {
      List<int> sortedKeys = this._sheetData.keys.toList()..sort();
      sortedKeys.forEach((rowKey) {
        if (_isContain(this._sheetData[rowKey]) &&
            this._sheetData[rowKey].isNotEmpty) {
          List<int> keys = this._sheetData[rowKey].keys.toList()..sort();
          if (keys != null && keys.isNotEmpty && keys.last > maximumColIndex) {
            maximumColIndex = keys.last;
          }
        }
      });

      if (sortedKeys != null && sortedKeys.isNotEmpty) {
        maximumRowIndex = sortedKeys.last;
      }
    }

    this._maxCols = maximumColIndex + 1;
    this._maxRows = maximumRowIndex + 1;
  }

  /// remove Column
  removeColumn(int colIndex) {
    _checkMaxCol(colIndex);
    if (colIndex < 0 || colIndex >= this.maxCols) {
      return;
    }
    _checkSheetArguments();

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object
    if (_spanList != null) {
      for (int i = 0; i < _spanList.length; i++) {
        _Span spanObj = _spanList[i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        if (colIndex <= endColumn) {
          _Span newSpanObj = _Span();
          if (colIndex < startColumn) {
            startColumn -= 1;
          }
          endColumn -= 1;
          if (/* startColumn >= endColumn */
              (colIndex == (endColumn + 1)) &&
                  (colIndex ==
                      (colIndex < startColumn
                          ? startColumn + 1
                          : startColumn))) {
            this._spanList[i] = null;
          } else {
            newSpanObj._start = [startRow, startColumn];
            newSpanObj._end = [endRow, endColumn];
            this._spanList[i] = newSpanObj;
          }
          updateSpanCell = true;
          _excel._mergeChanges = true;
        }

        if (_spanList[i] != null) {
          String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
          if (!this._spannedItems.contains(rc)) {
            this._spannedItems.add(rc);
          }
        }
      }
      _cleanUpSpanMap();
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = this.sheetName;
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
    _countRowAndCol();
  }

  /// insert Column at index = [colIndex]
  insertColumn(int colIndex) {
    if (colIndex < 0) {
      return;
    }
    _checkMaxCol(colIndex);
    _checkSheetArguments();

    bool updateSpanCell = false;

    if (this._spanList != null) {
      this._spannedItems = List<String>();
      for (int i = 0; i < _spanList.length; i++) {
        _Span spanObj = _spanList[i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        if (colIndex <= endColumn) {
          _Span newSpanObj = _Span();
          if (colIndex <= startColumn) {
            startColumn += 1;
          }
          endColumn += 1;
          newSpanObj._start = [startRow, startColumn];
          newSpanObj._end = [endRow, endColumn];
          this._spanList[i] = newSpanObj;
          updateSpanCell = true;
          _excel._mergeChanges = true;
        }
        String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!_spannedItems.contains(rc)) {
          this._spannedItems.add(rc);
        }
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = this.sheetName;
    }

    if (_isContain(this._sheetData) && this._sheetData.isNotEmpty) {
      Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      List<int> sortedKeys = this._sheetData.keys.toList()..sort();
      if (colIndex <= this.maxCols - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          Map<int, Data> colMap = Map<int, Data>();

          /// getting the cols keys in descending order so as to shifting becomes easy
          List<int> sortedColKeys = this._sheetData[rowKey].keys.toList()
            ..sort((a, b) => b.compareTo(a));
          sortedColKeys.forEach((colKey) {
            if (_isContain(this._sheetData[rowKey]) &&
                _isContain(this._sheetData[rowKey][colKey])) {
              if (colKey < colIndex) {
                colMap[colKey] = this._sheetData[rowKey][colKey];
              }
              if (colIndex <= colKey) {
                colMap[colKey + 1] = this._sheetData[rowKey][colKey];
              }
            }
          });
          colMap[colIndex] = Data.newData(this, rowKey, colIndex);
          _data[rowKey] = Map<int, Data>.from(colMap);
        });
        this._sheetData = Map<int, Map<int, Data>>.from(_data);
      } else {
        /// just put the data in the very first available row and
        /// in the desired Column index only one time as we will be using less space on internal implementatoin
        /// and mock the user as if the 2-D list is being saved
        ///
        /// As when user calls DataObject.cells then we will output 2-D list - pretending.
        this._sheetData[sortedKeys.first][colIndex] =
            Data.newData(this, sortedKeys.first, colIndex);
      }
    } else {
      /// here simply just take the first row and put the columnIndex as the _sheetData was previously null
      this._sheetData = Map<int, Map<int, Data>>();
      this._sheetData[0] = {colIndex: Data.newData(this, 0, colIndex)};
    }

    _countRowAndCol();
  }

  /// remove Row
  removeRow(int rowIndex) {
    if (rowIndex < 0 || rowIndex >= this._maxRows) {
      return;
    }
    _checkMaxRow(rowIndex);
    _checkSheetArguments();

    bool updateSpanCell = false;

    if (_spanList != null) {
      for (int i = 0; i < _spanList.length; i++) {
        _Span spanObj = _spanList[i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        if (rowIndex <= endRow) {
          _Span newSpanObj = _Span();
          if (rowIndex < startRow) {
            startRow -= 1;
          }
          endRow -= 1;
          if (/* startRow >= endRow */
              (rowIndex == (endRow + 1)) &&
                  (rowIndex ==
                      (rowIndex < startRow ? startRow + 1 : startRow))) {
            _spanList[i] = null;
          } else {
            newSpanObj._start = [startRow, startColumn];
            newSpanObj._end = [endRow, endColumn];
            _spanList[i] = newSpanObj;
          }
          updateSpanCell = true;
          _excel._mergeChanges = true;
        }
        if (this._spanList[i] != null) {
          String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
          if (!this._spannedItems.contains(rc)) {
            this._spannedItems.add(rc);
          }
        }
      }
      _cleanUpSpanMap();
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = this.sheetName;
    }

    if (_isContain(this._sheetData) && this._sheetData.isNotEmpty) {
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
      _countRowAndCol();
    } else {
      this._maxRows = 0;
      this._maxCols = 0;
    }
  }

  /// insert Row at index = [rowIndex]
  insertRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }
    _checkSheetArguments();
    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    if (_isContain(this._sheetData)) {
      this._spannedItems = List<String>();
      for (int i = 0; i < _spanList.length; i++) {
        _Span spanObj = _spanList[i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        if (rowIndex <= endRow) {
          _Span newSpanObj = _Span();
          if (rowIndex <= startRow) {
            startRow += 1;
          }
          endRow += 1;
          newSpanObj._start = [startRow, startColumn];
          newSpanObj._end = [endRow, endColumn];
          this._spanList[i] = newSpanObj;
          updateSpanCell = true;
          _excel._mergeChanges = true;
        }
        String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!this._spannedItems.contains(rc)) {
          this._spannedItems.add(rc);
        }
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = this.sheetName;
    }

    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (_isContain(this._sheetData) && this._sheetData.isNotEmpty) {
      List<int> sortedKeys = this._sheetData.keys.toList()
        ..sort((a, b) => b.compareTo(a));
      if (rowIndex <= this.maxRows - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          if (rowKey < rowIndex) {
            _data[rowKey] = this._sheetData[rowKey];
          }
          if (rowIndex <= rowKey) {
            _data[rowKey + 1] = this._sheetData[rowKey];
          }
        });
      }
    }
    _data[rowIndex] = {0: Data.newData(this, rowIndex, 0)};
    this._sheetData = Map<int, Map<int, Data>>.from(_data);
    _countRowAndCol();
  }

  _updateSheetClassCell(CellIndex cellIndex, dynamic value,
      {CellStyle cellStyle}) {
    _checkSheetArguments();
    int columnIndex = cellIndex._columnIndex;
    int rowIndex = cellIndex._rowIndex;
    if (columnIndex < 0 || rowIndex < 0) {
      return;
    }
    _checkMaxCol(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;
    if (this._spanList != null && this._spanList.isNotEmpty) {
      List updatedPosition = _isInsideSpanning(rowIndex, columnIndex);
      newRowIndex = updatedPosition[0];
      newColumnIndex = updatedPosition[1];
    }

    if (!_isContain(this._sheetData)) {
      this._sheetData = Map<int, Map<int, Data>>();
    }
    if (value != null) {
      print(value.runtimeType.toString());
    }

    if (this._sheetData.isNotEmpty && _isContain(this._sheetData[rowIndex])) {
      if (!_isContain(this._sheetData[rowIndex][columnIndex])) {
        this._sheetData[rowIndex][columnIndex] =
            Data.newData(this, rowIndex, columnIndex);
      }
      this._sheetData[rowIndex][columnIndex]._value =
          this._sheetData[rowIndex][columnIndex]._value ?? value;
      this._sheetData[rowIndex][columnIndex]._cellStyle =
          this._sheetData[rowIndex][columnIndex]._cellStyle ?? cellStyle;
    } else {
      this._sheetData[rowIndex] = {
        columnIndex: Data.newData(this, rowIndex, columnIndex)
      };
    }

    if (!_excel._sharedStrings.contains('$value')) {
      _excel._sharedStrings.add(value.toString());
    }
  }

  merge() {}

  /// returns true if the contents are successfully cleared else false
  bool clearRow(int rowIndex) {
    if (rowIndex < 0) {
      return false;
    }

    /// lets assume that this row is already cleared and is not inside spanList
    /// If this row exists then we check for the span condition
    bool isNotInside = true;

    if (_isContain(this._sheetData) &&
        _isContain(this._sheetData[rowIndex]) &&
        this._sheetData[rowIndex].isNotEmpty) {
      /// lets start iterating the spanList and check that if the row is inside the spanList or not
      /// we will expect that value of isNotInside should not be changed to false
      /// If it changes to false then we can't clear this row as it is inside the spanned Cells
      if (this._spanList != null) {
        for (int i = 0; i < this._spanList.length; i++) {
          _Span spanObj = this._spanList[i];
          if (rowIndex >= spanObj.rowSpanStart &&
              rowIndex <= spanObj.rowSpanEnd) {
            isNotInside = false;
            break;
          }
        }
      }

      /// As the row is not inside any SpanList so we can easily clear its content.
      if (isNotInside) {
        this._sheetData[rowIndex].keys.toList().forEach((key) {
          /// Main concern here is to [clear] and [not delete] the contents inside the cell
          this._sheetData[rowIndex][key] = Data.newData(this, rowIndex, key);
        });
      }
    }
    _countRowAndCol();
    return isNotInside;
  }

  List<int> _isInsideSpanning(int rowIndex, int columnIndex) {
    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    if (this._spanList != null && this._spanList.isNotEmpty) {
      for (int i = 0; i < this._spanList.length; i++) {
        _Span spanObj = this._spanList[i];

        if (rowIndex >= spanObj.rowSpanStart &&
            rowIndex <= spanObj.rowSpanEnd &&
            columnIndex >= spanObj.columnSpanStart &&
            columnIndex <= spanObj.columnSpanEnd) {
          newRowIndex = spanObj.rowSpanStart;
          newColumnIndex = spanObj.columnSpanStart;
          break;
        }
      }
    }

    return [newRowIndex, newColumnIndex];
  }

  /// Check if columnIndex is not out of Excel Column limits.
  _checkMaxCol(int colIndex) {
    if ((this._maxCols >= 16384) || colIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
  }

  /// Check if rowIndex is not out of Excel Row limits.
  _checkMaxRow(int rowIndex) {
    if ((this._maxRows >= 1048576) || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
  }

  _checkSheetArguments() {
    if (!_excel._update) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }
  }

  // Cleaning up the null values from the Span Map
  _cleanUpSpanMap() {
    if (this._spanList != null &&
        this.sheetName != null &&
        this._spanList.isNotEmpty) {
      this._spanList.removeWhere((value) => value == null);
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
