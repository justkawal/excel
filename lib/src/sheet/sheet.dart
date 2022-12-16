part of excel;

class Sheet {
  late Excel _excel;
  late String _sheet;
  late bool _isRTL;
  late int _maxRows;
  late int _maxCols;
  List<double> _colWidth = <double>[];
  List<bool> _colAutoFit = <bool>[];
  late FastList<String> _spannedItems;
  late List<_Span?> _spanList;
  late Map<int, Map<int, Data>> _sheetData;
  late HeaderFooter? _headerFooter;

  ///
  /// It will clone the object by changing the `this` reference of previous oldSheetObject and putting `new this` reference, with copying the values too
  ///
  Sheet._clone(Excel excel, String sheetName, Sheet oldSheetObject)
      : this._(excel, sheetName,
            sh: oldSheetObject._sheetData,
            spanL_: oldSheetObject._spanList,
            spanI_: oldSheetObject._spannedItems,
            maxRowsVal: oldSheetObject._maxRows,
            maxColsVal: oldSheetObject._maxCols,
            colWidthVal: oldSheetObject._colWidth,
            colAutoFitVal: oldSheetObject._colAutoFit,
            isRTLVal: oldSheetObject._isRTL,
            headerFooter: oldSheetObject._headerFooter);

  Sheet._(Excel excel, String sheetName,
      {Map<int, Map<int, Data>>? sh,
      List<_Span?>? spanL_,
      FastList<String>? spanI_,
      int? maxRowsVal,
      int? maxColsVal,
      bool? isRTLVal,
      List<double>? colWidthVal,
      List<bool>? colAutoFitVal,
      HeaderFooter? headerFooter}) {
    _excel = excel;
    _sheet = sheetName;
    _sheetData = <int, Map<int, Data>>{};
    _spanList = <_Span?>[];
    _spannedItems = FastList<String>();
    _isRTL = false;
    _maxRows = 0;
    _maxCols = 0;
    _headerFooter = headerFooter;

    if (spanL_ != null) {
      _spanList = List<_Span?>.from(spanL_);
      _excel._mergeChangeLookup = sheetName;
    }
    if (spanI_ != null) {
      _spannedItems = FastList<String>.from(spanI_);
    }
    if (maxColsVal != null) {
      _maxCols = maxColsVal;
    }
    if (maxRowsVal != null) {
      _maxRows = maxRowsVal;
    }
    if (isRTLVal != null) {
      _isRTL = isRTLVal;
      _excel._rtlChangeLookup = sheetName;
    }
    if (colWidthVal != null) {
      _colWidth = List<double>.from(colWidthVal);
    }
    if (colAutoFitVal != null) {
      _colAutoFit = List<bool>.from(colAutoFitVal);
    }

    /// copy the data objects into a temp folder and then while putting it into `_sheetData` change the data objects references.
    if (sh != null) {
      _sheetData = <int, Map<int, Data>>{};
      Map<int, Map<int, Data>> temp = Map<int, Map<int, Data>>.from(sh);
      temp.forEach((key, value) {
        if (_sheetData[key] == null) {
          _sheetData[key] = <int, Data>{};
        }
        temp[key]!.forEach((key1, oldDataObject) {
          _sheetData[key]![key1] = Data._clone(this, oldDataObject);
        });
      });
    }
    _countRowAndCol();
  }

  ///
  /// returns `true` is this sheet is `right-to-left` other-wise `false`
  ///
  bool get isRTL {
    return _isRTL;
  }

  ///
  /// set sheet-object to `true` for making it `right-to-left` otherwise `false`
  ///
  set isRTL(bool _u) {
    _isRTL = _u;
    _excel._rtlChangeLookup = sheetName;
  }

  ///
  /// returns the `DataObject` at position of `cellIndex`
  ///
  Data cell(CellIndex cellIndex) {
    _checkMaxCol(cellIndex.columnIndex);
    _checkMaxRow(cellIndex.rowIndex);
    if (cellIndex._columnIndex < 0 || cellIndex._rowIndex < 0) {
      _damagedExcel(
          text:
              '${cellIndex._columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex._columnIndex < 0 ? cellIndex._columnIndex : cellIndex._rowIndex} Negative index does not exist.');
    }

    /// increasing the rowCount
    if (_maxRows < (cellIndex._rowIndex + 1)) {
      _maxRows = cellIndex._rowIndex + 1;
    }

    /// increasing the colCount
    if (_maxCols < (cellIndex._columnIndex + 1)) {
      _maxCols = cellIndex._columnIndex + 1;
    }

    /// checking if the map has been already initialized or not?
    /// if the user has called this class by its own
    /* if (_sheetData == null) {
      _sheetData = Map<int, Map<int, Data>>();
    } */

    /// if the sheetData contains the row then start putting the column
    if (_sheetData[cellIndex._rowIndex] != null) {
      if (_sheetData[cellIndex._rowIndex]![cellIndex._columnIndex] == null) {
        _sheetData[cellIndex._rowIndex]![cellIndex._columnIndex] =
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex);
      }
    } else {
      /// else put the column with map showing.
      _sheetData[cellIndex._rowIndex] = {
        cellIndex._columnIndex:
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex)
      };
    }

    return _sheetData[cellIndex._rowIndex]![cellIndex._columnIndex]!;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements
  ///
  List<List<Data?>> get rows {
    var _data = <List<Data?>>[];

    if (_sheetData.isEmpty) {
      return _data;
    }

    if (_maxRows > 0 && maxCols > 0) {
      _data = List.generate(_maxRows, (rowIndex) {
        return List.generate(_maxCols, (colIndex) {
          if (_sheetData[rowIndex] != null &&
              _sheetData[rowIndex]![colIndex] != null) {
            return _sheetData[rowIndex]![colIndex];
          }
          return null;
        });
      });
    }

    return _data;
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<Data?>?> selectRangeWithString(String range) {
    List<List<Data?>?> _selectedRange = <List<Data?>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      _selectedRange = selectRange(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      _selectedRange = selectRange(start, end: end);
    }
    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet cell data in that range.
  ///
  List<List<Data?>?> selectRange(CellIndex start, {CellIndex? end}) {
    _checkMaxCol(start.columnIndex);
    _checkMaxRow(start.rowIndex);
    if (end != null) {
      _checkMaxCol(end.columnIndex);
      _checkMaxRow(end.rowIndex);
    }

    int _startColumn = start.columnIndex, _startRow = start.rowIndex;
    int? _endColumn = end?.columnIndex, _endRow = end?.rowIndex;

    if (_endColumn != null && _endRow != null) {
      if (_startRow > _endRow) {
        _startRow = end!.rowIndex;
        _endRow = start.rowIndex;
      }
      if (_endColumn < _startColumn) {
        _endColumn = start.columnIndex;
        _startColumn = end!.columnIndex;
      }
    }

    List<List<Data?>?> _selectedRange = <List<Data?>?>[];
    if (_sheetData.isEmpty) {
      return _selectedRange;
    }

    for (var i = _startRow; i <= (_endRow ?? maxRows); i++) {
      var mapData = _sheetData[i];
      if (mapData != null) {
        List<Data?> row = <Data?>[];
        for (var j = _startColumn; j <= (_endColumn ?? maxCols); j++) {
          row.add(mapData[j]);
        }
        _selectedRange.add(row);
      } else {
        _selectedRange.add(null);
      }
    }

    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  /// Ex. selectRange('D8:H12'); or selectRange('D8');
  ///
  List<List<dynamic>?> selectRangeValuesWithString(String range) {
    List<List<dynamic>?> _selectedRange = <List<dynamic>?>[];
    if (!range.contains(':')) {
      var start = CellIndex.indexByString(range);
      _selectedRange = selectRangeValues(start);
    } else {
      var rangeVars = range.split(':');
      var start = CellIndex.indexByString(rangeVars[0]);
      var end = CellIndex.indexByString(rangeVars[1]);
      _selectedRange = selectRangeValues(start, end: end);
    }
    return _selectedRange;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements in that range.
  ///
  List<List<dynamic>?> selectRangeValues(CellIndex start, {CellIndex? end}) {
    var _list =
        (end == null ? selectRange(start) : selectRange(start, end: end));
    return _list
        .map((List<Data?>? e) =>
            e?.map((e1) => e1 != null ? e1.value : null).toList())
        .toList();
  }

  ///
  /// updates count of rows and cols
  ///
  _countRowAndCol() {
    int maximumColIndex = -1, maximumRowIndex = -1;
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    sortedKeys.forEach((rowKey) {
      if (_sheetData[rowKey] != null && _sheetData[rowKey]!.isNotEmpty) {
        List<int> keys = _sheetData[rowKey]!.keys.toList()..sort();
        if (keys.isNotEmpty && keys.last > maximumColIndex) {
          maximumColIndex = keys.last;
        }
      }
    });

    if (sortedKeys.isNotEmpty) {
      maximumRowIndex = sortedKeys.last;
    }

    _maxCols = maximumColIndex + 1;
    _maxRows = maximumRowIndex + 1;
  }

  ///
  /// If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(int colIndex) {
    _checkMaxCol(colIndex);
    if (colIndex < 0 || colIndex >= maxCols) {
      return;
    }

    bool updateSpanCell = false;

    /// Do the shifting of the cell Id of span Object

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
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
                    (colIndex < startColumn ? startColumn + 1 : startColumn))) {
          _spanList[i] = null;
        } else {
          newSpanObj._start = [startRow, startColumn];
          newSpanObj._end = [endRow, endColumn];
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }

      if (_spanList[i] != null) {
        String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!_spannedItems.contains(rc)) {
          _spannedItems.add(rc);
        }
      }
    }
    _cleanUpSpanMap();

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (colIndex <= maxCols - 1) {
      /// do the shifting task
      List<int> sortedKeys = _sheetData.keys.toList()..sort();
      sortedKeys.forEach((rowKey) {
        Map<int, Data> colMap = Map<int, Data>();
        List<int> sortedColKeys = _sheetData[rowKey]!.keys.toList()..sort();
        sortedColKeys.forEach((colKey) {
          if (_sheetData[rowKey] != null &&
              _sheetData[rowKey]![colKey] != null) {
            if (colKey < colIndex) {
              colMap[colKey] = _sheetData[rowKey]![colKey]!;
            }
            if (colIndex == colKey) {
              _sheetData[rowKey]!.remove(colKey);
            }
            if (colIndex < colKey) {
              colMap[colKey - 1] = _sheetData[rowKey]![colKey]!;
              _sheetData[rowKey]!.remove(colKey);
            }
          }
        });
        _data[rowKey] = Map<int, Data>.from(colMap);
      });
      _sheetData = Map<int, Map<int, Data>>.from(_data);
    }

    if (_maxCols - 1 <= colIndex) {
      _maxCols -= 1;
    }

    //_countRowAndCol();
  }

  ///
  /// Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  /// If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(int colIndex) {
    if (colIndex < 0) {
      return;
    }
    _checkMaxCol(colIndex);

    bool updateSpanCell = false;

    _spannedItems = FastList<String>();
    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
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
        _spanList[i] = newSpanObj;
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
      if (!_spannedItems.contains(rc)) {
        _spannedItems.add(rc);
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    if (_sheetData.isNotEmpty) {
      Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      List<int> sortedKeys = _sheetData.keys.toList()..sort();
      if (colIndex <= maxCols - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          Map<int, Data> colMap = Map<int, Data>();

          /// getting the cols keys in descending order so as to shifting becomes easy
          List<int> sortedColKeys = _sheetData[rowKey]!.keys.toList()
            ..sort((a, b) {
              return b.compareTo(a);
            });
          sortedColKeys.forEach((colKey) {
            if (_sheetData[rowKey] != null &&
                _sheetData[rowKey]![colKey] != null) {
              if (colKey < colIndex) {
                colMap[colKey] = _sheetData[rowKey]![colKey]!;
              }
              if (colIndex <= colKey) {
                colMap[colKey + 1] = _sheetData[rowKey]![colKey]!;
              }
            }
          });
          colMap[colIndex] = Data.newData(this, rowKey, colIndex);
          _data[rowKey] = Map<int, Data>.from(colMap);
        });
        _sheetData = Map<int, Map<int, Data>>.from(_data);
      } else {
        /// just put the data in the very first available row and
        /// in the desired Column index only one time as we will be using less space on internal implementatoin
        /// and mock the user as if the 2-D list is being saved
        ///
        /// As when user calls DataObject.cells then we will output 2-D list - pretending.
        _sheetData[sortedKeys.first]![colIndex] =
            Data.newData(this, sortedKeys.first, colIndex);
      }
    } else {
      /// here simply just take the first row and put the columnIndex as the _sheetData was previously null
      _sheetData = Map<int, Map<int, Data>>();
      _sheetData[0] = {colIndex: Data.newData(this, 0, colIndex)};
    }
    if (_maxCols - 1 <= colIndex) {
      _maxCols += 1;
    } else {
      _maxCols = colIndex + 1;
    }

    //_countRowAndCol();
  }

  ///
  /// If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(int rowIndex) {
    if (rowIndex < 0 || rowIndex >= _maxRows) {
      return;
    }
    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
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
                (rowIndex == (rowIndex < startRow ? startRow + 1 : startRow))) {
          _spanList[i] = null;
        } else {
          newSpanObj._start = [startRow, startColumn];
          newSpanObj._end = [endRow, endColumn];
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      if (_spanList[i] != null) {
        String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!_spannedItems.contains(rc)) {
          _spannedItems.add(rc);
        }
      }
    }
    _cleanUpSpanMap();

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    if (_sheetData.isNotEmpty) {
      Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        List<int> sortedKeys = _sheetData.keys.toList()..sort();
        sortedKeys.forEach((rowKey) {
          if (rowKey < rowIndex && _sheetData[rowKey] != null) {
            _data[rowKey] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
          if (rowIndex == rowKey && _sheetData[rowKey] != null) {
            _sheetData.remove(rowKey);
          }
          if (rowIndex < rowKey && _sheetData[rowKey] != null) {
            _data[rowKey - 1] = Map<int, Data>.from(_sheetData[rowKey]!);
            _sheetData.remove(rowKey);
          }
        });
        _sheetData = Map<int, Map<int, Data>>.from(_data);
      }
      //_countRowAndCol();
    } else {
      _maxRows = 0;
      _maxCols = 0;
    }

    if (_maxRows - 1 <= rowIndex) {
      _maxRows -= 1;
    }
  }

  ///
  /// Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  /// If `rowIndex == null` or `rowIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);

    bool updateSpanCell = false;

    _spannedItems = FastList<String>();
    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
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
        _spanList[i] = newSpanObj;
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
      if (!_spannedItems.contains(rc)) {
        _spannedItems.add(rc);
      }
    }

    if (updateSpanCell) {
      _excel._mergeChangeLookup = sheetName;
    }

    Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
    if (_sheetData.isNotEmpty) {
      List<int> sortedKeys = _sheetData.keys.toList()
        ..sort((a, b) {
          return b.compareTo(a);
        });
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          if (rowKey < rowIndex) {
            _data[rowKey] = _sheetData[rowKey]!;
          }
          if (rowIndex <= rowKey) {
            _data[rowKey + 1] = _sheetData[rowKey]!;
          }
        });
      }
    }
    _data[rowIndex] = {0: Data.newData(this, rowIndex, 0)};
    _sheetData = Map<int, Map<int, Data>>.from(_data);

    if (_maxRows - 1 <= rowIndex) {
      _maxRows += 1;
    } else {
      _maxRows = rowIndex + 1;
    }

    //_countRowAndCol();
  }

  ///
  /// Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  /// ----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  /// Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  /// If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(CellIndex cellIndex, dynamic value, {CellStyle? cellStyle}) {
    int columnIndex = cellIndex._columnIndex;
    int rowIndex = cellIndex._rowIndex;
    if (columnIndex < 0 || rowIndex < 0) {
      return;
    }
    _checkMaxCol(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    /// Check if this is lying in merged-cell cross-section
    /// If yes then get the starting position of merged cells
    if (_spanList.isNotEmpty) {
      List updatedPosition = _isInsideSpanning(rowIndex, columnIndex);
      newRowIndex = updatedPosition[0];
      newColumnIndex = updatedPosition[1];
    }

    /// Puts Data
    _putData(newRowIndex, newColumnIndex, value);

    /// Puts the cellStyle
    if (cellStyle != null) {
      _sheetData[newRowIndex]![newColumnIndex]!._cellStyle = cellStyle;
      _excel._colorChanges = true;
    }
  }

  ///
  /// Merges the cells starting from `start` to `end`.
  ///
  /// If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  merge(CellIndex start, CellIndex end, {dynamic customValue}) {
    int startColumn = start._columnIndex,
        startRow = start._rowIndex,
        endColumn = end._columnIndex,
        endRow = end._rowIndex;

    _checkMaxCol(startColumn);
    _checkMaxCol(endColumn);
    _checkMaxRow(startRow);
    _checkMaxRow(endRow);

    if ((startColumn == endColumn && startRow == endRow) ||
        (startColumn < 0 || startRow < 0 || endColumn < 0 || endRow < 0) ||
        (_spannedItems.contains(
            getSpanCellId(startColumn, startRow, endColumn, endRow)))) {
      return;
    }

    List<int> gotPosition = _getSpanPosition(start, end);

    _excel._mergeChanges = true;

    startColumn = gotPosition[0];
    startRow = gotPosition[1];
    endColumn = gotPosition[2];
    endRow = gotPosition[3];

    bool getValue = true;

    Data value = Data.newData(this, startRow, startColumn);
    if (customValue != null) {
      value._value = customValue;
      getValue = false;
    }

    for (int j = startRow; j <= endRow; j++) {
      for (int k = startColumn; k <= endColumn; k++) {
        if (_sheetData[j] != null && _sheetData[j]![k] != null) {
          if (getValue &&
              _sheetData[j]![k]!.value != null &&
              _sheetData[j]![k]!.cellStyle != null) {
            value = _sheetData[j]![k]!;
            getValue = false;
          }
          _sheetData[j]!.remove(k);
        }
      }
    }

    if (_sheetData[startRow] != null) {
      _sheetData[startRow]![startColumn] = value;
    } else {
      _sheetData[startRow] = {startColumn: value};
    }

    String sp = getSpanCellId(startColumn, startRow, endColumn, endRow);

    if (!_spannedItems.contains(sp)) {
      _spannedItems.add(sp);
    }

    _Span s = _Span();
    s._start = [startRow, startColumn];
    s._end = [endRow, endColumn];

    _spanList.add(s);
    _excel._mergeChangeLookup = sheetName;
  }

  ///
  /// unMerge the merged cells.
  ///
  ///        var sheet = 'DesiredSheet';
  ///        List<String> spannedCells = excel.getMergedCells(sheet);
  ///        var cellToUnMerge = "A1:A2";
  ///        excel.unMerge(sheet, cellToUnMerge);
  ///
  unMerge(String unmergeCells) {
    if (_spannedItems.isNotEmpty &&
        _spanList.isNotEmpty &&
        _spannedItems.contains(unmergeCells)) {
      List<String> lis = unmergeCells.split(RegExp(r":"));
      if (lis.length == 2) {
        bool remove = false;
        List<int> start, end;
        start =
            _cellCoordsFromCellId(lis[0]); // [x,y] => [startRow, startColumn]
        end = _cellCoordsFromCellId(lis[1]); // [x,y] => [endRow, endColumn]
        for (int i = 0; i < _spanList.length; i++) {
          _Span? spanObject = _spanList[i];
          if (spanObject == null) {
            continue;
          }

          if (spanObject.columnSpanStart == start[1] &&
              spanObject.rowSpanStart == start[0] &&
              spanObject.columnSpanEnd == end[1] &&
              spanObject.rowSpanEnd == end[0]) {
            _spanList[i] = null;
            remove = true;
          }
        }
        if (remove) {
          _cleanUpSpanMap();
        }
      }
      _spannedItems.remove(unmergeCells);
      _excel._mergeChangeLookup = sheetName;
    }
  }

  ///
  /// Helps to find the interaction between the pre-existing span position and updates if with new span if there any interaction(Cross-Sectional Spanning) exists.
  ///
  List<int> _getSpanPosition(CellIndex start, CellIndex end) {
    int startColumn = start._columnIndex,
        startRow = start._rowIndex,
        endColumn = end._columnIndex,
        endRow = end._rowIndex;

    bool remove = false;

    if (startRow > endRow) {
      startRow = end._rowIndex;
      endRow = start._rowIndex;
    }
    if (endColumn < startColumn) {
      endColumn = start._columnIndex;
      startColumn = end._columnIndex;
    }

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      List locationChange = _isLocationChangeRequired(
          startColumn, startRow, endColumn, endRow, spanObj);
      List<int> gotPosition = locationChange[1];

      if (locationChange[0]) {
        startColumn = gotPosition[0];
        startRow = gotPosition[1];
        endColumn = gotPosition[2];
        endRow = gotPosition[3];
        String sp = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
            spanObj.columnSpanEnd, spanObj.rowSpanEnd);
        if (_spannedItems.contains(sp)) {
          _spannedItems.remove(sp);
        }
        remove = true;
        _spanList[i] = null;
      }
    }
    if (remove) {
      _cleanUpSpanMap();
    }

    return [startColumn, startRow, endColumn, endRow];
  }

  ///
  /// Appends [row] iterables just post the last filled `rowIndex`.
  ///
  appendRow(List<dynamic> row) {
    int targetRow = maxRows;
    insertRowIterables(row, targetRow);
  }

  /// getting the List of _Span Objects which have the rowIndex containing and
  /// also lower the range by giving the starting columnIndex
  List<_Span> _getSpannedObjects(int rowIndex, int startingColumnIndex) {
    List<_Span> obtained = <_Span>[];

    if (_spanList.isNotEmpty) {
      obtained = <_Span>[];
      _spanList.forEach((spanObject) {
        if (spanObject != null &&
            spanObject.rowSpanStart <= rowIndex &&
            rowIndex <= spanObject.rowSpanEnd &&
            startingColumnIndex <= spanObject.columnSpanEnd) {
          obtained.add(spanObject);
        }
      });
    }
    return obtained;
  }

  ///
  /// Checking if the columnIndex and the rowIndex passed is inside the spanObjectList which is got from calling function.
  ///
  bool _isInsideSpanObject(
      List<_Span> spanObjectList, int columnIndex, int rowIndex) {
    for (int i = 0; i < spanObjectList.length; i++) {
      _Span spanObject = spanObjectList[i];

      if (spanObject.columnSpanStart <= columnIndex &&
          columnIndex <= spanObject.columnSpanEnd &&
          spanObject.rowSpanStart <= rowIndex &&
          rowIndex <= spanObject.rowSpanEnd) {
        if (columnIndex < spanObject.columnSpanEnd) {
          return false;
        } else if (columnIndex == spanObject.columnSpanEnd) {
          return true;
        }
      }
    }
    return true;
  }

  ///
  /// Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
  ///
  /// [startingColumn] tells from where we should start putting the [row] iterables
  ///
  /// [overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
  ///
  /// [overwriteMergedCells] when set to [false] puts the cell value in next unique cell available and putting the value in merged cells only once.
  ///
  insertRowIterables(List<dynamic> row, int rowIndex,
      {int startingColumn = 0, bool overwriteMergedCells = true}) {
    if (row.length == 0 || rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);
    int columnIndex = 0;
    if (startingColumn > 0) {
      columnIndex = startingColumn;
    }
    _checkMaxCol(columnIndex + row.length);
    int rowsLength = _maxRows,
        maxIterationIndex = row.length - 1,
        currentRowPosition = 0; // position in [row] iterables

    if (overwriteMergedCells || rowIndex >= rowsLength) {
      // Normally iterating and putting the data present in the [row] as we are on the last index.

      while (currentRowPosition <= maxIterationIndex) {
        _putData(rowIndex, columnIndex, row[currentRowPosition]);
        currentRowPosition++;
        columnIndex++;
      }
    } else {
      // expensive function as per time complexity
      _selfCorrectSpanMap(_excel);
      List<_Span> _spanObjectsList = _getSpannedObjects(rowIndex, columnIndex);

      if (_spanObjectsList.length <= 0) {
        while (currentRowPosition <= maxIterationIndex) {
          _putData(rowIndex, columnIndex, row[currentRowPosition]);
          currentRowPosition++;
          columnIndex++;
        }
      } else {
        while (currentRowPosition <= maxIterationIndex) {
          if (_isInsideSpanObject(_spanObjectsList, columnIndex, rowIndex)) {
            _putData(rowIndex, columnIndex, row[currentRowPosition]);
            currentRowPosition++;
          }
          columnIndex++;
        }
      }
    }
    //int tempo_max_col = columnIndex + row.length - 1;

    //if (_maxCols - 1 < tempo_max_col) {
    //  _maxCols = tempo_max_col + 1;
    //}
  }

  ///
  /// Internal function for putting the data in `_sheetData`.
  ///
  _putData(int rowIndex, int columnIndex, dynamic value) {
    if (_sheetData[rowIndex] != null) {
      if (_sheetData[rowIndex]![columnIndex] == null) {
        _sheetData[rowIndex]![columnIndex] =
            Data.newData(this, rowIndex, columnIndex);
      }
    } else {
      _sheetData[rowIndex] = {
        columnIndex: Data.newData(this, rowIndex, columnIndex)
      };
    }

    if (value is String) {
      final sharedString = _excel._sharedStrings.addFromString(value);
      _sheetData[rowIndex]![columnIndex]!._value = sharedString;
    } else {
      _sheetData[rowIndex]![columnIndex]!._value = value;
    }

    /// Sets value of `isFormula` to true if this is `instance of Formula`.
    _sheetData[rowIndex]![columnIndex]!._isFormula =
        value is Formula || value.runtimeType == Formula;

    /// Sets type of the Data to `_cellType`
    _sheetData[rowIndex]![columnIndex]!._cellType =
        _getCellType(value.runtimeType);

    if ((_maxCols - 1) < columnIndex) {
      _maxCols = columnIndex + 1;
    }

    if ((_maxRows - 1) < rowIndex) {
      _maxRows = rowIndex + 1;
    }

    //_countRowAndCol();
  }

  ///
  /// returns list of auto fit columns
  ///
  List<bool> get getColAutoFits {
    return _colAutoFit;
  }

  ///
  /// returns list of custom width columns
  ///
  List<double> get getColWidths {
    return _colWidth;
  }

  ///
  /// Get Column AutoFit
  ///
  bool getColAutoFit(int colIndex) {
    _checkMaxCol(colIndex);
    return _colAutoFit[colIndex];
  }

  ///
  /// Get Column Width
  ///
  double getColWidth(int colIndex) {
    _checkMaxCol(colIndex);
    return _colWidth[colIndex];
  }

  ///
  /// Set Column AutoFit
  ///
  void setColAutoFit(int colIndex) {
    _checkMaxCol(colIndex);

    while (colIndex >= _colAutoFit.length) {
      _colAutoFit.add(false);
    }
    _colAutoFit[colIndex] = true;
  }

  ///
  /// Set Column Width
  ///
  void setColWidth(int colIndex, double colWidth) {
    _checkMaxCol(colIndex);
    if (colWidth < 0) return;

    while (colIndex >= _colWidth.length) {
      _colWidth.add(_defaultColumnWidth);
    }
    _colWidth[colIndex] = colWidth;
  }

  CellType _getCellType(var type) {
    switch (type) {
      case int:
        return CellType.int;
      case double:
        return CellType.double;
      case bool:
        return CellType.bool;
      case Formula:
        return CellType.Formula;
      default:
        return CellType.String;
    }
  }

  ///
  ///Returns the `count` of replaced `source` with `target`
  ///
  ///`source` is dynamic which allows you to pass your custom `RegExp` providing more control over it.
  ///
  ///optional argument `first` is used to replace the number of first earlier occurrences
  ///
  ///If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
  ///
  ///       excel.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///       or
  ///
  ///       var mySheet = excel['mySheetName'];
  ///       mySheet.findAndReplace('sad', 'happy', first: 3);
  ///
  ///In the above example it will replace all the occurences of `sad` with `happy` in the cells
  ///
  ///Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
  ///
  int findAndReplace(dynamic source, dynamic target,
      {int first = -1,
      int startingRow = -1,
      int endingRow = -1,
      int startingColumn = -1,
      int endingColumn = -1}) {
    int replaceCount = 0,
        _startingRow = 0,
        _endingRow = -1,
        _startingColumn = 0,
        _endingColumn = -1;

    if (startingRow != -1 && endingRow != -1) {
      if (startingRow > endingRow) {
        _endingRow = startingRow;
        _startingRow = endingRow;
      } else {
        _endingRow = endingRow;
        _startingRow = startingRow;
      }
    }

    if (startingColumn != -1 && endingColumn != -1) {
      if (startingColumn > endingColumn) {
        _endingColumn = startingColumn;
        _startingColumn = endingColumn;
      } else {
        _endingColumn = endingColumn;
        _startingColumn = startingColumn;
      }
    }

    int rowsLength = maxRows, columnLength = maxCols;
    RegExp sourceRegx;
    if (source.runtimeType == RegExp) {
      sourceRegx = source;
    } else {
      sourceRegx = RegExp(source.toString());
    }

    for (int i = _startingRow; i < rowsLength; i++) {
      if (_endingRow != -1 && i > _endingRow) {
        break;
      }
      for (int j = _startingColumn; j < columnLength; j++) {
        if (_endingColumn != -1 && j > _endingColumn) {
          break;
        }
        if (_sheetData.isNotEmpty &&
            _sheetData[i] != null &&
            _sheetData[i]![j] != null &&
            sourceRegx.hasMatch(_sheetData[i]![j]!.value.toString()) &&
            (first == -1 || first != replaceCount)) {
          _sheetData[i]![j]!
              .value
              .toString()
              .replaceAll(sourceRegx, target.toString());

          replaceCount += 1;
        }
      }
    }

    return replaceCount;
  }

  ///
  /// returns `true` if the contents are successfully `cleared` else `false`.
  ///
  /// If the row is having any spanned-cells then it will not be cleared and hence returns `false`.
  ///
  bool clearRow(int rowIndex) {
    if (rowIndex < 0) {
      return false;
    }

    /// lets assume that this row is already cleared and is not inside spanList
    /// If this row exists then we check for the span condition
    bool isNotInside = true;

    if (_sheetData[rowIndex] != null && _sheetData[rowIndex]!.isNotEmpty) {
      /// lets start iterating the spanList and check that if the row is inside the spanList or not
      /// we will expect that value of isNotInside should not be changed to false
      /// If it changes to false then we can't clear this row as it is inside the spanned Cells
      for (int i = 0; i < _spanList.length; i++) {
        _Span? spanObj = _spanList[i];
        if (spanObj == null) {
          continue;
        }
        if (rowIndex >= spanObj.rowSpanStart &&
            rowIndex <= spanObj.rowSpanEnd) {
          isNotInside = false;
          break;
        }
      }

      /// As the row is not inside any SpanList so we can easily clear its content.
      if (isNotInside) {
        _sheetData[rowIndex]!.keys.toList().forEach((key) {
          /// Main concern here is to [clear the contents] and [not remove] the entire row or the cell block
          _sheetData[rowIndex]![key] = Data.newData(this, rowIndex, key);
        });
      }
    }
    //_countRowAndCol();
    return isNotInside;
  }

  ///
  ///It is used to check if cell at rowIndex, columnIndex is inside any spanning cell or not ?
  ///
  ///If it exist then the very first index of than spanned cells is returned in order to point to the starting cell
  ///otherwise the parameters are returned back.
  ///
  List<int> _isInsideSpanning(int rowIndex, int columnIndex) {
    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      if (rowIndex >= spanObj.rowSpanStart &&
          rowIndex <= spanObj.rowSpanEnd &&
          columnIndex >= spanObj.columnSpanStart &&
          columnIndex <= spanObj.columnSpanEnd) {
        newRowIndex = spanObj.rowSpanStart;
        newColumnIndex = spanObj.columnSpanStart;
        break;
      }
    }

    return [newRowIndex, newColumnIndex];
  }

  ///
  ///Check if columnIndex is not out of `Excel Column limits`.
  ///
  _checkMaxCol(int colIndex) {
    if (_maxCols >= 16384 || colIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
    if (colIndex < 0) {
      throw ArgumentError('Negative colIndex found: $colIndex');
    }
  }

  ///
  ///Check if rowIndex is not out of `Excel Row limits`.
  ///
  _checkMaxRow(int rowIndex) {
    if (_maxRows >= 1048576 || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
    if (rowIndex < 0) {
      throw ArgumentError('Negative rowIndex found: $rowIndex');
    }
  }

  ///
  ///returns List of Spanned Cells as
  ///
  ///     ["A1:A2", "A4:G6", "Y4:Y6", ....]
  ///
  ///return type if String based cell-id
  ///
  List<String> get spannedItems {
    _spannedItems = FastList<String>();

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      String rC = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
          spanObj.columnSpanEnd, spanObj.rowSpanEnd);
      if (!_spannedItems.contains(rC)) {
        _spannedItems.add(rC);
      }
    }

    return _spannedItems.keys;
  }

  ///
  ///Cleans the `_SpanList` by removing the indexes where null value exists.
  ///
  _cleanUpSpanMap() {
    if (_spanList.isNotEmpty) {
      _spanList.removeWhere((value) {
        return value == null;
      });
    }
  }

  ///return `SheetName`
  String get sheetName {
    return _sheet;
  }

  ///returns row at index = `rowIndex`
  List<Data?> row(int rowIndex) {
    if (rowIndex < 0) {
      return <Data?>[];
    }
    if (rowIndex < _maxRows) {
      if (_sheetData[rowIndex] != null) {
        return List.generate(_maxCols, (colIndex) {
          if (_sheetData[rowIndex]![colIndex] != null) {
            return _sheetData[rowIndex]![colIndex]!;
          }
          return null;
        });
      } else {
        return List.generate(_maxCols, (_) => null);
      }
    }
    return <Data?>[];
  }

  ///
  ///returns count of `rows` having data in `sheet`
  ///
  int get maxRows {
    return _maxRows;
  }

  ///
  ///returns count of `cols` having data in `sheet`
  ///
  int get maxCols {
    return _maxCols;
  }

  HeaderFooter? get headerFooter {
    return _headerFooter;
  }

  set headerFooter(HeaderFooter? headerFooter) {
    _headerFooter = headerFooter;
  }
}
