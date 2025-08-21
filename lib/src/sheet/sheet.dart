part of excel;

class Sheet {
  final Excel _excel;
  final String _sheet;
  bool _isRTL = false;
  int _maxRows = 0;
  int _maxColumns = 0;
  double? _defaultColumnWidth;
  double? _defaultRowHeight;
  Map<int, double> _columnWidths = {};
  Map<int, double> _rowHeights = {};
  Map<int, bool> _columnAutoFit = {};
  FastList<String> _spannedItems = FastList<String>();
  List<_Span?> _spanList = [];
  Map<int, Map<int, Data>> _sheetData = {};
  HeaderFooter? _headerFooter;

  ///
  /// It will clone the object by changing the `this` reference of previous oldSheetObject and putting `new this` reference, with copying the values too
  ///
  Sheet._clone(Excel excel, String sheetName, Sheet oldSheetObject)
      : this._(excel, sheetName,
            sh: oldSheetObject._sheetData,
            spanL_: oldSheetObject._spanList,
            spanI_: oldSheetObject._spannedItems,
            maxRowsVal: oldSheetObject._maxRows,
            maxColumnsVal: oldSheetObject._maxColumns,
            columnWidthsVal: oldSheetObject._columnWidths,
            rowHeightsVal: oldSheetObject._rowHeights,
            columnAutoFitVal: oldSheetObject._columnAutoFit,
            isRTLVal: oldSheetObject._isRTL,
            headerFooter: oldSheetObject._headerFooter);

  Sheet._(this._excel, this._sheet,
      {Map<int, Map<int, Data>>? sh,
      List<_Span?>? spanL_,
      FastList<String>? spanI_,
      int? maxRowsVal,
      int? maxColumnsVal,
      bool? isRTLVal,
      Map<int, double>? columnWidthsVal,
      Map<int, double>? rowHeightsVal,
      Map<int, bool>? columnAutoFitVal,
      HeaderFooter? headerFooter}) {
    _headerFooter = headerFooter;

    if (spanL_ != null) {
      _spanList = List<_Span?>.from(spanL_);
      _excel._mergeChangeLookup = sheetName;
    }
    if (spanI_ != null) {
      _spannedItems = FastList<String>.from(spanI_);
    }
    if (maxColumnsVal != null) {
      _maxColumns = maxColumnsVal;
    }
    if (maxRowsVal != null) {
      _maxRows = maxRowsVal;
    }
    if (isRTLVal != null) {
      _isRTL = isRTLVal;
      _excel._rtlChangeLookup = sheetName;
    }
    if (columnWidthsVal != null) {
      _columnWidths = Map<int, double>.from(columnWidthsVal);
    }
    if (rowHeightsVal != null) {
      _rowHeights = Map<int, double>.from(rowHeightsVal);
    }
    if (columnAutoFitVal != null) {
      _columnAutoFit = Map<int, bool>.from(columnAutoFitVal);
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
    _countRowsAndColumns();
  }

  /// Removes a cell from the specified [rowIndex] and [columnIndex].
  ///
  /// If the specified [rowIndex] or [columnIndex] does not exist,
  /// no action is taken.
  ///
  /// If the removal of the cell results in an empty row, the entire row is removed.
  ///
  /// Parameters:
  ///   - [rowIndex]: The index of the row from which to remove the cell.
  ///   - [columnIndex]: The index of the column from which to remove the cell.
  ///
  /// Example:
  /// ```dart
  /// final sheet = Spreadsheet();
  /// sheet.removeCell(1, 2);
  /// ```
  void _removeCell(int rowIndex, int columnIndex) {
    _sheetData[rowIndex]?.remove(columnIndex);
    final rowIsEmptyAfterRemovalOfCell = _sheetData[rowIndex]?.isEmpty == true;
    if (rowIsEmptyAfterRemovalOfCell) {
      _sheetData.remove(rowIndex);
    }
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
    _checkMaxColumn(cellIndex.columnIndex);
    _checkMaxRow(cellIndex.rowIndex);
    if (cellIndex.columnIndex < 0 || cellIndex.rowIndex < 0) {
      _damagedExcel(
          text:
              '${cellIndex.columnIndex < 0 ? "Column" : "Row"} Index: ${cellIndex.columnIndex < 0 ? cellIndex.columnIndex : cellIndex.rowIndex} Negative index does not exist.');
    }

    /// increasing the row count
    if (_maxRows < (cellIndex.rowIndex + 1)) {
      _maxRows = cellIndex.rowIndex + 1;
    }

    /// increasing the column count
    if (_maxColumns < (cellIndex.columnIndex + 1)) {
      _maxColumns = cellIndex.columnIndex + 1;
    }

    /// checking if the map has been already initialized or not?
    /// if the user has called this class by its own
    /* if (_sheetData == null) {
      _sheetData = Map<int, Map<int, Data>>();
    } */

    /// if the sheetData contains the row then start putting the column
    if (_sheetData[cellIndex.rowIndex] != null) {
      if (_sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] == null) {
        _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex] =
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex);
      }
    } else {
      /// else put the column with map showing.
      _sheetData[cellIndex.rowIndex] = {
        cellIndex.columnIndex:
            Data.newData(this, cellIndex.rowIndex, cellIndex.columnIndex)
      };
    }

    return _sheetData[cellIndex.rowIndex]![cellIndex.columnIndex]!;
  }

  ///
  /// returns `2-D dynamic List` of the sheet elements
  ///
  List<List<Data?>> get rows {
    var _data = <List<Data?>>[];

    if (_sheetData.isEmpty) {
      return _data;
    }

    if (_maxRows > 0 && maxColumns > 0) {
      _data = List.generate(_maxRows, (rowIndex) {
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex] != null &&
              _sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex];
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
    _checkMaxColumn(start.columnIndex);
    _checkMaxRow(start.rowIndex);
    if (end != null) {
      _checkMaxColumn(end.columnIndex);
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
        for (var j = _startColumn; j <= (_endColumn ?? maxColumns); j++) {
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
  /// updates count of rows and columns
  ///
  _countRowsAndColumns() {
    int maximumColumnIndex = -1, maximumRowIndex = -1;
    List<int> sortedKeys = _sheetData.keys.toList()..sort();
    sortedKeys.forEach((rowKey) {
      if (_sheetData[rowKey] != null && _sheetData[rowKey]!.isNotEmpty) {
        List<int> keys = _sheetData[rowKey]!.keys.toList()..sort();
        if (keys.isNotEmpty && keys.last > maximumColumnIndex) {
          maximumColumnIndex = keys.last;
        }
      }
    });

    if (sortedKeys.isNotEmpty) {
      maximumRowIndex = sortedKeys.last;
    }

    _maxColumns = maximumColumnIndex + 1;
    _maxRows = maximumRowIndex + 1;
  }

  ///
  /// If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0 || columnIndex >= maxColumns) {
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

      if (columnIndex <= endColumn) {
        if (columnIndex < startColumn) {
          startColumn -= 1;
        }
        endColumn -= 1;
        if (/* startColumn >= endColumn */
            (columnIndex == (endColumn + 1)) &&
                (columnIndex ==
                    (columnIndex < startColumn
                        ? startColumn + 1
                        : startColumn))) {
          _spanList[i] = null;
        } else {
          _Span newSpanObj = _Span(
            rowSpanStart: startRow,
            columnSpanStart: startColumn,
            rowSpanEnd: endRow,
            columnSpanEnd: endColumn,
          );
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
    if (columnIndex <= maxColumns - 1) {
      /// do the shifting task
      List<int> sortedKeys = _sheetData.keys.toList()..sort();
      sortedKeys.forEach((rowKey) {
        Map<int, Data> columnMap = Map<int, Data>();
        List<int> sortedColumnKeys = _sheetData[rowKey]!.keys.toList()..sort();
        sortedColumnKeys.forEach((columnKey) {
          if (_sheetData[rowKey] != null &&
              _sheetData[rowKey]![columnKey] != null) {
            if (columnKey < columnIndex) {
              columnMap[columnKey] = _sheetData[rowKey]![columnKey]!;
            }
            if (columnIndex == columnKey) {
              _sheetData[rowKey]!.remove(columnKey);
            }
            if (columnIndex < columnKey) {
              columnMap[columnKey - 1] = _sheetData[rowKey]![columnKey]!;
              _sheetData[rowKey]!.remove(columnKey);
            }
          }
        });
        _data[rowKey] = Map<int, Data>.from(columnMap);
      });
      _sheetData = Map<int, Map<int, Data>>.from(_data);
    }

    if (_maxColumns - 1 <= columnIndex) {
      _maxColumns -= 1;
    }
  }

  ///
  /// Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  /// If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  /// If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);

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

      if (columnIndex <= endColumn) {
        if (columnIndex <= startColumn) {
          startColumn += 1;
        }
        endColumn += 1;
        _Span newSpanObj = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
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
      final Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      final List<int> sortedKeys = _sheetData.keys.toList()..sort();
      if (columnIndex <= maxColumns - 1) {
        /// do the shifting task
        sortedKeys.forEach((rowKey) {
          final Map<int, Data> columnMap = Map<int, Data>();

          /// getting the column keys in descending order so as to shifting becomes easy
          final List<int> sortedColumnKeys = _sheetData[rowKey]!.keys.toList()
            ..sort((a, b) {
              return b.compareTo(a);
            });
          sortedColumnKeys.forEach((columnKey) {
            if (_sheetData[rowKey] != null &&
                _sheetData[rowKey]![columnKey] != null) {
              if (columnKey < columnIndex) {
                columnMap[columnKey] = _sheetData[rowKey]![columnKey]!;
              }
              if (columnIndex <= columnKey) {
                columnMap[columnKey + 1] = _sheetData[rowKey]![columnKey]!;
              }
            }
          });
          columnMap[columnIndex] = Data.newData(this, rowKey, columnIndex);
          _data[rowKey] = Map<int, Data>.from(columnMap);
        });
        _sheetData = Map<int, Map<int, Data>>.from(_data);
      } else {
        /// just put the data in the very first available row and
        /// in the desired Column index only one time as we will be using less space on internal implementatoin
        /// and mock the user as if the 2-D list is being saved
        ///
        /// As when user calls DataObject.cells then we will output 2-D list - pretending.
        _sheetData[sortedKeys.first]![columnIndex] =
            Data.newData(this, sortedKeys.first, columnIndex);
      }
    } else {
      /// here simply just take the first row and put the columnIndex as the _sheetData was previously null
      _sheetData = Map<int, Map<int, Data>>();
      _sheetData[0] = {columnIndex: Data.newData(this, 0, columnIndex)};
    }
    if (_maxColumns - 1 <= columnIndex) {
      _maxColumns += 1;
    } else {
      _maxColumns = columnIndex + 1;
    }

    //_countRowsAndColumns();
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
      final _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (rowIndex <= endRow) {
        if (rowIndex < startRow) {
          startRow -= 1;
        }
        endRow -= 1;
        if (/* startRow >= endRow */
            (rowIndex == (endRow + 1)) &&
                (rowIndex == (rowIndex < startRow ? startRow + 1 : startRow))) {
          _spanList[i] = null;
        } else {
          final _Span newSpanObj = _Span(
            rowSpanStart: startRow,
            columnSpanStart: startColumn,
            rowSpanEnd: endRow,
            columnSpanEnd: endColumn,
          );
          _spanList[i] = newSpanObj;
        }
        updateSpanCell = true;
        _excel._mergeChanges = true;
      }
      if (_spanList[i] != null) {
        final String rc =
            getSpanCellId(startColumn, startRow, endColumn, endRow);
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
      final Map<int, Map<int, Data>> _data = Map<int, Map<int, Data>>();
      if (rowIndex <= maxRows - 1) {
        /// do the shifting task
        final List<int> sortedKeys = _sheetData.keys.toList()..sort();
        sortedKeys.forEach((rowKey) {
          if (rowKey < rowIndex && _sheetData[rowKey] != null) {
            _data[rowKey] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
          if (rowIndex < rowKey && _sheetData[rowKey] != null) {
            _data[rowKey - 1] = Map<int, Data>.from(_sheetData[rowKey]!);
          }
        });
        _sheetData = Map<int, Map<int, Data>>.from(_data);
      }
      //_countRowsAndColumns();
    } else {
      _maxRows = 0;
      _maxColumns = 0;
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
      final _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }
      int startColumn = spanObj.columnSpanStart,
          startRow = spanObj.rowSpanStart,
          endColumn = spanObj.columnSpanEnd,
          endRow = spanObj.rowSpanEnd;

      if (rowIndex <= endRow) {
        if (rowIndex <= startRow) {
          startRow += 1;
        }
        endRow += 1;
        final _Span newSpanObj = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
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
            _data[rowKey + 1]!.forEach((key, value) {
              value._rowIndex++;
            });
          }
        });
      } else {
        _data = _sheetData;
      }
    }
    _data[rowIndex] = {0: Data.newData(this, rowIndex, 0)};
    _sheetData = Map<int, Map<int, Data>>.from(_data);

    if (_maxRows - 1 <= rowIndex) {
      _maxRows = rowIndex + 1;
    } else {
      _maxRows += 1;
    }

    //_countRowsAndColumns();
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
  void updateCell(CellIndex cellIndex, CellValue? value,
      {CellStyle? cellStyle}) {
    int columnIndex = cellIndex.columnIndex;
    int rowIndex = cellIndex.rowIndex;
    if (columnIndex < 0 || rowIndex < 0) {
      return;
    }
    _checkMaxColumn(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    /// Check if this is lying in merged-cell cross-section
    /// If yes then get the starting position of merged cells
    if (_spanList.isNotEmpty) {
      (newRowIndex, newColumnIndex) = _isInsideSpanning(rowIndex, columnIndex);
    }

    /// Puts Data
    _putData(newRowIndex, newColumnIndex, value);

    // check if the numberFormat works with the value provided
    // otherwise fall back to the default for this value type
    if (cellStyle != null) {
      final numberFormat = cellStyle.numberFormat;
      if (!numberFormat.accepts(value)) {
        cellStyle =
            cellStyle.copyWith(numberFormat: NumFormat.defaultFor(value));
      }
    } else {
      final cellStyleBefore =
          _sheetData[cellIndex.rowIndex]?[cellIndex.columnIndex]?.cellStyle;
      if (cellStyleBefore != null &&
          !cellStyleBefore.numberFormat.accepts(value)) {
        cellStyle =
            cellStyleBefore.copyWith(numberFormat: NumFormat.defaultFor(value));
      }
    }

    /// Puts the cellStyle
    if (cellStyle != null) {
      _sheetData[newRowIndex]![newColumnIndex]!._cellStyle = cellStyle;
      _excel._styleChanges = true;
    }
  }

  ///
  /// Merges the cells starting from `start` to `end`.
  ///
  /// If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  void merge(CellIndex start, CellIndex end, {CellValue? customValue}) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    _checkMaxColumn(startColumn);
    _checkMaxColumn(endColumn);
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

    // Update maxColumns maxRows
    _maxColumns = _maxColumns > endColumn ? _maxColumns : endColumn + 1;
    _maxRows = _maxRows > endRow ? _maxRows : endRow + 1;

    bool getValue = true;

    Data value = Data.newData(this, startRow, startColumn);
    if (customValue != null) {
      value._value = customValue;
      getValue = false;
    }

    for (int j = startRow; j <= endRow; j++) {
      for (int k = startColumn; k <= endColumn; k++) {
        if (_sheetData[j] != null) {
          if (getValue && _sheetData[j]![k]?.value != null) {
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

    _Span s = _Span(
      rowSpanStart: startRow,
      columnSpanStart: startColumn,
      rowSpanEnd: endRow,
      columnSpanEnd: endColumn,
    );

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
  void unMerge(String unmergeCells) {
    if (_spannedItems.isNotEmpty &&
        _spanList.isNotEmpty &&
        _spannedItems.contains(unmergeCells)) {
      List<String> lis = unmergeCells.split(RegExp(r":"));
      if (lis.length == 2) {
        bool remove = false;
        CellIndex start = CellIndex.indexByString(lis[0]),
            end = CellIndex.indexByString(lis[1]);
        for (int i = 0; i < _spanList.length; i++) {
          _Span? spanObject = _spanList[i];
          if (spanObject == null) {
            continue;
          }

          if (spanObject.columnSpanStart == start.columnIndex &&
              spanObject.rowSpanStart == start.rowIndex &&
              spanObject.columnSpanEnd == end.columnIndex &&
              spanObject.rowSpanEnd == end.rowIndex) {
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
  /// Sets the cellStyle of the merged cells.
  ///
  /// It will get the merged cells only by giving the starting position of merged cells.
  ///
  void setMergedCellStyle(CellIndex start, CellStyle mergedCellStyle) {
    List<List<CellIndex>> _mergedCells = spannedItems
        .map(
          (e) => e.split(":").map((e) => CellIndex.indexByString(e)).toList(),
        )
        .toList();

    List<CellIndex> _startIndices = _mergedCells.map((e) => e[0]).toList();
    List<CellIndex> _endIndices = _mergedCells.map((e) => e[1]).toList();

    if (_mergedCells.isEmpty ||
        start.columnIndex < 0 ||
        start.rowIndex < 0 ||
        !_startIndices.contains(start)) {
      return;
    }

    CellIndex end = _endIndices[_startIndices.indexOf(start)];

    bool hasBorder = mergedCellStyle.topBorder != Border() ||
        mergedCellStyle.bottomBorder != Border() ||
        mergedCellStyle.leftBorder != Border() ||
        mergedCellStyle.rightBorder != Border() ||
        mergedCellStyle.diagonalBorderUp ||
        mergedCellStyle.diagonalBorderDown;
    if (hasBorder) {
      for (var i = start.rowIndex; i <= end.rowIndex; i++) {
        for (var j = start.columnIndex; j <= end.columnIndex; j++) {
          CellStyle cellStyle = mergedCellStyle.copyWith(
            topBorderVal: Border(),
            bottomBorderVal: Border(),
            leftBorderVal: Border(),
            rightBorderVal: Border(),
            diagonalBorderUpVal: false,
            diagonalBorderDownVal: false,
          );

          if (i == start.rowIndex) {
            cellStyle = cellStyle.copyWith(
              topBorderVal: mergedCellStyle.topBorder,
            );
          }
          if (i == end.rowIndex) {
            cellStyle = cellStyle.copyWith(
              bottomBorderVal: mergedCellStyle.bottomBorder,
            );
          }
          if (j == start.columnIndex) {
            cellStyle = cellStyle.copyWith(
              leftBorderVal: mergedCellStyle.leftBorder,
            );
          }
          if (j == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              rightBorderVal: mergedCellStyle.rightBorder,
            );
          }

          if (i == j ||
              start.rowIndex == end.rowIndex ||
              start.columnIndex == end.columnIndex) {
            cellStyle = cellStyle.copyWith(
              diagonalBorderUpVal: mergedCellStyle.diagonalBorderUp,
              diagonalBorderDownVal: mergedCellStyle.diagonalBorderDown,
            );
          }

          if (i == start.rowIndex && j == start.columnIndex) {
            cell(start).cellStyle = cellStyle;
          } else {
            _putData(i, j, null);
            _sheetData[i]![j]!.cellStyle = cellStyle;
          }
        }
      }
    }
  }

  ///
  /// Helps to find the interaction between the pre-existing span position and updates if with new span if there any interaction(Cross-Sectional Spanning) exists.
  ///
  List<int> _getSpanPosition(CellIndex start, CellIndex end) {
    int startColumn = start.columnIndex,
        startRow = start.rowIndex,
        endColumn = end.columnIndex,
        endRow = end.rowIndex;

    bool remove = false;

    if (startRow > endRow) {
      startRow = end.rowIndex;
      endRow = start.rowIndex;
    }
    if (endColumn < startColumn) {
      endColumn = start.columnIndex;
      startColumn = end.columnIndex;
    }

    for (int i = 0; i < _spanList.length; i++) {
      _Span? spanObj = _spanList[i];
      if (spanObj == null) {
        continue;
      }

      final locationChange = _isLocationChangeRequired(
          startColumn, startRow, endColumn, endRow, spanObj);

      if (locationChange.$1) {
        startColumn = locationChange.$2.$1;
        startRow = locationChange.$2.$2;
        endColumn = locationChange.$2.$3;
        endRow = locationChange.$2.$4;
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
  void appendRow(List<CellValue?> row) {
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
  void insertRowIterables(
    List<CellValue?> row,
    int rowIndex, {
    int startingColumn = 0,
    bool overwriteMergedCells = true,
  }) {
    if (row.isEmpty || rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);
    int columnIndex = 0;
    if (startingColumn > 0) {
      columnIndex = startingColumn;
    }
    _checkMaxColumn(columnIndex + row.length);
    int rowsLength = _maxRows,
        maxIterationIndex = row.length - 1,
        currentRowPosition = 0; // position in [row] iterables

    if (overwriteMergedCells || rowIndex >= rowsLength) {
      // Normally iterating and putting the data present in the [row] as we are on the last index.

      while (currentRowPosition <= maxIterationIndex) {
        _putData(rowIndex, columnIndex++, row[currentRowPosition++]);
      }
    } else {
      // expensive function as per time complexity
      _selfCorrectSpanMap(_excel);
      List<_Span> _spanObjectsList = _getSpannedObjects(rowIndex, columnIndex);

      if (_spanObjectsList.isEmpty) {
        while (currentRowPosition <= maxIterationIndex) {
          _putData(rowIndex, columnIndex++, row[currentRowPosition++]);
        }
      } else {
        while (currentRowPosition <= maxIterationIndex) {
          if (_isInsideSpanObject(_spanObjectsList, columnIndex, rowIndex)) {
            _putData(rowIndex, columnIndex, row[currentRowPosition++]);
          }
          columnIndex++;
        }
      }
    }
  }

  ///
  /// Internal function for putting the data in `_sheetData`.
  ///
  void _putData(int rowIndex, int columnIndex, CellValue? value) {
    var row = _sheetData[rowIndex];
    if (row == null) {
      _sheetData[rowIndex] = row = {};
    }
    var cell = row[columnIndex];
    if (cell == null) {
      row[columnIndex] = cell = Data.newData(this, rowIndex, columnIndex);
    }

    cell._value = value;
    cell._cellStyle = CellStyle(numberFormat: NumFormat.defaultFor(value));
    if (cell._cellStyle != NumFormat.standard_0) {
      _excel._styleChanges = true;
    }

    if ((_maxColumns - 1) < columnIndex) {
      _maxColumns = columnIndex + 1;
    }

    if ((_maxRows - 1) < rowIndex) {
      _maxRows = rowIndex + 1;
    }

    //_countRowsAndColumns();
  }

  double? get defaultRowHeight => _defaultRowHeight;

  double? get defaultColumnWidth => _defaultColumnWidth;

  ///
  /// returns map of auto fit columns
  ///
  Map<int, bool> get getColumnAutoFits => _columnAutoFit;

  ///
  /// returns map of custom width columns
  ///
  Map<int, double> get getColumnWidths => _columnWidths;

  ///
  /// returns map of custom height rows
  ///
  Map<int, double> get getRowHeights => _rowHeights;

  ///
  /// returns auto fit state of column index
  ///
  bool getColumnAutoFit(int columnIndex) {
    if (_columnAutoFit.containsKey(columnIndex)) {
      return _columnAutoFit[columnIndex]!;
    }
    return false;
  }

  ///
  /// returns width of column index
  ///
  double getColumnWidth(int columnIndex) {
    if (_columnWidths.containsKey(columnIndex)) {
      return _columnWidths[columnIndex]!;
    }
    return _defaultColumnWidth!;
  }

  ///
  /// returns height of row index
  ///
  double getRowHeight(int rowIndex) {
    if (_rowHeights.containsKey(rowIndex)) {
      return _rowHeights[rowIndex]!;
    }
    return _defaultRowHeight!;
  }

  ///
  /// Set the default column width.
  ///
  /// If both `setDefaultRowHeight` and `setDefaultColumnWidth` are not called,
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultColumnWidth([double columnWidth = _excelDefaultColumnWidth]) {
    if (columnWidth < 0) return;
    _defaultColumnWidth = columnWidth;
  }

  ///
  /// Set the default row height.
  ///
  /// If both `setDefaultRowHeight` and `setDefaultColumnWidth` are not called,
  /// then the default row height and column width will be set by Excel.
  ///
  /// The default row height is 15.0 and the default column width is 8.43.
  ///
  void setDefaultRowHeight([double rowHeight = _excelDefaultRowHeight]) {
    if (rowHeight < 0) return;
    _defaultRowHeight = rowHeight;
  }

  ///
  /// Set Column AutoFit
  ///
  void setColumnAutoFit(int columnIndex) {
    _checkMaxColumn(columnIndex);
    if (columnIndex < 0) return;
    _columnAutoFit[columnIndex] = true;
  }

  ///
  /// Set Column Width
  ///
  void setColumnWidth(int columnIndex, double columnWidth) {
    _checkMaxColumn(columnIndex);
    if (columnWidth < 0) return;
    _columnWidths[columnIndex] = columnWidth;
  }

  ///
  /// Set Row Height
  ///
  void setRowHeight(int rowIndex, double rowHeight) {
    _checkMaxRow(rowIndex);
    if (rowHeight < 0) return;
    _rowHeights[rowIndex] = rowHeight;
  }

  ///
  ///Returns the `count` of replaced `source` with `target`
  ///
  ///`source` is Pattern which allows you to pass your custom `RegExp` or a simple `String` providing more control over it.
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
  int findAndReplace(Pattern source, String target,
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

    int rowsLength = maxRows, columnLength = maxColumns;

    for (int i = _startingRow; i < rowsLength; i++) {
      if (_endingRow != -1 && i > _endingRow) {
        break;
      }
      for (int j = _startingColumn; j < columnLength; j++) {
        if (_endingColumn != -1 && j > _endingColumn) {
          break;
        }
        final sourceData = _sheetData[i]?[j]?.value;
        if (sourceData is! TextCellValue) {
          continue;
        }
        final result =
            sourceData.value.toString().replaceAllMapped(source, (match) {
          if (first == -1 || first != replaceCount) {
            ++replaceCount;
            return match.input.replaceRange(match.start, match.end, target);
          }
          return match.input;
        });
        _sheetData[i]![j]!.value = TextCellValue(result);
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
    //_countRowsAndColumns();
    return isNotInside;
  }

  ///
  ///It is used to check if cell at rowIndex, columnIndex is inside any spanning cell or not ?
  ///
  ///If it exist then the very first index of than spanned cells is returned in order to point to the starting cell
  ///otherwise the parameters are returned back.
  ///
  (int newRowIndex, int newColumnIndex) _isInsideSpanning(
      int rowIndex, int columnIndex) {
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

    return (newRowIndex, newColumnIndex);
  }

  ///
  ///Check if columnIndex is not out of `Excel Column limits`.
  ///
  void _checkMaxColumn(int columnIndex) {
    if (_maxColumns >= 16384 || columnIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
    if (columnIndex < 0) {
      throw ArgumentError('Negative columnIndex found: $columnIndex');
    }
  }

  ///
  ///Check if rowIndex is not out of `Excel Row limits`.
  ///
  void _checkMaxRow(int rowIndex) {
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
  void _cleanUpSpanMap() {
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
        return List.generate(_maxColumns, (columnIndex) {
          if (_sheetData[rowIndex]![columnIndex] != null) {
            return _sheetData[rowIndex]![columnIndex]!;
          }
          return null;
        });
      } else {
        return List.generate(_maxColumns, (_) => null);
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
  ///returns count of `columns` having data in `sheet`
  ///
  int get maxColumns {
    return _maxColumns;
  }

  HeaderFooter? get headerFooter {
    return _headerFooter;
  }

  set headerFooter(HeaderFooter? headerFooter) {
    _headerFooter = headerFooter;
  }
}
