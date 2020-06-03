part of excel;

class Sheet {
  Excel _excel;
  String _sheet;
  int _maxRows = 0;
  int _maxCols = 0;
  List<String> _spannedItems = List<String>();
  List<_Span> _spanList = List<_Span>();
  Map<int, Map<int, Data>> _sheetData = Map<int, Map<int, Data>>();

  /**
   * 
   * 
   * 
   * It will clone the object by changing the `this` reference of previous oldSheetObject and putting `new this` reference, with copying the values too
   * 
   * 
   * 
   */
  Sheet._clone(Excel excel, String sheetName, Sheet oldSheetObject)
      : this._(
          excel,
          sheetName,
          sh: oldSheetObject._sheetData,
          spanL_: oldSheetObject._spanList,
          spanI_: oldSheetObject._spannedItems,
          maxR_: oldSheetObject._maxRows,
          maxC_: oldSheetObject._maxCols,
        );

  Sheet._(
    Excel excel,
    String sheetName, {
    Map<int, Map<int, Data>> sh,
    List<_Span> spanL_,
    List<String> spanI_,
    int maxR_,
    int maxC_,
  }) {
    this._excel = excel;
    this._sheet = sheetName;
    this._spanList = spanL_ ?? List<_Span>();
    this._spannedItems = spanI_ ?? List<String>();
    this._maxCols = maxC_ ?? 0;
    this._maxRows = maxR_ ?? 0;
    this._sheetData = Map<int, Map<int, Data>>();

    if (spanL_ != null) {
      _excel._mergeChangeLookup = sheetName;
      this._spanList = List<_Span>.from(spanL_);
    }

    if (spanI_ != null) {
      this._spannedItems = List<String>.from(spanI_);
    }

    /// copy the data objects into a temp folder and then while putting it into `this._sheetData` change the data objects references.
    if (sh != null) {
      this._sheetData = Map<int, Map<int, Data>>.from(sh);
      Map<int, Map<int, Data>> temp = Map<int, Map<int, Data>>.from(sh);
      temp.forEach((key, value) {
        this._sheetData[key].forEach((key1, oldDataObject) {
          Data newDataObject = Data._clone(this, oldDataObject);
          this._sheetData[key][key1] = newDataObject;
        });
      });
    }
  }

  /**
   * 
   * 
   * returns the `DataObject` at position of `cellIndex`
   * 
   * 
   */
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

  /**
   * 
   * 
   * returns `2-D dynamic List` of the sheet elements
   * 
   * 
   */
  List<List<dynamic>> get rows {
    List<List<dynamic>> _data = List<List<dynamic>>();

    if (!_isContain(this._sheetData) || this._sheetData.isEmpty) {
      return _data;
    }

    if (this._maxRows > 0 && this.maxCols > 0) {
      _data = List.generate(this._maxRows, (rowIndex) {
        return List.generate(this._maxCols, (colIndex) {
          if (_isContain(this._sheetData[rowIndex]) &&
              _isContain(this._sheetData[rowIndex][colIndex])) {
            return this._sheetData[rowIndex][colIndex].value;
          }
          return null;
        });
      });
    }

    return _data;
  }

  /**
   * 
   * updates count of rows and cols
   * 
   */
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

  /**
   * 
   * 
   * If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
   * 
   * 
   */
  void removeColumn(int colIndex) {
    _checkMaxCol(colIndex);
    if (colIndex < 0 || colIndex >= this.maxCols) {
      return;
    }

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
          String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
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

  /**
   * 
   * 
   * Inserts an empty `column` in sheet at position = `columnIndex`.
   * 
   * If `columnIndex == null` or `columnIndex < 0` if will not execute 
   * 
   * If the `sheet` does not exists then it will be created automatically.
   * 
   * 
   */
  void insertColumn(int colIndex) {
    if (colIndex == null || colIndex < 0) {
      return;
    }
    _checkMaxCol(colIndex);

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
        String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
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

  /**
   * 
   * 
   * If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
   * 
   * 
   */
  void removeRow(int rowIndex) {
    if (rowIndex < 0 || rowIndex >= this._maxRows) {
      return;
    }
    _checkMaxRow(rowIndex);

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
          String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
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

    if (this._sheetData.isNotEmpty) {
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

  /**
   * 
   * 
   * 
   * Inserts an empty row in `sheet` at position = `rowIndex`.
   * 
   * If `rowIndex == null` or `rowIndex < 0` if will not execute 
   * 
   * If the `sheet` does not exists then it will be created automatically.
   * 
   * 
   * 
   */
  void insertRow(int rowIndex) {
    if (rowIndex < 0) {
      return;
    }

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
        String rc = getSpanCellId(startColumn, startRow, endColumn, endRow);
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

  /**
   * 
   * 
   * Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
   * 
   * ----or---- by `cellIndex: CellIndex.indexByString("A3");`.
   * 
   * Styling of cell can be done by passing the CellStyle object to `cellStyle`.
   * 
   * If `sheet` does not exist then it will be automatically created.
   * 
   * 
   */
  void updateCell(CellIndex cellIndex, dynamic value, {CellStyle cellStyle}) {
    int columnIndex = cellIndex._columnIndex;
    int rowIndex = cellIndex._rowIndex;
    if (columnIndex < 0 || rowIndex < 0 || value == null) {
      return;
    }
    _checkMaxCol(columnIndex);
    _checkMaxRow(rowIndex);

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    /// Check if this is lying in merged-cell cross-section
    /// If yes then get the starting position of merged cells
    if (this._spanList.isNotEmpty) {
      List updatedPosition = _isInsideSpanning(rowIndex, columnIndex);
      newRowIndex = updatedPosition[0];
      newColumnIndex = updatedPosition[1];
    }

    /// Puts Data
    _putData(newRowIndex, newColumnIndex, value);

    /// Puts the cellStyle
    if (cellStyle != null) {
      this._sheetData[newRowIndex][newColumnIndex]._cellStyle = cellStyle;
      _excel._colorChanges = true;
    }
  }

  /**
   * 
   * 
   * Merges the cells starting from `start` to `end`.
   * 
   * If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
   * 
   * 
   */
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
        (_spannedItems != null &&
            _spannedItems.contains(
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
        if (_isContain(this._sheetData) &&
            _isContain(this._sheetData[j]) &&
            _isContain(this._sheetData[j][k])) {
          if (getValue &&
              this._sheetData[j][k].value != null &&
              this._sheetData[j][k].cellStyle != null) {
            value = this._sheetData[j][k];
            getValue = false;
          }
          this._sheetData[j].remove(k);
        }
      }
    }

    if (_isContain(this._sheetData[startRow])) {
      this._sheetData[startRow][startColumn] = value;
    } else {
      this._sheetData[startRow] = {startColumn: value};
    }

    String sp = getSpanCellId(startColumn, startRow, endColumn, endRow);

    if (!_spannedItems.contains(sp)) {
      _spannedItems.add(sp);
    }

    _Span s = _Span();
    s._start = [startRow, startColumn];
    s._end = [endRow, endColumn];

    _spanList.add(s);
    _excel._mergeChangeLookup = this.sheetName;
  }

  /**
   * 
   * 
   * unMerge the merged cells.
   * 
   *        var sheet = 'DesiredSheet';
   *        List<String> spannedCells = excel.getMergedCells(sheet);
   *        var cellToUnMerge = "A1:A2";
   *        excel.unMerge(sheet, cellToUnMerge);
   * 
   * 
   */
  unMerge(String unmergeCells) {
    if (unmergeCells != null &&
        _spannedItems.isNotEmpty &&
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
          _Span spanObject = _spanList[i];

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
      _excel._mergeChangeLookup = this.sheetName;
    }
  }

  /**
   * 
   * 
   * 
   * Helps to find the interaction between the pre-existing span position and updates if with new span if there any interaction(Cross-Sectional Spanning) exists.
   * 
   * 
   * 
   */
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

    if (_spanList != null) {
      for (int i = 0; i < _spanList.length; i++) {
        _Span spanObj = _spanList[i];

        Map<String, List<int>> gotMap = _isLocationChangeRequired(
            startColumn, startRow, endColumn, endRow, spanObj);
        List<int> gotPosition = gotMap['gotPosition'];
        int changeValue = gotMap['changeValue'][0];

        if (changeValue == 1) {
          startColumn = gotPosition[0];
          startRow = gotPosition[1];
          endColumn = gotPosition[2];
          endRow = gotPosition[3];
          String sp = getSpanCellId(spanObj.columnSpanStart,
              spanObj.rowSpanStart, spanObj.columnSpanEnd, spanObj.rowSpanEnd);
          if (_spannedItems != null && _spannedItems.contains(sp)) {
            _spannedItems.remove(sp);
          }
          remove = true;
          _spanList[i] = null;
        }
      }
      if (remove) {
        _cleanUpSpanMap();
      }
    }
    return [startColumn, startRow, endColumn, endRow];
  }

  /**
   * 
   * 
   * Appends [row] iterables just post the last filled `rowIndex`.
   * 
   * 
   */
  appendRow(List<dynamic> row) {
    int targetRow = this.maxRows;
    insertRowIterables(row, targetRow);
  }

  /// getting the List of _Span Objects which have the rowIndex containing and
  /// also lower the range by giving the starting columnIndex
  List<_Span> _getSpannedObjects(int rowIndex, int startingColumnIndex) {
    List<_Span> obtained;

    if (this._spanList.isNotEmpty) {
      obtained = List<_Span>();
      this._spanList.forEach((spanObject) {
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

  /**
   * 
   * 
   * Checking if the columnIndex and the rowIndex passed is inside the spanObjectList which is got from calling function.
   * 
   * 
   */
  bool _isInsideSpanObject(
      List<_Span> spanObjectList, int columnIndex, int rowIndex) {
    for (int i = 0; i < spanObjectList.length; i++) {
      _Span spanObject = spanObjectList[i];

      if (spanObject != null &&
          spanObject.columnSpanStart <= columnIndex &&
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

  /**
   * 
   * 
   * Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
   * 
   * [startingColumn] tells from where we should start putting the [row] iterables
   * 
   * [overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
   * 
   * [overwriteMergedCells] when set to [false] puts the cell value in next unique cell available and putting the value in merged cells only once.
   * 
   * 
   */
  insertRowIterables(List<dynamic> row, int rowIndex,
      {int startingColumn = 0, bool overwriteMergedCells = true}) {
    if (row == null || row.length == 0 || rowIndex == null || rowIndex < 0) {
      return;
    }

    _checkMaxRow(rowIndex);
    int columnIndex = 0;
    if (startingColumn > 0) {
      columnIndex = startingColumn;
    }
    _checkMaxCol(columnIndex + row.length);
    int rowsLength = this.maxRows,
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

      if (_spanObjectsList == null || _spanObjectsList.length <= 0) {
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
  }

  /**
   * 
   * 
   * Internal function for putting the data in `_sheetData`.
   * 
   * 
   */
  _putData(int rowIndex, int columnIndex, dynamic value) {
    if (_isContain(this._sheetData[rowIndex])) {
      if (!_isContain(this._sheetData[rowIndex][columnIndex])) {
        this._sheetData[rowIndex][columnIndex] =
            Data.newData(this, rowIndex, columnIndex);
      }
    } else {
      this._sheetData[rowIndex] = {
        columnIndex: Data.newData(this, rowIndex, columnIndex)
      };
    }
    this._sheetData[rowIndex][columnIndex]._value = value;

    /// Sets value of `isFormula` to true if this is `instance of Formula`.
    this._sheetData[rowIndex][columnIndex]._isFormula =
        value is Formula || value.runtimeType == Formula;

    /// Sets type of the Data to `_cellType`
    this._sheetData[rowIndex][columnIndex]._cellType =
        _getCellType(value.runtimeType);

    if (value.runtimeType == String &&
        !_excel._sharedStrings.contains('$value')) {
      print(value + " is String");
      _excel._sharedStrings.add(value.toString());
    }
    _countRowAndCol();
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

  /**
   * 
   * 
   * Returns the `count` of replaced `source` with `target`
   *
   * `source` is dynamic which allows you to pass your custom `RegExp` providing more control over it.
   *
   * optional argument `first` is used to replace the number of first earlier occurrences
   *
   * If `first` is set to `3` then it will replace only first `3 occurrences` of the `source` with `target`.
   * 
   *        excel.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
   * 
   *        or
   * 
   *        var mySheet = excel['mySheetName'];
   *        mySheet.findAndReplace('sad', 'happy', first: 3);
   * 
   * In the above example it will replace all the occurences of `sad` with `happy` in the cells
   *
   * Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
   * 
   * 
   */
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

    int rowsLength = this.maxRows, columnLength = this.maxCols;
    RegExp sourceRegx;
    if (source.runtimeType == RegExp) {
      sourceRegx == source;
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
        if (this._sheetData.isNotEmpty &&
            _isContain(this._sheetData[i]) &&
            _isContain(this._sheetData[i][j]) &&
            sourceRegx.hasMatch(this._sheetData[i][j].toString()) &&
            (first == -1 || first != replaceCount)) {
          this
              ._sheetData[i][j]
              .value
              .toString()
              .replaceAll(sourceRegx, target.toString());

          replaceCount += 1;
        }
      }
    }

    return replaceCount;
  }

  /**
   * 
   * 
   * returns `true` if the contents are successfully `cleared` else `false`.
   * 
   * If the row is having any spanned-cells then it will not be cleared and hence returns `false`.
   * 
   * 
   */
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

  /**
   * 
   * 
   * It is used to check if cell at rowIndex, columnIndex is inside any spanning cell or not ?
   * 
   * If it exist then the very first index of than spanned cells is returned in order to point to the starting cell 
   * otherwise the parameters are returned back.
   * 
   * 
   */
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

  /**
   * 
   * 
   * Check if columnIndex is not out of `Excel Column limits`.
   * 
   * 
   */
  _checkMaxCol(int colIndex) {
    if ((this._maxCols >= 16384) || colIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
  }

  /**
   * 
   * 
   * Check if rowIndex is not out of `Excel Row limits`.
   * 
   * 
   */
  _checkMaxRow(int rowIndex) {
    if ((this._maxRows >= 1048576) || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
  }

  /**
   * 
   * 
   * returns List of Spanned Cells as 
   *      
   *      ["A1:A2", "A4:G6", "Y4:Y6", ....]
   * 
   * return type if String based cell-id
   * 
   * 
   */

  List<String> get spannedItems {
    this._spannedItems = List<String>();

    if (this._spanList != null && this._spanList.isNotEmpty) {
      this._spanList.forEach((spanObj) {
        String rC = getSpanCellId(spanObj.columnSpanStart, spanObj.rowSpanStart,
            spanObj.columnSpanEnd, spanObj.rowSpanEnd);
        if (!this._spannedItems.contains(rC)) {
          this._spannedItems.add(rC);
        }
      });
    }

    return this._spannedItems;
  }

  /**
   * 
   * 
   * Cleans the `_SpanList` by removing the indexes where null value exists.
   * 
   * 
   * 
   */
  _cleanUpSpanMap() {
    if (_spanList != null && _spanList.isNotEmpty) {
      this._spanList.removeWhere((value) => value == null);
    }
  }

  /**
   * 
   * 
   * return `SheetName`
   * 
   * 
   */
  String get sheetName {
    return this._sheet;
  }

  /**
   * 
   * 
   * returns row at index = `rowIndex`
   * 
   * 
   */
  List<Data> row(int rowIndex) {
    if (rowIndex == null || rowIndex < 0) {
      return [];
    }
    if (_isContain(this._sheetData[rowIndex])) {
      return List.generate(this.maxCols, (colIndex) {
        if (_isContain(this._sheetData[rowIndex][colIndex])) {
          return this._sheetData[rowIndex][colIndex];
        }
        return Data.newData(this, rowIndex, colIndex);
      });
    }
    return [];
  }

  /**
   * 
   * 
   * returns count of `rows` having data in `sheet`
   * 
   * 
   */
  int get maxRows => this._maxRows;

  /**
   * 
   * 
   * returns count of `cols` having data in `sheet`
   * 
   * 
   */
  int get maxCols => this._maxCols;
}
