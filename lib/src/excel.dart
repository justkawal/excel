part of excel;

Excel _newExcel(Archive archive) {
  // Lookup at file format
  var format;

  var mimetype = archive.findFile('mimetype');
  if (mimetype == null) {
    var xl = archive.findFile('xl/workbook.xml');
    if (xl != null) {
      format = _spreasheetXlsx;
    }
  }

  switch (format) {
    case _spreasheetXlsx:
      return Excel._(archive);
    default:
      throw UnsupportedError('Excel format unsupported.');
  }
}

/// Decode a excel file.
class Excel {
  late bool _colorChanges;
  late bool _mergeChanges;
  late bool _rtlChanges;

  late Archive _archive;

  late Map<String, XmlNode> _sheets;
  late Map<String, XmlDocument> _xmlFiles;
  late Map<String, String> _xmlSheetId;
  late Map<String, Map<String, int>> _cellStyleReferenced;
  late Map<String, Sheet> _sheetMap;

  late List<CellStyle> _cellStyleList;
  late List<String> _patternFill;
  late List<String> _mergeChangeLook;
  late List<String> _rtlChangeLook;
  late List<_FontStyle> _fontStyleList;
  late List<int> _numFormats;

  late _SharedStringsMaintainer _sharedStrings;

  late String _stylesTarget;
  late String _sharedStringsTarget;

  String? _defaultSheet;
  late Parser parser;

  Excel._(Archive archive) {
    _colorChanges = false;
    _mergeChanges = false;
    _rtlChanges = false;
    _sheets = <String, XmlNode>{};
    _xmlFiles = <String, XmlDocument>{};
    _xmlSheetId = <String, String>{};
    _cellStyleReferenced = <String, Map<String, int>>{};
    _sheetMap = <String, Sheet>{};
    _cellStyleList = <CellStyle>[];
    _patternFill = <String>[];
    _mergeChangeLook = <String>[];
    _rtlChangeLook = <String>[];
    _fontStyleList = <_FontStyle>[];
    _numFormats = <int>[];
    _stylesTarget = '';
    _sharedStringsTarget = '';

    _archive = archive;
    _sharedStrings = _SharedStringsMaintainer.instance;
    _sharedStrings.ensureReinitialize();
    parser = Parser._(this);
    parser._startParsing();
  }

  factory Excel.createExcel() {
    String newSheet =
        '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';
    return Excel.decodeBytes(Base64Decoder().convert(newSheet));
  }

  factory Excel.decodeBytes(List<int> data) {
    return _newExcel(ZipDecoder().decodeBytes(data));
  }

  factory Excel.decodeBuffer(InputStream input) {
    return _newExcel(ZipDecoder().decodeBuffer(input));
  }

  ///
  ///It will return `tables` as map in order to mimic the previous versions reading the data.
  ///
  Map<String, Sheet> get tables {
    if (this._sheetMap.isEmpty) {
      _damagedExcel(text: "Corrupted Excel file.");
    }
    return Map<String, Sheet>.from(this._sheetMap);
  }

  ///
  ///It will return the SheetObject of `sheet`.
  ///
  ///If the `sheet` does not exist then it will create `sheet` with `New Sheet Object`
  ///
  Sheet operator [](String sheet) {
    _availSheet(sheet);
    return _sheetMap[sheet]!;
  }

  ///
  ///Returns the `Map<String, Sheet>`
  ///
  ///where `key` is the `Sheet Name` and the `value` is the `Sheet Object`
  ///
  Map<String, Sheet> get sheets {
    return Map<String, Sheet>.from(_sheetMap);
  }

  ///
  ///If `sheet` does not exist then it will be automatically created with contents of `sheetObject`
  ///
  ///Newly created sheet with name = `sheet` will have seperate reference and will not be linked to sheetObject.
  ///
  operator []=(String sheet, Sheet sheetObject) {
    _availSheet(sheet);

    _sheetMap[sheet] = Sheet._clone(this, sheet, sheetObject);
  }

  ///
  ///`sheet2Object` will be linked with `sheet1`.
  ///
  ///If `sheet1` does not exist then it will be automatically created.
  ///
  ///Important Note: After linkage the operations performed on `sheet1`, will also get performed on `sheet2Object` and `vica-versa`.
  ///
  void link(String sheet1, Sheet existingSheetObject) {
    if (_sheetMap[existingSheetObject.sheetName] != null) {
      _availSheet(sheet1);

      _sheetMap[sheet1] = _sheetMap[existingSheetObject.sheetName]!;

      if (_cellStyleReferenced[existingSheetObject.sheetName] != null) {
        _cellStyleReferenced[sheet1] = Map<String, int>.from(
            _cellStyleReferenced[existingSheetObject.sheetName]!);
      }
    }
  }

  ///
  ///If `sheet` is linked with any other sheet's object then it's link will be broke
  ///
  void unLink(String sheet) {
    if (_sheetMap[sheet] != null) {
      ///
      /// copying the sheet into itself thus resulting in breaking the linkage as Sheet._clone() will provide new reference;
      copy(sheet, sheet);
    }
  }

  ///
  ///Copies the content of `fromSheet` into `toSheet`.
  ///
  ///In order to successfully copy: `fromSheet` should exist in `excel.tables.keys`.
  ///
  ///If `toSheet` does not exist then it will be automatically created.
  ///
  void copy(String fromSheet, String toSheet) {
    _availSheet(toSheet);

    if (_sheetMap[fromSheet] != null) {
      this[toSheet] = this[fromSheet];
    }
    if (_cellStyleReferenced[fromSheet] != null) {
      _cellStyleReferenced[toSheet] =
          Map<String, int>.from(_cellStyleReferenced[fromSheet]!);
    }
  }

  ///
  ///Changes the name from `oldSheetName` to `newSheetName`.
  ///
  ///In order to rename : `oldSheetName` should exist in `excel.tables.keys` and `newSheetName` must not exist.
  ///
  void rename(String oldSheetName, String newSheetName) {
    if (_sheetMap[oldSheetName] != null && _sheetMap[newSheetName] == null) {
      ///
      /// rename from _defaultSheet var also
      if (_defaultSheet == oldSheetName) {
        _defaultSheet = newSheetName;
      }

      copy(oldSheetName, newSheetName);

      ///
      /// delete the `oldSheetName` as sheet with `newSheetName` is having cloned `SheetObject of oldSheetName` with new reference,
      delete(oldSheetName);
    }
  }

  ///
  ///If `sheet` exist in `excel.tables.keys` and `excel.tables.keys.length >= 2` then it will be `deleted`.
  ///
  void delete(String sheet) {
    ///
    /// remove the sheet `name` or `key` from the below locations if they exist.

    ///
    /// If it is not the last sheet then `delete` otherwise `return`;
    if (_sheetMap.length <= 1) {
      return;
    }

    ///
    ///remove from _defaultSheet var also
    if (_defaultSheet == sheet) {
      _defaultSheet = null;
    }

    ///
    /// remove the `Sheet Object` from `_sheetMap`.
    if (_sheetMap[sheet] != null) {
      _sheetMap.remove(sheet);
    }

    ///
    /// remove from `_mergeChangeLook`.
    if (_mergeChangeLook.contains(sheet)) {
      _mergeChangeLook.remove(sheet);
    }

    ///
    /// remove from `_rtlChangeLook`.
    if (_rtlChangeLook.contains(sheet)) {
      _rtlChangeLook.remove(sheet);
    }

    ///
    /// remove from `_xmlSheetId`.
    if (_xmlSheetId[sheet] != null) {
      String sheetId1 = "worksheets" +
              _xmlSheetId[sheet].toString().split('worksheets')[1].toString(),
          sheetId2 = _xmlSheetId[sheet]!;

      _xmlFiles['xl/_rels/workbook.xml.rels']
          ?.rootElement
          .children
          .removeWhere((_sheetName) {
        return _sheetName.getAttribute('Target') != null &&
            _sheetName.getAttribute('Target').toString() == sheetId1;
      });

      _xmlFiles['[Content_Types].xml']
          ?.rootElement
          .children
          .removeWhere((_sheetName) {
        return _sheetName.getAttribute('PartName') != null &&
            _sheetName.getAttribute('PartName').toString() == '/' + sheetId2;
      });

      ///
      /// Remove from the `_archive` also
      _archive.files.removeWhere((file) {
        return file.name.toLowerCase() ==
            _xmlSheetId[sheet].toString().toLowerCase();
      });

      ///
      /// Also remove from the _xmlFiles list as we might want to create this sheet again from new starting.
      if (_xmlFiles[_xmlSheetId[sheet]] != null) {
        _xmlFiles.remove(_xmlSheetId[sheet]);
      }

      _xmlSheetId.remove(sheet);
    }

    ///
    /// remove from key = `sheet` from `_sheets`
    if (_sheets[sheet] != null) {
      ///
      /// Remove from `xl/workbook.xml`
      ///
      _xmlFiles['xl/workbook.xml']
          ?.findAllElements('sheets')
          .first
          .children
          .removeWhere((element) {
        return element.getAttribute('name') != null &&
            element.getAttribute('name').toString() == sheet;
      });

      _sheets.remove(sheet);
    }

    ///
    /// remove the cellStlye Referencing as it would be useless to have cellStyleReferenced saved
    if (_cellStyleReferenced[sheet] != null) {
      _cellStyleReferenced.remove(sheet);
    }
  }

  ///
  ///It will start setting the edited values of `sheets` into the `files` and then `exports the file`.
  ///
  List<int>? encode() {
    Save s = Save._(this, parser);
    return s._save();
  }

  /// Starts Saving the file.
  /// `On Web`
  /// ```
  /// // Call function save() to download the file
  /// var bytes = excel.save(fileName: "My_Excel_File_Name.xlsx");
  ///
  ///
  /// ```
  /// `On Android / iOS`
  ///
  /// For getting directory on Android or iOS, Use: [path_provider](https://pub.dev/packages/path_provider)
  /// ```
  /// // Call function save() to download the file
  /// var fileBytes = excel.save();
  /// var directory = await getApplicationDocumentsDirectory();
  ///
  /// File(join("$directory/output_file_name.xlsx"))
  ///   ..createSync(recursive: true)
  ///   ..writeAsBytesSync(fileBytes);
  ///
  ///```
  List<int>? save({String fileName = 'FlutterExcel.xlsx'}) {
    Save s = Save._(this, parser);
    var onValue = s._save();
    return helper.SavingHelper.saveFile(onValue, fileName);
  }

  ///
  ///returns the name of the `defaultSheet` (the sheet which opens firstly when xlsx file is opened in `excel based software`).
  ///
  String? getDefaultSheet() {
    if (_defaultSheet != null) {
      return _defaultSheet;
    } else {
      String? re = _getDefaultSheet();
      return re;
    }
  }

  ///
  ///Internal function which returns the defaultSheet-Name by reading from `workbook.xml`
  ///
  String? _getDefaultSheet() {
    Iterable<XmlElement>? elements =
        _xmlFiles['xl/workbook.xml']?.findAllElements('sheet');
    XmlElement? _sheet;
    if (elements?.isNotEmpty ?? false) {
      _sheet = elements?.first;
    }

    if (_sheet != null) {
      var defaultSheet = _sheet.getAttribute('name');
      if (defaultSheet != null) {
        return defaultSheet.toString();
      } else {
        _damagedExcel(
            text: 'Excel sheet corrupted!! Try creating new excel file.');
      }
    }
    return null;
  }

  ///
  ///It returns `true` if the passed `sheetName` is successfully set to `default opening sheet` otherwise returns `false`.
  ///
  bool setDefaultSheet(String sheetName) {
    if (_sheetMap[sheetName] != null) {
      _defaultSheet = sheetName;
      return true;
    }
    return false;
  }

  ///
  ///Inserts an empty `column` in sheet at position = `columnIndex`.
  ///
  ///If `columnIndex == null` or `columnIndex < 0` if will not execute
  ///
  ///If the `sheet` does not exists then it will be created automatically.
  ///
  void insertColumn(String sheet, int columnIndex) {
    if (columnIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertColumn(columnIndex);
  }

  ///
  ///If `sheet` exists and `columnIndex < maxColumns` then it removes column at index = `columnIndex`
  ///
  void removeColumn(String sheet, int columnIndex) {
    if (columnIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeColumn(columnIndex);
    }
  }

  ///
  ///Inserts an empty row in `sheet` at position = `rowIndex`.
  ///
  ///If `rowIndex == null` or `rowIndex < 0` if will not execute
  ///
  ///If the `sheet` does not exists then it will be created automatically.
  ///
  void insertRow(String sheet, int rowIndex) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap[sheet]!.insertRow(rowIndex);
  }

  ///
  ///If `sheet` exists and `rowIndex < maxRows` then it removes row at index = `rowIndex`
  ///
  void removeRow(String sheet, int rowIndex) {
    if (rowIndex >= 0 && _sheetMap[sheet] != null) {
      _sheetMap[sheet]!.removeRow(rowIndex);
    }
  }

  ///
  ///Appends [row] iterables just post the last filled index in the [sheet]
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void appendRow(String sheet, List<dynamic> row) {
    if (row.length == 0) {
      return;
    }
    _availSheet(sheet);
    int targetRow = _sheetMap[sheet]!.maxRows;
    insertRowIterables(sheet, row, targetRow);
  }

  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  ///Adds the [row] iterables in the given rowIndex = [rowIndex] in [sheet]
  ///
  ///[startingColumn] tells from where we should start putting the [row] iterables
  ///
  ///[overwriteMergedCells] when set to [true] will over-write mergedCell and does not jumps to next unqiue cell.
  ///
  ///[overwriteMergedCells] when set to [false] puts the cell value to next unique cell available by putting the value in merged cells only once and jumps to next unique cell.
  ///
  void insertRowIterables(String sheet, List<dynamic> row, int rowIndex,
      {int startingColumn = 0, bool overwriteMergedCells = true}) {
    if (rowIndex < 0) {
      return;
    }
    _availSheet(sheet);
    _sheetMap['$sheet']!.insertRowIterables(row, rowIndex,
        startingColumn: startingColumn,
        overwriteMergedCells: overwriteMergedCells);
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
  ///       mySheet.findAndReplace('MySheetName', 'sad', 'happy', first: 3);
  ///
  ///In the above example it will replace all the occurences of `sad` with `happy` in the cells
  ///
  ///Other `options` are used to `narrow down` the `starting and ending ranges of cells`.
  ///
  int findAndReplace(String sheet, dynamic source, dynamic target,
      {int first = -1,
      int startingRow = -1,
      int endingRow = -1,
      int startingColumn = -1,
      int endingColumn = -1}) {
    int replaceCount = 0;
    if (_sheetMap[sheet] == null) return replaceCount;

    _sheetMap['$sheet']!.findAndReplace(
      source,
      target,
      first: first,
      startingRow: startingRow,
      endingRow: endingRow,
      startingColumn: startingColumn,
      endingColumn: endingColumn,
    );

    return replaceCount;
  }

  ///
  ///Make `sheet` available if it does not exist in `_sheetMap`
  ///
  void _availSheet(String sheet) {
    if (_sheetMap[sheet] == null) {
      _sheetMap[sheet] = Sheet._(this, sheet);
    }
  }

  ///
  ///Updates the contents of `sheet` of the `cellIndex: CellIndex.indexByColumnRow(0, 0);` where indexing starts from 0
  ///
  ///----or---- by `cellIndex: CellIndex.indexByString("A3");`.
  ///
  ///Styling of cell can be done by passing the CellStyle object to `cellStyle`.
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void updateCell(String sheet, CellIndex cellIndex, dynamic value,
      {CellStyle? cellStyle}) {
    _availSheet(sheet);

    if (cellStyle != null) {
      _colorChanges = true;
      _sheetMap[sheet]!.updateCell(cellIndex, value, cellStyle: cellStyle);
    } else {
      _sheetMap[sheet]!.updateCell(cellIndex, value);
    }
  }

  ///
  ///Merges the cells starting from `start` to `end`.
  ///
  ///If `custom value` is not defined then it will look for the very first available value in range `start` to `end` by searching row-wise from left to right.
  ///
  ///If `sheet` does not exist then it will be automatically created.
  ///
  void merge(String sheet, CellIndex start, CellIndex end,
      {dynamic customValue}) {
    _availSheet(sheet);
    _sheetMap[sheet]!.merge(start, end, customValue: customValue);
  }

  ///
  ///returns an Iterable of `cell-Id` for the previously merged cell-Ids.
  ///
  List<String> getMergedCells(String sheet) {
    return List<String>.from(
        _sheetMap[sheet] != null ? _sheetMap[sheet]!.spannedItems : <String>[]);
  }

  ///
  ///unMerge the merged cells.
  ///
  ///       var sheet = 'DesiredSheet';
  ///       List<String> spannedCells = excel.getMergedCells(sheet);
  ///       var cellToUnMerge = "A1:A2";
  ///       excel.unMerge(sheet, cellToUnMerge);
  ///
  void unMerge(String sheet, String unmergeCells) {
    if (_sheetMap[sheet] != null) {
      _sheetMap[sheet]!.unMerge(unmergeCells);
    }
  }

  ///
  ///Internal function taking care of adding the `sheetName` to the `mergeChangeLook` List
  ///So that merging function will be only called on `sheetNames of mergeChangeLook`
  ///
  set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
      //_mergeChanges = true;
    }
  }

  set _rtlChangeLookup(String value) {
    if (!_rtlChangeLook.contains(value)) {
      _rtlChangeLook.add(value);
      _rtlChanges = true;
    }
  }
}
