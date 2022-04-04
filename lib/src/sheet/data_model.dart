part of excel;

// ignore: must_be_immutable
class Data extends Equatable {
  CellStyle? _cellStyle;
  late dynamic _value;
  CellType _cellType = CellType.String;
  late Sheet _sheet;
  late String _sheetName;
  bool _isFormula = false;
  late int _rowIndex;
  late int _colIndex;

  ///
  ///It will clone the object by changing the `this` reference of previous DataObject and putting `new this` reference, with copying the values too
  ///
  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.colIndex,
          value_: dataObject._value,
          cellStyleVal: dataObject._cellStyle,
          isFormulaVal: dataObject._isFormula,
          cellTypeVal: dataObject._cellType,
        );

  ///
  ///Initializes the new `Data Object`
  ///
  Data._(
    Sheet sheet,
    int row,
    int col, {
    dynamic value_,
    CellStyle? cellStyleVal,
    bool isFormulaVal = false,
    CellType cellTypeVal = CellType.String,
  }) {
    _sheet = sheet;
    _value = value_;
    _cellStyle = cellStyleVal;
    _isFormula = isFormulaVal;
    _cellType = cellTypeVal;
    _sheetName = sheet.sheetName;
    _rowIndex = row;
    _colIndex = col;
  }

  /// returns the newData object when called from Sheet Class
  static Data newData(Sheet sheet, int row, int col) {
    return Data._(sheet, row, col);
  }

  /// returns the cell type
  CellType get cellType {
    return _cellType;
  }

  /// returns true is the cellType is CellType.Formula
  bool get isFormula {
    return _isFormula;
  }

  /// returns the row Index
  int get rowIndex {
    return _rowIndex;
  }

  /// returns the column Index
  int get colIndex {
    return _colIndex;
  }

  /// returns the sheet-name
  String get sheetName {
    return _sheetName;
  }

  /// returns the string based cellId as A1, A2 or Z5
  CellIndex get cellIndex {
    return CellIndex.indexByColumnRow(
        columnIndex: _colIndex, rowIndex: _rowIndex);
  }

  /// Helps to set the formula
  ///```
  ///var sheet = excel['Sheet1'];
  ///var cell = sheet.cell(CellIndex.indexByString("E5"));
  ///cell.setFormula('=SUM(1,2)');
  ///```
  void setFormula(String formula) {
    _sheet.updateCell(cellIndex, Formula.custom(formula));
  }

  set value(dynamic val) {
    _sheet.updateCell(cellIndex, val);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  get value {
    return _value;
  }

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return _cellStyle;
  }

  /// sets the user defined CellStyle in this current cell
  set cellStyle(CellStyle? _) {
    _sheet._excel._colorChanges = true;
    _cellStyle = _;
  }

  @override
  List<Object?> get props => [
        _value,
        _colIndex,
        _rowIndex,
        _cellStyle,
        _sheetName,
      ];
}
