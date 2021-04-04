part of excel;

// ignore: must_be_immutable
class Data extends Equatable {
  CellStyle _cellStyle;
  dynamic _value;
  CellType _cellType;
  Sheet _sheet;
  String _sheetName;
  bool _isFormula;
  int _rowIndex;
  int _colIndex;

  ///It will clone the object by changing the `this` reference of previous DataObject and putting `new this` reference, with copying the values too
  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.colIndex,
          value_: dataObject._value,
          cellStyle_: dataObject._cellStyle,
          isFormula_: dataObject._isFormula,
          cellType_: dataObject._cellType,
        );

  ///Initializes the new `Data Object`
  Data._(
    Sheet sheet,
    int row,
    int col, {
    dynamic value_,
    CellStyle cellStyle_,
    bool isFormula_,
    CellType cellType_,
  }) {
    _sheet = sheet;
    _value = value_;
    _cellStyle = cellStyle_;
    _isFormula = isFormula_ ?? false;
    _cellType = cellType_ ?? CellType.String;
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
  String get cellId {
    return getCellId(_colIndex, _rowIndex);
  }

  set value(dynamic _value) {
    _sheet.updateCell(CellIndex.indexByString(cellId), _value);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  get value {
    return _value;
  }

  /// sets the user defined CellStyle in this current cell
  set style(CellStyle _style) {
    _cellStyle = _style;
  }

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle get cellStyle {
    return _cellStyle;
  }

  set cellStyle(CellStyle cellStyle_) {
    _sheet._excel._colorChanges = true;
    _cellStyle = cellStyle_;
  }

  @override
  List<Object> get props => [
        _value,
        _colIndex,
        _rowIndex,
        _cellStyle,
        _sheetName,
      ];
}
