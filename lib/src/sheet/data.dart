part of excel;

class Data {
  CellStyle _cellStyle;
  dynamic _value;
  CellType _cellType;
  bool _isFormula;

  Data._() {
    this._value = null;
    this._cellStyle = null;
    this._isFormula = false;
    this._cellType = CellType.String;
  }

  set value(dynamic _value) {
    if (_value != null) {
      if (_value is Formula || _value.runtimeType == Formula) {
      } else {
        this._value = _value;
        this._cellType =
            _value.runtimeType == int || _value.runtimeType == double
                ? CellType.int
                : CellType.String;
      }
    }
  }

  get value {
    return _value;
  }

  set style(CellStyle _style) {
    this._cellStyle = _style;
  }

  get style {
    return this._cellStyle;
  }
}
