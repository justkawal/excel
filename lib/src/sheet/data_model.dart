part of excel;

// ignore: must_be_immutable
class Data extends Equatable {
  CellStyle? _cellStyle;
  CellValue? _value;
  Sheet _sheet;
  String _sheetName;
  int _rowIndex;
  int _columnIndex;

  ///
  ///It will clone the object by changing the `this` reference of previous DataObject and putting `new this` reference, with copying the values too
  ///
  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.columnIndex,
          value: dataObject._value,
          cellStyleVal: dataObject._cellStyle,
        );

  ///
  ///Initializes the new `Data Object`
  ///
  Data._(
    Sheet sheet,
    int row,
    int column, {
    CellValue? value,
    NumFormat? numberFormat,
    CellStyle? cellStyleVal,
    bool isFormulaVal = false,
  })  : _sheet = sheet,
        _value = value,
        _cellStyle = cellStyleVal,
        _sheetName = sheet.sheetName,
        _rowIndex = row,
        _columnIndex = column;

  /// returns the newData object when called from Sheet Class
  static Data newData(Sheet sheet, int row, int column) {
    return Data._(sheet, row, column);
  }

  /// returns the row Index
  int get rowIndex {
    return _rowIndex;
  }

  /// returns the column Index
  int get columnIndex {
    return _columnIndex;
  }

  /// returns the sheet-name
  String get sheetName {
    return _sheetName;
  }

  /// returns the string based cellId as A1, A2 or Z5
  CellIndex get cellIndex {
    return CellIndex.indexByColumnRow(
        columnIndex: _columnIndex, rowIndex: _rowIndex);
  }

  /// Helps to set the formula
  ///```
  ///var sheet = excel['Sheet1'];
  ///var cell = sheet.cell(CellIndex.indexByString("E5"));
  ///cell.setFormula('=SUM(1,2)');
  ///```
  void setFormula(String formula) {
    _sheet.updateCell(cellIndex, FormulaCellValue(formula));
  }

  set value(CellValue? val) {
    _sheet.updateCell(cellIndex, val);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  CellValue? get value => _value;

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return _cellStyle;
  }

  /// sets the user defined CellStyle in this current cell
  set cellStyle(CellStyle? _) {
    _sheet._excel._styleChanges = true;
    _cellStyle = _;
  }

  @override
  List<Object?> get props => [
        _value,
        _columnIndex,
        _rowIndex,
        _cellStyle,
        _sheetName,
      ];
}

sealed class CellValue {
  const CellValue();
}

class FormulaCellValue extends CellValue {
  final String formula;

  const FormulaCellValue(this.formula);

  @override
  String toString() {
    return formula;
  }

  @override
  int get hashCode => Object.hash(runtimeType, formula);

  @override
  operator ==(Object other) {
    return other is FormulaCellValue && other.formula == formula;
  }
}

class IntCellValue extends CellValue {
  final int value;

  const IntCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is IntCellValue && other.value == value;
  }
}

class DoubleCellValue extends CellValue {
  final double value;

  const DoubleCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is DoubleCellValue && other.value == value;
  }
}

class DateCellValue extends CellValue {
  final int year;
  final int month;
  final int day;

  const DateCellValue({
    required this.year,
    required this.month,
    required this.day,
  })  : assert(month <= 12 && month >= 1),
        assert(day <= 31 && day >= 1);

  DateCellValue.fromDateTime(DateTime dt)
      : year = dt.year,
        month = dt.month,
        day = dt.day;

  DateTime asDateTimeLocal() {
    return DateTime(year, month, day);
  }

  DateTime asDateTimeUtc() {
    return DateTime.utc(year, month, day);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(runtimeType, year, month, day);

  @override
  operator ==(Object other) {
    return other is DateCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day;
  }
}

class TextCellValue extends CellValue {
  final TextSpan value;

  TextCellValue(String text) : value = TextSpan(text: text);
  TextCellValue.span(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is TextCellValue && other.value == value;
  }
}

class BoolCellValue extends CellValue {
  final bool value;

  const BoolCellValue(this.value);

  @override
  String toString() {
    return value.toString();
  }

  @override
  int get hashCode => Object.hash(runtimeType, value);

  @override
  operator ==(Object other) {
    return other is BoolCellValue && other.value == value;
  }
}

class TimeCellValue extends CellValue {
  final int hour;
  final int minute;
  final int second;
  final int millisecond;
  final int microsecond;

  const TimeCellValue({
    this.hour = 0,
    this.minute = 0,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  })  : assert(hour >= 0),
        assert(minute <= 60 && minute >= 0),
        assert(second <= 60 && second >= 0),
        assert(millisecond <= 1000 && millisecond >= 0),
        assert(microsecond <= 1000 && microsecond >= 0);

  /// [fractionOfDay]=1.0 is 24 hours, 0.5 is 12 hours and so on.
  factory TimeCellValue.fromFractionOfDay(num fractionOfDay) {
    var duration =
        Duration(milliseconds: (fractionOfDay * 24 * 3600 * 1000).round());
    return TimeCellValue.fromDuration(duration);
  }

  factory TimeCellValue.fromDuration(Duration duration) {
    final someUtcDate = DateTime.utc(0).add(duration);
    return TimeCellValue(
      hour: someUtcDate.hour,
      minute: someUtcDate.minute,
      second: someUtcDate.second,
      millisecond: someUtcDate.millisecond,
      microsecond: someUtcDate.microsecond,
    );
  }

  TimeCellValue.fromTimeOfDateTime(DateTime dt)
      : hour = dt.hour,
        minute = dt.minute,
        second = dt.second,
        millisecond = dt.millisecond,
        microsecond = dt.microsecond;

  Duration asDuration() {
    return Duration(
      hours: hour,
      minutes: minute,
      seconds: second,
      milliseconds: millisecond,
      microseconds: microsecond,
    );
  }

  @override
  String toString() {
    return '${_twoDigits(hour)}:${_twoDigits(minute)}:${_twoDigits(second)}';
  }

  @override
  int get hashCode => Object.hash(
        runtimeType,
        hour,
        minute,
        second,
        millisecond,
        microsecond,
      );

  @override
  operator ==(Object other) {
    return other is TimeCellValue &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}

/// Excel does not know if this is UTC or not. Use methods [asDateTimeLocal]
/// or [asDateTimeUtc] to get the DateTime object you prefer.
class DateTimeCellValue extends CellValue {
  final int year;
  final int month;
  final int day;
  final int hour;
  final int minute;
  final int second;
  final int millisecond;
  final int microsecond;

  const DateTimeCellValue({
    required this.year,
    required this.month,
    required this.day,
    required this.hour,
    required this.minute,
    this.second = 0,
    this.millisecond = 0,
    this.microsecond = 0,
  })  : assert(month <= 12 && month >= 1),
        assert(day <= 31 && day >= 1),
        assert(hour <= 24 && hour >= 0),
        assert(minute <= 60 && minute >= 0),
        assert(second <= 60 && second >= 0),
        assert(millisecond <= 1000 && millisecond >= 0),
        assert(microsecond <= 1000 && microsecond >= 0);

  DateTimeCellValue.fromDateTime(DateTime date)
      : year = date.year,
        month = date.month,
        day = date.day,
        hour = date.hour,
        minute = date.minute,
        second = date.second,
        millisecond = date.millisecond,
        microsecond = date.microsecond;

  DateTime asDateTimeLocal() {
    return DateTime(
        year, month, day, hour, minute, second, millisecond, microsecond);
  }

  DateTime asDateTimeUtc() {
    return DateTime.utc(
        year, month, day, hour, minute, second, millisecond, microsecond);
  }

  @override
  String toString() {
    return asDateTimeUtc().toIso8601String();
  }

  @override
  int get hashCode => Object.hash(
        runtimeType,
        year,
        month,
        day,
        hour,
        minute,
        second,
        millisecond,
        microsecond,
      );

  @override
  operator ==(Object other) {
    return other is DateTimeCellValue &&
        other.year == year &&
        other.month == month &&
        other.day == day &&
        other.hour == hour &&
        other.minute == minute &&
        other.second == second &&
        other.millisecond == millisecond &&
        other.microsecond == microsecond;
  }
}
