part of excel;

Map<V, K> _createInverseMap<K, V>(Map<K, V> map) {
  final inverse = <V, K>{};
  for (var entry in map.entries) {
    assert(!inverse.containsKey(entry.value), 'map values are not unique');
    inverse[entry.value] = entry.key;
  }
  return inverse;
}

class NumFormatMaintainer {
  static const int _firstCustomFmtId = 164;
  int _nextFmtId = _firstCustomFmtId;
  Map<int, NumFormat> _map = {..._standardNumFormats};
  Map<NumFormat, int> _inverseMap =
      _createInverseMap(_standardNumFormats);

  void add(int numFmtId, CustomNumFormat format) {
    if (_map.containsKey(numFmtId)) {
      throw Exception('numFmtId $numFmtId already exists');
    }
    if (numFmtId < _firstCustomFmtId) {
      throw Exception(
          'invalid numFmtId $numFmtId, custom numFmtId must be $_firstCustomFmtId or greater');
    }
    _map[numFmtId] = format;
    _inverseMap[format] = numFmtId;
    if (numFmtId >= _nextFmtId) {
      _nextFmtId = numFmtId + 1;
    }
  }

  int findOrAdd(CustomNumFormat format) {
    var fmtId = _inverseMap[format];
    if (fmtId != null) {
      return fmtId;
    }
    fmtId = _nextFmtId;
    _nextFmtId++;
    _map[fmtId] = format;
    return fmtId;
  }

  void clear() {
    _nextFmtId = _firstCustomFmtId;
    _map = {..._standardNumFormats};
    _inverseMap = _createInverseMap(_standardNumFormats);
  }

  NumFormat? getByNumFmtId(int numFmtId) {
    return _map[numFmtId];
  }
}

sealed class NumFormat {
  final String formatCode;

  static const defaultNumeric = standard_0;
  static const defaultBool = standard_0;
  static const defaultDate = standard_14;
  static const defaultTime = standard_20;
  static const defaultDateTime = standard_22;

  static const standard_0 =
      StandardNumericNumFormat._(numFmtId: 0, formatCode: 'General');
  static const standard_1 =
      StandardNumericNumFormat._(numFmtId: 1, formatCode: "0");
  static const standard_2 =
      StandardNumericNumFormat._(numFmtId: 2, formatCode: "0.00");
  static const standard_3 =
      StandardNumericNumFormat._(numFmtId: 3, formatCode: "#,##0");
  static const standard_4 =
      StandardNumericNumFormat._(numFmtId: 4, formatCode: "#,##0.00");
  static const standard_9 =
      StandardNumericNumFormat._(numFmtId: 9, formatCode: "0%");
  static const standard_10 =
      StandardNumericNumFormat._(numFmtId: 10, formatCode: "0.00%");
  static const standard_11 =
      StandardNumericNumFormat._(numFmtId: 11, formatCode: "0.00E+00");
  static const standard_12 =
      StandardNumericNumFormat._(numFmtId: 12, formatCode: "# ?/?");
  static const standard_13 =
      StandardNumericNumFormat._(numFmtId: 13, formatCode: "# ??/??");
  static const standard_14 =
      StandardDateTimeNumFormat._(numFmtId: 14, formatCode: "mm-dd-yy");
  static const standard_15 =
      StandardDateTimeNumFormat._(numFmtId: 15, formatCode: "d-mmm-yy");
  static const standard_16 =
      StandardDateTimeNumFormat._(numFmtId: 16, formatCode: "d-mmm");
  static const standard_17 =
      StandardDateTimeNumFormat._(numFmtId: 17, formatCode: "mmm-yy");
  static const standard_18 =
      StandardTimeNumFormat._(numFmtId: 18, formatCode: "h:mm AM/PM");
  static const standard_19 =
      StandardTimeNumFormat._(numFmtId: 19, formatCode: "h:mm:ss AM/PM");
  static const standard_20 =
      StandardTimeNumFormat._(numFmtId: 20, formatCode: "h:mm");
  static const standard_21 =
      StandardTimeNumFormat._(numFmtId: 21, formatCode: "h:mm:dd");
  static const standard_22 =
      StandardDateTimeNumFormat._(numFmtId: 22, formatCode: "m/d/yy h:mm");
  static const standard_37 =
      StandardNumericNumFormat._(numFmtId: 37, formatCode: "#,##0 ;(#,##0)");
  static const standard_38 = StandardNumericNumFormat._(
      numFmtId: 38, formatCode: "#,##0 ;[Red](#,##0)");
  static const standard_39 = StandardNumericNumFormat._(
      numFmtId: 39, formatCode: "#,##0.00;(#,##0.00)");
  static const standard_40 = StandardNumericNumFormat._(
      numFmtId: 40, formatCode: "#,##0.00;[Red](#,#)");
  static const standard_45 =
      StandardTimeNumFormat._(numFmtId: 45, formatCode: "mm:ss");
  static const standard_46 =
      StandardTimeNumFormat._(numFmtId: 46, formatCode: "[h]:mm:ss");
  static const standard_47 =
      StandardTimeNumFormat._(numFmtId: 47, formatCode: "mmss.0");
  static const standard_48 =
      StandardNumericNumFormat._(numFmtId: 48, formatCode: "##0.0");
  static const standard_49 =
      StandardNumericNumFormat._(numFmtId: 49, formatCode: "@");

  const NumFormat({
    required this.formatCode,
  });

  static CustomNumFormat custom({
    required String formatCode,
  }) {
    if (formatCode == 'General') {
      return CustomNumericNumFormat(formatCode: 'General');
    }

    //const dateParts = ['m', 'mm', 'mmm', 'mmmm', 'mmmmm', 'd', 'dd', 'ddd', 'yy', 'yyyy'];
    //const timeParts = ['h', 'hh', 'm', 'mm', 's', 'ss', 'AM/PM'];

    /// mm appears in dateParts and timeParts, about this from the microsoft website:
    /// > If you use "m" immediately after the "h" or "hh" code or immediately before
    /// > the "ss" code, Excel displays minutes instead of the month.

    /// a very rudamentary check if we're talking date/time/numeric
    /// https://support.microsoft.com/en-us/office/format-numbers-as-dates-or-times-418bd3fe-0577-47c8-8caa-b4d30c528309
    /// or: https://www.ablebits.com/office-addins-blog/custom-excel-number-format/
    /// about dates: https://www.ablebits.com/office-addins-blog/change-date-format-excel/#custom-date-format
    /// about times: https://www.ablebits.com/office-addins-blog/excel-time-format/#custom
    /// [Green]#,##0.00\ \X\X"POSITIV";[Red]\-#\ "Negativ"\.##0.00

    if (_formatCodeLooksLikeDateTime(formatCode)) {
      return CustomDateTimeNumFormat(formatCode: formatCode);
    } else {
      return CustomNumericNumFormat(formatCode: formatCode);
    }
  }

  CellValue read(String v);

  @override
  int get hashCode => Object.hash(runtimeType, formatCode);

  @override
  operator ==(Object other) =>
      other.runtimeType == runtimeType &&
      (other as NumFormat).formatCode == formatCode;

  bool accepts(CellValue? value);

  static NumFormat defaultFor(CellValue? value) => switch (value) {
        null ||
        FormulaCellValue() ||
        IntCellValue() ||
        DoubleCellValue() ||
        TextCellValue() =>
          NumFormat.standard_0,
        DateCellValue() => NumFormat.defaultDate,
        BoolCellValue() => NumFormat.defaultBool,
        TimeCellValue() => NumFormat.defaultTime,
        DateTimeCellValue() => NumFormat.defaultDateTime,
      };
}

const Map<int, NumFormat> _standardNumFormats = {
  0: NumFormat.standard_0,
  1: NumFormat.standard_1,
  2: NumFormat.standard_2,
  3: NumFormat.standard_3,
  4: NumFormat.standard_4,
  9: NumFormat.standard_9,
  10: NumFormat.standard_10,
  11: NumFormat.standard_11,
  12: NumFormat.standard_12,
  13: NumFormat.standard_13,
  14: NumFormat.standard_14,
  15: NumFormat.standard_15,
  16: NumFormat.standard_16,
  17: NumFormat.standard_17,
  18: NumFormat.standard_18,
  19: NumFormat.standard_19,
  20: NumFormat.standard_20,
  21: NumFormat.standard_21,
  22: NumFormat.standard_22,
  37: NumFormat.standard_37,
  38: NumFormat.standard_38,
  39: NumFormat.standard_39,
  40: NumFormat.standard_40,
  45: NumFormat.standard_45,
  46: NumFormat.standard_46,
  47: NumFormat.standard_47,
  48: NumFormat.standard_48,
  49: NumFormat.standard_49,
};

bool _formatCodeLooksLikeDateTime(String formatCode) {
  // for comparison, remove any character that is quoted or escaped
  var inEscape = false;
  var inQuotes = false;
  for (var i = 0; i < formatCode.length; ++i) {
    final c = formatCode[i];
    if (inEscape) {
      inEscape = false;
      continue;
    } else if (c == '\\') {
      inEscape = true;
      continue;
    }
    if (inQuotes) {
      if (c == '"') {
        inQuotes = false;
      }
      continue;
    } else if (c == '"') {
      inQuotes = true;
      continue;
    }

    switch (c) {
      case 'y':
      case 'm':
      case 'd':
      case 'h':
      case 's':
        return true;
      case ';':
        // separator only exists for decimal formats
        return false;
      default:
        break;
    }
  }
  return false;
}

sealed class StandardNumFormat implements NumFormat {
  int get numFmtId;
}

sealed class CustomNumFormat implements NumFormat {
  String get formatCode;
}

sealed class NumericNumFormat extends NumFormat {
  const NumericNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    final decimalSeparatorIdx = v.indexOf('.');
    if (decimalSeparatorIdx == -1) {
      return IntCellValue(int.parse(v));
    }

    // also read .0 (or even .00) as an int
    bool noActualDecimalPlaces = true;
    for (var idx = decimalSeparatorIdx + 1; idx < v.length; ++idx) {
      if (v[idx] != '0') {
        noActualDecimalPlaces = false;
        break;
      }
    }
    if (noActualDecimalPlaces) {
      return IntCellValue(int.parse(v.substring(0, decimalSeparatorIdx)));
    }

    return DoubleCellValue(double.parse(v));
  }

  String writeDouble(DoubleCellValue value) {
    return value.value.toString();
  }

  String writeInt(IntCellValue value) {
    return value.value.toString();
  }
}

class StandardNumericNumFormat extends NumericNumFormat
    implements StandardNumFormat {
  @override
  final int numFmtId;

  const StandardNumericNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => true,
        TextCellValue() => numFmtId == 0,
        BoolCellValue() => true,
        DoubleCellValue() => true,
        DateCellValue() => false,
        TimeCellValue() => false,
        DateTimeCellValue() => false,
      };

  @override
  String toString() {
    return 'StandardNumericNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomNumericNumFormat extends NumericNumFormat
    implements CustomNumFormat {
  const CustomNumericNumFormat({
    required super.formatCode,
  });

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => true,
        TextCellValue() => false,
        BoolCellValue() => true,
        DoubleCellValue() => true,
        DateCellValue() => false,
        TimeCellValue() => false,
        DateTimeCellValue() => false,
      };

  @override
  String toString() {
    return 'CustomNumericNumFormat("$formatCode")';
  }
}

sealed class DateTimeNumFormat extends NumFormat {
  const DateTimeNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    if (v.startsWith('0.')) {
      return TimeCellValue.fromFractionOfDay(double.parse(v));
    }
    var delta = num.parse(v) * 24 * 3600 * 1000;
    var dateOffset = DateTime.utc(1899, 12, 30);
    final utcDate = dateOffset.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue.fromDateTime(utcDate);
    } else {
      return DateTimeCellValue.fromDateTime(utcDate);
    }
  }

  String writeDate(DateCellValue value) {
    var dateOffset = DateTime.utc(1899, 12, 30);
    final delta = value.asDateTimeUtc().difference(dateOffset);
    final dayFractions = delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return dayFractions.toString();
  }

  String writeDateTime(DateTimeCellValue value) {
    var dateOffset = DateTime.utc(1899, 12, 30);
    final delta = value.asDateTimeUtc().difference(dateOffset);
    final dayFractions = delta.inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return dayFractions.toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => false,
        TextCellValue() => false,
        BoolCellValue() => false,
        DoubleCellValue() => false,
        DateCellValue() => true,
        DateTimeCellValue() => true,
        TimeCellValue() => false,
      };
}

class StandardDateTimeNumFormat extends DateTimeNumFormat
    implements StandardNumFormat {
  final int numFmtId;

  const StandardDateTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardDateTimeNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomDateTimeNumFormat extends DateTimeNumFormat
    implements CustomNumFormat {
  const CustomDateTimeNumFormat({
    required super.formatCode,
  });

  @override
  String toString() {
    return 'CustomDateTimeNumFormat("$formatCode")';
  }
}

sealed class TimeNumFormat extends NumFormat {
  const TimeNumFormat({
    required super.formatCode,
  });

  @override
  CellValue read(String v) {
    if (v == '0') {
      return const TimeCellValue(
        hour: 0,
        minute: 0,
        second: 0,
        millisecond: 0,
        microsecond: 0,
      );
    }
    if (v.startsWith('0.')) {
      var delta = num.parse(v) * 24 * 3600 * 1000;
      final time = Duration(milliseconds: delta.round());
      final date = DateTime.utc(0).add(time);
      return TimeCellValue(
        hour: date.hour,
        minute: date.minute,
        second: date.second,
        millisecond: date.millisecond,
        microsecond: date.microsecond,
      );
    }
    var delta = num.parse(v) * 24 * 3600 * 1000;
    var dateOffset = DateTime.utc(1899, 12, 30);
    final utcDate = dateOffset.add(Duration(milliseconds: delta.round()));
    if (!v.contains('.') || v.endsWith('.0')) {
      return DateCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
      );
    } else {
      return DateTimeCellValue(
        year: utcDate.year,
        month: utcDate.month,
        day: utcDate.day,
        hour: utcDate.hour,
        minute: utcDate.minute,
        second: utcDate.second,
        millisecond: utcDate.millisecond,
        microsecond: utcDate.microsecond,
      );
    }
  }

  String writeTime(TimeCellValue value) {
    final fractionOfDay =
        value.asDuration().inMilliseconds.toDouble() / (1000 * 3600 * 24);
    return fractionOfDay.toString();
  }

  @override
  bool accepts(CellValue? value) => switch (value) {
        null => true,
        FormulaCellValue() => true,
        IntCellValue() => false,
        TextCellValue() => false,
        BoolCellValue() => false,
        DoubleCellValue() => false,
        DateCellValue() => false,
        DateTimeCellValue() => false,
        TimeCellValue() => true,
      };
}

class StandardTimeNumFormat extends TimeNumFormat
    implements StandardNumFormat {
  final int numFmtId;

  const StandardTimeNumFormat._({
    required this.numFmtId,
    required super.formatCode,
  });

  @override
  String toString() {
    return 'StandardTimeNumFormat($numFmtId, "$formatCode")';
  }
}

class CustomTimeNumFormat extends TimeNumFormat
    implements CustomNumFormat {
  const CustomTimeNumFormat({
    required super.formatCode,
  });

  @override
  String toString() {
    return 'CustomTimeNumFormat("$formatCode")';
  }
}
