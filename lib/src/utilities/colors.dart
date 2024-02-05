part of excel;

String _decimalToHexadecimal(int decimalVal) {
  if (decimalVal == 0) {
    return '0';
  }
  bool negative = false;
  if (decimalVal < 0) {
    negative = true;
    decimalVal *= -1;
  }
  String hexString = '';
  while (decimalVal > 0) {
    String hexVal = '';
    final int remainder = decimalVal % 16;
    decimalVal = decimalVal ~/ 16;
    if (_hexTable.containsKey(remainder)) {
      hexVal = _hexTable[remainder]!;
    } else {
      hexVal = remainder.toString();
    }
    hexString = hexVal + hexString;
  }
  return negative ? '-$hexString' : hexString;
}

bool _assertHexString(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      return false;
    }
  }
  return true;
}

int _hexadecimalToDecimal(String hexString) {
  hexString = hexString.replaceAll('#', '').trim().toUpperCase();

  final bool isNegative = hexString[0] == '-';
  if (isNegative) hexString = hexString.substring(1);

  int decimalVal = 0;
  for (int i = 0; i < hexString.length; i++) {
    if (int.tryParse(hexString[i]) == null &&
        _hexTableReverse.containsKey(hexString[i]) == false) {
      throw Exception('Non-hex value was passed to the function');
    } else {
      decimalVal += (pow(16, hexString.length - i - 1) *
              (int.tryParse(hexString[i]) != null
                  ? int.parse(hexString[i])
                  : _hexTableReverse[hexString[i]]!))
          .toInt();
    }
  }
  return isNegative ? -1 * decimalVal : decimalVal;
}

const _hexTable = {
  10: 'A',
  11: 'B',
  12: 'C',
  13: 'D',
  14: 'E',
  15: 'F',
};

final _hexTableReverse = _hexTable.map((k, v) => MapEntry(v, k));

extension StringExt on String {
  /// Return [ExcelColor.black] if not a color hexadecimal
  ExcelColor get excelColor => this == 'none'
      ? ExcelColor.none
      : _assertHexString(this)
          ? ExcelColor._(this)
          : ExcelColor.black;
}

/// Copying from Flutter Material Color
class ExcelColor extends Equatable {
  const ExcelColor._([this._color, this._name, this._type]);

  final String? _color;
  final String? _name;
  final ColorType? _type;

  /// Return 'none' if [_color] is null, [black] if not match
  String get colorHex => _color == null
      ? 'none'
      : _assertHexString(_color!)
          ? _color!
          : black.colorHex;

  /// Return [black] if [_color] is null and not match
  int get colorInt => _color == null
      ? 0xFF000000
      : _assertHexString(_color!)
          ? _hexadecimalToDecimal(_color!)
          : 0xFF000000;

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromInt(int colorIntValue) =>
      ExcelColor._(_decimalToHexadecimal(colorIntValue));

  /// Warning! Highly unsafe method.
  /// Can break your excel file if you do not know what you are doing
  factory ExcelColor.fromHexString(String colorHexValue) =>
      ExcelColor._(colorHexValue);

  static const none = ExcelColor._();

  static const black = ExcelColor._('FF000000', 'black', ColorType.color);
  static const black12 = ExcelColor._('1F000000', 'black12', ColorType.color);
  static const black26 = ExcelColor._('42000000', 'black26', ColorType.color);
  static const black38 = ExcelColor._('61000000', 'black38', ColorType.color);
  static const black45 = ExcelColor._('73000000', 'black45', ColorType.color);
  static const black54 = ExcelColor._('8A000000', 'black54', ColorType.color);
  static const black87 = ExcelColor._('DD000000', 'black87', ColorType.color);
  static const white = ExcelColor._('FFFFFFFF', 'white', ColorType.color);
  static const white10 = ExcelColor._('1AFFFFFF', 'white10', ColorType.color);
  static const white12 = ExcelColor._('1FFFFFFF', 'white12', ColorType.color);
  static const white24 = ExcelColor._('3DFFFFFF', 'white24', ColorType.color);
  static const white30 = ExcelColor._('4DFFFFFF', 'white30', ColorType.color);
  static const white38 = ExcelColor._('62FFFFFF', 'white38', ColorType.color);
  static const white54 = ExcelColor._('8AFFFFFF', 'white54', ColorType.color);
  static const white60 = ExcelColor._('99FFFFFF', 'white60', ColorType.color);
  static const white70 = ExcelColor._('B3FFFFFF', 'white70', ColorType.color);
  static const redAccent =
      ExcelColor._('FFFF5252', 'redAccent', ColorType.materialAccent);
  static const pinkAccent =
      ExcelColor._('FFFF4081', 'pinkAccent', ColorType.materialAccent);
  static const purpleAccent =
      ExcelColor._('FFE040FB', 'purpleAccent', ColorType.materialAccent);
  static const deepPurpleAccent =
      ExcelColor._('FF7C4DFF', 'deepPurpleAccent', ColorType.materialAccent);
  static const indigoAccent =
      ExcelColor._('FF536DFE', 'indigoAccent', ColorType.materialAccent);
  static const blueAccent =
      ExcelColor._('FF448AFF', 'blueAccent', ColorType.materialAccent);
  static const lightBlueAccent =
      ExcelColor._('FF40C4FF', 'lightBlueAccent', ColorType.materialAccent);
  static const cyanAccent =
      ExcelColor._('FF18FFFF', 'cyanAccent', ColorType.materialAccent);
  static const tealAccent =
      ExcelColor._('FF64FFDA', 'tealAccent', ColorType.materialAccent);
  static const greenAccent =
      ExcelColor._('FF69F0AE', 'greenAccent', ColorType.materialAccent);
  static const lightGreenAccent =
      ExcelColor._('FFB2FF59', 'lightGreenAccent', ColorType.materialAccent);
  static const limeAccent =
      ExcelColor._('FFEEFF41', 'limeAccent', ColorType.materialAccent);
  static const yellowAccent =
      ExcelColor._('FFFFFF00', 'yellowAccent', ColorType.materialAccent);
  static const amberAccent =
      ExcelColor._('FFFFD740', 'amberAccent', ColorType.materialAccent);
  static const orangeAccent =
      ExcelColor._('FFFFAB40', 'orangeAccent', ColorType.materialAccent);
  static const deepOrangeAccent =
      ExcelColor._('FFFF6E40', 'deepOrangeAccent', ColorType.materialAccent);
  static const red = ExcelColor._('FFF44336', 'red', ColorType.material);
  static const pink = ExcelColor._('FFE91E63', 'pink', ColorType.material);
  static const purple = ExcelColor._('FF9C27B0', 'purple', ColorType.material);
  static const deepPurple =
      ExcelColor._('FF673AB7', 'deepPurple', ColorType.material);
  static const indigo = ExcelColor._('FF3F51B5', 'indigo', ColorType.material);
  static const blue = ExcelColor._('FF2196F3', 'blue', ColorType.material);
  static const lightBlue =
      ExcelColor._('FF03A9F4', 'lightBlue', ColorType.material);
  static const cyan = ExcelColor._('FF00BCD4', 'cyan', ColorType.material);
  static const teal = ExcelColor._('FF009688', 'teal', ColorType.material);
  static const green = ExcelColor._('FF4CAF50', 'green', ColorType.material);
  static const lightGreen =
      ExcelColor._('FF8BC34A', 'lightGreen', ColorType.material);
  static const lime = ExcelColor._('FFCDDC39', 'lime', ColorType.material);
  static const yellow = ExcelColor._('FFFFEB3B', 'yellow', ColorType.material);
  static const amber = ExcelColor._('FFFFC107', 'amber', ColorType.material);
  static const orange = ExcelColor._('FFFF9800', 'orange', ColorType.material);
  static const deepOrange =
      ExcelColor._('FFFF5722', 'deepOrange', ColorType.material);
  static const brown = ExcelColor._('FF795548', 'brown', ColorType.material);
  static const grey = ExcelColor._('FF9E9E9E', 'grey', ColorType.material);
  static const blueGrey =
      ExcelColor._('FF607D8B', 'blueGrey', ColorType.material);
  static const redAccent100 =
      ExcelColor._('FFFF8A80', 'redAccent100', ColorType.materialAccent);
  static const redAccent200 =
      ExcelColor._('FFFF5252', 'redAccent200', ColorType.materialAccent);
  static const redAccent400 =
      ExcelColor._('FFFF1744', 'redAccent400', ColorType.materialAccent);
  static const redAccent700 =
      ExcelColor._('FFD50000', 'redAccent700', ColorType.materialAccent);
  static const pinkAccent100 =
      ExcelColor._('FFFF80AB', 'pinkAccent100', ColorType.materialAccent);
  static const pinkAccent200 =
      ExcelColor._('FFFF4081', 'pinkAccent200', ColorType.materialAccent);
  static const pinkAccent400 =
      ExcelColor._('FFF50057', 'pinkAccent400', ColorType.materialAccent);
  static const pinkAccent700 =
      ExcelColor._('FFC51162', 'pinkAccent700', ColorType.materialAccent);
  static const purpleAccent100 =
      ExcelColor._('FFEA80FC', 'purpleAccent100', ColorType.materialAccent);
  static const purpleAccent200 =
      ExcelColor._('FFE040FB', 'purpleAccent200', ColorType.materialAccent);
  static const purpleAccent400 =
      ExcelColor._('FFD500F9', 'purpleAccent400', ColorType.materialAccent);
  static const purpleAccent700 =
      ExcelColor._('FFAA00FF', 'purpleAccent700', ColorType.materialAccent);
  static const deepPurpleAccent100 =
      ExcelColor._('FFB388FF', 'deepPurpleAccent100', ColorType.materialAccent);
  static const deepPurpleAccent200 =
      ExcelColor._('FF7C4DFF', 'deepPurpleAccent200', ColorType.materialAccent);
  static const deepPurpleAccent400 =
      ExcelColor._('FF651FFF', 'deepPurpleAccent400', ColorType.materialAccent);
  static const deepPurpleAccent700 =
      ExcelColor._('FF6200EA', 'deepPurpleAccent700', ColorType.materialAccent);
  static const indigoAccent100 =
      ExcelColor._('FF8C9EFF', 'indigoAccent100', ColorType.materialAccent);
  static const indigoAccent200 =
      ExcelColor._('FF536DFE', 'indigoAccent200', ColorType.materialAccent);
  static const indigoAccent400 =
      ExcelColor._('FF3D5AFE', 'indigoAccent400', ColorType.materialAccent);
  static const indigoAccent700 =
      ExcelColor._('FF304FFE', 'indigoAccent700', ColorType.materialAccent);
  static const blueAccent100 =
      ExcelColor._('FF82B1FF', 'blueAccent100', ColorType.materialAccent);
  static const blueAccent200 =
      ExcelColor._('FF448AFF', 'blueAccent200', ColorType.materialAccent);
  static const blueAccent400 =
      ExcelColor._('FF2979FF', 'blueAccent400', ColorType.materialAccent);
  static const blueAccent700 =
      ExcelColor._('FF2962FF', 'blueAccent700', ColorType.materialAccent);
  static const lightBlueAccent100 =
      ExcelColor._('FF80D8FF', 'lightBlueAccent100', ColorType.materialAccent);
  static const lightBlueAccent200 =
      ExcelColor._('FF40C4FF', 'lightBlueAccent200', ColorType.materialAccent);
  static const lightBlueAccent400 =
      ExcelColor._('FF00B0FF', 'lightBlueAccent400', ColorType.materialAccent);
  static const lightBlueAccent700 =
      ExcelColor._('FF0091EA', 'lightBlueAccent700', ColorType.materialAccent);
  static const cyanAccent100 =
      ExcelColor._('FF84FFFF', 'cyanAccent100', ColorType.materialAccent);
  static const cyanAccent200 =
      ExcelColor._('FF18FFFF', 'cyanAccent200', ColorType.materialAccent);
  static const cyanAccent400 =
      ExcelColor._('FF00E5FF', 'cyanAccent400', ColorType.materialAccent);
  static const cyanAccent700 =
      ExcelColor._('FF00B8D4', 'cyanAccent700', ColorType.materialAccent);
  static const tealAccent100 =
      ExcelColor._('FFA7FFEB', 'tealAccent100', ColorType.materialAccent);
  static const tealAccent200 =
      ExcelColor._('FF64FFDA', 'tealAccent200', ColorType.materialAccent);
  static const tealAccent400 =
      ExcelColor._('FF1DE9B6', 'tealAccent400', ColorType.materialAccent);
  static const tealAccent700 =
      ExcelColor._('FF00BFA5', 'tealAccent700', ColorType.materialAccent);
  static const greenAccent100 =
      ExcelColor._('FFB9F6CA', 'greenAccent100', ColorType.materialAccent);
  static const greenAccent200 =
      ExcelColor._('FF69F0AE', 'greenAccent200', ColorType.materialAccent);
  static const greenAccent400 =
      ExcelColor._('FF00E676', 'greenAccent400', ColorType.materialAccent);
  static const greenAccent700 =
      ExcelColor._('FF00C853', 'greenAccent700', ColorType.materialAccent);
  static const lightGreenAccent100 =
      ExcelColor._('FFCCFF90', 'lightGreenAccent100', ColorType.materialAccent);
  static const lightGreenAccent200 =
      ExcelColor._('FFB2FF59', 'lightGreenAccent200', ColorType.materialAccent);
  static const lightGreenAccent400 =
      ExcelColor._('FF76FF03', 'lightGreenAccent400', ColorType.materialAccent);
  static const lightGreenAccent700 =
      ExcelColor._('FF64DD17', 'lightGreenAccent700', ColorType.materialAccent);
  static const limeAccent100 =
      ExcelColor._('FFF4FF81', 'limeAccent100', ColorType.materialAccent);
  static const limeAccent200 =
      ExcelColor._('FFEEFF41', 'limeAccent200', ColorType.materialAccent);
  static const limeAccent400 =
      ExcelColor._('FFC6FF00', 'limeAccent400', ColorType.materialAccent);
  static const limeAccent700 =
      ExcelColor._('FFAEEA00', 'limeAccent700', ColorType.materialAccent);
  static const yellowAccent100 =
      ExcelColor._('FFFFFF8D', 'yellowAccent100', ColorType.materialAccent);
  static const yellowAccent200 =
      ExcelColor._('FFFFFF00', 'yellowAccent200', ColorType.materialAccent);
  static const yellowAccent400 =
      ExcelColor._('FFFFEA00', 'yellowAccent400', ColorType.materialAccent);
  static const yellowAccent700 =
      ExcelColor._('FFFFD600', 'yellowAccent700', ColorType.materialAccent);
  static const amberAccent100 =
      ExcelColor._('FFFFE57F', 'amberAccent100', ColorType.materialAccent);
  static const amberAccent200 =
      ExcelColor._('FFFFD740', 'amberAccent200', ColorType.materialAccent);
  static const amberAccent400 =
      ExcelColor._('FFFFC400', 'amberAccent400', ColorType.materialAccent);
  static const amberAccent700 =
      ExcelColor._('FFFFAB00', 'amberAccent700', ColorType.materialAccent);
  static const orangeAccent100 =
      ExcelColor._('FFFFD180', 'orangeAccent100', ColorType.materialAccent);
  static const orangeAccent200 =
      ExcelColor._('FFFFAB40', 'orangeAccent200', ColorType.materialAccent);
  static const orangeAccent400 =
      ExcelColor._('FFFF9100', 'orangeAccent400', ColorType.materialAccent);
  static const orangeAccent700 =
      ExcelColor._('FFFF6D00', 'orangeAccent700', ColorType.materialAccent);
  static const deepOrangeAccent100 =
      ExcelColor._('FFFF9E80', 'deepOrangeAccent100', ColorType.materialAccent);
  static const deepOrangeAccent200 =
      ExcelColor._('FFFF6E40', 'deepOrangeAccent200', ColorType.materialAccent);
  static const deepOrangeAccent400 =
      ExcelColor._('FFFF3D00', 'deepOrangeAccent400', ColorType.materialAccent);
  static const deepOrangeAccent700 =
      ExcelColor._('FFDD2C00', 'deepOrangeAccent700', ColorType.materialAccent);
  static const red50 = ExcelColor._('FFFFEBEE', 'red50', ColorType.material);
  static const red100 = ExcelColor._('FFFFCDD2', 'red100', ColorType.material);
  static const red200 = ExcelColor._('FFEF9A9A', 'red200', ColorType.material);
  static const red300 = ExcelColor._('FFE57373', 'red300', ColorType.material);
  static const red400 = ExcelColor._('FFEF5350', 'red400', ColorType.material);
  static const red500 = ExcelColor._('FFF44336', 'red500', ColorType.material);
  static const red600 = ExcelColor._('FFE53935', 'red600', ColorType.material);
  static const red700 = ExcelColor._('FFD32F2F', 'red700', ColorType.material);
  static const red800 = ExcelColor._('FFC62828', 'red800', ColorType.material);
  static const red900 = ExcelColor._('FFB71C1C', 'red900', ColorType.material);
  static const pink50 = ExcelColor._('FFFCE4EC', 'pink50', ColorType.material);
  static const pink100 =
      ExcelColor._('FFF8BBD0', 'pink100', ColorType.material);
  static const pink200 =
      ExcelColor._('FFF48FB1', 'pink200', ColorType.material);
  static const pink300 =
      ExcelColor._('FFF06292', 'pink300', ColorType.material);
  static const pink400 =
      ExcelColor._('FFEC407A', 'pink400', ColorType.material);
  static const pink500 =
      ExcelColor._('FFE91E63', 'pink500', ColorType.material);
  static const pink600 =
      ExcelColor._('FFD81B60', 'pink600', ColorType.material);
  static const pink700 =
      ExcelColor._('FFC2185B', 'pink700', ColorType.material);
  static const pink800 =
      ExcelColor._('FFAD1457', 'pink800', ColorType.material);
  static const pink900 =
      ExcelColor._('FF880E4F', 'pink900', ColorType.material);
  static const purple50 =
      ExcelColor._('FFF3E5F5', 'purple50', ColorType.material);
  static const purple100 =
      ExcelColor._('FFE1BEE7', 'purple100', ColorType.material);
  static const purple200 =
      ExcelColor._('FFCE93D8', 'purple200', ColorType.material);
  static const purple300 =
      ExcelColor._('FFBA68C8', 'purple300', ColorType.material);
  static const purple400 =
      ExcelColor._('FFAB47BC', 'purple400', ColorType.material);
  static const purple500 =
      ExcelColor._('FF9C27B0', 'purple500', ColorType.material);
  static const purple600 =
      ExcelColor._('FF8E24AA', 'purple600', ColorType.material);
  static const purple700 =
      ExcelColor._('FF7B1FA2', 'purple700', ColorType.material);
  static const purple800 =
      ExcelColor._('FF6A1B9A', 'purple800', ColorType.material);
  static const purple900 =
      ExcelColor._('FF4A148C', 'purple900', ColorType.material);
  static const deepPurple50 =
      ExcelColor._('FFEDE7F6', 'deepPurple50', ColorType.material);
  static const deepPurple100 =
      ExcelColor._('FFD1C4E9', 'deepPurple100', ColorType.material);
  static const deepPurple200 =
      ExcelColor._('FFB39DDB', 'deepPurple200', ColorType.material);
  static const deepPurple300 =
      ExcelColor._('FF9575CD', 'deepPurple300', ColorType.material);
  static const deepPurple400 =
      ExcelColor._('FF7E57C2', 'deepPurple400', ColorType.material);
  static const deepPurple500 =
      ExcelColor._('FF673AB7', 'deepPurple500', ColorType.material);
  static const deepPurple600 =
      ExcelColor._('FF5E35B1', 'deepPurple600', ColorType.material);
  static const deepPurple700 =
      ExcelColor._('FF512DA8', 'deepPurple700', ColorType.material);
  static const deepPurple800 =
      ExcelColor._('FF4527A0', 'deepPurple800', ColorType.material);
  static const deepPurple900 =
      ExcelColor._('FF311B92', 'deepPurple900', ColorType.material);
  static const indigo50 =
      ExcelColor._('FFE8EAF6', 'indigo50', ColorType.material);
  static const indigo100 =
      ExcelColor._('FFC5CAE9', 'indigo100', ColorType.material);
  static const indigo200 =
      ExcelColor._('FF9FA8DA', 'indigo200', ColorType.material);
  static const indigo300 =
      ExcelColor._('FF7986CB', 'indigo300', ColorType.material);
  static const indigo400 =
      ExcelColor._('FF5C6BC0', 'indigo400', ColorType.material);
  static const indigo500 =
      ExcelColor._('FF3F51B5', 'indigo500', ColorType.material);
  static const indigo600 =
      ExcelColor._('FF3949AB', 'indigo600', ColorType.material);
  static const indigo700 =
      ExcelColor._('FF303F9F', 'indigo700', ColorType.material);
  static const indigo800 =
      ExcelColor._('FF283593', 'indigo800', ColorType.material);
  static const indigo900 =
      ExcelColor._('FF1A237E', 'indigo900', ColorType.material);
  static const blue50 = ExcelColor._('FFE3F2FD', 'blue50', ColorType.material);
  static const blue100 =
      ExcelColor._('FFBBDEFB', 'blue100', ColorType.material);
  static const blue200 =
      ExcelColor._('FF90CAF9', 'blue200', ColorType.material);
  static const blue300 =
      ExcelColor._('FF64B5F6', 'blue300', ColorType.material);
  static const blue400 =
      ExcelColor._('FF42A5F5', 'blue400', ColorType.material);
  static const blue500 =
      ExcelColor._('FF2196F3', 'blue500', ColorType.material);
  static const blue600 =
      ExcelColor._('FF1E88E5', 'blue600', ColorType.material);
  static const blue700 =
      ExcelColor._('FF1976D2', 'blue700', ColorType.material);
  static const blue800 =
      ExcelColor._('FF1565C0', 'blue800', ColorType.material);
  static const blue900 =
      ExcelColor._('FF0D47A1', 'blue900', ColorType.material);
  static const lightBlue50 =
      ExcelColor._('FFE1F5FE', 'lightBlue50', ColorType.material);
  static const lightBlue100 =
      ExcelColor._('FFB3E5FC', 'lightBlue100', ColorType.material);
  static const lightBlue200 =
      ExcelColor._('FF81D4FA', 'lightBlue200', ColorType.material);
  static const lightBlue300 =
      ExcelColor._('FF4FC3F7', 'lightBlue300', ColorType.material);
  static const lightBlue400 =
      ExcelColor._('FF29B6F6', 'lightBlue400', ColorType.material);
  static const lightBlue500 =
      ExcelColor._('FF03A9F4', 'lightBlue500', ColorType.material);
  static const lightBlue600 =
      ExcelColor._('FF039BE5', 'lightBlue600', ColorType.material);
  static const lightBlue700 =
      ExcelColor._('FF0288D1', 'lightBlue700', ColorType.material);
  static const lightBlue800 =
      ExcelColor._('FF0277BD', 'lightBlue800', ColorType.material);
  static const lightBlue900 =
      ExcelColor._('FF01579B', 'lightBlue900', ColorType.material);
  static const cyan50 = ExcelColor._('FFE0F7FA', 'cyan50', ColorType.material);
  static const cyan100 =
      ExcelColor._('FFB2EBF2', 'cyan100', ColorType.material);
  static const cyan200 =
      ExcelColor._('FF80DEEA', 'cyan200', ColorType.material);
  static const cyan300 =
      ExcelColor._('FF4DD0E1', 'cyan300', ColorType.material);
  static const cyan400 =
      ExcelColor._('FF26C6DA', 'cyan400', ColorType.material);
  static const cyan500 =
      ExcelColor._('FF00BCD4', 'cyan500', ColorType.material);
  static const cyan600 =
      ExcelColor._('FF00ACC1', 'cyan600', ColorType.material);
  static const cyan700 =
      ExcelColor._('FF0097A7', 'cyan700', ColorType.material);
  static const cyan800 =
      ExcelColor._('FF00838F', 'cyan800', ColorType.material);
  static const cyan900 =
      ExcelColor._('FF006064', 'cyan900', ColorType.material);
  static const teal50 = ExcelColor._('FFE0F2F1', 'teal50', ColorType.material);
  static const teal100 =
      ExcelColor._('FFB2DFDB', 'teal100', ColorType.material);
  static const teal200 =
      ExcelColor._('FF80CBC4', 'teal200', ColorType.material);
  static const teal300 =
      ExcelColor._('FF4DB6AC', 'teal300', ColorType.material);
  static const teal400 =
      ExcelColor._('FF26A69A', 'teal400', ColorType.material);
  static const teal500 =
      ExcelColor._('FF009688', 'teal500', ColorType.material);
  static const teal600 =
      ExcelColor._('FF00897B', 'teal600', ColorType.material);
  static const teal700 =
      ExcelColor._('FF00796B', 'teal700', ColorType.material);
  static const teal800 =
      ExcelColor._('FF00695C', 'teal800', ColorType.material);
  static const teal900 =
      ExcelColor._('FF004D40', 'teal900', ColorType.material);
  static const green50 =
      ExcelColor._('FFE8F5E9', 'green50', ColorType.material);
  static const green100 =
      ExcelColor._('FFC8E6C9', 'green100', ColorType.material);
  static const green200 =
      ExcelColor._('FFA5D6A7', 'green200', ColorType.material);
  static const green300 =
      ExcelColor._('FF81C784', 'green300', ColorType.material);
  static const green400 =
      ExcelColor._('FF66BB6A', 'green400', ColorType.material);
  static const green500 =
      ExcelColor._('FF4CAF50', 'green500', ColorType.material);
  static const green600 =
      ExcelColor._('FF43A047', 'green600', ColorType.material);
  static const green700 =
      ExcelColor._('FF388E3C', 'green700', ColorType.material);
  static const green800 =
      ExcelColor._('FF2E7D32', 'green800', ColorType.material);
  static const green900 =
      ExcelColor._('FF1B5E20', 'green900', ColorType.material);
  static const lightGreen50 =
      ExcelColor._('FFF1F8E9', 'lightGreen50', ColorType.material);
  static const lightGreen100 =
      ExcelColor._('FFDCEDC8', 'lightGreen100', ColorType.material);
  static const lightGreen200 =
      ExcelColor._('FFC5E1A5', 'lightGreen200', ColorType.material);
  static const lightGreen300 =
      ExcelColor._('FFAED581', 'lightGreen300', ColorType.material);
  static const lightGreen400 =
      ExcelColor._('FF9CCC65', 'lightGreen400', ColorType.material);
  static const lightGreen500 =
      ExcelColor._('FF8BC34A', 'lightGreen500', ColorType.material);
  static const lightGreen600 =
      ExcelColor._('FF7CB342', 'lightGreen600', ColorType.material);
  static const lightGreen700 =
      ExcelColor._('FF689F38', 'lightGreen700', ColorType.material);
  static const lightGreen800 =
      ExcelColor._('FF558B2F', 'lightGreen800', ColorType.material);
  static const lightGreen900 =
      ExcelColor._('FF33691E', 'lightGreen900', ColorType.material);
  static const lime50 = ExcelColor._('FFF9FBE7', 'lime50', ColorType.material);
  static const lime100 =
      ExcelColor._('FFF0F4C3', 'lime100', ColorType.material);
  static const lime200 =
      ExcelColor._('FFE6EE9C', 'lime200', ColorType.material);
  static const lime300 =
      ExcelColor._('FFDCE775', 'lime300', ColorType.material);
  static const lime400 =
      ExcelColor._('FFD4E157', 'lime400', ColorType.material);
  static const lime500 =
      ExcelColor._('FFCDDC39', 'lime500', ColorType.material);
  static const lime600 =
      ExcelColor._('FFC0CA33', 'lime600', ColorType.material);
  static const lime700 =
      ExcelColor._('FFAFB42B', 'lime700', ColorType.material);
  static const lime800 =
      ExcelColor._('FF9E9D24', 'lime800', ColorType.material);
  static const lime900 =
      ExcelColor._('FF827717', 'lime900', ColorType.material);
  static const yellow50 =
      ExcelColor._('FFFFFDE7', 'yellow50', ColorType.material);
  static const yellow100 =
      ExcelColor._('FFFFF9C4', 'yellow100', ColorType.material);
  static const yellow200 =
      ExcelColor._('FFFFF59D', 'yellow200', ColorType.material);
  static const yellow300 =
      ExcelColor._('FFFFF176', 'yellow300', ColorType.material);
  static const yellow400 =
      ExcelColor._('FFFFEE58', 'yellow400', ColorType.material);
  static const yellow500 =
      ExcelColor._('FFFFEB3B', 'yellow500', ColorType.material);
  static const yellow600 =
      ExcelColor._('FFFDD835', 'yellow600', ColorType.material);
  static const yellow700 =
      ExcelColor._('FFFBC02D', 'yellow700', ColorType.material);
  static const yellow800 =
      ExcelColor._('FFF9A825', 'yellow800', ColorType.material);
  static const yellow900 =
      ExcelColor._('FFF57F17', 'yellow900', ColorType.material);
  static const amber50 =
      ExcelColor._('FFFFF8E1', 'amber50', ColorType.material);
  static const amber100 =
      ExcelColor._('FFFFECB3', 'amber100', ColorType.material);
  static const amber200 =
      ExcelColor._('FFFFE082', 'amber200', ColorType.material);
  static const amber300 =
      ExcelColor._('FFFFD54F', 'amber300', ColorType.material);
  static const amber400 =
      ExcelColor._('FFFFCA28', 'amber400', ColorType.material);
  static const amber500 =
      ExcelColor._('FFFFC107', 'amber500', ColorType.material);
  static const amber600 =
      ExcelColor._('FFFFB300', 'amber600', ColorType.material);
  static const amber700 =
      ExcelColor._('FFFFA000', 'amber700', ColorType.material);
  static const amber800 =
      ExcelColor._('FFFF8F00', 'amber800', ColorType.material);
  static const amber900 =
      ExcelColor._('FFFF6F00', 'amber900', ColorType.material);
  static const orange50 =
      ExcelColor._('FFFFF3E0', 'orange50', ColorType.material);
  static const orange100 =
      ExcelColor._('FFFFE0B2', 'orange100', ColorType.material);
  static const orange200 =
      ExcelColor._('FFFFCC80', 'orange200', ColorType.material);
  static const orange300 =
      ExcelColor._('FFFFB74D', 'orange300', ColorType.material);
  static const orange400 =
      ExcelColor._('FFFFA726', 'orange400', ColorType.material);
  static const orange500 =
      ExcelColor._('FFFF9800', 'orange500', ColorType.material);
  static const orange600 =
      ExcelColor._('FFFB8C00', 'orange600', ColorType.material);
  static const orange700 =
      ExcelColor._('FFF57C00', 'orange700', ColorType.material);
  static const orange800 =
      ExcelColor._('FFEF6C00', 'orange800', ColorType.material);
  static const orange900 =
      ExcelColor._('FFE65100', 'orange900', ColorType.material);
  static const deepOrange50 =
      ExcelColor._('FFFBE9E7', 'deepOrange50', ColorType.material);
  static const deepOrange100 =
      ExcelColor._('FFFFCCBC', 'deepOrange100', ColorType.material);
  static const deepOrange200 =
      ExcelColor._('FFFFAB91', 'deepOrange200', ColorType.material);
  static const deepOrange300 =
      ExcelColor._('FFFF8A65', 'deepOrange300', ColorType.material);
  static const deepOrange400 =
      ExcelColor._('FFFF7043', 'deepOrange400', ColorType.material);
  static const deepOrange500 =
      ExcelColor._('FFFF5722', 'deepOrange500', ColorType.material);
  static const deepOrange600 =
      ExcelColor._('FFF4511E', 'deepOrange600', ColorType.material);
  static const deepOrange700 =
      ExcelColor._('FFE64A19', 'deepOrange700', ColorType.material);
  static const deepOrange800 =
      ExcelColor._('FFD84315', 'deepOrange800', ColorType.material);
  static const deepOrange900 =
      ExcelColor._('FFBF360C', 'deepOrange900', ColorType.material);
  static const brown50 =
      ExcelColor._('FFEFEBE9', 'brown50', ColorType.material);
  static const brown100 =
      ExcelColor._('FFD7CCC8', 'brown100', ColorType.material);
  static const brown200 =
      ExcelColor._('FFBCAAA4', 'brown200', ColorType.material);
  static const brown300 =
      ExcelColor._('FFA1887F', 'brown300', ColorType.material);
  static const brown400 =
      ExcelColor._('FF8D6E63', 'brown400', ColorType.material);
  static const brown500 =
      ExcelColor._('FF795548', 'brown500', ColorType.material);
  static const brown600 =
      ExcelColor._('FF6D4C41', 'brown600', ColorType.material);
  static const brown700 =
      ExcelColor._('FF5D4037', 'brown700', ColorType.material);
  static const brown800 =
      ExcelColor._('FF4E342E', 'brown800', ColorType.material);
  static const brown900 =
      ExcelColor._('FF3E2723', 'brown900', ColorType.material);
  static const grey50 = ExcelColor._('FFFAFAFA', 'grey50', ColorType.material);
  static const grey100 =
      ExcelColor._('FFF5F5F5', 'grey100', ColorType.material);
  static const grey200 =
      ExcelColor._('FFEEEEEE', 'grey200', ColorType.material);
  static const grey300 =
      ExcelColor._('FFE0E0E0', 'grey300', ColorType.material);
  static const grey350 =
      ExcelColor._('FFD6D6D6', 'grey350', ColorType.material);
  static const grey400 =
      ExcelColor._('FFBDBDBD', 'grey400', ColorType.material);
  static const grey500 =
      ExcelColor._('FF9E9E9E', 'grey500', ColorType.material);
  static const grey600 =
      ExcelColor._('FF757575', 'grey600', ColorType.material);
  static const grey700 =
      ExcelColor._('FF616161', 'grey700', ColorType.material);
  static const grey800 =
      ExcelColor._('FF424242', 'grey800', ColorType.material);
  static const grey850 =
      ExcelColor._('FF303030', 'grey850', ColorType.material);
  static const grey900 =
      ExcelColor._('FF212121', 'grey900', ColorType.material);
  static const blueGrey50 =
      ExcelColor._('FFECEFF1', 'blueGrey50', ColorType.material);
  static const blueGrey100 =
      ExcelColor._('FFCFD8DC', 'blueGrey100', ColorType.material);
  static const blueGrey200 =
      ExcelColor._('FFB0BEC5', 'blueGrey200', ColorType.material);
  static const blueGrey300 =
      ExcelColor._('FF90A4AE', 'blueGrey300', ColorType.material);
  static const blueGrey400 =
      ExcelColor._('FF78909C', 'blueGrey400', ColorType.material);
  static const blueGrey500 =
      ExcelColor._('FF607D8B', 'blueGrey500', ColorType.material);
  static const blueGrey600 =
      ExcelColor._('FF546E7A', 'blueGrey600', ColorType.material);
  static const blueGrey700 =
      ExcelColor._('FF455A64', 'blueGrey700', ColorType.material);
  static const blueGrey800 =
      ExcelColor._('FF37474F', 'blueGrey800', ColorType.material);
  static const blueGrey900 =
      ExcelColor._('FF263238', 'blueGrey900', ColorType.material);

  static List<ExcelColor> get values => [
        black,
        black12,
        black26,
        black38,
        black45,
        black54,
        black87,
        white,
        white10,
        white12,
        white24,
        white30,
        white38,
        white54,
        white60,
        white70,
        redAccent,
        pinkAccent,
        purpleAccent,
        deepPurpleAccent,
        indigoAccent,
        blueAccent,
        lightBlueAccent,
        cyanAccent,
        tealAccent,
        greenAccent,
        lightGreenAccent,
        limeAccent,
        yellowAccent,
        amberAccent,
        orangeAccent,
        deepOrangeAccent,
        red,
        pink,
        purple,
        deepPurple,
        indigo,
        blue,
        lightBlue,
        cyan,
        teal,
        green,
        lightGreen,
        lime,
        yellow,
        amber,
        orange,
        deepOrange,
        brown,
        grey,
        blueGrey,
        redAccent100,
        redAccent200,
        redAccent400,
        redAccent700,
        pinkAccent100,
        pinkAccent200,
        pinkAccent400,
        pinkAccent700,
        purpleAccent100,
        purpleAccent200,
        purpleAccent400,
        purpleAccent700,
        deepPurpleAccent100,
        deepPurpleAccent200,
        deepPurpleAccent400,
        deepPurpleAccent700,
        indigoAccent100,
        indigoAccent200,
        indigoAccent400,
        indigoAccent700,
        blueAccent100,
        blueAccent200,
        blueAccent400,
        blueAccent700,
        lightBlueAccent100,
        lightBlueAccent200,
        lightBlueAccent400,
        lightBlueAccent700,
        cyanAccent100,
        cyanAccent200,
        cyanAccent400,
        cyanAccent700,
        tealAccent100,
        tealAccent200,
        tealAccent400,
        tealAccent700,
        greenAccent100,
        greenAccent200,
        greenAccent400,
        greenAccent700,
        lightGreenAccent100,
        lightGreenAccent200,
        lightGreenAccent400,
        lightGreenAccent700,
        limeAccent100,
        limeAccent200,
        limeAccent400,
        limeAccent700,
        yellowAccent100,
        yellowAccent200,
        yellowAccent400,
        yellowAccent700,
        amberAccent100,
        amberAccent200,
        amberAccent400,
        amberAccent700,
        orangeAccent100,
        orangeAccent200,
        orangeAccent400,
        orangeAccent700,
        deepOrangeAccent100,
        deepOrangeAccent200,
        deepOrangeAccent400,
        deepOrangeAccent700,
        red50,
        red100,
        red200,
        red300,
        red400,
        red500,
        red600,
        red700,
        red800,
        red900,
        pink50,
        pink100,
        pink200,
        pink300,
        pink400,
        pink500,
        pink600,
        pink700,
        pink800,
        pink900,
        purple50,
        purple100,
        purple200,
        purple300,
        purple400,
        purple500,
        purple600,
        purple700,
        purple800,
        purple900,
        deepPurple50,
        deepPurple100,
        deepPurple200,
        deepPurple300,
        deepPurple400,
        deepPurple500,
        deepPurple600,
        deepPurple700,
        deepPurple800,
        deepPurple900,
        indigo50,
        indigo100,
        indigo200,
        indigo300,
        indigo400,
        indigo500,
        indigo600,
        indigo700,
        indigo800,
        indigo900,
        blue50,
        blue100,
        blue200,
        blue300,
        blue400,
        blue500,
        blue600,
        blue700,
        blue800,
        blue900,
        lightBlue50,
        lightBlue100,
        lightBlue200,
        lightBlue300,
        lightBlue400,
        lightBlue500,
        lightBlue600,
        lightBlue700,
        lightBlue800,
        lightBlue900,
        cyan50,
        cyan100,
        cyan200,
        cyan300,
        cyan400,
        cyan500,
        cyan600,
        cyan700,
        cyan800,
        cyan900,
        teal50,
        teal100,
        teal200,
        teal300,
        teal400,
        teal500,
        teal600,
        teal700,
        teal800,
        teal900,
        green50,
        green100,
        green200,
        green300,
        green400,
        green500,
        green600,
        green700,
        green800,
        green900,
        lightGreen50,
        lightGreen100,
        lightGreen200,
        lightGreen300,
        lightGreen400,
        lightGreen500,
        lightGreen600,
        lightGreen700,
        lightGreen800,
        lightGreen900,
        lime50,
        lime100,
        lime200,
        lime300,
        lime400,
        lime500,
        lime600,
        lime700,
        lime800,
        lime900,
        yellow50,
        yellow100,
        yellow200,
        yellow300,
        yellow400,
        yellow500,
        yellow600,
        yellow700,
        yellow800,
        yellow900,
        amber50,
        amber100,
        amber200,
        amber300,
        amber400,
        amber500,
        amber600,
        amber700,
        amber800,
        amber900,
        orange50,
        orange100,
        orange200,
        orange300,
        orange400,
        orange500,
        orange600,
        orange700,
        orange800,
        orange900,
        deepOrange50,
        deepOrange100,
        deepOrange200,
        deepOrange300,
        deepOrange400,
        deepOrange500,
        deepOrange600,
        deepOrange700,
        deepOrange800,
        deepOrange900,
        brown50,
        brown100,
        brown200,
        brown300,
        brown400,
        brown500,
        brown600,
        brown700,
        brown800,
        brown900,
        grey50,
        grey100,
        grey200,
        grey300,
        grey350,
        grey400,
        grey500,
        grey600,
        grey700,
        grey800,
        grey850,
        grey900,
        blueGrey50,
        blueGrey100,
        blueGrey200,
        blueGrey300,
        blueGrey400,
        blueGrey500,
        blueGrey600,
        blueGrey700,
        blueGrey800,
        blueGrey900,
      ];

  @override
  List<Object?> get props => [
        _name,
        _color,
        _type,
        colorHex,
        colorInt,
      ];
}

enum ColorType {
  color,
  material,
  materialAccent,
  ;
}
