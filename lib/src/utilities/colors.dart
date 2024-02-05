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
  const ExcelColor._([this._color, this._name]);

  final String? _color;
  final String? _name;

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

  static const black = ExcelColor._('FF000000', 'black');
  static const black12 = ExcelColor._('1F000000', 'black12');
  static const black26 = ExcelColor._('42000000', 'black26');
  static const black38 = ExcelColor._('61000000', 'black38');
  static const black45 = ExcelColor._('73000000', 'black45');
  static const black54 = ExcelColor._('8A000000', 'black54');
  static const black87 = ExcelColor._('DD000000', 'black87');
  static const white = ExcelColor._('FFFFFFFF', 'white');
  static const white10 = ExcelColor._('1AFFFFFF', 'white10');
  static const white12 = ExcelColor._('1FFFFFFF', 'white12');
  static const white24 = ExcelColor._('3DFFFFFF', 'white24');
  static const white30 = ExcelColor._('4DFFFFFF', 'white30');
  static const white38 = ExcelColor._('62FFFFFF', 'white38');
  static const white54 = ExcelColor._('8AFFFFFF', 'white54');
  static const white60 = ExcelColor._('99FFFFFF', 'white60');
  static const white70 = ExcelColor._('B3FFFFFF', 'white70');
  static const redAccent = ExcelColor._('FFFF5252', 'redAccent');
  static const pinkAccent = ExcelColor._('FFFF4081', 'pinkAccent');
  static const purpleAccent = ExcelColor._('FFE040FB', 'purpleAccent');
  static const deepPurpleAccent = ExcelColor._('FF7C4DFF', 'deepPurpleAccent');
  static const indigoAccent = ExcelColor._('FF536DFE', 'indigoAccent');
  static const blueAccent = ExcelColor._('FF448AFF', 'blueAccent');
  static const lightBlueAccent = ExcelColor._('FF40C4FF', 'lightBlueAccent');
  static const cyanAccent = ExcelColor._('FF18FFFF', 'cyanAccent');
  static const tealAccent = ExcelColor._('FF64FFDA', 'tealAccent');
  static const greenAccent = ExcelColor._('FF69F0AE', 'greenAccent');
  static const lightGreenAccent = ExcelColor._('FFB2FF59', 'lightGreenAccent');
  static const limeAccent = ExcelColor._('FFEEFF41', 'limeAccent');
  static const yellowAccent = ExcelColor._('FFFFFF00', 'yellowAccent');
  static const amberAccent = ExcelColor._('FFFFD740', 'amberAccent');
  static const orangeAccent = ExcelColor._('FFFFAB40', 'orangeAccent');
  static const deepOrangeAccent = ExcelColor._('FFFF6E40', 'deepOrangeAccent');
  static const red = ExcelColor._('FFF44336', 'red');
  static const pink = ExcelColor._('FFE91E63', 'pink');
  static const purple = ExcelColor._('FF9C27B0', 'purple');
  static const deepPurple = ExcelColor._('FF673AB7', 'deepPurple');
  static const indigo = ExcelColor._('FF3F51B5', 'indigo');
  static const blue = ExcelColor._('FF2196F3', 'blue');
  static const lightBlue = ExcelColor._('FF03A9F4', 'lightBlue');
  static const cyan = ExcelColor._('FF00BCD4', 'cyan');
  static const teal = ExcelColor._('FF009688', 'teal');
  static const green = ExcelColor._('FF4CAF50', 'green');
  static const lightGreen = ExcelColor._('FF8BC34A', 'lightGreen');
  static const lime = ExcelColor._('FFCDDC39', 'lime');
  static const yellow = ExcelColor._('FFFFEB3B', 'yellow');
  static const amber = ExcelColor._('FFFFC107', 'amber');
  static const orange = ExcelColor._('FFFF9800', 'orange');
  static const deepOrange = ExcelColor._('FFFF5722', 'deepOrange');
  static const brown = ExcelColor._('FF795548', 'brown');
  static const grey = ExcelColor._('FF9E9E9E', 'grey');
  static const blueGrey = ExcelColor._('FF607D8B', 'blueGrey');
  static const redAccent100 = ExcelColor._('FFFF8A80', 'redAccent100');
  static const redAccent200 = ExcelColor._('FFFF5252', 'redAccent200');
  static const redAccent400 = ExcelColor._('FFFF1744', 'redAccent400');
  static const redAccent700 = ExcelColor._('FFD50000', 'redAccent700');
  static const pinkAccent100 = ExcelColor._('FFFF80AB', 'pinkAccent100');
  static const pinkAccent200 = ExcelColor._('FFFF4081', 'pinkAccent200');
  static const pinkAccent400 = ExcelColor._('FFF50057', 'pinkAccent400');
  static const pinkAccent700 = ExcelColor._('FFC51162', 'pinkAccent700');
  static const purpleAccent100 = ExcelColor._('FFEA80FC', 'purpleAccent100');
  static const purpleAccent200 = ExcelColor._('FFE040FB', 'purpleAccent200');
  static const purpleAccent400 = ExcelColor._('FFD500F9', 'purpleAccent400');
  static const purpleAccent700 = ExcelColor._('FFAA00FF', 'purpleAccent700');
  static const deepPurpleAccent100 =
      ExcelColor._('FFB388FF', 'deepPurpleAccent100');
  static const deepPurpleAccent200 =
      ExcelColor._('FF7C4DFF', 'deepPurpleAccent200');
  static const deepPurpleAccent400 =
      ExcelColor._('FF651FFF', 'deepPurpleAccent400');
  static const deepPurpleAccent700 =
      ExcelColor._('FF6200EA', 'deepPurpleAccent700');
  static const indigoAccent100 = ExcelColor._('FF8C9EFF', 'indigoAccent100');
  static const indigoAccent200 = ExcelColor._('FF536DFE', 'indigoAccent200');
  static const indigoAccent400 = ExcelColor._('FF3D5AFE', 'indigoAccent400');
  static const indigoAccent700 = ExcelColor._('FF304FFE', 'indigoAccent700');
  static const blueAccent100 = ExcelColor._('FF82B1FF', 'blueAccent100');
  static const blueAccent200 = ExcelColor._('FF448AFF', 'blueAccent200');
  static const blueAccent400 = ExcelColor._('FF2979FF', 'blueAccent400');
  static const blueAccent700 = ExcelColor._('FF2962FF', 'blueAccent700');
  static const lightBlueAccent100 =
      ExcelColor._('FF80D8FF', 'lightBlueAccent100');
  static const lightBlueAccent200 =
      ExcelColor._('FF40C4FF', 'lightBlueAccent200');
  static const lightBlueAccent400 =
      ExcelColor._('FF00B0FF', 'lightBlueAccent400');
  static const lightBlueAccent700 =
      ExcelColor._('FF0091EA', 'lightBlueAccent700');
  static const cyanAccent100 = ExcelColor._('FF84FFFF', 'cyanAccent100');
  static const cyanAccent200 = ExcelColor._('FF18FFFF', 'cyanAccent200');
  static const cyanAccent400 = ExcelColor._('FF00E5FF', 'cyanAccent400');
  static const cyanAccent700 = ExcelColor._('FF00B8D4', 'cyanAccent700');
  static const tealAccent100 = ExcelColor._('FFA7FFEB', 'tealAccent100');
  static const tealAccent200 = ExcelColor._('FF64FFDA', 'tealAccent200');
  static const tealAccent400 = ExcelColor._('FF1DE9B6', 'tealAccent400');
  static const tealAccent700 = ExcelColor._('FF00BFA5', 'tealAccent700');
  static const greenAccent100 = ExcelColor._('FFB9F6CA', 'greenAccent100');
  static const greenAccent200 = ExcelColor._('FF69F0AE', 'greenAccent200');
  static const greenAccent400 = ExcelColor._('FF00E676', 'greenAccent400');
  static const greenAccent700 = ExcelColor._('FF00C853', 'greenAccent700');
  static const lightGreenAccent100 =
      ExcelColor._('FFCCFF90', 'lightGreenAccent100');
  static const lightGreenAccent200 =
      ExcelColor._('FFB2FF59', 'lightGreenAccent200');
  static const lightGreenAccent400 =
      ExcelColor._('FF76FF03', 'lightGreenAccent400');
  static const lightGreenAccent700 =
      ExcelColor._('FF64DD17', 'lightGreenAccent700');
  static const limeAccent100 = ExcelColor._('FFF4FF81', 'limeAccent100');
  static const limeAccent200 = ExcelColor._('FFEEFF41', 'limeAccent200');
  static const limeAccent400 = ExcelColor._('FFC6FF00', 'limeAccent400');
  static const limeAccent700 = ExcelColor._('FFAEEA00', 'limeAccent700');
  static const yellowAccent100 = ExcelColor._('FFFFFF8D', 'yellowAccent100');
  static const yellowAccent200 = ExcelColor._('FFFFFF00', 'yellowAccent200');
  static const yellowAccent400 = ExcelColor._('FFFFEA00', 'yellowAccent400');
  static const yellowAccent700 = ExcelColor._('FFFFD600', 'yellowAccent700');
  static const amberAccent100 = ExcelColor._('FFFFE57F', 'amberAccent100');
  static const amberAccent200 = ExcelColor._('FFFFD740', 'amberAccent200');
  static const amberAccent400 = ExcelColor._('FFFFC400', 'amberAccent400');
  static const amberAccent700 = ExcelColor._('FFFFAB00', 'amberAccent700');
  static const orangeAccent100 = ExcelColor._('FFFFD180', 'orangeAccent100');
  static const orangeAccent200 = ExcelColor._('FFFFAB40', 'orangeAccent200');
  static const orangeAccent400 = ExcelColor._('FFFF9100', 'orangeAccent400');
  static const orangeAccent700 = ExcelColor._('FFFF6D00', 'orangeAccent700');
  static const deepOrangeAccent100 =
      ExcelColor._('FFFF9E80', 'deepOrangeAccent100');
  static const deepOrangeAccent200 =
      ExcelColor._('FFFF6E40', 'deepOrangeAccent200');
  static const deepOrangeAccent400 =
      ExcelColor._('FFFF3D00', 'deepOrangeAccent400');
  static const deepOrangeAccent700 =
      ExcelColor._('FFDD2C00', 'deepOrangeAccent700');
  static const red50 = ExcelColor._('FFFFEBEE', 'red50');
  static const red100 = ExcelColor._('FFFFCDD2', 'red100');
  static const red200 = ExcelColor._('FFEF9A9A', 'red200');
  static const red300 = ExcelColor._('FFE57373', 'red300');
  static const red400 = ExcelColor._('FFEF5350', 'red400');
  static const red500 = ExcelColor._('FFF44336', 'red500');
  static const red600 = ExcelColor._('FFE53935', 'red600');
  static const red700 = ExcelColor._('FFD32F2F', 'red700');
  static const red800 = ExcelColor._('FFC62828', 'red800');
  static const red900 = ExcelColor._('FFB71C1C', 'red900');
  static const pink50 = ExcelColor._('FFFCE4EC', 'pink50');
  static const pink100 = ExcelColor._('FFF8BBD0', 'pink100');
  static const pink200 = ExcelColor._('FFF48FB1', 'pink200');
  static const pink300 = ExcelColor._('FFF06292', 'pink300');
  static const pink400 = ExcelColor._('FFEC407A', 'pink400');
  static const pink500 = ExcelColor._('FFE91E63', 'pink500');
  static const pink600 = ExcelColor._('FFD81B60', 'pink600');
  static const pink700 = ExcelColor._('FFC2185B', 'pink700');
  static const pink800 = ExcelColor._('FFAD1457', 'pink800');
  static const pink900 = ExcelColor._('FF880E4F', 'pink900');
  static const purple50 = ExcelColor._('FFF3E5F5', 'purple50');
  static const purple100 = ExcelColor._('FFE1BEE7', 'purple100');
  static const purple200 = ExcelColor._('FFCE93D8', 'purple200');
  static const purple300 = ExcelColor._('FFBA68C8', 'purple300');
  static const purple400 = ExcelColor._('FFAB47BC', 'purple400');
  static const purple500 = ExcelColor._('FF9C27B0', 'purple500');
  static const purple600 = ExcelColor._('FF8E24AA', 'purple600');
  static const purple700 = ExcelColor._('FF7B1FA2', 'purple700');
  static const purple800 = ExcelColor._('FF6A1B9A', 'purple800');
  static const purple900 = ExcelColor._('FF4A148C', 'purple900');
  static const deepPurple50 = ExcelColor._('FFEDE7F6', 'deepPurple50');
  static const deepPurple100 = ExcelColor._('FFD1C4E9', 'deepPurple100');
  static const deepPurple200 = ExcelColor._('FFB39DDB', 'deepPurple200');
  static const deepPurple300 = ExcelColor._('FF9575CD', 'deepPurple300');
  static const deepPurple400 = ExcelColor._('FF7E57C2', 'deepPurple400');
  static const deepPurple500 = ExcelColor._('FF673AB7', 'deepPurple500');
  static const deepPurple600 = ExcelColor._('FF5E35B1', 'deepPurple600');
  static const deepPurple700 = ExcelColor._('FF512DA8', 'deepPurple700');
  static const deepPurple800 = ExcelColor._('FF4527A0', 'deepPurple800');
  static const deepPurple900 = ExcelColor._('FF311B92', 'deepPurple900');
  static const indigo50 = ExcelColor._('FFE8EAF6', 'indigo50');
  static const indigo100 = ExcelColor._('FFC5CAE9', 'indigo100');
  static const indigo200 = ExcelColor._('FF9FA8DA', 'indigo200');
  static const indigo300 = ExcelColor._('FF7986CB', 'indigo300');
  static const indigo400 = ExcelColor._('FF5C6BC0', 'indigo400');
  static const indigo500 = ExcelColor._('FF3F51B5', 'indigo500');
  static const indigo600 = ExcelColor._('FF3949AB', 'indigo600');
  static const indigo700 = ExcelColor._('FF303F9F', 'indigo700');
  static const indigo800 = ExcelColor._('FF283593', 'indigo800');
  static const indigo900 = ExcelColor._('FF1A237E', 'indigo900');
  static const blue50 = ExcelColor._('FFE3F2FD', 'blue50');
  static const blue100 = ExcelColor._('FFBBDEFB', 'blue100');
  static const blue200 = ExcelColor._('FF90CAF9', 'blue200');
  static const blue300 = ExcelColor._('FF64B5F6', 'blue300');
  static const blue400 = ExcelColor._('FF42A5F5', 'blue400');
  static const blue500 = ExcelColor._('FF2196F3', 'blue500');
  static const blue600 = ExcelColor._('FF1E88E5', 'blue600');
  static const blue700 = ExcelColor._('FF1976D2', 'blue700');
  static const blue800 = ExcelColor._('FF1565C0', 'blue800');
  static const blue900 = ExcelColor._('FF0D47A1', 'blue900');
  static const lightBlue50 = ExcelColor._('FFE1F5FE', 'lightBlue50');
  static const lightBlue100 = ExcelColor._('FFB3E5FC', 'lightBlue100');
  static const lightBlue200 = ExcelColor._('FF81D4FA', 'lightBlue200');
  static const lightBlue300 = ExcelColor._('FF4FC3F7', 'lightBlue300');
  static const lightBlue400 = ExcelColor._('FF29B6F6', 'lightBlue400');
  static const lightBlue500 = ExcelColor._('FF03A9F4', 'lightBlue500');
  static const lightBlue600 = ExcelColor._('FF039BE5', 'lightBlue600');
  static const lightBlue700 = ExcelColor._('FF0288D1', 'lightBlue700');
  static const lightBlue800 = ExcelColor._('FF0277BD', 'lightBlue800');
  static const lightBlue900 = ExcelColor._('FF01579B', 'lightBlue900');
  static const cyan50 = ExcelColor._('FFE0F7FA', 'cyan50');
  static const cyan100 = ExcelColor._('FFB2EBF2', 'cyan100');
  static const cyan200 = ExcelColor._('FF80DEEA', 'cyan200');
  static const cyan300 = ExcelColor._('FF4DD0E1', 'cyan300');
  static const cyan400 = ExcelColor._('FF26C6DA', 'cyan400');
  static const cyan500 = ExcelColor._('FF00BCD4', 'cyan500');
  static const cyan600 = ExcelColor._('FF00ACC1', 'cyan600');
  static const cyan700 = ExcelColor._('FF0097A7', 'cyan700');
  static const cyan800 = ExcelColor._('FF00838F', 'cyan800');
  static const cyan900 = ExcelColor._('FF006064', 'cyan900');
  static const teal50 = ExcelColor._('FFE0F2F1', 'teal50');
  static const teal100 = ExcelColor._('FFB2DFDB', 'teal100');
  static const teal200 = ExcelColor._('FF80CBC4', 'teal200');
  static const teal300 = ExcelColor._('FF4DB6AC', 'teal300');
  static const teal400 = ExcelColor._('FF26A69A', 'teal400');
  static const teal500 = ExcelColor._('FF009688', 'teal500');
  static const teal600 = ExcelColor._('FF00897B', 'teal600');
  static const teal700 = ExcelColor._('FF00796B', 'teal700');
  static const teal800 = ExcelColor._('FF00695C', 'teal800');
  static const teal900 = ExcelColor._('FF004D40', 'teal900');
  static const green50 = ExcelColor._('FFE8F5E9', 'green50');
  static const green100 = ExcelColor._('FFC8E6C9', 'green100');
  static const green200 = ExcelColor._('FFA5D6A7', 'green200');
  static const green300 = ExcelColor._('FF81C784', 'green300');
  static const green400 = ExcelColor._('FF66BB6A', 'green400');
  static const green500 = ExcelColor._('FF4CAF50', 'green500');
  static const green600 = ExcelColor._('FF43A047', 'green600');
  static const green700 = ExcelColor._('FF388E3C', 'green700');
  static const green800 = ExcelColor._('FF2E7D32', 'green800');
  static const green900 = ExcelColor._('FF1B5E20', 'green900');
  static const lightGreen50 = ExcelColor._('FFF1F8E9', 'lightGreen50');
  static const lightGreen100 = ExcelColor._('FFDCEDC8', 'lightGreen100');
  static const lightGreen200 = ExcelColor._('FFC5E1A5', 'lightGreen200');
  static const lightGreen300 = ExcelColor._('FFAED581', 'lightGreen300');
  static const lightGreen400 = ExcelColor._('FF9CCC65', 'lightGreen400');
  static const lightGreen500 = ExcelColor._('FF8BC34A', 'lightGreen500');
  static const lightGreen600 = ExcelColor._('FF7CB342', 'lightGreen600');
  static const lightGreen700 = ExcelColor._('FF689F38', 'lightGreen700');
  static const lightGreen800 = ExcelColor._('FF558B2F', 'lightGreen800');
  static const lightGreen900 = ExcelColor._('FF33691E', 'lightGreen900');
  static const lime50 = ExcelColor._('FFF9FBE7', 'lime50');
  static const lime100 = ExcelColor._('FFF0F4C3', 'lime100');
  static const lime200 = ExcelColor._('FFE6EE9C', 'lime200');
  static const lime300 = ExcelColor._('FFDCE775', 'lime300');
  static const lime400 = ExcelColor._('FFD4E157', 'lime400');
  static const lime500 = ExcelColor._('FFCDDC39', 'lime500');
  static const lime600 = ExcelColor._('FFC0CA33', 'lime600');
  static const lime700 = ExcelColor._('FFAFB42B', 'lime700');
  static const lime800 = ExcelColor._('FF9E9D24', 'lime800');
  static const lime900 = ExcelColor._('FF827717', 'lime900');
  static const yellow50 = ExcelColor._('FFFFFDE7', 'yellow50');
  static const yellow100 = ExcelColor._('FFFFF9C4', 'yellow100');
  static const yellow200 = ExcelColor._('FFFFF59D', 'yellow200');
  static const yellow300 = ExcelColor._('FFFFF176', 'yellow300');
  static const yellow400 = ExcelColor._('FFFFEE58', 'yellow400');
  static const yellow500 = ExcelColor._('FFFFEB3B', 'yellow500');
  static const yellow600 = ExcelColor._('FFFDD835', 'yellow600');
  static const yellow700 = ExcelColor._('FFFBC02D', 'yellow700');
  static const yellow800 = ExcelColor._('FFF9A825', 'yellow800');
  static const yellow900 = ExcelColor._('FFF57F17', 'yellow900');
  static const amber50 = ExcelColor._('FFFFF8E1', 'amber50');
  static const amber100 = ExcelColor._('FFFFECB3', 'amber100');
  static const amber200 = ExcelColor._('FFFFE082', 'amber200');
  static const amber300 = ExcelColor._('FFFFD54F', 'amber300');
  static const amber400 = ExcelColor._('FFFFCA28', 'amber400');
  static const amber500 = ExcelColor._('FFFFC107', 'amber500');
  static const amber600 = ExcelColor._('FFFFB300', 'amber600');
  static const amber700 = ExcelColor._('FFFFA000', 'amber700');
  static const amber800 = ExcelColor._('FFFF8F00', 'amber800');
  static const amber900 = ExcelColor._('FFFF6F00', 'amber900');
  static const orange50 = ExcelColor._('FFFFF3E0', 'orange50');
  static const orange100 = ExcelColor._('FFFFE0B2', 'orange100');
  static const orange200 = ExcelColor._('FFFFCC80', 'orange200');
  static const orange300 = ExcelColor._('FFFFB74D', 'orange300');
  static const orange400 = ExcelColor._('FFFFA726', 'orange400');
  static const orange500 = ExcelColor._('FFFF9800', 'orange500');
  static const orange600 = ExcelColor._('FFFB8C00', 'orange600');
  static const orange700 = ExcelColor._('FFF57C00', 'orange700');
  static const orange800 = ExcelColor._('FFEF6C00', 'orange800');
  static const orange900 = ExcelColor._('FFE65100', 'orange900');
  static const deepOrange50 = ExcelColor._('FFFBE9E7', 'deepOrange50');
  static const deepOrange100 = ExcelColor._('FFFFCCBC', 'deepOrange100');
  static const deepOrange200 = ExcelColor._('FFFFAB91', 'deepOrange200');
  static const deepOrange300 = ExcelColor._('FFFF8A65', 'deepOrange300');
  static const deepOrange400 = ExcelColor._('FFFF7043', 'deepOrange400');
  static const deepOrange500 = ExcelColor._('FFFF5722', 'deepOrange500');
  static const deepOrange600 = ExcelColor._('FFF4511E', 'deepOrange600');
  static const deepOrange700 = ExcelColor._('FFE64A19', 'deepOrange700');
  static const deepOrange800 = ExcelColor._('FFD84315', 'deepOrange800');
  static const deepOrange900 = ExcelColor._('FFBF360C', 'deepOrange900');
  static const brown50 = ExcelColor._('FFEFEBE9', 'brown50');
  static const brown100 = ExcelColor._('FFD7CCC8', 'brown100');
  static const brown200 = ExcelColor._('FFBCAAA4', 'brown200');
  static const brown300 = ExcelColor._('FFA1887F', 'brown300');
  static const brown400 = ExcelColor._('FF8D6E63', 'brown400');
  static const brown500 = ExcelColor._('FF795548', 'brown500');
  static const brown600 = ExcelColor._('FF6D4C41', 'brown600');
  static const brown700 = ExcelColor._('FF5D4037', 'brown700');
  static const brown800 = ExcelColor._('FF4E342E', 'brown800');
  static const brown900 = ExcelColor._('FF3E2723', 'brown900');
  static const grey50 = ExcelColor._('FFFAFAFA', 'grey50');
  static const grey100 = ExcelColor._('FFF5F5F5', 'grey100');
  static const grey200 = ExcelColor._('FFEEEEEE', 'grey200');
  static const grey300 = ExcelColor._('FFE0E0E0', 'grey300');
  static const grey350 = ExcelColor._('FFD6D6D6', 'grey350');
  static const grey400 = ExcelColor._('FFBDBDBD', 'grey400');
  static const grey500 = ExcelColor._('FF9E9E9E', 'grey500');
  static const grey600 = ExcelColor._('FF757575', 'grey600');
  static const grey700 = ExcelColor._('FF616161', 'grey700');
  static const grey800 = ExcelColor._('FF424242', 'grey800');
  static const grey850 = ExcelColor._('FF303030', 'grey850');
  static const grey900 = ExcelColor._('FF212121', 'grey900');
  static const blueGrey50 = ExcelColor._('FFECEFF1', 'blueGrey50');
  static const blueGrey100 = ExcelColor._('FFCFD8DC', 'blueGrey100');
  static const blueGrey200 = ExcelColor._('FFB0BEC5', 'blueGrey200');
  static const blueGrey300 = ExcelColor._('FF90A4AE', 'blueGrey300');
  static const blueGrey400 = ExcelColor._('FF78909C', 'blueGrey400');
  static const blueGrey500 = ExcelColor._('FF607D8B', 'blueGrey500');
  static const blueGrey600 = ExcelColor._('FF546E7A', 'blueGrey600');
  static const blueGrey700 = ExcelColor._('FF455A64', 'blueGrey700');
  static const blueGrey800 = ExcelColor._('FF37474F', 'blueGrey800');
  static const blueGrey900 = ExcelColor._('FF263238', 'blueGrey900');

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
        colorHex,
        colorInt,
      ];
}
