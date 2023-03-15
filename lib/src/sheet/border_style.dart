part of excel;

class Border extends Equatable {
  late final BorderStyle? borderStyle;
  late final String? borderColorHex;

  Border({this.borderStyle, String? borderColorHex}) {
    this.borderColorHex =
        borderColorHex != null ? _isColorAppropriate(borderColorHex) : null;
  }

  @override
  String toString() {
    return 'Border(borderStyle: $borderStyle, borderColorHex: $borderColorHex)';
  }

  @override
  List<Object?> get props => [
        borderStyle,
        borderColorHex,
      ];
}

class BorderSet extends Equatable {
  late final Border? leftBorder;
  late final Border? rightBorder;
  late final Border? topBorder;
  late final Border? bottomBorder;
  late final Border? diagonalBorder;
  late final bool diagonalBorderUp;
  late final bool diagonalBorderDown;

  BorderSet({
    this.leftBorder,
    this.rightBorder,
    this.topBorder,
    this.bottomBorder,
    this.diagonalBorder,
    this.diagonalBorderUp = false,
    this.diagonalBorderDown = false,
  });

  BorderSet copyWith({
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    bool? diagonalBorderUp,
    bool? diagonalBorderDown,
  }) {
    return BorderSet(
      leftBorder: leftBorder ?? this.leftBorder,
      rightBorder: rightBorder ?? this.rightBorder,
      topBorder: topBorder ?? this.topBorder,
      bottomBorder: bottomBorder ?? this.bottomBorder,
      diagonalBorder: diagonalBorder ?? this.diagonalBorder,
      diagonalBorderUp: diagonalBorderUp ?? this.diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDown ?? this.diagonalBorderDown,
    );
  }

  @override
  List<Object?> get props => [
        leftBorder,
        rightBorder,
        topBorder,
        bottomBorder,
        diagonalBorder,
        diagonalBorderUp,
        diagonalBorderDown,
      ];
}

enum BorderStyle {
  None,
  DashDot,
  DashDotDot,
  Dashed,
  Dotted,
  Double,
  Hair,
  Medium,
  MediumDashDot,
  MediumDashDotDot,
  MediumDashed,
  SlantDashDot,
  Thick,
  Thin,
}

Map<BorderStyle, String> _borderStyleStringMap = {};

Map<BorderStyle, String> _getBorderStyleStringMap() {
  if (_borderStyleStringMap.isEmpty) {
    for (BorderStyle borderStyle in BorderStyle.values) {
      var borderStyleStr =
          borderStyle.toString().replaceAll('BorderStyle.', '');
      borderStyleStr =
          borderStyleStr[0].toLowerCase() + borderStyleStr.substring(1);
      _borderStyleStringMap[borderStyle] = borderStyleStr;
    }
  }

  return _borderStyleStringMap;
}

BorderStyle? getBorderStyleByName(String name) {
  var borderStyleStringMap = _getBorderStyleStringMap();
  return borderStyleStringMap.keys
      .firstWhereOrNull((k) => borderStyleStringMap[k] == name);
}

String getBorderStyleName(BorderStyle borderStyle) =>
    _getBorderStyleStringMap()[borderStyle]!;
