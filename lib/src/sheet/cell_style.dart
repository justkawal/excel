part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class CellStyle extends Equatable {
  String _fontColorHex = ExcelColor.black.colorHex;
  String _backgroundColorHex = ExcelColor.none.colorHex;
  String? _fontFamily;
  FontScheme _fontScheme;
  HorizontalAlign _horizontalAlign = HorizontalAlign.Left;
  VerticalAlign _verticalAlign = VerticalAlign.Bottom;
  TextWrapping? _textWrapping;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;
  int _rotation = 0;
  Border _leftBorder;
  Border _rightBorder;
  Border _topBorder;
  Border _bottomBorder;
  Border _diagonalBorder;
  bool _diagonalBorderUp = false;
  bool _diagonalBorderDown = false;
  NumFormat numberFormat;

  CellStyle({
    ExcelColor fontColorHex = ExcelColor.black,
    ExcelColor backgroundColorHex = ExcelColor.none,
    int? fontSize,
    String? fontFamily,
    FontScheme? fontScheme,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping? textWrapping,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
    int rotation = 0,
    Border? leftBorder,
    Border? rightBorder,
    Border? topBorder,
    Border? bottomBorder,
    Border? diagonalBorder,
    bool diagonalBorderUp = false,
    bool diagonalBorderDown = false,
    this.numberFormat = NumFormat.standard_0,
  })  : _textWrapping = textWrapping,
        _bold = bold,
        _fontSize = fontSize,
        _italic = italic,
        _fontFamily = fontFamily,
        _fontScheme = fontScheme ?? FontScheme.Unset,
        _rotation = rotation,
        _fontColorHex = _isColorAppropriate(fontColorHex.colorHex),
        _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex),
        _verticalAlign = verticalAlign,
        _horizontalAlign = horizontalAlign,
        _leftBorder = leftBorder ?? Border(),
        _rightBorder = rightBorder ?? Border(),
        _topBorder = topBorder ?? Border(),
        _bottomBorder = bottomBorder ?? Border(),
        _diagonalBorder = diagonalBorder ?? Border(),
        _diagonalBorderUp = diagonalBorderUp,
        _diagonalBorderDown = diagonalBorderDown;

  CellStyle copyWith({
    ExcelColor? fontColorHexVal,
    ExcelColor? backgroundColorHexVal,
    String? fontFamilyVal,
    FontScheme? fontSchemeVal,
    HorizontalAlign? horizontalAlignVal,
    VerticalAlign? verticalAlignVal,
    TextWrapping? textWrappingVal,
    bool? boldVal,
    bool? italicVal,
    Underline? underlineVal,
    int? fontSizeVal,
    int? rotationVal,
    Border? leftBorderVal,
    Border? rightBorderVal,
    Border? topBorderVal,
    Border? bottomBorderVal,
    Border? diagonalBorderVal,
    bool? diagonalBorderUpVal,
    bool? diagonalBorderDownVal,
    NumFormat? numberFormat,
  }) {
    return CellStyle(
      fontColorHex: fontColorHexVal ?? this._fontColorHex.excelColor,
      backgroundColorHex:
          backgroundColorHexVal ?? this._backgroundColorHex.excelColor,
      fontFamily: fontFamilyVal ?? this._fontFamily,
      fontScheme: fontSchemeVal ?? this._fontScheme,
      horizontalAlign: horizontalAlignVal ?? this._horizontalAlign,
      verticalAlign: verticalAlignVal ?? this._verticalAlign,
      textWrapping: textWrappingVal ?? this._textWrapping,
      bold: boldVal ?? this._bold,
      italic: italicVal ?? this._italic,
      underline: underlineVal ?? this._underline,
      fontSize: fontSizeVal ?? this._fontSize,
      rotation: rotationVal ?? this._rotation,
      leftBorder: leftBorderVal ?? this._leftBorder,
      rightBorder: rightBorderVal ?? this._rightBorder,
      topBorder: topBorderVal ?? this._topBorder,
      bottomBorder: bottomBorderVal ?? this._bottomBorder,
      diagonalBorder: diagonalBorderVal ?? this._diagonalBorder,
      diagonalBorderUp: diagonalBorderUpVal ?? this._diagonalBorderUp,
      diagonalBorderDown: diagonalBorderDownVal ?? this._diagonalBorderDown,
      numberFormat: numberFormat ?? this.numberFormat,
    );
  }

  ///Get Font Color
  ///
  ExcelColor get fontColor {
    return _fontColorHex.excelColor;
  }

  ///Set Font Color
  ///
  set fontColor(ExcelColor fontColorHex) {
    _fontColorHex = _isColorAppropriate(fontColorHex.colorHex);
  }

  ///Get Background Color
  ///
  ExcelColor get backgroundColor {
    return _backgroundColorHex.excelColor;
  }

  ///Set Background Color
  ///
  set backgroundColor(ExcelColor backgroundColorHex) {
    _backgroundColorHex = _isColorAppropriate(backgroundColorHex.colorHex);
  }

  ///Get Horizontal Alignment
  ///
  HorizontalAlign get horizontalAlignment {
    return _horizontalAlign;
  }

  ///Set Horizontal Alignment
  ///
  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    _horizontalAlign = horizontalAlign;
  }

  ///Get Vertical Alignment
  ///
  VerticalAlign get verticalAlignment {
    return _verticalAlign;
  }

  ///Set Vertical Alignment
  ///
  set verticalAlignment(VerticalAlign verticalAlign) {
    _verticalAlign = verticalAlign;
  }

  ///`Get Wrapping`
  ///
  TextWrapping? get wrap {
    return _textWrapping;
  }

  ///`Set Wrapping`
  ///
  set wrap(TextWrapping? textWrapping) {
    _textWrapping = textWrapping;
  }

  ///`Get FontFamily`
  ///
  String? get fontFamily {
    return _fontFamily;
  }

  ///`Set FontFamily`
  ///
  set fontFamily(String? family) {
    _fontFamily = family;
  }

  ///`Get FontScheme`
  ///
  FontScheme get fontScheme {
    return _fontScheme;
  }

  ///`Set FontScheme`
  ///
  set fontScheme(FontScheme scheme) {
    _fontScheme = scheme;
  }

  ///Get Font Size
  ///
  int? get fontSize {
    return _fontSize;
  }

  ///Set Font Size
  ///
  set fontSize(int? _fs) {
    _fontSize = _fs;
  }

  ///Get Rotation
  ///
  int get rotation {
    return _rotation;
  }

  ///Rotation varies from [90 to -90]
  ///
  set rotation(int _rotate) {
    if (_rotate > 90 || _rotate < -90) {
      _rotate = 0;
    }
    if (_rotate < 0) {
      /// The value is from 0 to -90 so now make it absolute and add it to 90
      ///
      /// -(_rotate) + 90
      _rotate = -(_rotate) + 90;
    }
    _rotation = _rotate;
  }

  ///Get `Underline`
  ///
  Underline get underline {
    return _underline;
  }

  ///set `Underline`
  ///
  set underline(Underline _) {
    _underline = _;
  }

  ///Get `Bold`
  ///
  bool get isBold {
    return _bold;
  }

  ///Set `Bold`
  set isBold(bool bold) {
    _bold = bold;
  }

  ///Get `Italic`
  ///
  bool get isItalic {
    return _italic;
  }

  ///Set `Italic`
  ///
  set isItalic(bool italic) {
    _italic = italic;
  }

  ///Get `LeftBorder`
  ///
  Border get leftBorder {
    return _leftBorder;
  }

  ///Set `LeftBorder`
  ///
  set leftBorder(Border? leftBorder) {
    _leftBorder = leftBorder ?? Border();
  }

  ///Get `RightBorder`
  ///
  Border get rightBorder {
    return _rightBorder;
  }

  ///Set `RightBorder`
  ///
  set rightBorder(Border? rightBorder) {
    _rightBorder = rightBorder ?? Border();
  }

  ///Get `TopBorder`
  ///
  Border get topBorder {
    return _topBorder;
  }

  ///Set `TopBorder`
  ///
  set topBorder(Border? topBorder) {
    _topBorder = topBorder ?? Border();
  }

  ///Get `BottomBorder`
  ///
  Border get bottomBorder {
    return _bottomBorder;
  }

  ///Set `BottomBorder`
  ///
  set bottomBorder(Border? bottomBorder) {
    _bottomBorder = bottomBorder ?? Border();
  }

  ///Get `DiagonalBorder`
  ///
  Border get diagonalBorder {
    return _diagonalBorder;
  }

  ///Set `DiagonalBorder`
  ///
  set diagonalBorder(Border? diagonalBorder) {
    _diagonalBorder = diagonalBorder ?? Border();
  }

  ///Get `DiagonalBorderUp`
  ///
  bool get diagonalBorderUp {
    return _diagonalBorderUp;
  }

  ///Set `DiagonalBorderUp`
  ///
  set diagonalBorderUp(bool diagonalBorderUp) {
    _diagonalBorderUp = diagonalBorderUp;
  }

  ///Get `DiagonalBorderDown`
  ///
  bool get diagonalBorderDown {
    return _diagonalBorderDown;
  }

  ///Set `DiagonalBorderDown`
  ///
  set diagonalBorderDown(bool diagonalBorderDown) {
    _diagonalBorderDown = diagonalBorderDown;
  }

  @override
  List<Object?> get props => [
        _bold,
        _rotation,
        _italic,
        _underline,
        _fontSize,
        _fontFamily,
        _fontScheme,
        _textWrapping,
        _verticalAlign,
        _horizontalAlign,
        _fontColorHex,
        _backgroundColorHex,
        _leftBorder,
        _rightBorder,
        _topBorder,
        _bottomBorder,
        _diagonalBorder,
        _diagonalBorderUp,
        _diagonalBorderDown,
        numberFormat,
      ];
}
