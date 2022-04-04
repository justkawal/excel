part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class CellStyle extends Equatable {
  String _fontColorHex = 'FF000000', _backgroundColorHex = 'none';
  String? _fontFamily;
  HorizontalAlign _horizontalAlign = HorizontalAlign.Left;
  VerticalAlign _verticalAlign = VerticalAlign.Bottom;
  TextWrapping? _textWrapping;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;
  int _rotation = 0;

  CellStyle({
    String fontColorHex = 'FF000000',
    String backgroundColorHex = 'none',
    int? fontSize,
    String? fontFamily,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping? textWrapping,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
    int rotation = 0,
  }) {
    _textWrapping = textWrapping;

    _bold = bold;

    _fontSize = fontSize;

    _italic = italic;

    _fontFamily = fontFamily;

    _rotation = rotation;

    _fontColorHex = _isColorAppropriate(fontColorHex);

    _backgroundColorHex = _isColorAppropriate(backgroundColorHex);

    _verticalAlign = verticalAlign;

    _horizontalAlign = horizontalAlign;
  }

  CellStyle copyWith({
    String? fontColorHexVal,
    String? backgroundColorHexVal,
    String? fontFamilyVal,
    HorizontalAlign? horizontalAlignVal,
    VerticalAlign? verticalAlignVal,
    TextWrapping? textWrappingVal,
    bool? boldVal,
    bool? italicVal,
    Underline? underlineVal,
    int? fontSizeVal,
    int? rotationVal,
  }) {
    return CellStyle(
      fontColorHex: fontColorHexVal ?? this._fontColorHex,
      backgroundColorHex: backgroundColorHexVal ?? this._backgroundColorHex,
      fontFamily: fontFamilyVal ?? this._fontFamily,
      horizontalAlign: horizontalAlignVal ?? this._horizontalAlign,
      verticalAlign: verticalAlignVal ?? this._verticalAlign,
      textWrapping: textWrappingVal ?? this._textWrapping,
      bold: boldVal ?? this._bold,
      italic: italicVal ?? this._italic,
      underline: underlineVal ?? this._underline,
      fontSize: fontSizeVal ?? this._fontSize,
      rotation: rotationVal ?? this._rotation,
    );
  }

  ///Get Font Color
  ///
  String get fontColor {
    return _fontColorHex;
  }

  ///Set Font Color
  ///
  set fontColor(String fontColorHex) {
    _fontColorHex = _isColorAppropriate(fontColorHex);
  }

  ///Get Background Color
  ///
  String get backgroundColor {
    return _backgroundColorHex;
  }

  ///Set Background Color
  ///
  set backgroundColor(String backgroundColorHex) {
    _backgroundColorHex = _isColorAppropriate(backgroundColorHex);
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

  @override
  List<Object?> get props => [
        _bold,
        _rotation,
        _italic,
        _underline,
        _fontSize,
        _fontFamily,
        _textWrapping,
        _verticalAlign,
        _horizontalAlign,
        _fontColorHex,
        _backgroundColorHex,
      ];
}
