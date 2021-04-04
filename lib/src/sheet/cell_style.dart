part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class CellStyle extends Equatable {
  String _fontColorHex, _backgroundColorHex, _fontFamily;
  HorizontalAlign _horizontalAlign;
  VerticalAlign _verticalAlign;
  TextWrapping _textWrapping;
  bool _bold, _italic;
  Underline _underline;
  int _fontSize, _rotation;

  CellStyle({
    String fontColorHex = "FF000000",
    String backgroundColorHex = "none",
    int fontSize,
    String fontFamily,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping textWrapping,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
    int rotation = 0,
  }) {
    _textWrapping = textWrapping;

    _bold = bold ?? false;

    fontSize = fontSize;

    _italic = italic ?? false;

    fontFamily = fontFamily;

    _rotation = rotation ?? 0;

    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      _fontColorHex = "FF000000";
    }

    if (backgroundColorHex != null) {
      _backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      _backgroundColorHex = "none";
    }

    _verticalAlign = verticalAlign ?? VerticalAlign.Bottom;

    _horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  ///Get Font Color
  String get fontColor {
    return _fontColorHex;
  }

  ///Set Font Color
  set fontColor(String fontColorHex) {
    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      _fontColorHex = "FF000000";
    }
  }

  ///Get Background Color
  String get backgroundColor {
    return _backgroundColorHex;
  }

  ///Set Background Color
  set backgroundColor(String backgroundColorHex) {
    if (backgroundColorHex != null) {
      _backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      _backgroundColorHex = "none";
    }
  }

  ///Get Horizontal Alignment
  HorizontalAlign get horizontalAlignment {
    return _horizontalAlign;
  }

  ///Set Horizontal Alignment
  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    _horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  ///Get Vertical Alignment
  VerticalAlign get verticalAlignment {
    return _verticalAlign;
  }

  ///Set Vertical Alignment
  set verticalAlignment(VerticalAlign verticalAlign) {
    _verticalAlign = verticalAlign ?? VerticalAlign.Bottom;
  }

  ///`Get Wrapping`
  TextWrapping get wrap {
    return _textWrapping;
  }

  ///`Set Wrapping`
  set wrap(TextWrapping textWrapping) {
    _textWrapping = textWrapping;
  }

  ///`Get FontFamily`
  String get fontFamily {
    return _fontFamily;
  }

  ///`Set FontFamily`
  set fontFamily(String family) {
    _fontFamily = family;
  }

  ///Get Font Size
  int get fontSize {
    return _fontSize;
  }

  ///Set Font Size
  set fontSize(int _font_Size) {
    _fontSize = _font_Size;
  }

  ///Get Rotation
  int get rotation {
    return _rotation;
  }

  ///Rotation varies from [90 to -90]
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
  Underline get underline {
    return _underline;
  }

  ///Set `Underline`
  set underline(Underline underline_) {
    _underline = underline_ ?? Underline.None;
  }

  ///Get `Bold`
  bool get isBold {
    return _bold;
  }

  ///Set `Bold`
  set isBold(bool bold) {
    _bold = bold ?? false;
  }

  ///Get `Italic`
  bool get isItalic {
    return _italic;
  }

  ///Set `Italic`
  set isItalic(bool italic) {
    _italic = italic ?? false;
  }

  @override
  List<Object> get props => [
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
