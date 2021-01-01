part of excel;

/// Styling class for cells
class CellStyle {
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
    this._textWrapping = textWrapping;

    this._bold = bold ?? false;

    this.fontSize = fontSize;

    this._italic = italic ?? false;

    this.fontFamily = fontFamily;

    this._rotation = rotation ?? 0;

    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }

    if (backgroundColorHex != null) {
      this._backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this._backgroundColorHex = "none";
    }

    this._verticalAlign = verticalAlign ?? VerticalAlign.Bottom;

    this._horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  ///
  ///Get Font Color
  ///
  String get fontColor {
    return this._fontColorHex;
  }

  ///
  ///Set Font Color
  ///
  set fontColor(String fontColorHex) {
    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }
  }

  ///
  ///Get Background Color
  ///
  String get backgroundColor {
    return this._backgroundColorHex;
  }

  ///
  ///Set Background Color
  ///
  set backgroundColor(String backgroundColorHex) {
    if (backgroundColorHex != null) {
      this._backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this._backgroundColorHex = "none";
    }
  }

  ///
  ///Get Horizontal Alignment
  ///
  HorizontalAlign get horizontalAlignment {
    return this._horizontalAlign;
  }

  ///
  ///Set Horizontal Alignment
  ///
  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    this._horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  ///
  ///Get Vertical Alignment
  ///
  VerticalAlign get verticalAlignment {
    return this._verticalAlign;
  }

  ///
  ///Set Vertical Alignment
  ///
  set verticalAlignment(VerticalAlign verticalAlign) {
    this._verticalAlign = verticalAlign ?? VerticalAlign.Bottom;
  }

  ///
  ///`Get Wrapping`
  ///
  TextWrapping get wrap {
    return this._textWrapping;
  }

  ///
  ///`Set Wrapping`
  ///
  set wrap(TextWrapping textWrapping) {
    this._textWrapping = textWrapping;
  }

  ///
  ///`Get FontFamily`
  ///
  String get fontFamily {
    return this._fontFamily;
  }

  ///
  ///`Set FontFamily`
  ///
  set fontFamily(String family) {
    this._fontFamily = family;
  }

  ///
  ///Get Font Size
  ///
  int get fontSize {
    return this._fontSize;
  }

  ///
  ///Set Font Size
  ///
  set fontSize(int _font_Size) {
    this._fontSize = _font_Size;
  }

  ///
  ///Get Rotation
  ///
  int get rotation {
    return this._rotation;
  }

  ///
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
    this._rotation = _rotate;
  }

  ///
  ///Get `Underline`
  ///
  Underline get underline {
    return this._underline;
  }

  ///
  ///set `Underline`
  ///
  set underline(Underline underline_) {
    this._underline = underline_ ?? Underline.None;
  }

  ///
  ///Get `Bold`
  ///
  bool get isBold {
    return this._bold;
  }

  ///
  ///Set `Bold`
  ///
  set isBold(bool bold) {
    this._bold = bold ?? false;
  }

  ///
  ///Get `Italic`
  ///
  bool get isItalic {
    return this._italic;
  }

  ///
  ///Set `Italic`
  ///
  set isItalic(bool italic) {
    this._italic = italic ?? false;
  }

  @override
  bool operator ==(o) {
    return o.isBold == this.isBold && // bold
        o.rotation == this.rotation && // rotation
        o.isItalic == this.isItalic && // italic
        o.fontSize == this.fontSize && // Font Size
        o.fontFamily == this.fontFamily &&
        o.runtimeType == this.runtimeType && // runtimeType
        o.wrap == this.wrap && // Font Wrapping
        o.fontColor == this.fontColor && // Font Color
        o.verticalAlignment == this.verticalAlignment && // Vertical Align
        o.horizontalAlignment == this.horizontalAlignment && // Horizontal Align
        o.backgroundColor == this.backgroundColor; // Background Color
  }

  @override
  String toString() {
    String b = "Background Color: " + this._backgroundColorHex;
    String f = "Font Color: " + this._fontColorHex;
    return b + "\n" + f;
  }
}
