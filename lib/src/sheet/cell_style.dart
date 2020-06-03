part of excel;

/// Styling class for cells
class CellStyle {
  String _fontColorHex, _backgroundColorHex, _fontFamily;
  HorizontalAlign _horizontalAlign;
  VerticalAlign _verticalAlign;
  TextWrapping _textWrapping;
  bool _bold, _italic;
  Underline _underline;
  int _fontSize;

  CellStyle(
      {String fontColorHex = "FF000000",
      String backgroundColorHex = "none",
      int fontSize = 12,
      String fontFamily = "Arial",
      HorizontalAlign horizontalAlign = HorizontalAlign.Left,
      VerticalAlign verticalAlign = VerticalAlign.Bottom,
      TextWrapping textWrapping,
      bool bold = false,
      Underline underline = Underline.None,
      bool italic = false}) {
    this._textWrapping = textWrapping;

    this._bold = bold ?? false;

    this.fontSize = fontSize ?? 12;

    this._italic = italic ?? false;

    this.fontFamily = fontFamily ?? FontFamily.Arial;

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

  /**
   * 
   * Get Font Color
   * 
   */
  String get fontColor {
    return this._fontColorHex;
  }

  /**
   * 
   * Set Font Color
   * 
   */
  set fontColor(String fontColorHex) {
    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }
  }

  /**
   * 
   * Get Background Color
   * 
   */
  String get backgroundColor {
    return this._backgroundColorHex;
  }

  /**
   * 
   * Set Background Color
   * 
   */
  set backgroundColor(String backgroundColorHex) {
    if (backgroundColorHex != null) {
      this._backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this._backgroundColorHex = "none";
    }
  }

  /**
   * 
   * Get Horizontal Alignment
   * 
   */
  HorizontalAlign get horizontalAlignment {
    return this._horizontalAlign;
  }

  /**
   * 
   * Set Horizontal Alignment
   * 
   */
  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    this._horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  /**
   * 
   * Get Vertical Alignment
   * 
   */
  VerticalAlign get verticalAlignment {
    return this._verticalAlign;
  }

  /**
   * 
   * Set Vertical Alignment
   * 
   */
  set verticalAlignment(VerticalAlign verticalAlign) {
    this._verticalAlign = verticalAlign ?? VerticalAlign.Bottom;
  }

  /**
   * 
   * `Get Wrapping`
   * 
   */
  TextWrapping get wrap => this._textWrapping;

  /**
   * 
   * `Set Wrapping`
   * 
   */
  set wrap(TextWrapping textWrapping) {
    this._textWrapping = textWrapping;
  }

  /**
   * 
   * `Get FontFamily`
   * 
   */
  String get fontFamily {
    return this._fontFamily;
  }

  /**
   * 
   * `Set FontFamily`
   * 
   */
  set fontFamily(String family) {
    this._fontFamily = family ?? "Arial";
  }

  /**
   * 
   * Get Font Size
   * 
   */
  int get fontSize {
    return this._fontSize;
  }

  /**
   * 
   * Set Font Size
   * 
   */
  set fontSize(int _font_Size) {
    this._fontSize = _font_Size ?? 12;
  }

  /**
   * 
   * Get `Underline`
   * 
   */
  get underline {
    return this._underline;
  }

  /**
   * 
   * set `Underline`
   * 
   */
  set underline(Underline underline) {
    this._underline = underline ?? Underline.None;
  }

  /**
   * 
   * Get `Bold`
   * 
   */
  get isBold {
    this._bold;
  }

  /**
   * 
   * Set `Bold`
   * 
   */
  set isBold(bool bold) {
    this._bold = bold ?? false;
  }

  /**
   * 
   * Get `Italic`
   * 
   */
  get isItalic {
    this._italic;
  }

  /**
   * 
   * Set `Italic`
   * 
   */
  set isItalic(bool italic) {
    this._italic = italic ?? false;
  }

  @override
  bool operator ==(o) =>
      o.bold == this._bold && // bold
      o.italic == this._italic && // italic
      o.fontSize == this._fontSize && // Font Size
      o.fontFamily == this._fontFamily &&
      o.runtimeType == this.runtimeType && // runtimeType
      o.textWrapping == this._textWrapping && // Font Wrapping
      o.fontColorHex == this._fontColorHex && // Font Color
      o.verticalAlign == this._verticalAlign && // Vertical Align
      o.horizontalAlign == this._horizontalAlign && // Horizontal Align
      o.backgroundColorHex == this._backgroundColorHex; // Background Color

  @override
  String toString() {
    String b = "Background Color: " + this._backgroundColorHex;
    String f = "Font Color: " + this._fontColorHex;

    return b + "\n" + f;
  }
}
