part of excel;

/// Styling class for cells
class _FontStyle {
  String _fontColorHex, _fontFamily;
  bool _bold, _italic;
  Underline _underline;
  int _fontSize;

  _FontStyle(
      {String fontColorHex = "FF000000",
      int fontSize = 12,
      String fontFamily = "Arial",
      bool bold = false,
      Underline underline = Underline.None,
      bool italic = false}) {
    this._bold = bold ?? false;

    this.fontSize = fontSize ?? 12;

    this._italic = italic ?? false;

    this.fontFamily = fontFamily ?? FontFamily.Arial;

    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }
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
      o.fontColorHex == this._fontColorHex; // Font Color

  @override
  String toString() {
    String f = "Font Color :" + this._fontColorHex,
        fs = "Font Size  :" + this.fontSize.toString(),
        bold = "Bold       :" + this._bold.toString(),
        italic = "Italic     :" + this._italic.toString(),
        fontFamily = "Font Family:" + this.fontFamily.toString();

    return f + "\n" + fs + "\n" + bold + "\n" + italic + "\n" + fontFamily;
  }
}
