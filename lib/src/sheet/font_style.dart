part of excel;

/// Styling class for cells
class _FontStyle {
  String _fontColorHex, _fontFamily;
  bool _bold, _italic;
  Underline _underline;
  int _fontSize;

  _FontStyle(
      {String fontColorHex = "FF000000",
      int fontSize,
      String fontFamily,
      bool bold = false,
      Underline underline = Underline.None,
      bool italic = false}) {
    this._bold = bold ?? false;

    this.fontSize = fontSize;

    this._italic = italic ?? false;

    this.fontFamily = fontFamily;

    this._underline = underline ?? Underline.None;

    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }
  }

  /// Get Font Color
  ///
  ///
  String get fontColor {
    return this._fontColorHex;
  }

  /// Set Font Color
  ///
  ///
  set fontColor(String fontColorHex) {
    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }
  }

  /// `Get FontFamily`
  ///
  ///
  String get fontFamily {
    return this._fontFamily;
  }

  /// `Set FontFamily`
  ///
  ///
  set fontFamily(String family) {
    this._fontFamily = family;
  }

  /// Get Font Size
  ///
  ///
  int get fontSize {
    return this._fontSize;
  }

  /// Set Font Size
  ///
  ///
  set fontSize(int _font_Size) {
    this._fontSize = _font_Size;
  }

  /// Get `Underline`
  ///
  ///
  get underline {
    return this._underline;
  }

  /// set `Underline`
  ///
  ///
  set underline(Underline underline) {
    this._underline = underline ?? Underline.None;
  }

  /// Get `Bold`
  ///
  ///
  get isBold {
    return this._bold;
  }

  /// Set `Bold`
  ///
  ///
  set isBold(bool bold) {
    this._bold = bold ?? false;
  }

  /// Get `Italic`
  ///
  ///
  get isItalic {
    return this._italic;
  }

  /// Set `Italic`
  ///
  ///
  set isItalic(bool italic) {
    this._italic = italic ?? false;
  }

  @override
  bool operator ==(o) {
    return o.isBold == this.isBold && // bold
        o.isItalic == this.isItalic && // italic
        o.fontSize == this.fontSize && // Font Size
        o.underline == this.underline && // Underline
        o.fontFamily == this.fontFamily && // font Family
        o.runtimeType == this.runtimeType && // runtimeType
        o.fontColor == this.fontColor; // Font Color
  }

  /* @override
  String toString() {
    String f = "Font Color :" + this.fontColor,
        fs = "Font Size  :" + this.fontSize.toString(),
        bold = "Bold       :" + this.isBold.toString(),
        underline = "Underline  :" + this.underline.toString(),
        italic = "Italic     :" + this.isItalic.toString(),
        fontFamily = "Font Family:" + this.fontFamily.toString();

    return f + "\n" + fs + "\n" + bold + "\n" + italic + "\n" + fontFamily;
  } */
}
