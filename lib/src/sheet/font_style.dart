part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class _FontStyle extends Equatable {
  String _fontColorHex, _fontFamily;
  bool _bold, _italic;
  Underline _underline;
  int _fontSize;

  _FontStyle({
    String fontColorHex = 'FF000000',
    int fontSize,
    String fontFamily,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
  }) {
    _bold = bold ?? false;

    fontSize = fontSize;

    _italic = italic ?? false;

    fontFamily = fontFamily;

    _underline = underline ?? Underline.None;

    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      _fontColorHex = 'FF000000';
    }
  }

  /// Get Font Color
  String get fontColor {
    return _fontColorHex;
  }

  /// Set Font Color
  set fontColor(String fontColorHex) {
    if (fontColorHex != null) {
      _fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      _fontColorHex = 'FF000000';
    }
  }

  /// `Get FontFamily`
  String get fontFamily {
    return _fontFamily;
  }

  /// `Set FontFamily`
  set fontFamily(String family) {
    _fontFamily = family;
  }

  /// Get Font Size
  int get fontSize {
    return _fontSize;
  }

  /// Set Font Size
  set fontSize(int _font_Size) {
    _fontSize = _font_Size;
  }

  /// Get `Underline`
  Underline get underline {
    return _underline;
  }

  /// Set `Underline`
  set underline(Underline underline) {
    _underline = underline ?? Underline.None;
  }

  /// Get `Bold`
  bool get isBold {
    return _bold;
  }

  /// Set `Bold`
  set isBold(bool bold) {
    _bold = bold ?? false;
  }

  /// Get `Italic`
  bool get isItalic {
    return _italic;
  }

  /// Set `Italic`
  set isItalic(bool italic) {
    _italic = italic ?? false;
  }

  @override
  List<Object> get props => [
        _bold,
        _italic,
        _fontSize,
        _underline,
        _fontFamily,
        _fontColorHex,
      ];
}
