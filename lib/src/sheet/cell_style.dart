part of excel;

/// Styling class for cells
class CellStyle {
  String fontColorHex;
  String backgroundColorHex;
  HorizontalAlign horizontalAlign;
  VerticalAlign verticalAlign;
  TextWrapping textWrapping;

  CellStyle({
    this.fontColorHex,
    this.backgroundColorHex,
    this.horizontalAlign,
    this.verticalAlign,
    this.textWrapping,
  }) {
    this.textWrapping = textWrapping;

    if (fontColorHex != null) {
      this.fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this.fontColorHex = "FF000000";
    }

    if (backgroundColorHex != null) {
      this.backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this.backgroundColorHex = "none";
    }

    this.verticalAlign = verticalAlign ?? VerticalAlign.Bottom;

    this.horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  String get fontColor => this.fontColorHex;

  set fontColor(String fontColorHex) {
    if (fontColorHex != null) {
      this.fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this.fontColorHex = "FF000000";
    }
  }

  String get backgroundColor => this.backgroundColorHex;

  set backgroundColor(String backgroundColorHex) {
    if (backgroundColorHex != null) {
      this.backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this.backgroundColorHex = "none";
    }
  }

  HorizontalAlign get horizontalAlignment => this.horizontalAlign;

  set horizontalAlignment(HorizontalAlign horizontalAlign) {
    this.horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  VerticalAlign get verticalAlignment => this.verticalAlign;

  set verticalAlignment(VerticalAlign verticalAlign) {
    this.verticalAlign = verticalAlign ?? VerticalAlign.Bottom;
  }

  TextWrapping get wrap => this.textWrapping;

  set wrap(TextWrapping textWrapping) => this.textWrapping = textWrapping;

  @override
  bool operator ==(o) =>
      o.runtimeType == this.runtimeType &&
      o.textWrapping == this.textWrapping &&
      o.fontColorHex == this.fontColorHex &&
      o.verticalAlign == this.verticalAlign &&
      o.horizontalAlign == this.horizontalAlign &&
      o.backgroundColorHex == this.backgroundColorHex;

  @override
  String toString() {
    String b = "Background Color: " + this.backgroundColorHex;
    String f = "Font Color: " + this.fontColorHex;

    return b + "\n" + f;
  }
}
