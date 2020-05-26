part of excel;

/// Styling class for cells
class CellStyle {
  String fontColorHex;
  String backgroundColorHex;
  HorizontalAlign horizontalAlign;
  VerticalAlign verticalAlign;
  TextWrapping textWrapping;

  CellStyle({
    this.fontColorHex = "FF000000",
    this.backgroundColorHex = "none",
    this.horizontalAlign,
    this.verticalAlign,
    this.textWrapping,
  }) {
    this.textWrapping = textWrapping;

    this.fontColorHex == fontColorHex ?? "FF000000";

    this.backgroundColorHex = backgroundColorHex ?? "none";

    this.verticalAlign = verticalAlign ?? VerticalAlign.Bottom;

    this.horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  String get getFontColorHex => this.fontColorHex;

  set setFontColorHex(String fontColorHex) {
    this.fontColorHex == fontColorHex ?? "FF000000";
  }

  String get getBackgroundColorHex => this.backgroundColorHex;

  set setBackgroundColorHex(String backgroundColorHex) {
    this.backgroundColorHex = backgroundColorHex ?? "none";
  }

  HorizontalAlign get getHorizontalAlignment => this.horizontalAlign;

  set setHorizontalAlignment(HorizontalAlign horizontalAlign) {
    this.horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  VerticalAlign get getVericalAlignment => this.verticalAlign;

  set setVericalAlignment(VerticalAlign verticalAlign) {
    this.verticalAlign = verticalAlign ?? VerticalAlign.Bottom;
  }

  TextWrapping get getTextWrapping => this.textWrapping;

  set setTextWrapping(TextWrapping textWrapping) =>
      this.textWrapping = textWrapping;

  @override
  bool operator ==(o) =>
      o.runtimeType == this.runtimeType &&
      o.textWrapping == this.textWrapping &&
      o.fontColorHex == this.fontColorHex &&
      o.verticalAlign == this.verticalAlign &&
      o.horizontalAlign == this.horizontalAlign &&
      o.backgroundColorHex == this.backgroundColorHex;
}
