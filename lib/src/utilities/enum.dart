part of excel;

///enum for `wrapping` up the text
///
enum TextWrapping {
  WrapText,
  Clip,
}

///
///enum for setting `vertical alignment`
///
enum VerticalAlign {
  Top,
  Center,
  Bottom,
}

///
///enum for setting `horizontal alignment`
///
enum HorizontalAlign {
  Left,
  Center,
  Right,
}

///
///`Cell Type`
///
enum CellType {
  String,
  int,
  Formula,
  double,
  bool,
}

///
///`Underline`
///
enum Underline {
  None,
  Single,
  Double,
}
