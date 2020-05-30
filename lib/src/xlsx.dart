part of excel;

/// Read and parse XSLX file
class XlsxDecoder {
  String get extension {
    return ".xlsx";
  }

  List<String> _rId;

  XlsxDecoder(Archive archive, {bool update = false}) {}
}
