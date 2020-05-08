part of excel;

/// Table of a spreadsheet file
class DataTable {
  final String name;
  DataTable(this.name);

  int _maxRows = 0, _maxCols = 0;

  List<List> _rows = List<List>();

  /// List of table's rows
  List<List> get rows => _rows;

  /// Get max rows
  int get maxRows => _maxRows;

  /// Get max cols
  int get maxCols => _maxCols;
}
