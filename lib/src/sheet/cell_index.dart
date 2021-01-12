part of excel;

class CellIndex {
  CellIndex._(String cell, int col, int row) {
    this._columnIndex = col;
    this._rowIndex = row;
    this._cellId = cell;
  }

  ///
  ///````
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///````
  static CellIndex indexByColumnRow({int columnIndex, int rowIndex}) {
    return CellIndex._(
        _stringIndex(columnIndex, rowIndex), columnIndex, rowIndex);
  }

  ///
  ///````
  /// CellIndex.indexByColumnRow('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByColumnRow('A2'); // columnIndex: 0, rowIndex: 1
  ///````
  static CellIndex indexByString(String cellIndex) {
    List<int> li = _cellCoordsFromCellId(cellIndex);

    return CellIndex._(cellIndex, li[0], li[1]);
  }

  static String _stringIndex(int colIndex, int rowIndex) {
    return getCellId(colIndex, rowIndex);
  }

  String _cellId;

  String get cellId {
    return this._cellId;
  }

  int _rowIndex;

  int get rowIndex {
    return this._rowIndex;
  }

  int _columnIndex;

  int get columnIndex {
    return this._columnIndex;
  }
}
