part of excel;

class CellIndex {
  CellIndex._(String cell, int col, int row) {
    this._columnIndex = col;
    this._rowIndex = row;
    this._cellId = cell;
  }
  static CellIndex indexByColumnRow({int columnIndex, int rowIndex}) {
    return CellIndex._(
        _stringIndex(columnIndex, rowIndex), columnIndex, rowIndex);
  }

  static CellIndex indexByString(String cellIndex) {
    List<int> li = cellCoordsFromCellId(cellIndex);

    return CellIndex._(cellIndex, li[1], li[0]);
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
