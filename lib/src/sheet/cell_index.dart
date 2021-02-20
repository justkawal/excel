part of excel;

class CellIndex {
  CellIndex._({int col, int row}) {
    this._columnIndex = col;
    this._rowIndex = row;
  }

  ///
  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///```
  static CellIndex indexByColumnRow({int columnIndex, int rowIndex}) {
    return CellIndex._(col: columnIndex, row: rowIndex);
  }

  ///
  ///```
  /// CellIndex.indexByColumnRow('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByColumnRow('A2'); // columnIndex: 0, rowIndex: 1
  ///```
  static CellIndex indexByString(String cellIndex) {
    List<int> li = _cellCoordsFromCellId(cellIndex);
    return CellIndex._(row: li[0], col: li[1]);
  }

  String get cellId {
    return getCellId(this.columnIndex, this.rowIndex);
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
