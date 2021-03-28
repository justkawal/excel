part of excel;

// ignore: must_be_immutable
class CellIndex extends Equatable {
  CellIndex._({int? col, int? row}) {
    assert(col != null && row != null);
    this._columnIndex = col!;
    this._rowIndex = row!;
  }

  ///
  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///```
  static CellIndex indexByColumnRow({int? columnIndex, int? rowIndex}) {
    assert(columnIndex != null && rowIndex != null);
    return CellIndex._(col: columnIndex!, row: rowIndex!);
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

  /// Avoid using it as it is very process expensive function.
  ///
  /// ```
  /// var cellIndex = CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 );
  /// var cell = cellIndex.cellId; // A1
  String get cellId {
    return getCellId(this.columnIndex, this.rowIndex);
  }

  late int _rowIndex;

  int get rowIndex {
    return this._rowIndex;
  }

  late int _columnIndex;

  int get columnIndex {
    return this._columnIndex;
  }

  @override
  List<Object?> get props => [_rowIndex, _columnIndex];
}
