part of excel;

class CellIndex {
  const CellIndex.indexByColumnRow({this.columnIndex, this.rowIndex});

  static CellIndex indexByString(String cellIndex) {
    return CellIndex.indexByColumnRow(
        rowIndex: cellCoordsFromCellId(cellIndex)[0],
        columnIndex: cellCoordsFromCellId(cellIndex)[1]);
  }

  final int rowIndex;

  int get _rowIndex {
    return rowIndex;
  }

  final int columnIndex;

  int get _columnIndex {
    return columnIndex;
  }
}
