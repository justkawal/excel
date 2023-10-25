import 'dart:io';

import '../lib/excel.dart';

void main(List<String> args) {
  var excel = Excel.createExcel();
  final Sheet sheet = excel[excel.getDefaultSheet()!];

  sheet.merge(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
      CellIndex.indexByColumnRow(columnIndex: 10, rowIndex: 5));

  sheet.merge(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 10),
      CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 10));

  Border border = Border(
    borderColorHex: "#FF000000",
    borderStyle: BorderStyle.Thin,
  );

  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
    "Merged cell border",
  );

  sheet.setMergedCellStyle(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
    CellStyle(
      fontSize: 25,
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
      diagonalBorderDown: true,
    ),
  );

  sheet.setMergedCellStyle(
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 10),
    CellStyle(
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
      diagonalBorderDown: true,
      diagonalBorderUp: true,
    ),
  );

  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1),
    "Normal border",
    cellStyle: CellStyle(
      fontSize: 25,
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
    ),
  );

  sheet.setColumnWidth(0, 50);

  // Create the example excel file in the current directory
  String outputFile = "excel_custom.xlsx";

  List<int>? fileBytes = excel.save();
  if (fileBytes != null) {
    File(outputFile)
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
