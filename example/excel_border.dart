import 'dart:convert';
import 'dart:io';
import 'dart:math';

import '../lib/excel.dart';

void main(List<String> args) {
  var excel = Excel.createExcel();
  final Sheet sheet = excel[excel.getDefaultSheet()!];

  sheet.merge(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
      CellIndex.indexByColumnRow(columnIndex: 10, rowIndex: 5));

  sheet.merge(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 10),
      CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 10));

  Border border = Border(
    borderColorHex: "#000000",
    borderStyle: BorderStyle.Thin,
  );

  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 1),
    "Merged cell border",
    cellStyle: CellStyle(
      fontSize: 25,
      topBorder: border,
      bottomBorder: border,
      leftBorder: border,
      rightBorder: border,
      diagonalBorder: border,
      diagonalBorderDown: true,
    ),
  );

  sheet.updateCell(
    CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 10),
    "",
    cellStyle: CellStyle(
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

String getRandString() {
  final random = Random.secure();
  final len = random.nextInt(20);
  final values = List<int>.generate(len, (i) => random.nextInt(255));
  return base64UrlEncode(values);
}
