import 'dart:convert';
import 'dart:io';
import 'dart:math';
import 'package:path/path.dart';

import '../lib/excel.dart';

void main(List<String> args) {
  var excel = Excel.createExcel();
  final Sheet sheet = excel[excel.getDefaultSheet()!];

  for (var row = 0; row < 100; row++) {
    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: row))
        .value = getRandString();

    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row))
        .value = getRandString();

    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: row))
        .value = getRandString();

    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: row))
        .value = getRandString();

    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 7, rowIndex: row))
        .value = getRandString();

    sheet
        .cell(CellIndex.indexByColumnRow(columnIndex: 50, rowIndex: row))
        .value = getRandString();
  }

  sheet.setColWidth(0, 10.0);
  sheet.setColWidth(1, 10.0);
  sheet.setColAutoFit(0);
  sheet.setColAutoFit(1);
  sheet.setColAutoFit(2);
  sheet.setColWidth(50, 10.0);

  sheet.merge(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0),
      CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 10));

  String outputFile =
      "/Users/igdmit/Downloads/excel_custom-${DateTime.now().toIso8601String()}.xlsx";

  List<int>? fileBytes = excel.save();
  if (fileBytes != null) {
    File(join(outputFile))
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
