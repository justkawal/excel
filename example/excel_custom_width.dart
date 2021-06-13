import 'dart:io';
import 'package:path/path.dart';

import '../lib/excel.dart';

void main(List<String> args) {
  //final file = "/Users/igdmit/Downloads/reference_template_v1.xlsx";
  //final bytes = File(file).readAsBytesSync();
  //final excel = Excel.decodeBytes(bytes);

  var excel = Excel.createExcel();
  final Sheet sheet = excel[excel.getDefaultSheet()!];

  sheet.cell(CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0)).value =
      'abcdefghijklmnopqrstuvwxyzABCD123';
  sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 5)).value =
      'abcdefghijklm';
  sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 8)).value =
      'abcdefghijklm01234567890';
  sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 10)).value =
      'abcdefghijklm01234567890'.toUpperCase();

  sheet.setColWidth(0, 50.0);
  sheet.setColWidth(10000, 10.0);
  sheet.setColAutoFit(5);

  String outputFile =
      "/Users/igdmit/Downloads/excel_custom-${DateTime.now().toIso8601String()}.xlsx";

  List<int>? fileBytes = excel.save();
  if (fileBytes != null) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
