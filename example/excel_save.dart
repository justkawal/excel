import 'dart:io';

import 'package:excel/excel.dart';

void main() {
  final excel = Excel.createExcel();
  int count = 0;
  const defaultSheetName = 'Sheet1';
  const testSheetToKeep = 'Sheet To Keep';
  const testSheetToKeepRename = 'Rename Of Sheet To Keep';
  const testSheetToRemove = 'Sheet To Remove';
  (List<List<dynamic>>.generate(3, (_) => List<int>.generate(3, (i) => i + 1))
        ..insert(0, [
          'A',
          'B',
          'C',
        ]))
      .forEach((el) {
    excel.insertRowIterables(
        testSheetToKeep,
        el.map((e) {
          final string = e.toString();
          return int.tryParse(string) != null
              ? IntCellValue(int.parse(string))
              : TextCellValue(string);
        }).toList(),
        count);
    excel.insertRowIterables(
        testSheetToRemove,
        el.map((e) {
          final string = e.toString();
          return int.tryParse(string) != null
              ? IntCellValue(int.parse(string))
              : TextCellValue(string);
        }).toList(),
        count);
    count++;
  });

  ///
  assert(excel.sheets.keys.contains(defaultSheetName));
  assert(excel.getDefaultSheet() == defaultSheetName);
  excel.delete(excel.getDefaultSheet()!);
  assert(!excel.sheets.keys.contains(defaultSheetName));

  ///
  assert(excel.sheets.keys.contains(testSheetToRemove));
  excel.delete(testSheetToRemove);
  assert(!excel.sheets.keys.contains(testSheetToRemove));

  ///
  excel.rename(testSheetToKeep, testSheetToKeepRename);
  excel.setDefaultSheet(testSheetToKeepRename);
  assert(excel.getDefaultSheet() == testSheetToKeepRename);

  final bytes = excel.encode();
  if (bytes != null) {
    File('example/example.xlsx')
      ..createSync()
      ..writeAsBytesSync(bytes);
  }
}
