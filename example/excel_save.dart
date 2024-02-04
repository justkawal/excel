import 'dart:io';

import 'package:excel/excel.dart';

void main() {
  final excel = Excel.createExcel();
  int count = 0;
  const testSheet = 'Sheet To Remove';
  (List<List<dynamic>>.generate(3, (_) => List<int>.generate(3, (i) => i + 1))
        ..insert(0, [
          'A',
          'B',
          'C',
        ]))
      .forEach((el) {
    excel.insertRowIterables(
        'Test Sheet',
        el.map((e) {
          final string = e.toString();
          return int.tryParse(string) != null
              ? IntCellValue(int.parse(string))
              : TextCellValue(string);
        }).toList(),
        count);
    excel.insertRowIterables(
        testSheet,
        el.map((e) {
          final string = e.toString();
          return int.tryParse(string) != null
              ? IntCellValue(int.parse(string))
              : TextCellValue(string);
        }).toList(),
        count);
    count++;
  });

  excel.delete(excel.getDefaultSheet()!);
  excel.delete(testSheet);
  excel.rename('Test Sheet', 'Test Sheet Rename');
  excel.setDefaultSheet('Test Sheet Rename');

  final bytes = excel.encode();
  if (bytes != null) {
    File('example/example.xlsx')
      ..createSync()
      ..writeAsBytesSync(bytes);
  }
}
