import 'dart:io';
import 'package:excel/excel.dart';
import 'package:flutter_test/flutter_test.dart';

void main() {
  test("read file and not null", () {
    final filePath = 'assets/mylekha.xlsx';
    final bytes = File(filePath).readAsBytesSync();
    final excel = Excel.decodeBytes(bytes);
    expect(excel, isNotNull);

    for (final sheetName in excel.tables.keys) {
      final sheet = excel.tables[sheetName]!;
      print(
        '=== Sheet: $sheetName (${sheet.maxRows} rows x ${sheet.maxColumns} cols) ===',
      );

      for (var rowIdx = 0; rowIdx < sheet.rows.length; rowIdx++) {
        final row = sheet.rows[rowIdx];
        final rowValues = <String>[];

        for (final cell in row) {
          if (cell == null) {
            rowValues.add('');
            continue;
          }
          final String display = switch (cell.value) {
            TextCellValue() => cell.value!.toString(),
            IntCellValue(value: final v) => v.toString(),
            DoubleCellValue(value: final v) => v.toString(),
            BoolCellValue(value: final v) => v.toString(),
            FormulaCellValue(formula: final f) => '=$f',
            DateCellValue(year: final y, month: final mo, day: final d) =>
              '$y-$mo-$d',
            DateTimeCellValue() =>
              (cell.value as DateTimeCellValue).asDateTimeLocal().toString(),
            TimeCellValue(hour: final h, minute: final m, second: final s) =>
              '$h:$m:$s',
            null => '',
          };
          rowValues.add(display);
        }

        print('Row ${rowIdx + 1}: ${rowValues.join(' | ')}');
      }
    }
  });
}
