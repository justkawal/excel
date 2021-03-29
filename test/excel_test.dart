import 'dart:io';

import 'package:excel/excel.dart';
import 'package:test/test.dart';

void main() {
  test('Create New XLSX File', () {
    var excel = Excel.createExcel();
    expect(excel.sheets.entries.length, equals(1));
    expect(excel.sheets.entries.first.key, equals('Sheet1'));
  });

  test('Read XLSX File', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(excel.tables['Sheet1']!.maxCols, equals(3));
    expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(), equals('Washington'));
  });

  group('Sheet Operations', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    Excel excel = Excel.decodeBytes(bytes);
    test('create Sheet', () {
      Sheet sheetObject = excel['SheetTmp'];
      sheetObject.insertRowIterables(['Country', 'Capital', 'Head'], 0);
      sheetObject.insertRowIterables(['Russia', 'Moscow', 'Putin'], 1);
      expect(excel.sheets.entries.length, equals(2));
      expect(
          excel.tables['Sheet1']!.rows[1][1]!.value.toString(), equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxCols, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(), equals('Putin'));
    });

    test('copy Sheet', () {
      excel.copy('SheetTmp', 'SheetTmp2');
      expect(excel.sheets.entries.length, equals(3));
      expect(
          excel.tables['Sheet1']!.rows[1][1]!.value.toString(), equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxCols, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(), equals('Putin'));
      expect(excel.tables['SheetTmp2']!.rows[1][2]!.value.toString(), equals('Putin'));
    });

    test('rename Sheet', () {
      excel.rename('SheetTmp2', 'SheetTmp3');
      expect(excel.sheets.entries.length, equals(3));
      expect(excel.tables['Sheettmp2'], equals(null));
      expect(
          excel.tables['Sheet1']!.rows[1][1]!.value.toString(), equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxCols, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(), equals('Putin'));
      expect(excel.tables['SheetTmp3']!.rows[1][2]!.value.toString(), equals('Putin'));
    });

    test('delete Sheet', () {
      excel.delete('SheetTmp3');
      excel.delete('SheetTmp');
      expect(excel.sheets.entries.length, equals(1));
      expect(
          excel.tables['Sheet1']!.rows[1][1]!.value.toString(), equals('Washington'));
    });
  });

  test('Saving XLSX File', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    Sheet? sheetObject = excel.tables['Sheet1']!;
    sheetObject.insertRowIterables(['Russia', 'Moscow', 'Putin'], 4);
    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/exampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/exampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);
    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));
    expect(
        newExcel.tables['Sheet1']!.rows[1][1]!.value.toString(), equals('Washington'));
    expect(newExcel.tables['Sheet1']!.maxCols, equals(3));
    expect(newExcel.tables['Sheet1']!.rows[4][1]!.value.toString(), equals('Moscow'));
  });
}
