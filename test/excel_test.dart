import 'dart:io';
import 'dart:math';

import 'package:excel/excel.dart';
import 'package:test/test.dart';

void main() {
  test('Create New XLSX File', () {
    var excel = Excel.createExcel();
    expect(excel.sheets.entries.length, equals(1));
    expect(excel.sheets.entries.first.key, equals('Sheet1'));
  });

  test('Read XLSX File', () {
    var file = "./test/test_resources/example.xlsx";
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(excel.tables['Sheet1'].maxCols, equals(3));
    expect(excel.tables["Sheet1"].rows[1][1].toString(), equals('Washington'));
  });

  test('Convert XLSX File -> toJson()', () {
    var file = "./test/test_resources/example.xlsx";
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    var json  = excel.toJson();

    expect(json, isMap);
    expect(json, isNotEmpty);
    expect(json.values.first.length, excel.sheets.values.first.maxRows - 1); // NO header row
    expect(json.values.first.first.values.first, excel.sheets.values.first.rows[1].first);
  });

  group('Sheet Operations', () {
    var file = "./test/test_resources/example.xlsx";
    var bytes = File(file).readAsBytesSync();
    Excel excel = Excel.decodeBytes(bytes);
    test('create Sheet', () {
      Sheet sheetObject = excel['SheetTmp'];
      sheetObject.insertRowIterables(["Country", "Capital", "Head"], 0);
      sheetObject.insertRowIterables(["Russia", "Moscow", "Putin"], 1);
      expect(excel.sheets.entries.length, equals(2));
      expect(
          excel.tables["Sheet1"].rows[1][1].toString(), equals('Washington'));
      expect(excel.tables['SheetTmp'].maxCols, equals(3));
      expect(excel.tables["SheetTmp"].rows[1][2].toString(), equals('Putin'));
    });

    test('copy Sheet', () {
      excel.copy('SheetTmp', 'SheetTmp2');
      expect(excel.sheets.entries.length, equals(3));
      expect(
          excel.tables["Sheet1"].rows[1][1].toString(), equals('Washington'));
      expect(excel.tables['SheetTmp'].maxCols, equals(3));
      expect(excel.tables["SheetTmp"].rows[1][2].toString(), equals('Putin'));
      expect(excel.tables["SheetTmp2"].rows[1][2].toString(), equals('Putin'));
    });

    test('rename Sheet', () {
      excel.rename('SheetTmp2', 'SheetTmp3');
      expect(excel.sheets.entries.length, equals(3));
      expect(excel.tables['Sheettmp2'], equals(null));
      expect(
          excel.tables["Sheet1"].rows[1][1].toString(), equals('Washington'));
      expect(excel.tables['SheetTmp'].maxCols, equals(3));
      expect(excel.tables["SheetTmp"].rows[1][2].toString(), equals('Putin'));
      expect(excel.tables["SheetTmp3"].rows[1][2].toString(), equals('Putin'));
    });

    test('delete Sheet', () {
      excel.delete("SheetTmp3");
      excel.delete("SheetTmp");
      expect(excel.sheets.entries.length, equals(1));
      expect(
          excel.tables["Sheet1"].rows[1][1].toString(), equals('Washington'));
    });
  });

  test('Saving XLSX File', () {
    var file = "./test/test_resources/example.xlsx";
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    Sheet sheetObject = excel.tables['Sheet1'];
    sheetObject.insertRowIterables(["Russia", "Moscow", "Putin"], 4);
    var onValue = excel.encode();
    File(Directory.current.path + "/tmp/exampleOut.xlsx")
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);

    var file1 = "./tmp/exampleOut.xlsx";
    var bytes1 = File(file1).readAsBytesSync();
    var excel1 = Excel.decodeBytes(bytes1);
    // delete tmp folder
    new Directory("./tmp").delete(recursive: true);
    expect(excel1.sheets.entries.length, equals(1));
    expect(excel1.tables["Sheet1"].rows[1][1].toString(), equals('Washington'));
    expect(excel1.tables['Sheet1'].maxCols, equals(3));
    expect(excel1.tables["Sheet1"].rows[4][1].toString(), equals('Moscow'));
  });
}
