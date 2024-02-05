import 'dart:convert';
import 'dart:io';
import 'package:archive/archive.dart';
import 'package:excel/excel.dart';
import 'package:test/test.dart';
import 'package:xml/xml.dart';

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
    expect(excel.tables['Sheet1']!.maxColumns, equals(3));
    expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'));
  });

  test('Cell Data-Types from Microsoft Excel 365 Destkop', () {
    var file = './test/test_resources/dataTypesUsingMsExcel365Desktop.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(
      excel.tables['Tabelle1']!.rows[2][1]?.value,
      equals(TextCellValue('Some text')),
    );
    expect(
      excel.tables['Tabelle1']?.rows[3][1]?.value,
      equals(IntCellValue(42)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[4][1]?.value,
      equals(DoubleCellValue(12.3)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[5][1]?.value,
      equals(DateCellValue(year: 2023, month: 4, day: 20)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[6][1]?.value,
      equals(DateTimeCellValue(
          year: 2023, month: 4, day: 20, hour: 15, minute: 44, second: 13)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[7][1]?.value,
      equals(BoolCellValue(true)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[8][1]?.value,
      equals(BoolCellValue(false)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[9][1]?.value,
      equals(DoubleCellValue(15.99)),
    );
    expect(
      excel.tables['Tabelle1']?.rows[10][1]?.value,
      equals(DoubleCellValue(0.05)),
    );
  });

  test('Cell Data-Types from Google Spreadsheet', () {
    var file = './test/test_resources/dataTypesUsingGoogleSpreadsheet.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(
      excel.tables['Sheet1']?.rows[2][1]?.value,
      equals(TextCellValue('Some text')),
    );
    expect(
      excel.tables['Sheet1']?.rows[3][1]?.value,
      equals(IntCellValue(42)),
    );
    expect(
      excel.tables['Sheet1']?.rows[4][1]?.value,
      equals(DoubleCellValue(12.3)),
    );
    expect(
      excel.tables['Sheet1']?.rows[5][1]?.value,
      equals(DateCellValue(year: 2023, month: 4, day: 20)),
    );
    expect(
      excel.tables['Sheet1']?.rows[6][1]?.value,
      equals(DateTimeCellValue(
          year: 2023, month: 4, day: 20, hour: 15, minute: 44, second: 13)),
    );
    expect(
      excel.tables['Sheet1']?.rows[7][1]?.value,
      equals(BoolCellValue(true)),
    );
    expect(
      excel.tables['Sheet1']?.rows[8][1]?.value,
      equals(BoolCellValue(false)),
    );
    expect(
      excel.tables['Sheet1']?.rows[9][1]?.value,
      equals(DoubleCellValue(15.99)),
    );
    expect(
      excel.tables['Sheet1']?.rows[10][1]?.value,
      equals(DoubleCellValue(0.05)),
    );
  });

  test('Cell Data-Types from LibreOffice', () {
    var file = './test/test_resources/dataTypesUsingLibreoffice.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    expect(
      excel.tables['Sheet1']?.rows[2][1]?.value,
      equals(TextCellValue('Some text')),
    );
    expect(
      excel.tables['Sheet1']?.rows[3][1]?.value,
      equals(IntCellValue(42)),
    );
    expect(
      excel.tables['Sheet1']?.rows[4][1]?.value,
      equals(DoubleCellValue(12.3)),
    );
    expect(
      excel.tables['Sheet1']?.rows[5][1]?.value,
      equals(DateCellValue(year: 2023, month: 4, day: 20)),
    );
    expect(
      excel.tables['Sheet1']?.rows[6][1]?.value,
      equals(DateTimeCellValue(
          year: 2023, month: 4, day: 20, hour: 15, minute: 44, second: 13)),
    );
    expect(
      excel.tables['Sheet1']?.rows[7][1]?.value,
      equals(BoolCellValue(true)),
    );
    expect(
      excel.tables['Sheet1']?.rows[8][1]?.value,
      equals(BoolCellValue(false)),
    );
    expect(
      excel.tables['Sheet1']?.rows[9][1]?.value,
      equals(DoubleCellValue(15.99)),
    );
    expect(
      excel.tables['Sheet1']?.rows[10][1]?.value,
      equals(DoubleCellValue(0.05)),
    );
  });

  test('Read/Write various data types', () {
    var file = './test/test_resources/dataTypesUsingMsExcel365Desktop.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    {
      final sheet = excel.tables['Tabelle1']!;
      sheet.updateCell(
        CellIndex.indexByString('B4'),
        DoubleCellValue(13.37),
      );
      sheet.updateCell(
        CellIndex.indexByString('B5'),
        DateCellValue(year: 2025, month: 11, day: 28),
      );
      sheet.updateCell(
        CellIndex.indexByString('B6'),
        null,
      );
      sheet.updateCell(
        CellIndex.indexByString('B7'),
        TimeCellValue(hour: 20, minute: 15),
      );
      sheet.updateCell(
        CellIndex.indexByString('B8'),
        DoubleCellValue(42),
        cellStyle: CellStyle(numberFormat: NumFormat.standard_11),
      );

      final b10 = sheet.cell(CellIndex.indexByString('B10'));
      b10.cellStyle = (b10.cellStyle ?? CellStyle()).copyWith(
        numberFormat: CustomNumericNumFormat(formatCode: r'0\m\²'),
      );
    }

    final bytesWritten = excel.encode()!;
    final excelAgain = Excel.decodeBytes(bytesWritten);
    {
      final sheet = excelAgain.tables['Tabelle1']!;
      final b3 = sheet.cell(CellIndex.indexByString('B3'));
      expect(b3.value, equals(TextCellValue('Some text')));
      expect(
        b3.cellStyle?.numberFormat ?? NumFormat.standard_0,
        equals(NumFormat.standard_0),
      );

      final b4 = sheet.cell(CellIndex.indexByString('B4'));
      expect(b4.value, equals(DoubleCellValue(13.37)));
      expect(
        b4.cellStyle?.numberFormat ?? NumFormat.defaultFloat,
        equals(NumFormat.defaultFloat),
      );

      final b5 = sheet.cell(CellIndex.indexByString('B5'));
      expect(b5.value, equals(DateCellValue(year: 2025, month: 11, day: 28)));
      expect(
        b5.cellStyle?.numberFormat,
        equals(NumFormat.defaultDate),
      );

      final b6 = sheet.cell(CellIndex.indexByString('B6'));
      expect(b6.value, equals(null));
      expect(
        b6.cellStyle?.numberFormat,
        equals(NumFormat.standard_0),
      );

      final b7 = sheet.cell(CellIndex.indexByString('B7'));
      expect(b7.value, equals(TimeCellValue(hour: 20, minute: 15)));
      expect(
        b7.cellStyle?.numberFormat,
        equals(NumFormat.defaultTime),
      );

      final b8 = sheet.cell(CellIndex.indexByString('B8'));
      expect(b8.value, equals(IntCellValue(42)));
      expect(
        b8.cellStyle?.numberFormat,
        equals(NumFormat.standard_11),
      );

      final b10 = sheet.cell(CellIndex.indexByString('B10'));
      expect(b10.value, equals(DoubleCellValue(15.99)));
      expect(
        b10.cellStyle?.numberFormat,
        equals(CustomNumericNumFormat(formatCode: r'0\m\²')),
      );
    }
  });

  group('Sheet Operations', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    Excel excel = Excel.decodeBytes(bytes);
    test('create Sheet', () {
      Sheet sheetObject = excel['SheetTmp'];
      sheetObject.insertRowIterables([
        TextCellValue('Country'),
        TextCellValue('Capital'),
        TextCellValue('Head')
      ], 0);
      sheetObject.insertRowIterables([
        TextCellValue('Russia'),
        TextCellValue('Moscow'),
        TextCellValue('Putin')
      ], 1);
      expect(excel.sheets.entries.length, equals(2));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxColumns, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(),
          equals('Putin'));
    });

    test('copy Sheet', () {
      excel.copy('SheetTmp', 'SheetTmp2');
      expect(excel.sheets.entries.length, equals(3));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxColumns, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(),
          equals('Putin'));
      expect(excel.tables['SheetTmp2']!.rows[1][2]!.value.toString(),
          equals('Putin'));
    });

    test('rename Sheet', () {
      excel.rename('SheetTmp2', 'SheetTmp3');
      expect(excel.sheets.entries.length, equals(3));
      expect(excel.tables['Sheettmp2'], equals(null));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
      expect(excel.tables['SheetTmp']!.maxColumns, equals(3));
      expect(excel.tables['SheetTmp']!.rows[1][2]!.value.toString(),
          equals('Putin'));
      expect(excel.tables['SheetTmp3']!.rows[1][2]!.value.toString(),
          equals('Putin'));
    });

    test('delete Sheet', () {
      excel.delete('SheetTmp3');
      excel.delete('SheetTmp');
      expect(excel.sheets.entries.length, equals(1));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Washington'));
    });
  });

  test('Saving XLSX File', () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    Sheet? sheetObject = excel.tables['Sheet1']!;
    sheetObject.insertRowIterables([
      TextCellValue('Russia'),
      TextCellValue('Moscow'),
      TextCellValue('Putin')
    ], 4);
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
    expect(newExcel.tables['Sheet1']!.rows[1][1]!.value.toString(),
        equals('Washington'));
    expect(newExcel.tables['Sheet1']!.maxColumns, equals(3));
    expect(newExcel.tables['Sheet1']!.rows[4][1]!.value.toString(),
        equals('Moscow'));
  });

  test('Saving XLSX File with superscript', () {
    var file = './test/test_resources/superscriptExample.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/superscriptExampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/superscriptExampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);
    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));

    expect(newExcel.tables['Sheet1']!.rows[0][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[1][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[2][0]!.value.toString(),
        equals('Text in A3'));
  });

  test(
      'Add already shared strings and make sure that they are reused by checking increased usage count but equal unique count',
      () {
    var file = './test/test_resources/example.xlsx';
    var bytes = File(file).readAsBytesSync();
    var archive = ZipDecoder().decodeBytes(bytes);
    var sharedStringsArchive = archive.findFile('xl/sharedStrings.xml')!;

    var oldSharedStringsDocument =
        XmlDocument.parse(utf8.decode(sharedStringsArchive.content));
    var oldCount = oldSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("count");
    var oldUniqueCount = oldSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("uniqueCount");

    var excel = Excel.decodeBytes(bytes);

    Sheet? sheetObject = excel.tables['Sheet1']!;
    sheetObject.insertRowIterables([
      TextCellValue('ISRAEL'),
      TextCellValue('Jerusalem'),
      TextCellValue('Benjamin Netanyahu')
    ], 4);
    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/exampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/exampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    expect(() => Excel.decodeBytes(newFileBytes), returnsNormally);

    var newArchive = ZipDecoder().decodeBytes(newFileBytes);
    var newSharedStringsArchive = newArchive.findFile('xl/sharedStrings.xml')!;

    var newSharedStringsDocument =
        XmlDocument.parse(utf8.decode(newSharedStringsArchive.content));
    var newCount = newSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("count");
    var newUniqueCount = newSharedStringsDocument
        .findAllElements('sst')
        .first
        .getAttributeNode("uniqueCount");

    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);

    expect(oldUniqueCount!.value, equals(newUniqueCount!.value));
    expect(oldCount!.value, "12");
    expect(newCount!.value, "15");
  });

  test('Saving XLSX File with superscript', () {
    var file = './test/test_resources/superscriptExample.xlsx';
    var bytes = File(file).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    var fileBytes = excel.encode();
    if (fileBytes != null) {
      File(Directory.current.path + '/tmp/superscriptExampleOut.xlsx')
        ..createSync(recursive: true)
        ..writeAsBytesSync(fileBytes);
    }
    var newFile = './tmp/superscriptExampleOut.xlsx';
    var newFileBytes = File(newFile).readAsBytesSync();
    var newExcel = Excel.decodeBytes(newFileBytes);
    // delete tmp folder
    new Directory('./tmp').delete(recursive: true);
    expect(newExcel.sheets.entries.length, equals(1));

    expect(newExcel.tables['Sheet1']!.rows[0][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[1][0]!.value.toString(),
        equals('Text and superscript text'));
    expect(newExcel.tables['Sheet1']!.rows[2][0]!.value.toString(),
        equals('Text in A3'));
  });

  group('Header/Footer', () {
    test("Update header/footer", () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      Sheet? sheetObject = excel.tables['Sheet1']!;

      sheetObject.headerFooter!.oddHeader = "Foo";
      sheetObject.headerFooter!.oddFooter = "Bar";

      var fileBytes = excel.encode();
      if (fileBytes != null) {
        File(Directory.current.path + '/tmp/exampleOut.xlsx')
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }
      var newFile = './tmp/exampleOut.xlsx';
      var newFileBytes = File(newFile).readAsBytesSync();
      var newExcel = Excel.decodeBytes(newFileBytes);
      expect(
          newExcel.tables['Sheet1']!.headerFooter!.oddHeader!, equals('Foo'));
      expect(
          newExcel.tables['Sheet1']!.headerFooter!.oddFooter!, equals('Bar'));

      // delete tmp folder only when test is successful (diagnosis)
      new Directory('./tmp').delete(recursive: true);
    });

    test("Save empty Workbook", () {
      var excel = Excel.createExcel();
      excel.save();
    });

    test("Clone header/footer of existing Workbook", () {
      var file = './test/test_resources/example.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      Sheet? sheetObject = excel.tables['Sheet1']!;

      sheetObject.headerFooter!.oddHeader = "Foo";
      sheetObject.headerFooter!.oddFooter = "Bar";

      excel.copy('Sheet1', 'test_sheet');

      Sheet? testSheet = excel.tables['test_sheet'];

      expect(testSheet!.headerFooter!.oddHeader!, equals('Foo'));
      expect(testSheet.headerFooter!.oddFooter!, equals('Bar'));
    });

    test("Remove header/footer from Workbook", () {});

    test("Reader headerFooter attributes", () {
      var file = './test/test_resources/headerFooter.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      Sheet? sheetObject = excel.tables['Sheet1']!;

      var headerFooter = sheetObject.headerFooter!;

      expect(headerFooter.alignWithMargins, isFalse);
      expect(headerFooter.differentFirst, isTrue);
      expect(headerFooter.differentOddEven, isTrue);
      expect(headerFooter.scaleWithDoc, isFalse);
    });
  });

  group('Borders', () {
    test('read file with borders', () {
      final file = './test/test_resources/borders.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;

      final borderEmpty = Border();
      final borderMedium = Border(borderStyle: BorderStyle.Medium);
      final borderMediumRed = Border(
          borderStyle: BorderStyle.Medium,
          borderColorHex: 'FFFF0000'.excelColor);
      final borderHair = Border(borderStyle: BorderStyle.Hair);
      final borderDouble = Border(borderStyle: BorderStyle.Double);

      final cellStyleA1 =
          sheetObject.cell(CellIndex.indexByString('A1')).cellStyle;
      expect(cellStyleA1?.leftBorder, equals(borderMedium));
      expect(cellStyleA1?.rightBorder, equals(borderMedium));
      expect(cellStyleA1?.topBorder, anyOf(isNull, equals(borderEmpty)));
      expect(cellStyleA1?.bottomBorder, equals(borderMediumRed));
      expect(cellStyleA1?.diagonalBorder, anyOf(isNull, equals(borderEmpty)));
      expect(cellStyleA1?.diagonalBorderUp, isFalse);
      expect(cellStyleA1?.diagonalBorderDown, isFalse);

      final cellStyleB3 =
          sheetObject.cell(CellIndex.indexByString('B3')).cellStyle;
      expect(cellStyleB3?.leftBorder, equals(borderMedium));
      expect(cellStyleB3?.rightBorder, equals(borderMedium));
      expect(cellStyleB3?.topBorder, equals(borderHair));
      expect(cellStyleB3?.bottomBorder, equals(borderHair));

      final cellStyleA5 =
          sheetObject.cell(CellIndex.indexByString('A5')).cellStyle;
      expect(cellStyleA5?.diagonalBorder, equals(borderDouble));
      expect(cellStyleA5?.diagonalBorderUp, isFalse);
      expect(cellStyleA5?.diagonalBorderDown, isTrue);

      final cellStyleC5 =
          sheetObject.cell(CellIndex.indexByString('C5')).cellStyle;
      expect(cellStyleC5?.diagonalBorder, equals(borderDouble));
      expect(cellStyleC5?.diagonalBorderUp, isTrue);
      expect(cellStyleC5?.diagonalBorderDown, isFalse);
    });

    test('test support all border styles', () {
      final file = './test/test_resources/borders2.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;

      final borderStyles = <BorderStyle>[
        BorderStyle.None,
        BorderStyle.DashDot,
        BorderStyle.DashDotDot,
        BorderStyle.Dashed,
        BorderStyle.Dotted,
        BorderStyle.Double,
        BorderStyle.Hair,
        BorderStyle.Medium,
        BorderStyle.MediumDashDot,
        BorderStyle.MediumDashDotDot,
        BorderStyle.MediumDashed,
        BorderStyle.SlantDashDot,
        BorderStyle.Thick,
        BorderStyle.Thin,
      ];

      for (var i = 1; i < borderStyles.length; ++i) {
        // Loop from i = 1, as Excel does not set None type.
        final border = Border(borderStyle: borderStyles[i]);

        final cellStyle = sheetObject
            .cell(CellIndex.indexByString('B${2 * (i + 1)}'))
            .cellStyle;

        expect(cellStyle?.leftBorder, equals(border));
        expect(cellStyle?.rightBorder, equals(border));
        expect(cellStyle?.topBorder, equals(border));
        expect(cellStyle?.bottomBorder, equals(border));
      }
    });

    test('test support for merged cells with borders', () {
      final file = './test/test_resources/mergedBorders.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      final Sheet sheetObject = excel.tables['Sheet1']!;

      final borderStyles = <BorderStyle>[
        BorderStyle.None,
        BorderStyle.DashDot,
        BorderStyle.DashDotDot,
        BorderStyle.Dashed,
        BorderStyle.Dotted,
        BorderStyle.Double,
        BorderStyle.Hair,
        BorderStyle.Medium,
        BorderStyle.MediumDashDot,
        BorderStyle.MediumDashDotDot,
        BorderStyle.MediumDashed,
        BorderStyle.SlantDashDot,
        BorderStyle.Thick,
        BorderStyle.Thin,
      ];

      sheetObject.merge(
          CellIndex.indexByString('B2'), CellIndex.indexByString('D4'));

      for (var i = 1; i < borderStyles.length; ++i) {
        // Loop from i = 1, as Excel does not set None type.
        final border = Border(
            borderStyle: borderStyles[i],
            borderColorHex: "FF000000".excelColor);
        final start = CellIndex.indexByString('B${(4 * i + 2)}');
        final end = CellIndex.indexByString('D${(4 * i + 4)}');

        sheetObject.merge(start, end);

        sheetObject.setMergedCellStyle(
          start,
          CellStyle(
            leftBorder: border,
            rightBorder: border,
            topBorder: border,
            bottomBorder: border,
          ),
        );
      }

      for (var i = 1; i < borderStyles.length; ++i) {
        CellIndex cellIndexStart = CellIndex.indexByString('B${(4 * i + 2)}');
        CellIndex cellIndexEnd = CellIndex.indexByString('D${(4 * i + 4)}');

        for (var j = cellIndexStart.rowIndex; j <= cellIndexEnd.rowIndex; j++) {
          for (var k = cellIndexStart.columnIndex;
              k <= cellIndexEnd.columnIndex;
              k++) {
            final cellStyle = sheetObject
                .cell(CellIndex.indexByColumnRow(columnIndex: k, rowIndex: j))
                .cellStyle;

            final borderStyle = Border(
              borderStyle: borderStyles[i],
              borderColorHex: "FF000000".excelColor,
            );

            if (j == cellIndexStart.rowIndex) {
              expect(cellStyle?.topBorder, equals(borderStyle));
            }

            if (j == cellIndexEnd.rowIndex) {
              expect(cellStyle?.bottomBorder, equals(borderStyle));
            }

            if (k == cellIndexStart.columnIndex) {
              expect(cellStyle?.leftBorder, equals(borderStyle));
            }

            if (k == cellIndexEnd.columnIndex) {
              expect(cellStyle?.rightBorder, equals(borderStyle));
            }
          }
        }
      }
    });

    test('saving XLSX File with borders', () {
      final file = './test/test_resources/borders.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);

      final outFilePath = Directory.current.path + '/tmp/bordersOut.xlsx';
      final fileBytes = excel.encode();
      if (fileBytes != null) {
        File(outFilePath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }

      final newFileBytes = File(outFilePath).readAsBytesSync();
      final newExcel = Excel.decodeBytes(newFileBytes);
      expect(newExcel.sheets.entries.length, equals(1));

      final borderEmpty = Border();
      final borderMedium = Border(borderStyle: BorderStyle.Medium);
      final borderMediumRed = Border(
          borderStyle: BorderStyle.Medium,
          borderColorHex: 'FFFF0000'.excelColor);

      final Sheet sheetObject = newExcel.tables['Sheet1']!;
      final cellStyleB1 =
          sheetObject.cell(CellIndex.indexByString('B1')).cellStyle;
      expect(cellStyleB1?.leftBorder, equals(borderMedium));
      expect(cellStyleB1?.rightBorder, equals(borderMedium));
      expect(cellStyleB1?.topBorder, equals(borderEmpty));
      expect(cellStyleB1?.bottomBorder, equals(borderMediumRed));

      // delete tmp folder only when test is successful (diagnosis)
      new Directory('./tmp').delete(recursive: true);
    });
  });

  group('rPh tag', () {
    test('Read Cell shared text without rPh elements', () {
      var file = './test/test_resources/rphSample.xlsx';
      var bytes = File(file).readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      expect(excel.tables['Sheet1']!.rows[1][0]!.value.toString(),
          equals('plainText'));
      expect(excel.tables['Sheet1']!.rows[1][1]!.value.toString(),
          equals('Hellow world'));
      expect(excel.tables['Sheet1']!.rows[1][2]!.value.toString(),
          equals('世界よこんにちは'));
      expect(excel.tables['Sheet1']!.rows[2][2]!.value.toString(),
          equals('ようこそユーザー'));
      expect(excel.tables['Sheet1']!.rows[3][2]!.value.toString(),
          equals('ロケール選択'));
      expect(excel.tables['Sheet1']!.rows[4][2]!.value.toString(),
          equals('ロケール選択'));
    });

    test('saving XLSX File without rPh elements', () {
      final file = './test/test_resources/rphSample.xlsx';
      final bytes = File(file).readAsBytesSync();
      final excel = Excel.decodeBytes(bytes);
      excel.tables['Sheet1']!.rows[3][2]!.value = TextCellValue('ロケール選択');

      final outFilePath = Directory.current.path + '/tmp/rphSampleOut.xlsx';
      final fileBytes = excel.encode();
      if (fileBytes != null) {
        File(outFilePath)
          ..createSync(recursive: true)
          ..writeAsBytesSync(fileBytes);
      }

      final newFileBytes = File(outFilePath).readAsBytesSync();
      final newExcel = Excel.decodeBytes(newFileBytes);
      expect(newExcel.tables['Sheet1']!.rows[3][2]!.value.toString(),
          equals('ロケール選択'));

      // delete tmp folder only when test is successful (diagnosis)
      new Directory('./tmp').delete(recursive: true);
    });
  });

  group(".xls file handling", () {
    test("Exception when opening old .xls file", () {
      final file = './test/test_resources/oldXLSFile.xls';
      final bytes = File(file).readAsBytesSync();
      try {
        Excel.decodeBytes(bytes);
      } catch (e) {
        expect(e, isA<UnsupportedError>());
        expect(
            e.toString(),
            equals(
                'Unsupported operation: Excel format unsupported. Only .xlsx files are supported'));
      }
    });

    test("Exception when opening new .xls file", () {
      final file = './test/test_resources/newXLSFile.xls';
      final bytes = File(file).readAsBytesSync();
      try {
        Excel.decodeBytes(bytes);
      } catch (e) {
        expect(e, isA<UnsupportedError>());
        expect(
            e.toString(),
            equals(
                'Unsupported operation: Excel format unsupported. Only .xlsx files are supported'));
      }
    });
  });
}
