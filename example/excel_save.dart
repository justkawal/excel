import 'dart:io';

import 'package:excel/excel.dart';

void main() {
  final excel = Excel.createExcel();

  const defaultSheetName = 'Sheet1';
  const testSheetToKeep = 'Sheet To Keep';
  const testSheetToKeepRename = 'Rename Of Sheet To Keep';

  var listDynamic = (List<List<dynamic>>.generate(
      5, (_) => List<int>.generate(5, (i) => i + 1))
    ..insert(0, [
      'A',
      'B',
      'C',
      'D',
      'E',
    ]));

  for (var row = 0; row < listDynamic.length; row++) {
    for (var column = 0; column < listDynamic[row].length; column++) {
      final cellIndex = CellIndex.indexByColumnRow(
        columnIndex: column,
        rowIndex: row,
      );
      var colorList = List.of(ExcelColor.values);
      final border = Border(
          borderColorHex: (colorList..shuffle()).first,
          borderStyle: BorderStyle.Thin);

      final string = listDynamic[row][column].toString();

      var cellValue = int.tryParse(string) != null
          ? IntCellValue(int.parse(string))
          : TextCellValue(string);

      excel.updateCell(
        testSheetToKeep,
        cellIndex,
        cellValue,
        cellStyle: CellStyle()
          ..backgroundColor = (colorList..shuffle()).first
          ..topBorder = border
          ..bottomBorder = border
          ..leftBorder = border
          ..rightBorder = border
          ..fontColor = (colorList..shuffle()).first
          ..fontFamily = 'Arial',
      );
    }
  }

  ///
  assert(excel.sheets.keys.contains(defaultSheetName));
  assert(excel.getDefaultSheet() == defaultSheetName);
  excel.delete(excel.getDefaultSheet()!);
  assert(!excel.sheets.keys.contains(defaultSheetName));

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
