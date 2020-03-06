import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/excel.xlsx";
  var bytes = File(file).readAsBytesSync();
  var updater = Excel.createExcel(); //.decodeBytes(bytes, update: true);

  for (var table in updater.tables.keys) {
    print(table);
    print(updater.tables[table].maxCols);
    print(updater.tables[table].maxRows);
    for (var row in updater.tables[table].rows) {
      print("$row");
    }
  }

  // if sheet with name = Sheet24 does not exist then it will be automatically created.
  var sheet = 'Sheet24';

  updater
    ..updateCell(sheet, CellIndex.indexByString("A1"), "Here Value of A1",
        fontColorHex: "#1AFF1A", verticalAlign: VerticalAlign.Top)
    ..updateCell(sheet, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
        "Here Value of C1", wrap: TextWrapping.WrapText)
    ..updateCell(sheet, CellIndex.indexByString("A2"), "Here Value of A2",
        backgroundColorHex: "#1AFF1A", wrap: TextWrapping.Clip)
    ..updateCell(sheet, CellIndex.indexByString("E5"), " E5",
        horizontalAlign: HorizontalAlign.Right, wrap: TextWrapping.Clip)
    ..merge(
        sheet, CellIndex.indexByString("A1"), CellIndex.indexByString("B11"));

  for (var table in updater.tables.keys) {
    print(table);
    print(updater.tables[table].maxCols);
    print(updater.tables[table].maxRows);
    for (var row in updater.tables[table].rows) {
      print("$row");
    }
  }

  updater.encode().then((onValue) {
    File(join("/Users/kawal/Desktop/excel_out.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  }).then((_) {
    print(
        "\n****************************** Printing Updated Data Directly by reading output file ******************************\n");
    var fileOut = "/Users/kawal/Desktop/excel_out.xlsx";
    var bytesOut = File(fileOut).readAsBytesSync();
    var updaterOut = Excel.decodeBytes(bytesOut, update: true);

    for (var table in updaterOut.tables.keys) {
      print(table);
      print(updaterOut.tables[table].maxCols);
      print(updaterOut.tables[table].maxRows);
      for (var row in updaterOut.tables[table].rows) {
        print("$row");
      }
    }
  });
}

/* List<String> spannedCells = updater.getSpannedItems(sheet);
  var cellToUnMerge = "A1:A2";
  updater.unMerge(sheet, spannedCells.indexOf(cellToUnMerge)); */
