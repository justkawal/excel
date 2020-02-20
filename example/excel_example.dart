import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "Path_to_input/excel.xlsx";
  var bytes = File(file).readAsBytesSync();
  var updater = Excel.decodeBytes(bytes, update: true);

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
        backgroundColorHex: "#1AFF1A")
    ..updateCell(sheet, CellIndex.indexByString("E5"), "Here Value of E5",
        horizontalAlign: HorizontalAlign.Right);

  updater.encode().then((onValue) {
    File(join("Path_to_destination/excel_out.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  }).then((_) {
    print(
        "\n****************************** Printing Updated Data Directly by reading output file ******************************\n");
    var fileOut = "Path_to_destination/excel_out.xlsx";
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
