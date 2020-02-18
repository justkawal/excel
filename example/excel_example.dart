import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/excel.xlsx";
  var bytes = File(file).readAsBytesSync();
  var decoder = Excel.decodeBytes(bytes, update: true);

  for (var table in decoder.tables.keys) {
    print(table);
    print(decoder.tables[table].maxCols);
    print(decoder.tables[table].maxRows);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  // if sheet with name = Sheet24 does not exist then it will be automatically created.
  var sheet = 'Sheet24';

  decoder
    ..updateCell(sheet, CellIndex.indexByString("A1"), "Here Value of A1",
        fontColorHex: "#1AFF1A", verticalAlign: VerticalAlign.Top)
    ..updateCell(sheet, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
        "Here Value of C1", wrap: TextWrapping.WrapText)
    ..updateCell(sheet, CellIndex.indexByString("A2"), "Here Value of A2",
        backgroundColorHex: "#1AFF1A")
    ..updateCell(sheet, CellIndex.indexByString("E5"), "Here Value of E5",
        horizontalAlign: HorizontalAlign.Right);

  decoder.encode().then((onValue) {
    File(join("/Users/kawal/Desktop/excel_out.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  }).then((_) {
    print(
        "\n****************************** Printing Updated Data Directly by reading output file ******************************\n");
    var fileOut = "/Users/kawal/Desktop/excel_out.xlsx";
    var bytesOut = File(fileOut).readAsBytesSync();
    var decoderOut = Excel.decodeBytes(bytesOut, update: true);

    for (var table in decoderOut.tables.keys) {
      print(table);
      print(decoderOut.tables[table].maxCols);
      print(decoderOut.tables[table].maxRows);
      for (var row in decoderOut.tables[table].rows) {
        print("$row");
      }
    }
  });
}
