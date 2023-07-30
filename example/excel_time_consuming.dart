import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  Stopwatch stopwatch = new Stopwatch()..start();

  Excel excel = Excel.createExcel();
  Sheet sh = excel['Sheet1'];
  for (int i = 0; i < 8; i++) {
    sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).value =
        'Column $i';
    //sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).cellStyle =CellStyle(bold: true);
  }
  for (int row = 1; row < 9000; row++) {
    for (int column = 0; column < 80; column++) {
      sh
          .cell(CellIndex.indexByColumnRow(rowIndex: row, columnIndex: column))
          .value = '$row$column value';
    }
  }
  print('Generating executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  var fileBytes = excel.encode();

  print('Encoding executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  if (fileBytes != null) {
    File(join("/Users/kawal/Desktop/r2.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
  print('Downloaded executed in ${stopwatch.elapsed}');
}
