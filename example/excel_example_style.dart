import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  //var file = "/Users/kawal/Desktop/excel/test/test_resources/example.xlsx";
  //var bytes = File(file).readAsBytesSync();
  var excel = Excel.createExcel();
  // or
  //var excel = Excel.decodeBytes(bytes);

  var sheet = excel['Sheet1'];

  sheet.appendRow([
    IntCellValue(8),
    DoubleCellValue(999.62221),
    DateCellValue(
      year: DateTime.now().year,
      month: DateTime.now().month,
      day: DateTime.now().day,
    ),
    DateTimeCellValue.fromDateTime(DateTime.now()),
  ]);

  // Saving the file

  // String outputFile = "/Users/kawal/Desktop/git_projects/r.xlsx";
  String outputFile = './example/example.xlsx';

  //stopwatch.reset();
  List<int>? fileBytes = excel.save();
  //print('saving executed in ${stopwatch.elapsed}');
  if (fileBytes != null) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
