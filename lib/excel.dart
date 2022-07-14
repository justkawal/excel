library excel;

import 'dart:convert';
import 'dart:math';
import 'dart:typed_data';
import 'package:archive/archive.dart';
import 'package:equatable/equatable.dart';
import 'package:xml/xml.dart';
import 'src/web_helper/client_save_excel.dart'
    if (dart.library.html) 'src/web_helper/web_save_excel_browser.dart'
    as helper;

/// main directory
part 'src/excel.dart';

/// sharedStrigns
part 'src/sharedStrings/shared_strings.dart';

/// Utilities
part 'src/utilities/span.dart';
part 'src/utilities/fast_list.dart';
part 'src/utilities/utility.dart';
part 'src/utilities/constants.dart';
part 'src/utilities/enum.dart';

/// Save
part 'src/save/save_file.dart';
part 'src/save/self_correct_span.dart';
part 'src/parser/parse.dart';

/// Sheet
part 'src/sheet/sheet.dart';
part 'src/sheet/font_family.dart';
part 'src/sheet/data_model.dart';
part 'src/sheet/formula.dart';
part 'src/sheet/cell_index.dart';
part 'src/sheet/cell_style.dart';
part 'src/sheet/font_style.dart';
part 'src/sheet/header_footer.dart';
