part of excel;

List<String> _noCompression = <String>['mimetype', 'Thumbnails/thumbnail.png'];

String getCellId(int colI, int rowI) {
  return '${_numericToLetters(colI + 1)}${rowI + 1}';
}

String _isColorAppropriate(String value) {
  switch (value.length) {
    case 7:
      return value.replaceAll(RegExp(r'#'), 'FF');
    case 9:
      return value.replaceAll(RegExp(r'#'), '');
    default:
      return value;
  }
}

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0, mul = 1, n = 1;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    n = 1;
    if (65 <= c && c <= 90) {
      n += c - 65;
    } else if (97 <= c && c <= 122) {
      n += c - 97;
    }
    sum += n * mul;
    mul = mul * 26;
  }
  return sum;
}

Iterable<XmlElement> _findRows(XmlElement table) {
  return table.findElements('row');
}

Iterable<XmlElement> _findCells(XmlElement row) {
  return row.findElements('c');
}

int? _getCellNumber(XmlElement cell) {
  var r = cell.getAttribute('r');
  if (r == null) {
    return null;
  }
  return _cellCoordsFromCellId(r)[1];
}

int? _getRowNumber(XmlElement row) {
  return int.tryParse(row.getAttribute('r').toString());
}

int _checkPosition(List<CellStyle> list, CellStyle cellStyle) {
  return list.indexOf(cellStyle);
}

int _letterOnly(int rune) {
  if (65 <= rune && rune <= 90) {
    return rune;
  } else if (97 <= rune && rune <= 122) {
    return rune - 32;
  }
  return 0;
}

String _twoDigits(int n) {
  if (n > 9) {
    return '$n';
  }
  return '0$n';
}

/// Convert a number to character based column
String _numericToLetters(int number) {
  var letters = '';

  while (number != 0) {
    // Set remainder from 1..26
    var remainder = number % 26;

    if (remainder == 0) {
      remainder = 26;
    }

    // Convert the remainder to a character.
    var letter = String.fromCharCode(65 + remainder - 1);

    // Accumulate the column letters, right to left.
    letters = letter + letters;

    // Get the next order of magnitude.
    number = (number - 1) ~/ 26;
  }
  return letters;
}

/// Normalize line
String _normalizeNewLine(String text) {
  return text.replaceAll('\r\n', '\n');
}

///
///Returns the coordinates from a cell name.
///
///       cellCoordsFromCellId("A2"); // returns [2, 1]
///       cellCoordsFromCellId("B3"); // returns [3, 2]
///
///It is useful to convert CellId to Indexing.
///
List<int> _cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart = utf8.decode(letters.where((rune) {
    return rune > 0;
  }).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);

  return [
    int.parse(numericsPart) - 1,
    lettersToNumeric(lettersPart) - 1
  ]; // [x , y]
}

///
///Throw error at situation where further processing is not possible
///It is also called when important parts of excel files are missing as corrupted excel file is used
///
void _damagedExcel({String text = ''}) {
  throw ArgumentError('\nDamaged Excel file: $text\n');
}

///
///return A2:B2 for spanning storage in unmerge list when [0,2] [2,2] is passed
///
String getSpanCellId(int startColumn, int startRow, int endColumn, int endRow) {
  return '${getCellId(startColumn, startRow)}:${getCellId(endColumn, endRow)}';
}

///
///returns updated SpanObject location as there might be cross-sectional interaction between the two spanning objects.
///
List _isLocationChangeRequired(
    int startColumn, int startRow, int endColumn, int endRow, _Span spanObj) {
  bool changeValue = (
          // Overlapping checker
          startRow <= spanObj.rowSpanStart &&
              startColumn <= spanObj.columnSpanStart &&
              endRow >= spanObj.rowSpanEnd &&
              endColumn >= spanObj.columnSpanEnd)
      // first check starts here
      ||
      ( // outwards checking
          ((startColumn < spanObj.columnSpanStart &&
                      endColumn >= spanObj.columnSpanStart) ||
                  (startColumn <= spanObj.columnSpanEnd &&
                      endColumn > spanObj.columnSpanEnd))
              // inwards checking
              &&
              ((startRow >= spanObj.rowSpanStart &&
                      startRow <= spanObj.rowSpanEnd) ||
                  (endRow >= spanObj.rowSpanStart &&
                      endRow <= spanObj.rowSpanEnd)))

      // second check starts here
      ||
      (
          // outwards checking
          ((startRow < spanObj.rowSpanStart &&
                      endRow >= spanObj.rowSpanStart) ||
                  (startRow <= spanObj.rowSpanEnd &&
                      endRow > spanObj.rowSpanEnd))
              // inwards checking
              &&
              ((startColumn >= spanObj.columnSpanStart &&
                      startColumn <= spanObj.columnSpanEnd) ||
                  (endColumn >= spanObj.columnSpanStart &&
                      endColumn <= spanObj.columnSpanEnd)));

  if (changeValue) {
    if (startColumn > spanObj.columnSpanStart) {
      startColumn = spanObj.columnSpanStart;
    }
    if (endColumn < spanObj.columnSpanEnd) {
      endColumn = spanObj.columnSpanEnd;
    }
    if (startRow > spanObj.rowSpanStart) {
      startRow = spanObj.rowSpanStart;
    }
    if (endRow < spanObj.rowSpanEnd) {
      endRow = spanObj.rowSpanEnd;
    }
  }

  return List.from([
    changeValue,
    [startColumn, startRow, endColumn, endRow]
  ]);
}

///
///Returns Column based String alphabet when column index is passed
///
///     `getColumnAlphabet(0); // returns A`
///     `getColumnAlphabet(5); // returns F`
///
String getColumnAlphabet(int collIndex) {
  return '${_numericToLetters(collIndex + 1)}';
}

///
///Returns Column based int index when column alphabet is passed
///
///    `getColumnAlphabet("A"); // returns 0`
///    `getColumnAlphabet("F"); // returns 5`
///
int getColumnIndex(String columnAlphabet) {
  return _cellCoordsFromCellId('${columnAlphabet}')[1];
}

///
///Checks if the fontStyle is already present in the list or not
///
int _fontStyleIndex(List<_FontStyle> list, _FontStyle fontStyle) {
  return list.indexOf(fontStyle);
}
