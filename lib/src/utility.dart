part of excel;

bool _isContain(dynamic d) {
  return (d ?? null) != null;
}

List<String> _noCompression = <String>[
  'mimetype',
  'Thumbnails/thumbnail.png',
];

String getCellId(int colI, int rowI) =>
    '${numericToLetters(colI + 1)}${rowI + 1}';

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0, mul = 1, n;
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
    return "$n";
  }
  return "0$n";
}

/// Convert a number to character based column
String numericToLetters(int number) {
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

// Normalize line
String _normalizeNewLine(String text) {
  return text.replaceAll('\r\n', '\n');
}

/// Returns the coordinates from a cell name.
/// "A2" returns [2, 1] and the "B3" return [3, 2].
List<int> cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart =
      utf8.decode(letters.where((rune) => rune > 0).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);

  return [
    int.parse(numericsPart) - 1,
    lettersToNumeric(lettersPart) - 1
  ]; // [x , y]
}

/// Throw error at situation where further more can't be processed as
/// important parts of excel are missing indicating it to be broken or corrupted.
_damagedExcel({String text}) {
  String t = '\nDamaged Excel file:';
  if (text != null) {
    t += ' $text';
  }
  throw ArgumentError(t + '\n');
}

/// return A2:B2 for spanning storage in unmerge list when [0,2] [2,2] is passed
String _getSpanCellId(
    int startColumn, int startRow, int endColumn, int endRow) {
  return '${getCellId(startColumn, startRow)}:${getCellId(endColumn, endRow)}';
}
