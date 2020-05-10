part of excel;

const String _relationshipsStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const String _relationshipsWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const String _relationshipsSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

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

/// Read and parse XSLX spreadsheet
class XlsxDecoder extends Excel {
  String get _mediaType {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }

  String get extension {
    return ".xlsx";
  }

  List<String> _rId;

  XlsxDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    _colorChanges = false;
    _mergeChanges = false;
    if (_update) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = <String, XmlDocument>{};
    }
    _worksheetTargets = <String, String>{};
    _xmlSheetId = <String, String>{};
    _cellStyleOther = <String, Map<String, CellStyle>>{};
    _cellStyleReferenced = <String, Map<String, int>>{};
    _fontColorHex = List<String>();
    _patternFill = List<String>();
    _tables = <String, DataTable>{};
    _sharedStrings = List<String>();
    _cellStyleList = List<CellStyle>();
    _innerCellStyle = List<CellStyle>();
    _rId = List<String>();
    _mergeChangeLook = List<String>();
    _spannedItems = <String, List<String>>{};
    _spanMap = <String, List<_Span>>{};
    _numFormats = List<int>();
    _putContentXml();
    _parseRelations();
    _parseStyles(_stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _parseMergedCells();
  }

  void updateCell(String sheet, CellIndex cellIndex, dynamic value,
      {String fontColorHex = "#000000",
      String backgroundColorHex = "none",
      TextWrapping wrap,
      VerticalAlign verticalAlign = VerticalAlign.Bottom,
      HorizontalAlign horizontalAlign = HorizontalAlign.Left}) {
    super.updateCell(sheet, cellIndex, value);
    int columnIndex = cellIndex._columnIndex;
    int rowIndex = cellIndex._rowIndex;

    String rC = '${numericToLetters(columnIndex + 1)}${rowIndex + 1}';

    if (fontColorHex != null ||
        backgroundColorHex != null ||
        wrap != null ||
        verticalAlign != VerticalAlign.Bottom ||
        horizontalAlign != HorizontalAlign.Left) {
      if (fontColorHex != null) {
        fontColorHex = isColorAppropriate(fontColorHex);
      }
      if (backgroundColorHex != null && backgroundColorHex != "none") {
        backgroundColorHex = isColorAppropriate(backgroundColorHex);
      }
      CellStyle cellStyle = CellStyle(
          fontColorHex: fontColorHex,
          backgroundColorHex: backgroundColorHex,
          horizontalAlign: horizontalAlign,
          verticalAlign: verticalAlign,
          textWrapping: wrap);

      if (_cellStyleOther.containsKey(sheet)) {
        _cellStyleOther[sheet][rC] = cellStyle;
      } else {
        _cellStyleOther[sheet] = {rC: cellStyle};
      }
      _colorChanges = true;
    }
  }

  String isColorAppropriate(String value) {
    String hex;
    if (value.length != 7) {
      throw ArgumentError(
          "InAppropriate Color provided. Use colorHex as example of: #FF0000");
    }
    hex = value.replaceAll(RegExp(r'#'), 'FF').toString();
    return hex;
  }

  _putContentXml() {
    var file = _archive.findFile("[Content_Types].xml");

    if (_xmlFiles != null) {
      if (file == null) {
        _damagedExcel();
      }
      file.decompress();
      _xmlFiles["[Content_Types].xml"] = parse(utf8.decode(file.content));
    }
  }

  _parseRelations() {
    var relations = _archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = parse(utf8.decode(relations.content));
      if (_xmlFiles != null) {
        _xmlFiles["xl/_rels/workbook.xml.rels"] = document;
      }
      document.findAllElements('Relationship').forEach((node) {
        String id = node.getAttribute('Id');
        switch (node.getAttribute('Type')) {
          case _relationshipsStyles:
            _stylesTarget = node.getAttribute('Target');
            break;
          case _relationshipsWorksheet:
            _worksheetTargets[id] = node.getAttribute('Target');
            break;
          case _relationshipsSharedStrings:
            _sharedStringsTarget = node.getAttribute('Target');
            break;
        }
        if (!_rId.contains(id)) {
          _rId.add(id);
        }
      });
    } else {
      _damagedExcel();
    }
  }

  _parseSharedStrings() {
    var sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    if (sharedStrings == null) {
      _sharedStringsTarget = 'sharedStrings.xml';

      // Running it with false will collect all the rid and will
      // help us to get the available rid to assign it to sharedStrings.xml
      _parseContent(run: false);

      if (_xmlFiles.containsKey("xl/_rels/workbook.xml.rels")) {
        int rIdNumber = _getAvailableRid();

        _xmlFiles["xl/_rels/workbook.xml.rels"]
            .findAllElements('Relationships')
            .first
            .children
            .add(XmlElement(
              XmlName('Relationship'),
              <XmlAttribute>[
                XmlAttribute(XmlName('Id'), 'rId$rIdNumber'),
                XmlAttribute(XmlName('Type'),
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),
                XmlAttribute(XmlName('Target'), 'sharedStrings.xml')
              ],
            ));
        if (!_rId.contains('rId$rIdNumber')) {
          _rId.add('rId$rIdNumber');
        }
        String content =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        bool contain = true;

        _xmlFiles["[Content_Types].xml"]
            .findAllElements('Override')
            .forEach((node) {
          var value = node.getAttribute('ContentType');
          if (value == content) {
            contain = false;
          }
        });
        if (contain) {
          _xmlFiles["[Content_Types].xml"]
              .findAllElements('Types')
              .first
              .children
              .add(XmlElement(
                XmlName('Override'),
                <XmlAttribute>[
                  XmlAttribute(XmlName('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(XmlName('ContentType'), content),
                ],
              ));
        }
      }

      var content = utf8.encode(
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>");
      _archive.addFile(
          ArchiveFile('xl/$_sharedStringsTarget', content.length, content));
      sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    }
    sharedStrings.decompress();
    var document = parse(utf8.decode(sharedStrings.content));
    if (_xmlFiles != null) {
      _xmlFiles["xl/$_sharedStringsTarget"] = document;
    }
    document.findAllElements('si').forEach((node) {
      _parseSharedString(node);
    });
  }

  _parseSharedString(XmlElement node) {
    var list = List();
    node.findAllElements('t').forEach((child) {
      list.add(_parseValue(child));
    });
    _sharedStrings.add(list.join(''));
  }

  _parseContent({bool run = true}) {
    var workbook = _archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
    }
    workbook.decompress();
    var document = parse(utf8.decode(workbook.content));
    if (_xmlFiles != null) {
      _xmlFiles["xl/workbook.xml"] = document;
    }
    document.findAllElements('sheet').forEach((node) {
      if (run) {
        _parseTable(node);
      } else {
        var rid = node.getAttribute('r:id');
        if (!_rId.contains(rid)) {
          _rId.add(rid);
        }
      }
    });
  }

  _parseMergedCells() {
    Map spannedCells = Map<String, List<String>>();
    _sheets.forEach((key, node) {
      XmlElement elementNode = node;
      List spanList = List<String>(),
          itemList = List<String>(),
          mapList = List<String>();

      if (spannedCells.containsKey(key) && spannedCells[key].isNotEmpty) {
        spanList = List<String>.from(spannedCells[key]);
      }
      if (_spannedItems.containsKey(key) && _spannedItems[key].isNotEmpty) {
        itemList = List<String>.from(_spannedItems[key]);
      }
      if (_spanMap.containsKey(key) && _spanMap[key].isNotEmpty) {
        mapList = List<String>.from(_spanMap[key]);
      }

      elementNode.findAllElements('mergeCell').forEach((elemen) {
        String ref = elemen.getAttribute('ref');
        if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
          if (!itemList.contains(ref)) {
            itemList.add(ref);
          }
          _spannedItems[key] = itemList;

          String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

          if (!spanList.contains(startCell)) {
            spanList.add(startCell);
          }
          spannedCells[key] = spanList;

          List<int> startIndex = cellCoordsFromCellId(startCell),
              endIndex = cellCoordsFromCellId(endCell);
          _Span spanObj = _Span();
          spanObj._start = [startIndex[0], startIndex[1]];
          spanObj._end = [endIndex[0], endIndex[1]];

          if (!mapList.contains(spanObj)) {
            mapList.add(spanObj);
          }
          _spanMap[key] = mapList;
          _mergeChangeLookup = key;
        }
      });
    });

    // Empty the elements of the tables if they are in merging area except the very first left cellId.
    _tables.keys.forEach((name) {
      if (spannedCells.containsKey(name)) {
        for (int row = 0; row < _tables[name].maxRows; row++) {
          for (int col = 0; col < _tables[name].maxCols; col++) {
            if (!(spannedCells[name].contains(getCellId(col, row)))) {
              _tables[name].rows[row][col] = null;
            }
          }
        }
      }
    });
  }
}
