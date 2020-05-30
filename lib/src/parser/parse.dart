part of excel;

class Parse {
  Excel _excel;
  Parse._(Excel excel) {
    this._excel = excel;
  }

  _startParsing() {
    _putContentXml();
    _parseRelations();
    _parseStyles(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _parseMergedCells();
  }

  _normalizeTable(DataTable table) {
    if (table._maxRows == 0) {
      table._rows.clear();
    } else if (table._maxRows < table._rows.length) {
      table._rows.removeRange(table._maxRows, table._rows.length);
    }

    for (var row = 0; row < table._rows.length; row++) {
      if (table._maxCols == 0) {
        table._rows[row].clear();
      } else if (table._maxCols < table._rows[row].length) {
        table._rows[row].removeRange(table._maxCols, table._rows[row].length);
      } else if (table._maxCols > table._rows[row].length) {
        var repeat = table._maxCols - table._rows[row].length;
        for (var index = 0; index < repeat; index++) {
          table._rows[row].add(null);
        }
      }
    }
  }

  _putContentXml() {
    var file = _excel._archive.findFile("[Content_Types].xml");

    if (_excel._xmlFiles != null) {
      if (file == null) {
        _damagedExcel();
      }
      file.decompress();
      _excel._xmlFiles["[Content_Types].xml"] =
          parse(utf8.decode(file.content));
    }
  }

  _parseRelations() {
    var relations = _excel._archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = parse(utf8.decode(relations.content));
      if (_excel._xmlFiles != null) {
        _excel._xmlFiles["xl/_rels/workbook.xml.rels"] = document;
      }
      document.findAllElements('Relationship').forEach((node) {
        String id = node.getAttribute('Id');
        switch (node.getAttribute('Type')) {
          case _relationshipsStyles:
            _excel._stylesTarget = node.getAttribute('Target');
            break;
          case _relationshipsWorksheet:
            _excel._worksheetTargets[id] = node.getAttribute('Target');
            break;
          case _relationshipsSharedStrings:
            _excel._sharedStringsTarget = node.getAttribute('Target');
            break;
        }
        if (!_excel._rId.contains(id)) {
          _excel._rId.add(id);
        }
      });
    } else {
      _damagedExcel();
    }
  }

  _parseSharedStrings() {
    var sharedStrings =
        _excel._archive.findFile('xl/${_excel._sharedStringsTarget}');
    if (sharedStrings == null) {
      _excel._sharedStringsTarget = 'sharedStrings.xml';

      // Running it with false will collect all the rid and will
      // help us to get the available rid to assign it to sharedStrings.xml
      _parseContent(run: false);

      if (_excel._xmlFiles.containsKey("xl/_rels/workbook.xml.rels")) {
        int rIdNumber = _excel._getAvailableRid();

        _excel._xmlFiles["xl/_rels/workbook.xml.rels"]
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
        if (!_excel._rId.contains('rId$rIdNumber')) {
          _excel._rId.add('rId$rIdNumber');
        }
        String content =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        bool contain = true;

        _excel._xmlFiles["[Content_Types].xml"]
            .findAllElements('Override')
            .forEach((node) {
          var value = node.getAttribute('ContentType');
          if (value == content) {
            contain = false;
          }
        });
        if (contain) {
          _excel._xmlFiles["[Content_Types].xml"]
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
      _excel._archive.addFile(ArchiveFile(
          'xl/${_excel._sharedStringsTarget}', content.length, content));
      sharedStrings =
          _excel._archive.findFile('xl/${_excel._sharedStringsTarget}');
    }
    sharedStrings.decompress();
    var document = parse(utf8.decode(sharedStrings.content));
    if (_excel._xmlFiles != null) {
      _excel._xmlFiles["xl/$_excel._sharedStringsTarget}"] = document;
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
    _excel._sharedStrings.add(list.join(''));
  }

  _parseContent({bool run = true}) {
    var workbook = _excel._archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
    }
    workbook.decompress();
    var document = parse(utf8.decode(workbook.content));
    if (_excel._xmlFiles != null) {
      _excel._xmlFiles["xl/workbook.xml"] = document;
    }
    document.findAllElements('sheet').forEach((node) {
      if (run) {
        _parseTable(node);
      } else {
        var rid = node.getAttribute('r:id');
        if (!_excel._rId.contains(rid)) {
          _excel._rId.add(rid);
        }
      }
    });
  }

  _parseMergedCells() {
    Map spannedCells = Map<String, List<String>>();
    _excel._sheets.forEach((sheetName, node) {
      _excel._availSheet(sheetName);
      XmlElement elementNode = node;
      List spanList = List<String>();

      elementNode.findAllElements('mergeCell').forEach((elemen) {
        String ref = elemen.getAttribute('ref');
        if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
          if (!_excel._sheetMap['$sheetName']._spannedItems.contains(ref)) {
            _excel._sheetMap['$sheetName']._spannedItems.add(ref);
          }

          String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

          if (!spanList.contains(startCell)) {
            spanList.add(startCell);
          }
          spannedCells[sheetName] = spanList;

          List<int> startIndex = _cellCoordsFromCellId(startCell),
              endIndex = _cellCoordsFromCellId(endCell);
          _Span spanObj = _Span();
          spanObj._start = [startIndex[0], startIndex[1]];
          spanObj._end = [endIndex[0], endIndex[1]];
          if (!_excel._sheetMap['$sheetName']._spanList.contains(spanObj)) {
            _excel._sheetMap['$sheetName']._spanList.add(spanObj);
          }
          _excel._mergeChangeLookup = sheetName;
        }
      });
    });

    // Remove those cells which are present inside the
    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (spannedCells.containsKey(sheetName)) {
        sheetObject._sheetData.forEach((row, colMap) {
          colMap.forEach((col, dataObject) {
            if (!(spannedCells[sheetName].contains(getCellId(col, row)))) {
              _excel['$sheetName']._sheetData[row].remove(col);
            }
          });
        });
      }
    });
  }

  // Reading the styles from the excel file.
  _parseStyles(String _stylesTarget) {
    var styles = _excel._archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = parse(utf8.decode(styles.content));
      if (_excel._xmlFiles != null) {
        _excel._xmlFiles['xl/$_stylesTarget'] = document;
      }
      _excel._fontColorHex = List<String>();
      _excel._patternFill = List<String>();
      _excel._cellStyleList = List<CellStyle>();
      int fontIndex = 0;
      document
          .findAllElements('font')
          .first
          .findElements('color')
          .forEach((child) {
        var colorHex = child.getAttribute('rgb');
        if (colorHex != null &&
            !_excel._fontColorHex.contains(colorHex.toString())) {
          _excel._fontColorHex.add(colorHex.toString());
          fontIndex = 1;
        } else if (fontIndex == 0 &&
            !_excel._fontColorHex.contains("FF000000")) {
          _excel._fontColorHex.add("FF000000");
        }
      });
      document.findAllElements('patternFill').forEach((node) {
        String patternType = node.getAttribute('patternType').toString(), rgb;
        if (node.children.isNotEmpty) {
          node.findElements('fgColor').forEach((child) {
            rgb = node.getAttribute('rgb').toString();
            _excel._patternFill.add(rgb);
          });
        } else {
          _excel._patternFill.add(patternType);
        }
      });

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          _excel._numFormats.add(_getFontIndex(node, 'numFmtId'));

          String fontColor = "FF000000", backgroundColor = "none";
          HorizontalAlign horizontalAlign = HorizontalAlign.Left;
          VerticalAlign verticalAlign = VerticalAlign.Bottom;
          TextWrapping textWrapping;
          int fontId = _getFontIndex(node, 'fontId');
          if (fontId < _excel._fontColorHex.length) {
            fontColor = _excel._fontColorHex[fontId];
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _excel._patternFill.length) {
            backgroundColor = _excel._patternFill[fillId];
          }

          if (node.children.isNotEmpty) {
            node.findElements('alignment').forEach((child) {
              if (_getFontIndex(child, 'wrapText') == 1) {
                textWrapping = TextWrapping.WrapText;
              } else if (_getFontIndex(child, 'shrinkToFit') == 1) {
                textWrapping = TextWrapping.Clip;
              }

              var vertical = node.getAttribute('vertical');
              if (vertical != null) {
                if (vertical.toString() == 'top') {
                  verticalAlign = VerticalAlign.Top;
                } else if (vertical.toString() == 'center') {
                  verticalAlign = VerticalAlign.Center;
                }
              }

              var horizontal = node.getAttribute('horizontal');
              if (horizontal != null) {
                if (horizontal.toString() == 'center') {
                  horizontalAlign = HorizontalAlign.Center;
                } else if (horizontal.toString() == 'right') {
                  horizontalAlign = HorizontalAlign.Right;
                }
              }
            });
          }

          CellStyle cellStyle = CellStyle(
              fontColorHex: fontColor,
              backgroundColorHex: backgroundColor,
              horizontalAlign: horizontalAlign,
              verticalAlign: verticalAlign,
              textWrapping: textWrapping);

          _excel._cellStyleList.add(cellStyle);
        });
      });
    } else {
      _damagedExcel(text: 'styles');
    }
  }

  int _getFontIndex(var node, String text) {
    int applyFontInt = 0;
    var applyFont = node.getAttribute(text);
    if (applyFont != null) {
      try {
        applyFontInt = int.parse(applyFont.toString());
      } catch (_) {}
    }
    return applyFontInt;
  }

  _parseTable(XmlElement node) {
    var name = node.getAttribute('name');
    var target = _excel._worksheetTargets[node.getAttribute('r:id')];

    _excel.tables[name] = DataTable(name);
    var table = _excel.tables[name];

    var file = _excel._archive.findFile('xl/$target');
    file.decompress();

    var content = parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, table, name);
    });

    if (_excel._update) {
      _excel._sheets[name] = sheet;
      _excel._xmlFiles['xl/$target'] = content;
    }
    _excel._xmlSheetId[name] = 'xl/$target';

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, DataTable table, String name) {
    var row = List();

    _findCells(node).forEach((child) {
      _parseCell(child, table, row, name);
    });

    var rowIndex = _getRowNumber(node) - 1;
    if (_isNotEmptyRow(row) && rowIndex > table._rows.length) {
      var repeat = rowIndex - table._rows.length;
      for (var index = 0; index < repeat; index++) {
        table._rows.add(List());
      }
    }

    if (_isNotEmptyRow(row)) {
      table._rows.add(row);
    } else {
      table._rows.add(List());
    }

    _countFilledRow(table, row);
  }

  _countFilledRow(DataTable table, List row) {
    if (_isNotEmptyRow(row) && table._maxRows < table._rows.length) {
      table._maxRows = table._rows.length;
    }
  }

  _countFilledColumn(DataTable table, List row, dynamic value) {
    if (value != null && table._maxCols < row.length) {
      table._maxCols = row.length;
    }
  }

  _parseCell(XmlElement node, DataTable table, List row, String name) {
    var colIndex = _getCellNumber(node);
    if (colIndex > row.length) {
      var repeat = colIndex - row.length;
      for (var index = 0; index < repeat; index++) {
        row.add(null);
      }
    }

    var s1 = node.getAttribute('s');
    int s = 0;
    if (s1 != null) {
      try {
        s = int.parse(s1.toString());
      } catch (_) {}

      String rC = node.getAttribute('r').toString();

      if (_excel._cellStyleReferenced.containsKey(name)) {
        _excel._cellStyleReferenced[name][rC] = s;
      } else {
        _excel._cellStyleReferenced[name] = {rC: s};
      }
    }

    if (node.children.isEmpty) {
      return;
    }

    var value, type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        value = _excel._sharedStrings[
            int.parse(_parseValue(node.findElements('v').first))];
        break;
      // boolean
      case 'b':
        value = _parseValue(node.findElements('v').first) == '1';
        break;
      // error
      case 'e':
      // formula
      case 'str':
        value = _parseValue(node.findElements('v').first);
        break;
      // inline string
      case 'inlineStr':
        // <c r='B2' t='inlineStr'>
        // <is><t>Dartonico</t></is>
        // </c>
        value = _parseValue(node.findAllElements('t').first);
        break;
      // number
      case 'n':
      default:
        var valueNode = node.findElements('v');
        var content = valueNode.first;
        if (s1 != null) {
          var fmtId = _excel._numFormats[s];
          // date
          if (((fmtId >= 14) && (fmtId <= 17)) || (fmtId == 22)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(1899, 12, 30);
            value = date
                .add(Duration(milliseconds: delta.toInt()))
                .toIso8601String();
            // time
          } else if (((fmtId >= 18) && (fmtId <= 21)) ||
              ((fmtId >= 45) && (fmtId <= 47))) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(0);
            date = date.add(Duration(milliseconds: delta.toInt()));
            value =
                '${_twoDigits(date.hour)}:${_twoDigits(date.minute)}:${_twoDigits(date.second)}';
            // number
          } else {
            value = num.parse(_parseValue(content));
          }
        } else {
          value = num.parse(_parseValue(content));
        }
    }
    row.add(value);
    if (!_excel._sharedStrings.contains('$value')) {
      _excel._sharedStrings.add('$value');
    }

    _countFilledColumn(table, row, value);
  }

  _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
      }
    });

    return buffer.toString();
  }
}
