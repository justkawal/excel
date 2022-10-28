part of excel;

class Save {
  final Excel _excel;
  late Map<String, ArchiveFile> _archiveFiles;
  late List<CellStyle> _innerCellStyle;
  final Parser parser;
  Save._(this._excel, this.parser) {
    _archiveFiles = <String, ArchiveFile>{};
    _innerCellStyle = <CellStyle>[];
  }

  List<int>? _save() {
    if (_excel._colorChanges) {
      _processStylesFile();
    }
    _setSheetElements();
    if (_excel._defaultSheet != null) {
      _setDefaultSheet(_excel._defaultSheet);
    }
    _setSharedStrings();

    if (_excel._mergeChanges) {
      _setMerge();
    }

    if (_excel._rtlChanges) {
      _setRTL();
    }

    for (var xmlFile in _excel._xmlFiles.keys) {
      var xml = _excel._xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_excel._archive));
  }

  Archive _cloneArchive(Archive archive) {
    var clone = Archive();
    archive.files.forEach((file) {
      if (file.isFile) {
        ArchiveFile copy;
        if (_archiveFiles.containsKey(file.name)) {
          copy = _archiveFiles[file.name]!;
        } else {
          var content = file.content as Uint8List;
          var compress = !_noCompression.contains(file.name);
          copy = ArchiveFile(file.name, content.length, content)
            ..compress = compress;
        }
        clone.addFile(copy);
      }
    });
    return clone;
  }

  bool _setDefaultSheet(String? sheetName) {
    if (sheetName == null || _excel._xmlFiles['xl/workbook.xml'] == null) {
      return false;
    }
    List<XmlElement> sheetList =
        _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').toList();
    XmlElement elementFound = XmlElement(XmlName(''));

    int position = -1;
    for (int i = 0; i < sheetList.length; i++) {
      var _sheetName = sheetList[i].getAttribute('name');
      if (_sheetName != null && _sheetName.toString() == sheetName) {
        elementFound = sheetList[i];
        position = i;
        break;
      }
    }

    if (position == -1) {
      return false;
    }
    if (position == 0) {
      return true;
    }

    _excel._xmlFiles['xl/workbook.xml']!
        .findAllElements('sheets')
        .first
        .children
      ..removeAt(position)
      ..insert(0, elementFound);

    String? expectedSheet = _excel._getDefaultSheet();

    return expectedSheet == sheetName;
  }

  /// Writing cell contained text into the excel sheet files.
  _setSheetElements() {
    _excel._sharedStrings = _SharedStringsMaintainer.instance;
    _excel._sharedStrings.clear();

    _excel._sheetMap.forEach((sheet, value) {
      ///
      /// Create the sheet's xml file if it does not exist.
      if (_excel._sheets[sheet] == null) {
        parser._createSheet(sheet);
      }

      /// Clear the previous contents of the sheet if it exists,
      /// in order to reduce the time to find and compare with the sheet rows
      /// and hence just do the work of putting the data only i.e. creating new rows
      if (_excel._sheets[sheet]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheet]!.children.clear();
      }

      _setColumnWidth(sheet);

      /// `Above function is important in order to wipe out the old contents of the sheet.`
      for (var rowIndex = 0; rowIndex < value._maxRows; rowIndex++) {
        if (value._sheetData[rowIndex] == null) {
          continue;
        }
        var foundRow =
            _createNewRow(_excel._sheets[sheet]! as XmlElement, rowIndex);
        for (var colIndex = 0; colIndex < value._maxCols; colIndex++) {
          var data = value._sheetData[rowIndex]![colIndex];
          if (data == null) {
            continue;
          }
          if (data.value != null) {
            _updateCell(sheet, foundRow, colIndex, rowIndex, data.value);
          }
        }
      }

      _setHeaderFooter(sheet);
    });
  }

  _setColumnWidth(String sheetName) {
    final sheetObject = _excel._sheetMap[sheetName];
    if (sheetObject == null) return;

    var xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final colElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColWidths.isEmpty &&
        sheetObject.getColAutoFits.isEmpty) {
      if (colElements.isEmpty) {
        return;
      }

      final cols = colElements.first;
      final worksheet = xmlFile.findAllElements('worksheet').first;
      worksheet.children.remove(cols);
      return;
    }

    if (colElements.isEmpty) {
      final worksheet = xmlFile.findAllElements('worksheet').first;
      final sheetData = xmlFile.findAllElements('sheetData').first;
      final index = worksheet.children.indexOf(sheetData);

      worksheet.children.insert(index, XmlElement(XmlName('cols'), [], []));
    }

    var cols = colElements.first;

    if (cols.children.isNotEmpty) {
      cols.children.clear();
    }

    final autoFits = sheetObject.getColAutoFits.asMap();
    final customWidths = sheetObject.getColWidths.asMap();

    final columnCount = max(autoFits.length, customWidths.length);

    List<double> colWidths = <double>[];
    int min = 0;

    for (var index = 0; index < columnCount; index++) {
      double value = _defaultColumnWidth;

      if (autoFits.containsKey(index) &&
          autoFits[index] == true &&
          (!customWidths.containsKey(index) ||
              customWidths[index] == _defaultColumnWidth)) {
        value = _calcAutoFitColWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          value = customWidths[index]!;
        }
      }

      colWidths.add(value);

      if (index != 0 && colWidths[index - 1] != value) {
        _addNewCol(cols, min, index - 1, colWidths[index - 1]);
        min = index;
      }

      if (index == (columnCount - 1)) {
        _addNewCol(cols, index, index, value);
      }
    }
  }

  void _addNewCol(XmlElement cols, int min, int max, double value) {
    cols.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), value.toStringAsFixed(2)),
      XmlAttribute(XmlName('bestFit'), "1"),
      XmlAttribute(XmlName('customWidth'), "1"),
    ], []));
  }

  double _calcAutoFitColWidth(Sheet sheet, int col) {
    var maxNumOfCharacters = 0;
    sheet._sheetData.forEach((key, value) {
      if (value.containsKey(col) && value[col]!._isFormula == false) {
        maxNumOfCharacters =
            max(value[col]!.value.toString().length, maxNumOfCharacters);
      }
    });

    return ((maxNumOfCharacters * 7.0 + 9.0) / 7.0 * 256).truncate() / 256;
  }

  _setRTL() {
    _excel._rtlChangeLook.forEach((s) {
      var sheetObject = _excel._sheetMap[s];
      if (sheetObject != null &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        var itrSheetViewsRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('sheetViews');

        if (itrSheetViewsRTLElement?.isNotEmpty ?? false) {
          var itrSheetViewRTLElement = _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetView');

          if (itrSheetViewRTLElement?.isNotEmpty ?? false) {
            /// clear all the children of the sheetViews here

            _excel._xmlFiles[_excel._xmlSheetId[s]]
                ?.findAllElements('sheetViews')
                .first
                .children
                .clear();
          }

          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('sheetViews')
              .first
              .children
              .add(XmlElement(
                XmlName('sheetView'),
                [
                  if (sheetObject.isRTL)
                    XmlAttribute(XmlName('rightToLeft'), '1'),
                  XmlAttribute(XmlName('workbookViewId'), '0'),
                ],
              ));
        } else {
          _excel._xmlFiles[_excel._xmlSheetId[s]]
              ?.findAllElements('worksheet')
              .first
              .children
              .add(XmlElement(XmlName('sheetViews'), [], [
                XmlElement(
                  XmlName('sheetView'),
                  [
                    if (sheetObject.isRTL)
                      XmlAttribute(XmlName('rightToLeft'), '1'),
                    XmlAttribute(XmlName('workbookViewId'), '0'),
                  ],
                )
              ]));
        }
      }
    });
  }

  /// Writing the merged cells information into the excel properties files.
  _setMerge() {
    _selfCorrectSpanMap(_excel);
    _excel._mergeChangeLook.forEach((s) {
      if (_excel._sheetMap[s] != null &&
          _excel._sheetMap[s]!._spanList.isNotEmpty &&
          _excel._xmlSheetId.containsKey(s) &&
          _excel._xmlFiles.containsKey(_excel._xmlSheetId[s])) {
        Iterable<XmlElement>? iterMergeElement = _excel
            ._xmlFiles[_excel._xmlSheetId[s]]
            ?.findAllElements('mergeCells');
        late XmlElement mergeElement;
        if (iterMergeElement?.isNotEmpty ?? false) {
          mergeElement = iterMergeElement!.first;
        } else {
          if ((_excel._xmlFiles[_excel._xmlSheetId[s]]
                      ?.findAllElements('worksheet')
                      .length ??
                  0) >
              0) {
            int index = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .indexOf(_excel._xmlFiles[_excel._xmlSheetId[s]]!
                    .findAllElements("sheetData")
                    .first);
            if (index == -1) {
              _damagedExcel();
            }
            _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('worksheet')
                .first
                .children
                .insert(
                    index + 1,
                    XmlElement(XmlName('mergeCells'),
                        [XmlAttribute(XmlName('count'), '0')]));

            mergeElement = _excel._xmlFiles[_excel._xmlSheetId[s]]!
                .findAllElements('mergeCells')
                .first;
          } else {
            _damagedExcel();
          }
        }

        List<String> _spannedItems =
            List<String>.from(_excel._sheetMap[s]!.spannedItems);

        [
          ['count', _spannedItems.length.toString()],
        ].forEach((value) {
          if (mergeElement.getAttributeNode(value[0]) == null) {
            mergeElement.attributes
                .add(XmlAttribute(XmlName(value[0]), value[1]));
          } else {
            mergeElement.getAttributeNode(value[0])!.value = value[1];
          }
        });

        mergeElement.children.clear();

        _spannedItems.forEach((value) {
          mergeElement.children.add(XmlElement(XmlName('mergeCell'),
              [XmlAttribute(XmlName('ref'), '$value')], []));
        });
      }
    });
  }

  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.

  _processStylesFile() {
    _innerCellStyle = <CellStyle>[];
    List<String> innerPatternFill = <String>[];
    List<_FontStyle> innerFontStyle = <_FontStyle>[];

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      sheetObject._sheetData.forEach((_, colMap) {
        colMap.forEach((_, dataObject) {
          if (dataObject.cellStyle != null) {
            int pos = _checkPosition(_innerCellStyle, dataObject.cellStyle!);
            if (pos == -1) {
              _innerCellStyle.add(dataObject.cellStyle!);
            }
          }
        });
      });
    });

    _innerCellStyle.forEach((cellStyle) {
      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily);

      /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fs`
      if (_fontStyleIndex(_excel._fontStyleList, _fs) == -1 &&
          _fontStyleIndex(innerFontStyle, _fs) == -1) {
        innerFontStyle.add(_fs);
      }

      /// Filling the inner usable extra list of backgroung color
      String backgroundColor = cellStyle.backgroundColor;
      if (!_excel._patternFill.contains(backgroundColor) &&
          !innerPatternFill.contains(backgroundColor)) {
        innerPatternFill.add(backgroundColor);
      }
    });

    XmlElement fonts =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fonts').first;

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value =
          '${_excel._fontStyleList.length + innerFontStyle.length}';
    } else {
      fonts.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._fontStyleList.length + innerFontStyle.length}'));
    }

    innerFontStyle.forEach((fontStyleElement) {
      fonts.children.add(XmlElement(XmlName('font'), [], [
        /// putting color
        if (fontStyleElement._fontColorHex != null &&
            fontStyleElement._fontColorHex != "FF000000")
          XmlElement(
              XmlName('color'),
              [XmlAttribute(XmlName('rgb'), fontStyleElement._fontColorHex!)],
              []),

        /// putting bold
        if (fontStyleElement.isBold) XmlElement(XmlName('b'), [], []),

        /// putting italic
        if (fontStyleElement.isItalic) XmlElement(XmlName('i'), [], []),

        /// putting single underline
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline == Underline.Single)
          XmlElement(XmlName('u'), [], []),

        /// putting double underline
        if (fontStyleElement.underline != Underline.None &&
            fontStyleElement.underline != Underline.Single &&
            fontStyleElement.underline == Underline.Double)
          XmlElement(
              XmlName('u'), [XmlAttribute(XmlName('val'), 'double')], []),

        /// putting fontFamily
        if (fontStyleElement.fontFamily != null &&
            fontStyleElement.fontFamily!.toLowerCase().toString() != 'null' &&
            fontStyleElement.fontFamily != '' &&
            fontStyleElement.fontFamily!.isNotEmpty)
          XmlElement(XmlName('name'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontFamily.toString())
          ], []),

        /// putting fontSize
        if (fontStyleElement.fontSize != null &&
            fontStyleElement.fontSize.toString().isNotEmpty)
          XmlElement(XmlName('sz'), [
            XmlAttribute(XmlName('val'), fontStyleElement.fontSize.toString())
          ], []),
      ]));
    });

    XmlElement fills =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fills').first;

    var fillAttribute = fills.getAttributeNode('count');

    if (fillAttribute != null) {
      fillAttribute.value =
          '${_excel._patternFill.length + innerPatternFill.length}';
    } else {
      fills.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._patternFill.length + innerPatternFill.length}'));
    }

    innerPatternFill.forEach((color) {
      if (color.length >= 2) {
        if (color.substring(0, 2).toUpperCase() == 'FF') {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'), [
              XmlAttribute(XmlName('patternType'), 'solid')
            ], [
              XmlElement(XmlName('fgColor'),
                  [XmlAttribute(XmlName('rgb'), color)], []),
              XmlElement(
                  XmlName('bgColor'), [XmlAttribute(XmlName('rgb'), color)], [])
            ])
          ]));
        } else if (color == "none" ||
            color == "gray125" ||
            color == "lightGray") {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'),
                [XmlAttribute(XmlName('patternType'), color)], [])
          ]));
        }
      } else {
        _damagedExcel(
            text:
                "Corrupted Styles Found. Can't process further, Open up issue in github.");
      }
    });

    XmlElement celx =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('cellXfs').first;
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value =
          '${_excel._cellStyleList.length + _innerCellStyle.length}';
    } else {
      celx.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._cellStyleList.length + _innerCellStyle.length}'));
    }

    _innerCellStyle.forEach((cellStyle) {
      String backgroundColor = cellStyle.backgroundColor;

      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily);

      HorizontalAlign horizontalALign = cellStyle.horizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.verticalAlignment;
      int rotation = cellStyle.rotation;
      TextWrapping? textWrapping = cellStyle.wrap;
      int backgroundIndex = innerPatternFill.indexOf(backgroundColor),
          fontIndex = _fontStyleIndex(innerFontStyle, _fs);

      var attributes = <XmlAttribute>[
        XmlAttribute(XmlName('borderId'), '0'),
        XmlAttribute(XmlName('fillId'),
            '${backgroundIndex == -1 ? 0 : backgroundIndex + _excel._patternFill.length}'),
        XmlAttribute(XmlName('fontId'),
            '${fontIndex == -1 ? 0 : fontIndex + _excel._fontStyleList.length}'),
        XmlAttribute(XmlName('numFmtId'), '0'),
        XmlAttribute(XmlName('xfId'), '0'),
      ];

      if ((_excel._patternFill.contains(backgroundColor) ||
              innerPatternFill.contains(backgroundColor)) &&
          backgroundColor != "none" &&
          backgroundColor != "gray125" &&
          backgroundColor.toLowerCase() != "lightgray") {
        attributes.add(XmlAttribute(XmlName('applyFill'), '1'));
      }

      if (_fontStyleIndex(_excel._fontStyleList, _fs) != -1 &&
          _fontStyleIndex(innerFontStyle, _fs) != -1) {
        attributes.add(XmlAttribute(XmlName('applyFont'), '1'));
      }

      var children = <XmlElement>[];

      if (horizontalALign != HorizontalAlign.Left ||
          textWrapping != null ||
          verticalAlign != VerticalAlign.Bottom ||
          rotation != 0) {
        attributes.add(XmlAttribute(XmlName('applyAlignment'), '1'));
        var childAttributes = <XmlAttribute>[];

        if (textWrapping != null) {
          childAttributes.add(XmlAttribute(
              XmlName(textWrapping == TextWrapping.Clip
                  ? 'shrinkToFit'
                  : 'wrapText'),
              '1'));
        }

        if (verticalAlign != VerticalAlign.Bottom) {
          String ver = verticalAlign == VerticalAlign.Top ? 'top' : 'center';
          childAttributes.add(XmlAttribute(XmlName('vertical'), '$ver'));
        }

        if (horizontalALign != HorizontalAlign.Left) {
          String hor =
              horizontalALign == HorizontalAlign.Right ? 'right' : 'center';
          childAttributes.add(XmlAttribute(XmlName('horizontal'), '$hor'));
        }
        if (rotation != 0) {
          childAttributes
              .add(XmlAttribute(XmlName('textRotation'), '$rotation'));
        }

        children.add(XmlElement(XmlName('alignment'), childAttributes, []));
      }

      celx.children.add(XmlElement(XmlName('xf'), attributes, children));
    });
  }

  /// Writing the value of excel cells into the separate
  /// sharedStrings file so as to minimize the size of excel files.
  _setSharedStrings() {
    var uniqueCount = 0;
    var count = 0;

    XmlElement shareString = _excel
        ._xmlFiles['xl/${_excel._sharedStringsTarget}']!
        .findAllElements('sst')
        .first;

    shareString.children.clear();

    _excel._sharedStrings._map.forEach((string, ss) {
      uniqueCount += 1;
      count += ss.count;

      shareString.children.add(string.node);
    });

    [
      ['count', '$count'],
      ['uniqueCount', '$uniqueCount']
    ].forEach((value) {
      if (shareString.getAttributeNode(value[0]) == null) {
        shareString.attributes.add(XmlAttribute(XmlName(value[0]), value[1]));
      } else {
        shareString.getAttributeNode(value[0])!.value = value[1];
      }
    });
  }

  // slow implementation
  /*XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = 0;
    for (var currentRow in rows) {
      currentIndex = _getRowNumber(currentRow) - 1;
      if (currentIndex >= rowIndex) {
        row = currentRow;
        break;
      }
    }

    // Create row if required
    if (row == null || currentIndex != rowIndex) {
      row = __insertRow(table, row, rowIndex);
    }

    return row;
  }
  
  XmlElement _createRow(int rowIndex) {
    return XmlElement(XmlName('row'),
        [XmlAttribute(XmlName('r'), (rowIndex + 1).toString())], []);
  } 
  
  XmlElement __insertRow(XmlElement table, XmlElement lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null) {
      table.children.add(row);
    } else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }*/

  ///
  XmlElement _createNewRow(XmlElement table, int rowIndex) {
    var row = XmlElement(XmlName('row'),
        [XmlAttribute(XmlName('r'), (rowIndex + 1).toString())], []);
    table.children.add(row);
    return row;
  }

/*   XmlElement _replaceCell(String sheet, XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var index = lastCell == null ? 0 : row.children.indexOf(lastCell);
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  } */

  // Manage value's type
  XmlElement _createCell(
      String sheet, int columnIndex, int rowIndex, dynamic value) {
    if (value is SharedString) {
      _excel._sharedStrings.add(value);
    }

    String rC = getCellId(columnIndex, rowIndex);

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      if (value is SharedString) XmlAttribute(XmlName('t'), 's'),
    ];

    if (_excel._colorChanges &&
        (_excel._sheetMap[sheet]?._sheetData != null) &&
        _excel._sheetMap[sheet]!._sheetData[rowIndex] != null &&
        _excel._sheetMap[sheet]!._sheetData[rowIndex]![columnIndex]
                ?.cellStyle !=
            null) {
      CellStyle cellStyle = _excel
          ._sheetMap[sheet]!._sheetData[rowIndex]![columnIndex]!.cellStyle!;
      int upperLevelPos = _checkPosition(_excel._cellStyleList, cellStyle);
      if (upperLevelPos == -1) {
        int lowerLevelPos = _checkPosition(_innerCellStyle, cellStyle);
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _excel._cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      attributes.insert(
        1,
        XmlAttribute(XmlName('s'), '$upperLevelPos'),
      );
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(
            XmlName('s'), '${_excel._cellStyleReferenced[sheet]![rC]}'),
      );
    }

    var children = value == null
        ? <XmlElement>[]
        : <XmlElement>[
            if (value is Formula)
              XmlElement(XmlName('f'), [], [XmlText(value.formula.toString())]),
            XmlElement(XmlName('v'), [], [
              XmlText(value is SharedString
                  ? _excel._sharedStrings.indexOf(value).toString()
                  : value is Formula
                      ? ''
                      : value.toString())
            ]),
          ];
    return XmlElement(XmlName('c'), attributes, children);
  }

// slow implementation
/*   XmlElement _updateCell(String sheet, XmlElement node, int columnIndex,
      int rowIndex, dynamic value) {
    XmlElement cell;
    var cells = _findCells(node);

    var currentIndex = 0; // cells could be empty
    for (var currentCell in cells) {
      currentIndex = _getCellNumber(currentCell);
      if (currentIndex >= columnIndex) {
        cell = currentCell;
        break;
      }
    }

    if (cell == null || currentIndex != columnIndex) {
      cell = _insertCell(sheet, node, cell, columnIndex, rowIndex, value);
    } else {
      cell = _replaceCell(sheet, node, cell, columnIndex, rowIndex, value);
    }

    return cell;
  } */
  XmlElement _updateCell(String sheet, XmlElement row, int columnIndex,
      int rowIndex, dynamic value) {
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    row.children.add(cell);
    return cell;
  }

  void _setHeaderFooter(String sheetName) {
    final sheet = _excel._sheetMap[sheetName];
    if (sheet == null) return;

    final xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
    if (xmlFile == null) return;

    final sheetXmlElement = xmlFile.findAllElements("worksheet").first;

    final results = sheetXmlElement.findAllElements("headerFooter");
    if (results.isNotEmpty) {
      sheetXmlElement.children.remove(results.first);
    }

    if (sheet.headerFooter == null) return;

    sheetXmlElement.children.add(sheet.headerFooter!.toXmlElement());
  }
}
