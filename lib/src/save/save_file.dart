part of excel;

class Save {
  final Excel _excel;
  late Map<String, ArchiveFile> _archiveFiles;
  late List<CellStyle> _innerCellStyle;
  late List<CellStyle> _mergedCellStyle;

  final Parser parser;
  Save._(this._excel, this.parser) {
    _archiveFiles = <String, ArchiveFile>{};
    _innerCellStyle = <CellStyle>[];
  }

  void _addNewCol(XmlElement cols, int min, int max, double width) {
    cols.children.add(XmlElement(XmlName('col'), [
      XmlAttribute(XmlName('min'), (min + 1).toString()),
      XmlAttribute(XmlName('max'), (max + 1).toString()),
      XmlAttribute(XmlName('width'), width.toStringAsFixed(2)),
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
  XmlNode _createCell(String sheet, XmlElement row, int columnIndex,
      int rowIndex, dynamic value,
      [bool addToRow = true]) {
    if (value is SharedString) {
      _excel._sharedStrings.add(value);
    }

    String rC = getCellId(columnIndex, rowIndex);

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      if (value is SharedString) XmlAttribute(XmlName('t'), 's'),
    ];

    if (_excel._mergedCellStyleReferenced.containsKey(sheet) &&
        _excel._mergedCellStyleReferenced[sheet]!.containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(
            XmlName('s'), '${_excel._mergedCellStyleReferenced[sheet]![rC]}'),
      );
    } else if (_excel._cellStyleReferenced.containsKey(sheet) &&
        _excel._cellStyleReferenced[sheet]!.containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(
            XmlName('s'), '${_excel._cellStyleReferenced[sheet]![rC]}'),
      );
    } else if (_excel._styleChanges &&
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

    XmlNode cell = XmlElement(XmlName('c'), attributes, children);
    if (addToRow) row.children.add(cell);
    return cell;
  }

  /// Create a new row in the sheet.
  XmlElement _createNewRow(XmlElement table, int rowIndex, double? height) {
    var row = XmlElement(XmlName('row'), [
      XmlAttribute(XmlName('r'), (rowIndex + 1).toString()),
      if (height != null)
        XmlAttribute(XmlName('ht'), height.toStringAsFixed(2)),
      if (height != null) XmlAttribute(XmlName('customHeight'), '1'),
    ], []);
    table.children.add(row);
    return row;
  }

  void _setMergeCellsStyles(String sheetName, Sheet sheetObject) {
    List<List<CellIndex>> _mergedCells = sheetObject.spannedItems
        .map(
          (e) => e.split(":").map((e) => CellIndex.indexByString(e)).toList(),
        )
        .toList();

    for (var i = 0; i < _mergedCells.length; i++) {
      CellIndex _cellIndexStart = _mergedCells[i][0];
      CellIndex _cellIndexEnd = _mergedCells[i][1];

      Data? mergedCellData = sheetObject._sheetData[_cellIndexStart.rowIndex]
          ?[_cellIndexStart.columnIndex];

      CellStyle? mergedCellStyle = mergedCellData?.cellStyle;

      if (mergedCellStyle != null) {
        bool hasBorder = mergedCellStyle.topBorder != Border() ||
            mergedCellStyle.bottomBorder != Border() ||
            mergedCellStyle.leftBorder != Border() ||
            mergedCellStyle.rightBorder != Border() ||
            mergedCellStyle.diagonalBorderUp ||
            mergedCellStyle.diagonalBorderDown;
        if (hasBorder) {
          for (var j = _cellIndexStart.rowIndex;
              j <= _cellIndexEnd.rowIndex;
              j++) {
            for (var k = _cellIndexStart.columnIndex;
                k <= _cellIndexEnd.columnIndex;
                k++) {
              CellStyle cellStyle = mergedCellStyle.copyWith(
                topBorderVal: Border(),
                bottomBorderVal: Border(),
                leftBorderVal: Border(),
                rightBorderVal: Border(),
                diagonalBorderUpVal: false,
                diagonalBorderDownVal: false,
              );

              if (j == _cellIndexStart.rowIndex) {
                cellStyle = cellStyle.copyWith(
                  topBorderVal: mergedCellStyle.topBorder,
                );
              }
              if (j == _cellIndexEnd.rowIndex) {
                cellStyle = cellStyle.copyWith(
                  bottomBorderVal: mergedCellStyle.bottomBorder,
                );
              }
              if (k == _cellIndexStart.columnIndex) {
                cellStyle = cellStyle.copyWith(
                  leftBorderVal: mergedCellStyle.leftBorder,
                );
              }
              if (k == _cellIndexEnd.columnIndex) {
                cellStyle = cellStyle.copyWith(
                  rightBorderVal: mergedCellStyle.rightBorder,
                );
              }

              if (j == k ||
                  _cellIndexStart.rowIndex - _cellIndexEnd.rowIndex == 0 ||
                  _cellIndexStart.columnIndex - _cellIndexEnd.columnIndex ==
                      0) {
                cellStyle = cellStyle.copyWith(
                  diagonalBorderUpVal: mergedCellStyle.diagonalBorderUp,
                  diagonalBorderDownVal: mergedCellStyle.diagonalBorderDown,
                );
              }

              if (j == _cellIndexStart.rowIndex &&
                  k == _cellIndexStart.columnIndex) {
                mergedCellData!._cellStyle = cellStyle;
              } else {
                CellStyle? savedCellStyle = _mergedCellStyle
                    .firstWhereOrNull((element) => element == cellStyle);

                if (savedCellStyle == null) {
                  _mergedCellStyle.add(cellStyle);
                }

                if (_excel._mergedCellStyleReferenced[sheetName] == null) {
                  _excel._mergedCellStyleReferenced[sheetName] = {};
                }

                _excel._mergedCellStyleReferenced[sheetName]![getCellId(k, j)] =
                    _mergedCellStyle.indexOf(cellStyle);
              }
            }
          }
        }
      }
    }
  }

  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.

  _processStylesFile() {
    _innerCellStyle = <CellStyle>[];
    _mergedCellStyle = <CellStyle>[];
    List<String> patternFill = <String>[];
    List<_FontStyle> fontStyle = <_FontStyle>[];
    List<_BorderSet> borderSet = <_BorderSet>[];

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      _setMergeCellsStyles(sheetName, sheetObject);
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

    List<CellStyle> cellStyles = _innerCellStyle + _mergedCellStyle;

    _excel._mergedCellStyleReferenced.forEach(
      (_, cellStyleRreference) => cellStyleRreference.forEach(
        (key, _) {
          cellStyleRreference[key] =
              cellStyleRreference[key]! + _innerCellStyle.length + 1;
        },
      ),
    );

    cellStyles.forEach((cellStyle) {
      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily);

      /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fs`
      if (_fontStyleIndex(_excel._fontStyleList, _fs) == -1 &&
          _fontStyleIndex(fontStyle, _fs) == -1) {
        fontStyle.add(_fs);
      }

      /// Filling the inner usable extra list of background color
      String backgroundColor = cellStyle.backgroundColor;
      if (!_excel._patternFill.contains(backgroundColor) &&
          !patternFill.contains(backgroundColor)) {
        patternFill.add(backgroundColor);
      }

      final _bs = _createBorderSetFromCellStyle(cellStyle);
      if (!_excel._borderSetList.contains(_bs) && !borderSet.contains(_bs)) {
        borderSet.add(_bs);
      }
    });

    XmlElement fonts =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('fonts').first;

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value =
          '${_excel._fontStyleList.length + fontStyle.length}';
    } else {
      fonts.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._fontStyleList.length + fontStyle.length}'));
    }

    fontStyle.forEach((fontStyleElement) {
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
          '${_excel._patternFill.length + patternFill.length}';
    } else {
      fills.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._patternFill.length + patternFill.length}'));
    }

    patternFill.forEach((color) {
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

    XmlElement borders =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('borders').first;
    var borderAttribute = borders.getAttributeNode('count');

    if (borderAttribute != null) {
      borderAttribute.value =
          '${_excel._borderSetList.length + borderSet.length}';
    } else {
      borders.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._borderSetList.length + borderSet.length}'));
    }

    borderSet.forEach((border) {
      var borderElement = XmlElement(XmlName('border'));
      if (border.diagonalBorderDown) {
        borderElement.attributes
            .add(XmlAttribute(XmlName('diagonalDown'), '1'));
      }
      if (border.diagonalBorderUp) {
        borderElement.attributes.add(XmlAttribute(XmlName('diagonalUp'), '1'));
      }
      final Map<String, Border> borderMap = {
        'left': border.leftBorder,
        'right': border.rightBorder,
        'top': border.topBorder,
        'bottom': border.bottomBorder,
        'diagonal': border.diagonalBorder,
      };
      for (var key in borderMap.keys) {
        final borderValue = borderMap[key]!;

        final element = XmlElement(XmlName(key));
        final style = borderValue.borderStyle;
        if (style != null) {
          element.attributes.add(XmlAttribute(XmlName('style'), style.style));
        }
        final color = borderValue.borderColorHex;
        if (color != null) {
          element.children.add(XmlElement(
              XmlName('color'), [XmlAttribute(XmlName('rgb'), color)]));
        }
        borderElement.children.add(element);
      }

      borders.children.add(borderElement);
    });

    XmlElement celx =
        _excel._xmlFiles['xl/styles.xml']!.findAllElements('cellXfs').first;
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value =
          '${_excel._cellStyleList.length + cellStyles.length}';
    } else {
      celx.attributes.add(XmlAttribute(XmlName('count'),
          '${_excel._cellStyleList.length + cellStyles.length}'));
    }

    cellStyles.forEach((cellStyle) {
      String backgroundColor = cellStyle.backgroundColor;

      _FontStyle _fs = _FontStyle(
          bold: cellStyle.isBold,
          italic: cellStyle.isItalic,
          fontColorHex: cellStyle.fontColor,
          underline: cellStyle.underline,
          fontSize: cellStyle.fontSize,
          fontFamily: cellStyle.fontFamily);

      HorizontalAlign horizontalAlign = cellStyle.horizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.verticalAlignment;
      int rotation = cellStyle.rotation;
      TextWrapping? textWrapping = cellStyle.wrap;
      int backgroundIndex = patternFill.indexOf(backgroundColor),
          fontIndex = _fontStyleIndex(fontStyle, _fs);
      _BorderSet _bs = _createBorderSetFromCellStyle(cellStyle);
      int borderIndex = borderSet.indexOf(_bs);

      var attributes = <XmlAttribute>[
        XmlAttribute(XmlName('borderId'),
            '${borderIndex == -1 ? 0 : borderIndex + _excel._borderSetList.length}'),
        XmlAttribute(XmlName('fillId'),
            '${backgroundIndex == -1 ? 0 : backgroundIndex + _excel._patternFill.length}'),
        XmlAttribute(XmlName('fontId'),
            '${fontIndex == -1 ? 0 : fontIndex + _excel._fontStyleList.length}'),
        XmlAttribute(XmlName('numFmtId'), '0'),
        XmlAttribute(XmlName('xfId'), '0'),
      ];

      if ((_excel._patternFill.contains(backgroundColor) ||
              patternFill.contains(backgroundColor)) &&
          backgroundColor != "none" &&
          backgroundColor != "gray125" &&
          backgroundColor.toLowerCase() != "lightgray") {
        attributes.add(XmlAttribute(XmlName('applyFill'), '1'));
      }

      if (_fontStyleIndex(_excel._fontStyleList, _fs) != -1 &&
          _fontStyleIndex(fontStyle, _fs) != -1) {
        attributes.add(XmlAttribute(XmlName('applyFont'), '1'));
      }

      var children = <XmlElement>[];

      if (horizontalAlign != HorizontalAlign.Left ||
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

        if (horizontalAlign != HorizontalAlign.Left) {
          String hor =
              horizontalAlign == HorizontalAlign.Right ? 'right' : 'center';
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

  List<int>? _save() {
    if (_excel._styleChanges) {
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

  void _setColumns(Sheet sheetObject, XmlDocument xmlFile) {
    final colElements = xmlFile.findAllElements('cols');

    if (sheetObject.getColWidths.isEmpty &&
        sheetObject.getColumnAutoFit.isEmpty) {
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

    final autoFits = sheetObject.getColumnAutoFit;
    final customWidths = sheetObject.getColWidths;

    final columnCount = max(autoFits.length, customWidths.length);

    List<double> colWidths = <double>[];
    int min = 0;

    double defaultColumnWidth =
        sheetObject.defaultColumnWidth ?? _excelDefaultColumnWidth;

    for (var index = 0; index < columnCount; index++) {
      double width = defaultColumnWidth;

      if (autoFits.containsKey(index) && (!customWidths.containsKey(index))) {
        width = _calcAutoFitColWidth(sheetObject, index);
      } else {
        if (customWidths.containsKey(index)) {
          width = customWidths[index]!;
        }
      }

      colWidths.add(width);

      if (index != 0 && colWidths[index - 1] != width) {
        _addNewCol(cols, min, index - 1, colWidths[index - 1]);
        min = index;
      }

      if (index == (columnCount - 1)) {
        _addNewCol(cols, index, index, width);
      }
    }
  }

  void _setRows(String sheetName, Sheet sheetObject) {
    final customHeights = sheetObject.getRowHeights;

    List<CellIndex> mergedCells =
        (_excel._mergedCellStyleReferenced[sheetName] ?? <String, int>{})
            .keys
            .map((e) => CellIndex.indexByString(e))
            .toList();

    List<int> mergedCellPositionRowRange = [];

    for (var cell in mergedCells) {
      if (!mergedCellPositionRowRange.contains(cell.rowIndex)) {
        mergedCellPositionRowRange.add(cell.rowIndex);
      }
    }

    for (var rowIndex = 0; rowIndex < sheetObject._maxRows; rowIndex++) {
      double? height;

      if (customHeights.containsKey(rowIndex)) {
        height = customHeights[rowIndex];
      }

      XmlElement? foundRow;

      if (mergedCellPositionRowRange.contains(rowIndex)) {
        foundRow = _createNewRow(
            _excel._sheets[sheetName]! as XmlElement, rowIndex, height);

        if (mergedCells.isNotEmpty) {
          for (var cell in mergedCells) {
            if (cell.rowIndex == rowIndex) {
              _createCell(
                  sheetName, foundRow, cell.columnIndex, rowIndex, null);
            }
          }
        }
      }

      if (sheetObject._sheetData[rowIndex] == null) {
        continue;
      } else {
        foundRow ??= _createNewRow(
            _excel._sheets[sheetName]! as XmlElement, rowIndex, height);
      }

      for (var colIndex = 0; colIndex < sheetObject._maxCols; colIndex++) {
        var data = sheetObject._sheetData[rowIndex]![colIndex];
        if (data == null) {
          continue;
        }
        if (data.value != null) {
          _createCell(sheetName, foundRow, colIndex, rowIndex, data.value);
          foundRow.children.sort((a, b) {
            var aColumnIndex =
                CellIndex.indexByString(a.getAttribute('r')!).columnIndex;
            var bColumnIndex =
                CellIndex.indexByString(b.getAttribute('r')!).columnIndex;
            return aColumnIndex.compareTo(bColumnIndex);
          });
        }
      }
    }
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

        XmlDocument? sheetFile = _excel._xmlFiles[_excel._xmlSheetId[s]];

        if (iterMergeElement?.isNotEmpty ?? false) {
          mergeElement = iterMergeElement!.first;
        } else {
          if ((sheetFile?.findAllElements('worksheet').length ?? 0) > 0) {
            int index = sheetFile!
                .findAllElements('worksheet')
                .first
                .children
                .indexOf(sheetFile.findAllElements("sheetData").first);
            if (index == -1) {
              _damagedExcel();
            }
            sheetFile.findAllElements('worksheet').first.children.insert(
                index + 1,
                XmlElement(XmlName('mergeCells'),
                    [XmlAttribute(XmlName('count'), '0')]));

            mergeElement = sheetFile.findAllElements('mergeCells').first;
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

  /// Writing cell contained text into the excel sheet files.
  _setSheetElements() {
    _excel._sharedStrings = _SharedStringsMaintainer.instance;
    _excel._sharedStrings.clear();

    _excel._sheetMap.forEach((sheetName, sheetObject) {
      ///
      /// Create the sheet's xml file if it does not exist.
      if (_excel._sheets[sheetName] == null) {
        parser._createSheet(sheetName);
      }

      /// Clear the previous contents of the sheet if it exists,
      /// in order to reduce the time to find and compare with the sheet rows
      /// and hence just do the work of putting the data only i.e. creating new rows
      if (_excel._sheets[sheetName]?.children.isNotEmpty ?? false) {
        _excel._sheets[sheetName]!.children.clear();
      }

      /// `Above function is important in order to wipe out the old contents of the sheet.`

      XmlDocument? xmlFile = _excel._xmlFiles[_excel._xmlSheetId[sheetName]];
      if (xmlFile == null) return;

      double? defaultRowHeight = sheetObject.defaultRowHeight;
      double? defaultColumnWidth = sheetObject.defaultColumnWidth;

      // Set default column width and height for the sheet.
      XmlElement sheetFormatPrElement = xmlFile
          .findAllElements('worksheet')
          .first
          .findElements('sheetFormatPr')
          .first;

      sheetFormatPrElement.attributes.clear();
      if (defaultRowHeight != null) {
        sheetFormatPrElement.attributes.add(XmlAttribute(
            XmlName('defaultRowHeight'), defaultRowHeight.toStringAsFixed(2)));
      }
      if (defaultColumnWidth != null) {
        sheetFormatPrElement.attributes.add(XmlAttribute(
            XmlName('defaultColWidth'), defaultColumnWidth.toStringAsFixed(2)));
      }

      if (defaultRowHeight == null && defaultColumnWidth == null) {
        xmlFile
            .findAllElements('worksheet')
            .first
            .children
            .remove(sheetFormatPrElement);
      }

      _setColumns(sheetObject, xmlFile);

      _setRows(sheetName, sheetObject);

      _setHeaderFooter(sheetName);
    });
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
  void _updateCell(String sheet, XmlElement row, int columnIndex, int rowIndex,
      dynamic value) {
    XmlNode? cell = row.children.firstWhereOrNull(
        (cell) => cell.getAttribute('r') == getCellId(columnIndex, rowIndex));
    if (cell == null) return;
    XmlNode newCell =
        _createCell(sheet, row, columnIndex, rowIndex, value, false);
    if (value == null) {
      newCell.children.clear();
      newCell.children.addAll(cell.children.map((e) => e.copy()));
    }
    row.children[row.children.indexOf(cell)] = newCell;
  }

  _BorderSet _createBorderSetFromCellStyle(CellStyle cellStyle) => _BorderSet(
        leftBorder: cellStyle.leftBorder,
        rightBorder: cellStyle.rightBorder,
        topBorder: cellStyle.topBorder,
        bottomBorder: cellStyle.bottomBorder,
        diagonalBorder: cellStyle.diagonalBorder,
        diagonalBorderUp: cellStyle.diagonalBorderUp,
        diagonalBorderDown: cellStyle.diagonalBorderDown,
      );
}
