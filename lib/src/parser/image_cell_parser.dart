part of excel;

class _ImageCellParser {
  final Excel _excel;

  _ImageCellParser(this._excel);

  void parseImageCell(XmlElement node, Sheet sheet, String rId, String name) {
    final drawingPath = _getDrawingPath(sheet, rId);
    if (drawingPath == null) {
      return;
    }

    final drawing = _excel._archive.findFile('xl/$drawingPath');
    if (drawing == null) {
      return;
    }

    final anchors = _parseDrawingAnchors(drawing);
    if (anchors.isEmpty) {
      return;
    }

    for (final anchor in anchors) {
      final cellLocation = _getCellLocation(anchor);
      if (cellLocation == null) {
        continue;
      }

      final imageInfo = _getImageInfo(anchor, drawingPath);
      if (imageInfo == null) {
        continue;
      }

      final imageCellValue = _createImageCellValue(imageInfo);
      if (imageCellValue == null) {
        continue;
      }

      sheet.updateCell(
        CellIndex.indexByColumnRow(
          columnIndex: cellLocation.$1,
          rowIndex: cellLocation.$2,
        ),
        imageCellValue,
      );
      return;
    }
  }

  String? _getDrawingPath(Sheet sheet, String rId) {
    final sheetRelsPath =
        'xl/worksheets/_rels/${sheet.sheetName.toLowerCase()}.xml.rels';
    final sheetRels = _excel._archive.findFile(sheetRelsPath);
    if (sheetRels == null) return null;

    final sheetRelsContent = utf8.decode(sheetRels.content);
    return XmlDocument.parse(sheetRelsContent)
        .findAllElements('Relationship')
        .firstWhere((e) => e.getAttribute('Id') == rId)
        .getAttribute('Target')
        ?.replaceAll('../', '');
  }

  Iterable<XmlElement> _parseDrawingAnchors(ArchiveFile drawing) {
    final drawingContent = utf8.decode(drawing.content);
    final doc = XmlDocument.parse(drawingContent);
    return doc.findAllElements('xdr:oneCellAnchor');
  }

  (int, int)? _getCellLocation(XmlElement anchor) {
    final from = anchor.findElements('xdr:from').firstOrNull;
    if (from == null) return null;

    final cellColumnIndex =
        int.tryParse(from.findElements('xdr:col').firstOrNull?.innerText ?? '');
    final rowIndex =
        int.tryParse(from.findElements('xdr:row').firstOrNull?.innerText ?? '');

    if (cellColumnIndex == null || rowIndex == null) return null;
    return (cellColumnIndex, rowIndex);
  }

  ({String path, List<int> bytes, String format, int? width, int? height})?
      _getImageInfo(XmlElement anchor, String drawingPath) {
    final blip = anchor.findAllElements('a:blip').firstOrNull;
    if (blip == null) return null;

    final imageRId = blip.getAttribute('r:embed');
    if (imageRId == null) return null;

    final dimensions = _getImageDimensions(anchor);
    final imagePath = _resolveImagePath(imageRId, drawingPath);
    if (imagePath == null) return null;

    final imageFile = _excel._archive.findFile(imagePath);
    if (imageFile == null) return null;

    final format = imagePath.split('.').last.toLowerCase();
    if (!['png', 'jpg', 'jpeg', 'gif'].contains(format)) return null;

    return (
      path: imagePath,
      bytes: imageFile.content,
      format: format,
      width: dimensions.$1,
      height: dimensions.$2,
    );
  }

  (int?, int?) _getImageDimensions(XmlElement anchor) {
    final ext = anchor.findAllElements('xdr:ext').firstOrNull ??
        anchor.findAllElements('a:ext').firstOrNull;
    if (ext == null) return (null, null);

    final width = (int.tryParse(ext.getAttribute('cx') ?? '') ?? 0) ~/
        9525; // Convert EMUs to pixels
    final height = (int.tryParse(ext.getAttribute('cy') ?? '') ?? 0) ~/ 9525;

    return (width, height);
  }

  String? _resolveImagePath(String imageRId, String drawingPath) {
    final drawingRelsPath =
        'xl/drawings/_rels/${drawingPath.split('/').last}.rels';
    final drawingRels = _excel._archive.findFile(drawingRelsPath);
    if (drawingRels == null) return null;

    final relationships = XmlDocument.parse(utf8.decode(drawingRels.content))
        .findAllElements('Relationship');

    final relationship = relationships
        .where((e) => e.getAttribute('Id') == imageRId)
        .firstOrNull;
    if (relationship == null) return null;

    final imagePath = relationship.getAttribute('Target');
    if (imagePath == null) return null;

    return imagePath.startsWith('../')
        ? 'xl/${imagePath.substring(3)}'
        : 'xl/media/$imagePath';
  }

  ImageCellValue? _createImageCellValue(
      ({
        String path,
        List<int> bytes,
        String format,
        int? width,
        int? height
      }) info) {
    return ImageCellValue(
      bytes: info.bytes,
      format: info.format,
      width: info.width,
      height: info.height,
    );
  }
}
