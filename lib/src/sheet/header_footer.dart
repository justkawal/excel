part of excel;

class HeaderFooter {
  late bool? _alignWithMargins;
  late bool? _differentFirst;
  late bool? _differentOddEven;
  late bool? _scaleWithDoc;

  late String? _evenFooter;
  late String? _evenHeader;
  late String? _firstFooter;
  late String? _firstHeader;
  late String? _oddFooter;
  late String? _oddHeader;

  HeaderFooter({
    bool? alignWithMargins = null,
    bool? differentFirst = null,
    bool? differentOddEven = null,
    bool? scaleWithDoc = null,
    String? evenFooter = null,
    String? evenHeader = null,
    String? firstFooter = null,
    String? firstHeader = null,
    String? oddFooter = null,
    String? oddHeader = null,
  })  : _alignWithMargins = alignWithMargins,
        _differentFirst = differentFirst,
        _differentOddEven = differentOddEven,
        _scaleWithDoc = scaleWithDoc,
        _evenFooter = evenFooter,
        _evenHeader = evenHeader,
        _firstFooter = firstFooter,
        _firstHeader = firstHeader,
        _oddFooter = oddFooter,
        _oddHeader = oddHeader;

  bool? get alignWithMargins {
    return _alignWithMargins;
  }

  set alignWithMargins(bool? alignWithMargins) {
    _alignWithMargins = alignWithMargins;
  }

  bool? get differentFirst {
    return _differentFirst;
  }

  set differentFist(bool? differentFirst) {
    _differentFirst = differentFirst;
  }

  bool? get differentOddEven {
    return _differentOddEven;
  }

  set differentOddEven(bool? differentOddEven) {
    _differentOddEven = differentOddEven;
  }

  bool? get scaleWithDoc {
    return _scaleWithDoc;
  }

  set scaleWithDoc(bool? scaleWithDoc) {
    _scaleWithDoc = scaleWithDoc;
  }

  String? get evenFooter {
    return _evenFooter;
  }

  set evenFooter(String? evenFooter) {
    _evenFooter = evenFooter;
  }

  String? get evenHeader {
    return _evenHeader;
  }

  set evenHeader(String? evenHeader) {
    _evenHeader = evenHeader;
  }

  String? get firstFooter {
    return _firstFooter;
  }

  set firstFooter(String? firstFooter) {
    _firstFooter = firstFooter;
  }

  String? get firstHeader {
    return _firstHeader;
  }

  set firstHeader(String? firstHeader) {
    _firstHeader = firstHeader;
  }

  String? get oddFooter {
    return _oddFooter;
  }

  set oddFooter(String? oddFooter) {
    _oddFooter = oddFooter;
  }

  String? get oddHeader {
    return _oddHeader;
  }

  set oddHeader(String? oddHeader) {
    _oddHeader = oddHeader;
  }

  XmlNode toXmlElement() {
    final attributes = <XmlAttribute>[];
    if (_alignWithMargins != null) {
      attributes.add(XmlAttribute(
          XmlName("alignWithMargins"), _alignWithMargins.toString()));
    }
    if (_differentFirst != null) {
      attributes.add(
          XmlAttribute(XmlName("differentFirst"), _differentFirst.toString()));
    }
    if (_differentOddEven != null) {
      attributes.add(XmlAttribute(
          XmlName("differentOddEven"), _differentOddEven.toString()));
    }
    if (_scaleWithDoc != null) {
      attributes
          .add(XmlAttribute(XmlName("scaleWithDoc"), _scaleWithDoc.toString()));
    }

    final children = <XmlNode>[];
    if (_evenFooter != null) {
      children
          .add(XmlElement(XmlName("evenFooter"), [], [XmlText(_evenFooter!)]));
    }
    if (_evenHeader != null) {
      children
          .add(XmlElement(XmlName("evenHeader"), [], [XmlText(_evenHeader!)]));
    }
    if (_firstFooter != null) {
      children.add(
          XmlElement(XmlName("firstFooter"), [], [XmlText(_firstFooter!)]));
    }
    if (_firstHeader != null) {
      children.add(
          XmlElement(XmlName("firstHeader"), [], [XmlText(_firstHeader!)]));
    }
    if (_oddFooter != null) {
      children
          .add(XmlElement(XmlName("oddFooter"), [], [XmlText(_oddFooter!)]));
    }
    if (_oddHeader != null) {
      children
          .add(XmlElement(XmlName("oddHeader"), [], [XmlText(_oddHeader!)]));
    }

    return XmlElement(XmlName("headerFooter"), attributes, children);
  }

  static HeaderFooter fromXmlElement(XmlElement headerFooterElement) {
    return HeaderFooter(
        alignWithMargins:
            headerFooterElement.getAttribute("alignWithMargins")?.parseBool(),
        differentFirst:
            headerFooterElement.getAttribute("differentFirst")?.parseBool(),
        differentOddEven:
            headerFooterElement.getAttribute("differentOddEven")?.parseBool(),
        scaleWithDoc:
            headerFooterElement.getAttribute("scaleWithDoc")?.parseBool(),
        evenFooter: headerFooterElement.getElement("evenFooter")?.innerXml,
        evenHeader: headerFooterElement.getElement("evenHeader")?.innerXml,
        firstFooter: headerFooterElement.getElement("firstFooter")?.innerXml,
        firstHeader: headerFooterElement.getElement("firstHeader")?.innerXml,
        oddFooter: headerFooterElement.getElement("oddFooter")?.innerXml,
        oddHeader: headerFooterElement.getElement("oddHeader")?.innerXml);
  }
}

extension BoolParsing on String {
  bool parseBool() {
    if (this.toLowerCase() == 'true') {
      return true;
    } else if (this.toLowerCase() == 'false') {
      return false;
    }

    throw '"$this" can not be parsed to boolean.';
  }
}
