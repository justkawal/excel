part of excel;

class HeaderFooter {
  bool? alignWithMargins;
  bool? differentFirst;
  bool? differentOddEven;
  bool? scaleWithDoc;

  String? evenFooter;
  String? evenHeader;
  String? firstFooter;
  String? firstHeader;
  String? oddFooter;
  String? oddHeader;

  HeaderFooter({
    this.alignWithMargins,
    this.differentFirst,
    this.differentOddEven,
    this.scaleWithDoc,
    this.evenFooter,
    this.evenHeader,
    this.firstFooter,
    this.firstHeader,
    this.oddFooter,
    this.oddHeader,
  });

  XmlNode toXmlElement() {
    final attributes = <XmlAttribute>[];
    if (alignWithMargins != null) {
      attributes.add(XmlAttribute(
          XmlName.parts("alignWithMargins"), alignWithMargins.toString()));
    }
    if (differentFirst != null) {
      attributes.add(XmlAttribute(
          XmlName.parts("differentFirst"), differentFirst.toString()));
    }
    if (differentOddEven != null) {
      attributes.add(XmlAttribute(
          XmlName.parts("differentOddEven"), differentOddEven.toString()));
    }
    if (scaleWithDoc != null) {
      attributes.add(
          XmlAttribute(XmlName.parts("scaleWithDoc"), scaleWithDoc.toString()));
    }

    final children = <XmlNode>[];
    if (evenHeader != null) {
      children.add(XmlElement(XmlName.parts("evenHeader"), [],
          [XmlText(evenHeader!.simplifyText())]));
    }
    if (evenFooter != null) {
      children.add(XmlElement(XmlName.parts("evenFooter"), [],
          [XmlText(evenFooter!.simplifyText())]));
    }
    if (firstHeader != null) {
      children.add(XmlElement(XmlName.parts("firstHeader"), [],
          [XmlText(firstHeader!.simplifyText())]));
    }
    if (firstFooter != null) {
      children.add(XmlElement(XmlName.parts("firstFooter"), [],
          [XmlText(firstFooter!.simplifyText())]));
    }
    if (oddHeader != null) {
      children.add(XmlElement(XmlName.parts("oddHeader"), [],
          [XmlText(oddHeader!.simplifyText())]));
    }
    if (oddFooter != null) {
      children.add(XmlElement(XmlName.parts("oddFooter"), [],
          [XmlText(oddFooter!.simplifyText())]));
    }

    return XmlElement(XmlName.parts("headerFooter"), attributes, children);
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
        evenHeader: headerFooterElement.getElement("evenHeader")?.innerText,
        evenFooter: headerFooterElement.getElement("evenFooter")?.innerText,
        firstHeader: headerFooterElement.getElement("firstHeader")?.innerText,
        firstFooter: headerFooterElement.getElement("firstFooter")?.innerText,
        oddFooter: headerFooterElement.getElement("oddFooter")?.innerText,
        oddHeader: headerFooterElement.getElement("oddHeader")?.innerText);
  }
}

extension BoolParsing on String {
  bool parseBool() {
    var value = this.toLowerCase();
    if (value == 'true' || value == '1') {
      return true;
    } else if (value == 'false' || value == '0') {
      return false;
    }

    throw '"$this" can not be parsed to boolean.';
  }

  String simplifyText() {
    String value = this.replaceAll('&amp', '&');
    value = value.replaceAll('amp', '&');
    value = value.replaceAll('&', '&amp;');
    value = value.replaceAll('"', '&quot;');
    return value;
  }
}
