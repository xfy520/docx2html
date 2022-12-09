import { Length, LengthUsage, LengthUsageType } from '../types';
import { convertBoolean, convertLength } from '../utils';

export function parseXmlString(xmlString: string, trimXmlDeclaration = false): Document {
  if (trimXmlDeclaration) {
    // eslint-disable-next-line no-param-reassign
    xmlString = xmlString.replace(/<[?].*[?]>/, '');
  }

  const result = new DOMParser().parseFromString(xmlString, 'application/xml');
  const errorText = hasXmlParserError(result);

  if (errorText) { throw new Error(errorText); }

  return result;
}

function hasXmlParserError(doc: Document) {
  return doc.getElementsByTagName('parsererror')[0]?.textContent;
}

export function serializeXmlString(elem: Node): string {
  return new XMLSerializer().serializeToString(elem);
}

export class XmlParser {
  static elements(elem: Element, localName: string = null): Element[] {
    const result = [];

    for (let i = 0, l = elem.childNodes.length; i < l;) {
      const c = elem.childNodes.item(i);

      if (c.nodeType === 1 && (localName == null || (c as Element).localName === localName)) { result.push(c); }
      i += 1;
    }

    return result;
  }

  static element(elem: Element, localName: string): Element {
    for (let i = 0, l = elem.childNodes.length; i < l;) {
      const c = elem.childNodes.item(i);

      if (c.nodeType === 1 && (c as Element).localName === localName) { return c as Element; }
      i += 1;
    }

    return null;
  }

  static elementAttr(elem: Element, localName: string, attrLocalName: string): string {
    const el = XmlParser.element(elem, localName);
    return el ? XmlParser.attr(el, attrLocalName) : undefined;
  }

  static attrs(elem: Element) {
    return Array.from(elem.attributes);
  }

  static attr(elem: Element, localName: string): string {
    for (let i = 0, l = elem.attributes.length; i < l;) {
      const a = elem.attributes.item(i);

      if (a.localName === localName) { return a.value; }
      i += 1;
    }

    return null;
  }

  static intAttr(node: Element, attrName: string, defaultValue: number = null): number {
    const val = XmlParser.attr(node, attrName);
    return val ? parseInt(val, 10) : defaultValue;
  }

  static hexAttr(node: Element, attrName: string, defaultValue: number = null): number {
    const val = XmlParser.attr(node, attrName);
    return val ? parseInt(val, 16) : defaultValue;
  }

  static floatAttr(node: Element, attrName: string, defaultValue: number = null): number {
    const val = XmlParser.attr(node, attrName);
    return val ? parseFloat(val) : defaultValue;
  }

  static boolAttr(node: Element, attrName: string, defaultValue: boolean = null) {
    return convertBoolean(XmlParser.attr(node, attrName), defaultValue);
  }

  static lengthAttr(node: Element, attrName: string, usage: LengthUsageType = LengthUsage.Dxa): Length {
    return convertLength(XmlParser.attr(node, attrName), usage);
  }
}
