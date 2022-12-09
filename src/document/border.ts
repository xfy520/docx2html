import { XmlParser } from '../parser/xml-parser';
import { Length, LengthUsage } from '../types';

export interface Border {
  color: string;
  type: string;
  size: Length;
  frame: boolean;
  shadow: boolean;
  offset: Length;
}

export interface Borders {
  top: Border;
  left: Border;
  right: Border;
  bottom: Border;
}

export function parseBorder(elem: Element): Border {
  return {
    type: XmlParser.attr(elem, 'val'),
    color: XmlParser.attr(elem, 'color'),
    size: XmlParser.lengthAttr(elem, 'sz', LengthUsage.Border),
    offset: XmlParser.lengthAttr(elem, 'space', LengthUsage.Point),
    frame: XmlParser.boolAttr(elem, 'frame'),
    shadow: XmlParser.boolAttr(elem, 'shadow'),
  };
}

export function parseBorders(elem: Element): Borders {
  const result = <Borders>{};
  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const element = elements[index];
    switch (element.localName) {
      case 'left': result.left = parseBorder(element); break;
      case 'top': result.top = parseBorder(element); break;
      case 'right': result.right = parseBorder(element); break;
      case 'bottom': result.bottom = parseBorder(element); break;
      default: break;
    }
    index += 1;
  }
  return result;
}
