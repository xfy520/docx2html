import { DomType, LengthUsage, XmlElement } from '../types';
import { XmlParser } from '../parser/xml-parser';

export class VmlElement implements XmlElement {
  type: DomType = DomType.VmlElement;

  tagName: string;

  cssStyleText?: string;

  attrs: Record<string, string> = {};

  chidren: VmlElement[] = [];

  wrapType?: string;

  imageHref?: {
    id: string,
    title: string
  };
}

function parseStroke(el: Element): Record<string, string> {
  return {
    stroke: XmlParser.attr(el, 'color'),
    'stroke-width': XmlParser.lengthAttr(el, 'weight', LengthUsage.Emu) ?? '1px',
  };
}

function parseFill(_el: Element): Record<string, string> {
  return {
    // 'fill': XmlParser.attr(el, "color2")
  };
}

function parsePoint(val: string): string[] {
  return val.split(',');
}

export function parseVmlElement(elem: Element): VmlElement {
  const result = new VmlElement();

  switch (elem.localName) {
    case 'rect':
      result.tagName = 'rect';
      Object.assign(result.attrs, { width: '100%', height: '100%' });
      break;

    case 'oval':
      result.tagName = 'ellipse';
      Object.assign(result.attrs, {
        cx: '50%', cy: '50%', rx: '50%', ry: '50%',
      });
      break;

    case 'line':
      result.tagName = 'line';
      break;

    case 'shape':
      result.tagName = 'g';
      break;

    default:
      return null;
  }

  const attrs = XmlParser.attrs(elem);
  for (let index = 0; index < attrs.length;) {
    const at = attrs[index];
    switch (at.localName) {
      case 'style':
        result.cssStyleText = at.value;
        break;

      case 'fillcolor':
        result.attrs.fill = at.value;
        break;

      case 'from':
        // eslint-disable-next-line no-case-declarations
        const [x1, y1] = parsePoint(at.value);
        Object.assign(result.attrs, { x1, y1 });
        break;

      case 'to':
        // eslint-disable-next-line no-case-declarations
        const [x2, y2] = parsePoint(at.value);
        Object.assign(result.attrs, { x2, y2 });
        break;
      default: break;
    }
    index += 1;
  }

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    let child = null;
    switch (el.localName) {
      case 'stroke':
        Object.assign(result.attrs, parseStroke(el));
        break;

      case 'fill':
        Object.assign(result.attrs, parseFill(el));
        break;

      case 'imagedata':
        result.tagName = 'image';
        Object.assign(result.attrs, { width: '100%', height: '100%' });
        result.imageHref = {
          id: XmlParser.attr(el, 'id'),
          title: XmlParser.attr(el, 'title'),
        };
        break;

      default:
        child = parseVmlElement(el);
        if (child) {
          result.chidren.push(child);
        }
        break;
    }
    index += 1;
  }

  return result;
}
