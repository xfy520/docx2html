import { XmlParser } from '../parser/xml-parser';

const embedFontTypeMap = {
  embedRegular: 'regular',
  embedBold: 'bold',
  embedItalic: 'italic',
  embedBoldItalic: 'boldItalic',
};

interface EmbedFontRef {
  id: string;
  key: string;
  type: 'regular' | 'bold' | 'italic' | 'boldItalic';
}

export interface FontDeclaration {
  name: string,
  altName: string,
  family: string,
  embedFontRefs: EmbedFontRef[];
}

export function parseFonts(root: Element): FontDeclaration[] {
  return XmlParser.elements(root).map((el) => parseFont(el));
}

export function parseFont(elem: Element): FontDeclaration {
  const result = <FontDeclaration>{
    name: XmlParser.attr(elem, 'name'),
    embedFontRefs: [],
  };

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    switch (el.localName) {
      case 'family':
        result.family = XmlParser.attr(el, 'val');
        break;

      case 'altName':
        result.altName = XmlParser.attr(el, 'val');
        break;

      case 'embedRegular':
      case 'embedBold':
      case 'embedItalic':
      case 'embedBoldItalic':
        result.embedFontRefs.push(parseEmbedFontRef(el));
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}

function parseEmbedFontRef(elem: Element): EmbedFontRef {
  return {
    id: XmlParser.attr(elem, 'id'),
    key: XmlParser.attr(elem, 'fontKey'),
    type: embedFontTypeMap[elem.localName],
  };
}
