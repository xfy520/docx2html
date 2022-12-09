import { XmlParser } from '../parser/xml-parser';

export class DmlTheme {
  colorScheme: DmlColorScheme;

  fontScheme: DmlFontScheme;
}

interface DmlColorScheme {
  name: string;
  colors: Record<string, string>;
}

interface DmlFontScheme {
  name: string;
  majorFont: DmlFormInfo,
  minorFont: DmlFormInfo
}

interface DmlFormInfo {
  latinTypeface: string;
  eaTypeface: string;
  csTypeface: string;
}

export function parseTheme(elem: Element) {
  const result = new DmlTheme();
  const themeElements = XmlParser.element(elem, 'themeElements');

  const elements = XmlParser.elements(themeElements);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    switch (el.localName) {
      case 'clrScheme': result.colorScheme = parseColorScheme(el); break;
      case 'fontScheme': result.fontScheme = parseFontScheme(el); break;
      default: break;
    }
    index += 1;
  }

  return result;
}

function parseColorScheme(elem: Element) {
  const result: DmlColorScheme = {
    name: XmlParser.attr(elem, 'name'),
    colors: {},
  };

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    const srgbClr = XmlParser.element(el, 'srgbClr');
    const sysClr = XmlParser.element(el, 'sysClr');

    if (srgbClr) {
      result.colors[el.localName] = XmlParser.attr(srgbClr, 'val');
    } else if (sysClr) {
      result.colors[el.localName] = XmlParser.attr(sysClr, 'lastClr');
    }
    index += 1;
  }

  return result;
}

function parseFontScheme(elem: Element) {
  const result: DmlFontScheme = {
    name: XmlParser.attr(elem, 'name'),
  } as DmlFontScheme;

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    switch (el.localName) {
      case 'majorFont': result.majorFont = parseFontInfo(el); break;
      case 'minorFont': result.minorFont = parseFontInfo(el); break;
      default: break;
    }
    index += 1;
  }

  return result;
}

function parseFontInfo(elem: Element): DmlFormInfo {
  return {
    latinTypeface: XmlParser.elementAttr(elem, 'latin', 'typeface'),
    eaTypeface: XmlParser.elementAttr(elem, 'ea', 'typeface'),
    csTypeface: XmlParser.elementAttr(elem, 'cs', 'typeface'),
  };
}
