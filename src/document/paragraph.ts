import { XmlParser } from '../parser/xml-parser';
import {
  CommonProperties, Length, ns, XmlElement,
} from '../types';
import { parseCommonProperty } from '../utils';
import { Borders } from './border';
import { LineSpacing, parseLineSpacing } from './line-spacing';
import { parseRunProperties, RunProperties } from './run';
import { parseSectionProperties, SectionProperties } from './section';

export interface ParagraphTab {
  style: 'bar' | 'center' | 'clear' | 'decimal' | 'end' | 'num' | 'start' | 'left' | 'right';
  leader: 'none' | 'dot' | 'heavy' | 'hyphen' | 'middleDot' | 'underscore';
  position: Length;
}

export interface ParagraphNumbering {
  id: string;
  level: number;
}

export interface ParagraphProperties extends CommonProperties {
  sectionProps: SectionProperties;
  tabs: ParagraphTab[];
  numbering: ParagraphNumbering;

  border: Borders;
  textAlignment: 'auto' | 'baseline' | 'bottom' | 'center' | 'top' | string;
  lineSpacing: LineSpacing;
  keepLines: boolean;
  keepNext: boolean;
  pageBreakBefore: boolean;
  outlineLevel: number;
  styleName?: string;

  runProps: RunProperties;
}

export interface WmlParagraph extends XmlElement, ParagraphProperties {
}

function parseTabs(elem: Element): ParagraphTab[] {
  return XmlParser.elements(elem, 'tab')
    .map((e) => <ParagraphTab>{
      position: XmlParser.lengthAttr(e, 'pos'),
      leader: XmlParser.attr(e, 'leader'),
      style: XmlParser.attr(e, 'val'),
    });
}

function parseNumbering(elem: Element): ParagraphNumbering {
  const result = <ParagraphNumbering>{};

  const elements = XmlParser.elements(elem);

  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    switch (el.localName) {
      case 'numId':
        result.id = XmlParser.attr(el, 'val');
        break;

      case 'ilvl':
        result.level = XmlParser.intAttr(el, 'val');
        break;
      default: break;
    }
    index += 1;
  }
  return result;
}

export function parseParagraphProperty(elem: Element, props: ParagraphProperties) {
  if (elem.namespaceURI !== ns.wordml) { return false; }

  if (parseCommonProperty(elem, props)) { return true; }

  switch (elem.localName) {
    case 'tabs':
      props.tabs = parseTabs(elem);
      break;

    case 'sectPr':
      props.sectionProps = parseSectionProperties(elem);
      break;

    case 'numPr':
      props.numbering = parseNumbering(elem);
      break;

    case 'spacing':
      props.lineSpacing = parseLineSpacing(elem);
      return false;

    case 'textAlignment':
      props.textAlignment = XmlParser.attr(elem, 'val');
      return false;

    case 'keepNext':
      props.keepLines = XmlParser.boolAttr(elem, 'val', true);
      props.keepNext = XmlParser.boolAttr(elem, 'val', true);
      break;

    case 'pageBreakBefore':
      props.pageBreakBefore = XmlParser.boolAttr(elem, 'val', true);
      break;

    case 'outlineLvl':
      props.outlineLevel = XmlParser.intAttr(elem, 'val');
      break;

    case 'pStyle':
      props.styleName = XmlParser.attr(elem, 'val');
      break;

    case 'rPr':
      props.runProps = parseRunProperties(elem);
      break;

    default:
      return false;
  }

  return true;
}

export function parseParagraphProperties(elem: Element): ParagraphProperties {
  const result = <ParagraphProperties>{};
  const elements = XmlParser.elements(elem);

  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    parseParagraphProperty(el, result);
    index += 1;
  }
  return result;
}
