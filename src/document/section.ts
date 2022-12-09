import { XmlParser } from '../parser/xml-parser';
import { Length } from '../types';
import { Borders, parseBorders } from './border';

interface Column {
  space: Length;
  width: Length;
}

interface Columns {
  space: Length;
  numberOfColumns: number;
  separator: boolean;
  equalWidth: boolean;
  columns: Column[];
}

interface PageSize {
  width: Length,
  height: Length,
  orientation: 'landscape' | string
}

interface PageNumber {
  start: number;
  chapSep: 'colon' | 'emDash' | 'endash' | 'hyphen' | 'period' | string;
  chapStyle: string;
  format: 'none' | 'cardinalText' | 'decimal' | 'decimalEnclosedCircle' | 'decimalEnclosedFullstop'
  | 'decimalEnclosedParen' | 'decimalZero' | 'lowerLetter' | 'lowerRoman'
  | 'ordinalText' | 'upperLetter' | 'upperRoman' | string;
}

interface PageMargins {
  top: Length;
  right: Length;
  bottom: Length;
  left: Length;
  header: Length;
  footer: Length;
  gutter: Length;
}

enum SectionType {
  Continuous = 'continuous',
  NextPage = 'nextPage',
  NextColumn = 'nextColumn',
  EvenPage = 'evenPage',
  OddPage = 'oddPage',
}

export interface FooterHeaderReference {
  id: string;
  type: string | 'first' | 'even' | 'default';
}

export interface SectionProperties {
  type: SectionType | string;
  pageSize: PageSize,
  pageMargins: PageMargins,
  pageBorders: Borders;
  pageNumber: PageNumber;
  columns: Columns;
  footerRefs: FooterHeaderReference[];
  headerRefs: FooterHeaderReference[];
  titlePage: boolean;
}

function parseColumns(elem: Element): Columns {
  return {
    numberOfColumns: XmlParser.intAttr(elem, 'num'),
    space: XmlParser.lengthAttr(elem, 'space'),
    separator: XmlParser.boolAttr(elem, 'sep'),
    equalWidth: XmlParser.boolAttr(elem, 'equalWidth', true),
    columns: XmlParser.elements(elem, 'col')
      .map((e) => <Column>{
        width: XmlParser.lengthAttr(e, 'w'),
        space: XmlParser.lengthAttr(e, 'space'),
      }),
  };
}

function parsePageNumber(elem: Element): PageNumber {
  return {
    chapSep: XmlParser.attr(elem, 'chapSep'),
    chapStyle: XmlParser.attr(elem, 'chapStyle'),
    format: XmlParser.attr(elem, 'fmt'),
    start: XmlParser.intAttr(elem, 'start'),
  };
}

function parseFooterHeaderReference(elem: Element): FooterHeaderReference {
  return {
    id: XmlParser.attr(elem, 'id'),
    type: XmlParser.attr(elem, 'type'),
  };
}

export function parseSectionProperties(elem: Element): SectionProperties {
  const section = <SectionProperties>{};

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const element = elements[index];
    switch (element.localName) {
      case 'pgSz':
        section.pageSize = {
          width: XmlParser.lengthAttr(element, 'w'),
          height: XmlParser.lengthAttr(element, 'h'),
          orientation: XmlParser.attr(element, 'orient'),
        };
        break;

      case 'type':
        section.type = XmlParser.attr(element, 'val');
        break;

      case 'pgMar':
        section.pageMargins = {
          left: XmlParser.lengthAttr(element, 'left'),
          right: XmlParser.lengthAttr(element, 'right'),
          top: XmlParser.lengthAttr(element, 'top'),
          bottom: XmlParser.lengthAttr(element, 'bottom'),
          header: XmlParser.lengthAttr(element, 'header'),
          footer: XmlParser.lengthAttr(element, 'footer'),
          gutter: XmlParser.lengthAttr(element, 'gutter'),
        };
        break;

      case 'cols':
        section.columns = parseColumns(element);
        break;

      case 'headerReference':
        (section.headerRefs ?? (section.headerRefs = [])).push(parseFooterHeaderReference(element));
        break;

      case 'footerReference':
        (section.footerRefs ?? (section.footerRefs = [])).push(parseFooterHeaderReference(element));
        break;

      case 'titlePg':
        section.titlePage = XmlParser.boolAttr(element, 'val', true);
        break;

      case 'pgBorders':
        section.pageBorders = parseBorders(element);
        break;

      case 'pgNumType':
        section.pageNumber = parsePageNumber(element);
        break;
      default: break;
    }
    index += 1;
  }

  return section;
}
