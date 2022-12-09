import { XmlParser } from '../parser/xml-parser';
import { Length } from '../types';

export interface LineSpacing {
  after: Length;
  before: Length;
  line: number;
  lineRule: 'atLeast' | 'exactly' | 'auto';
}

export function parseLineSpacing(elem: Element): LineSpacing {
  return {
    before: XmlParser.lengthAttr(elem, 'before'),
    after: XmlParser.lengthAttr(elem, 'after'),
    line: XmlParser.intAttr(elem, 'line'),
    lineRule: XmlParser.attr(elem, 'lineRule'),
  } as LineSpacing;
}
