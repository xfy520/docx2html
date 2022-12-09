import { XmlParser } from '../parser/xml-parser';
import { DomType, XmlElement } from '../types';

export interface WmlBookmarkStart extends XmlElement {
  id: string;
  name: string;
  colFirst: number;
  colLast: number;
}

export interface WmlBookmarkEnd extends XmlElement {
  id: string;
}

export function parseBookmarkStart(elem: Element): WmlBookmarkStart {
  return {
    type: DomType.BookmarkStart,
    id: XmlParser.attr(elem, 'id'),
    name: XmlParser.attr(elem, 'name'),
    colFirst: XmlParser.intAttr(elem, 'colFirst'),
    colLast: XmlParser.intAttr(elem, 'colLast'),
  };
}

export function parseBookmarkEnd(elem: Element): WmlBookmarkEnd {
  return {
    type: DomType.BookmarkEnd,
    id: XmlParser.attr(elem, 'id'),
  };
}
