import { XmlParser } from '../parser/xml-parser';
import { Length } from '../types';

export interface WmlSettings {
  defaultTabStop: Length;
  footnoteProps: NoteProperties;
  endnoteProps: NoteProperties;
  autoHyphenation: boolean;
}

interface NoteProperties {
  nummeringFormat: string;
  defaultNoteIds: string[];
}

export function parseSettings(elem: Element) {
  const result = {} as WmlSettings;

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    switch (el.localName) {
      case 'defaultTabStop': result.defaultTabStop = XmlParser.lengthAttr(el, 'val'); break;
      case 'footnotePr': result.footnoteProps = parseNoteProperties(el); break;
      case 'endnotePr': result.endnoteProps = parseNoteProperties(el); break;
      case 'autoHyphenation': result.autoHyphenation = XmlParser.boolAttr(el, 'val'); break;
      default: break;
    }
    index += 1;
  }

  return result;
}

function parseNoteProperties(elem: Element) {
  const result = {
    defaultNoteIds: [],
  } as NoteProperties;

  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    switch (el.localName) {
      case 'numFmt':
        result.nummeringFormat = XmlParser.attr(el, 'val');
        break;

      case 'footnote':
      case 'endnote':
        result.defaultNoteIds.push(XmlParser.attr(el, 'id'));
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}
