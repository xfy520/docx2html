import Part from '../base/part';
import Xml from '../base/xml';
import Parser from '../parser';
import { WmlBaseNote, WmlEndnote, WmlFootnote } from './elements';

class BaseNotePart<T extends WmlBaseNote> extends Part {
  protected _documentParser: Parser;

  notes: T[];

  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path);
    this._documentParser = parser;
  }
}

export class FootnotesPart extends BaseNotePart<WmlFootnote> {
  // eslint-disable-next-line no-useless-constructor
  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path, parser);
  }

  parseXml(root: Element) {
    this.notes = this._documentParser.parseNotes(root, 'footnote', WmlFootnote);
  }
}

export class EndnotesPart extends BaseNotePart<WmlEndnote> {
  // eslint-disable-next-line no-useless-constructor
  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path, parser);
  }

  parseXml(root: Element) {
    this.notes = this._documentParser.parseNotes(root, 'endnote', WmlEndnote);
  }
}
