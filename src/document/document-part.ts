import Part from '../base/part';
import Xml from '../base/xml';
import Parser from '../parser';
import { DocumentElement } from './document';

export default class DocumentPart extends Part {
  private _parser: Parser;

  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path);
    this._parser = parser;
  }

  body: DocumentElement;

  parseXml(root: Element) {
    this.body = this._parser.parseDocumentFile(root);
  }
}
