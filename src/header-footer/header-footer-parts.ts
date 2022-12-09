import Part from '../base/part';
import Xml from '../base/xml';
import Parser from '../parser';
import { XmlElement } from '../types';
import { WmlFooter, WmlHeader } from './elements';

export abstract class BaseHeaderFooterPart<T extends XmlElement = XmlElement> extends Part {
  rootElement: T;

  private _documentParser: Parser;

  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path);
    this._documentParser = parser;
  }

  parseXml(root: Element) {
    this.rootElement = this.createRootElement();
    this.rootElement.children = this._documentParser.parseBodyElements(root);
  }

  protected abstract createRootElement(): T;
}

export class HeaderPart extends BaseHeaderFooterPart<WmlHeader> {
  // eslint-disable-next-line class-methods-use-this
  protected createRootElement(): WmlHeader {
    return new WmlHeader();
  }
}

export class FooterPart extends BaseHeaderFooterPart<WmlFooter> {
  // eslint-disable-next-line class-methods-use-this
  protected createRootElement(): WmlFooter {
    return new WmlFooter();
  }
}
