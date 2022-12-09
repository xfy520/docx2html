import Part from '../base/part';
import Xml from '../base/xml';
import { ParagraphProperties } from '../document/paragraph';
import { RunProperties } from '../document/run';
import Parser from '../parser';

export interface IDomStyle {
  id: string;
  name?: string;
  cssName?: string;
  aliases?: string[];
  target: string;
  basedOn?: string;
  isDefault?: boolean;
  styles: IDomSubStyle[];
  linked?: string;
  next?: string;

  paragraphProps: ParagraphProperties;
  runProps: RunProperties;
}

export interface IDomSubStyle {
  target: string;
  mod?: string;
  values: Record<string, string>;
}

export default class StylesPart extends Part {
  styles: IDomStyle[];

  private _documentParser: Parser;

  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path);
    this._documentParser = parser;
  }

  parseXml(root: Element) {
    this.styles = this._documentParser.parseStylesFile(root);
  }
}
