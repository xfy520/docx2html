import Part from '../base/part';
import Xml from '../base/xml';
import Parser from '../parser';
import { IDomNumbering } from '../types';
import {
  AbstractNumbering, Numbering, NumberingBulletPicture, NumberingPartProperties, parseNumberingPart,
} from './numbering';

export default class NumberingPart extends Part implements NumberingPartProperties {
  private _documentParser: Parser;

  constructor(xml: Xml, path: string, parser: Parser) {
    super(xml, path);
    this._documentParser = parser;
  }

  numberings: Numbering[];

  abstractNumberings: AbstractNumbering[];

  bulletPictures: NumberingBulletPicture[];

  domNumberings: IDomNumbering[];

  parseXml(root: Element) {
    Object.assign(this, parseNumberingPart(root));
    this.domNumberings = this._documentParser.parseNumberingFile(root);
  }
}
