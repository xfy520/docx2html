import Part from '../base/part';
import { FontDeclaration, parseFonts } from './fonts';

export default class FontTablePart extends Part {
  fonts: FontDeclaration[];

  parseXml(root: Element) {
    this.fonts = parseFonts(root);
  }
}
