import Part from '../base/part';
import Xml from '../base/xml';
import { DmlTheme, parseTheme } from './theme';

export default class ThemePart extends Part {
  theme: DmlTheme;

  // eslint-disable-next-line no-useless-constructor
  constructor(xml: Xml, path: string) {
    super(xml, path);
  }

  parseXml(root: Element) {
    this.theme = parseTheme(root);
  }
}
