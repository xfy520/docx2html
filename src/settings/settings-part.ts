import Part from '../base/part';
import Xml from '../base/xml';
import { WmlSettings, parseSettings } from './settings';

export default class SettingsPart extends Part {
  settings: WmlSettings;

  // eslint-disable-next-line no-useless-constructor
  constructor(xml: Xml, path: string) {
    super(xml, path);
  }

  parseXml(root: Element) {
    this.settings = parseSettings(root);
  }
}
