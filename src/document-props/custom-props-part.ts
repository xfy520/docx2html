import Part from '../base/part';
import { CustomProperty, parseCustomProps } from './custom-props';

export default class CustomPropsPart extends Part {
  props: CustomProperty[];

  parseXml(root: Element) {
    this.props = parseCustomProps(root);
  }
}
