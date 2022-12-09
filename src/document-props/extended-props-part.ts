import Part from '../base/part';
import { ExtendedPropsDeclaration, parseExtendedProps } from './extended-props';

export default class ExtendedPropsPart extends Part {
  props: ExtendedPropsDeclaration;

  parseXml(root: Element) {
    this.props = parseExtendedProps(root);
  }
}
