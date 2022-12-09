import Part from '../base/part';
import { CorePropsDeclaration, parseCoreProps } from './core-props';

export default class CorePropsPart extends Part {
  props: CorePropsDeclaration;

  parseXml(root: Element) {
    this.props = parseCoreProps(root);
  }
}
