import { XmlElement } from '../types';
import { SectionProperties } from './section';

export interface DocumentElement extends XmlElement {
  props: SectionProperties;
}
