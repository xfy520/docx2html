import { XmlParser } from '../parser/xml-parser';
import { CommonProperties, XmlElement } from '../types';
import { parseCommonProperty } from '../utils';

export type RunProperties = CommonProperties

export interface WmlRun extends XmlElement, RunProperties {
  id?: string;
  verticalAlign?: string;
  fieldRun?: boolean;
}

export function parseRunProperty(elem: Element, props: RunProperties) {
  if (parseCommonProperty(elem, props)) { return true; }
  return false;
}

export function parseRunProperties(elem: Element): RunProperties {
  const result = <RunProperties>{};

  const elements = XmlParser.elements(elem);

  for (let index = 0; index < elements.length;) {
    const el = elements[index];
    parseRunProperty(el, result);
    index += 1;
  }

  return result;
}
