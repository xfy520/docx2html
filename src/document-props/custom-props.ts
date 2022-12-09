import { XmlParser } from '../parser/xml-parser';

export interface CustomProperty {
  formatId: string;
  name: string;
  type: string;
  value: string;
}

export function parseCustomProps(root: Element): CustomProperty[] {
  return XmlParser.elements(root, 'property').map((e) => {
    const { firstChild } = e;

    return {
      formatId: XmlParser.attr(e, 'fmtid'),
      name: XmlParser.attr(e, 'name'),
      type: firstChild.nodeName,
      value: firstChild.textContent,
    };
  });
}
