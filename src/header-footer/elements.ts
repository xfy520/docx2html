import { DomType, XmlElement } from '../types';

export class WmlHeader implements XmlElement {
  type: DomType = DomType.Header;

  children?: XmlElement[] = [];

  cssStyle?: Record<string, string> = {};

  className?: string;

  parent?: XmlElement;
}

export class WmlFooter implements XmlElement {
  type: DomType = DomType.Footer;

  children?: XmlElement[] = [];

  cssStyle?: Record<string, string> = {};

  className?: string;

  parent?: XmlElement;
}
