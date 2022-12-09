import { DomType, XmlElement } from '../types';

export abstract class WmlBaseNote implements XmlElement {
  id: string;

  type: DomType;

  noteType: string;

  children?: XmlElement[] = [];

  cssStyle?: Record<string, string> = {};

  className?: string;

  parent?: XmlElement;
}

export class WmlFootnote extends WmlBaseNote {
  type = DomType.Footnote;
}

export class WmlEndnote extends WmlBaseNote {
  type = DomType.Endnote;
}
