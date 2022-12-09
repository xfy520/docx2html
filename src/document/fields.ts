import { XmlElement } from '../types';

export interface WmlInstructionText extends XmlElement {
  text: string;
}

export interface WmlFieldChar extends XmlElement {
  charType: 'begin' | 'end' | 'separate' | string;
  lock: boolean;
}

export interface WmlFieldSimple extends XmlElement {
  instruction: string;
  lock: boolean;
  dirty: boolean;
}
