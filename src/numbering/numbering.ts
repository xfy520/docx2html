import { ParagraphProperties, parseParagraphProperties } from '../document/paragraph';
import { parseRunProperties, RunProperties } from '../document/run';
import { XmlParser } from '../parser/xml-parser';

export interface NumberingPartProperties {
  numberings: Numbering[];
  abstractNumberings: AbstractNumbering[];
  bulletPictures: NumberingBulletPicture[];
}

export interface Numbering {
  id: string;
  abstractId: string;
  overrides: NumberingLevelOverride[];
}

interface NumberingLevelOverride {
  level: number;
  start: number;
  numberingLevel: NumberingLevel;
}

export interface AbstractNumbering {
  id: string;
  name: string;
  multiLevelType: 'singleLevel' | 'multiLevel' | 'hybridMultilevel' | string;
  levels: NumberingLevel[];
  numberingStyleLink: string;
  styleLink: string;
}

interface NumberingLevel {
  level: number;
  start: string;
  restart: number;
  format: 'lowerRoman' | 'lowerLetter' | string;
  text: string;
  justification: string;
  bulletPictureId: string;
  paragraphStyle: string;
  paragraphProps: ParagraphProperties;
  runProps: RunProperties;
}

export interface NumberingBulletPicture {
  id: string;
  referenceId: string;
  style: string;
}

function parseNumberingLevelOverrride(elem: Element): NumberingLevelOverride {
  const result = <NumberingLevelOverride>{
    level: XmlParser.intAttr(elem, 'ilvl'),
  };
  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const e = elements[index];
    switch (e.localName) {
      case 'startOverride':
        result.start = XmlParser.intAttr(e, 'val');
        break;
      case 'lvl':
        result.numberingLevel = parseNumberingLevel(e);
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}

function parseNumberingBulletPicture(elem: Element): NumberingBulletPicture {
  const pict = XmlParser.element(elem, 'pict');
  const shape = pict && XmlParser.element(pict, 'shape');
  const imagedata = shape && XmlParser.element(shape, 'imagedata');

  return imagedata ? {
    id: XmlParser.attr(elem, 'numPicBulletId'),
    referenceId: XmlParser.attr(imagedata, 'id'),
    style: XmlParser.attr(shape, 'style'),
  } : null;
}

export function parseNumberingPart(elem: Element): NumberingPartProperties {
  const result: NumberingPartProperties = {
    numberings: [],
    abstractNumberings: [],
    bulletPictures: [],
  };
  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const e = elements[index];
    switch (e.localName) {
      case 'num':
        result.numberings.push(parseNumbering(e));
        break;
      case 'abstractNum':
        result.abstractNumberings.push(parseAbstractNumbering(e));
        break;
      case 'numPicBullet':
        result.bulletPictures.push(parseNumberingBulletPicture(e));
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}

export function parseNumbering(elem: Element): Numbering {
  const result = <Numbering>{
    id: XmlParser.attr(elem, 'numId'),
    overrides: [],
  };
  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const e = elements[index];
    switch (e.localName) {
      case 'abstractNumId':
        result.abstractId = XmlParser.attr(e, 'val');
        break;
      case 'lvlOverride':
        result.overrides.push(parseNumberingLevelOverrride(e));
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}

export function parseAbstractNumbering(elem: Element): AbstractNumbering {
  const result = <AbstractNumbering>{
    id: XmlParser.attr(elem, 'abstractNumId'),
    levels: [],
  };
  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const e = elements[index];
    switch (e.localName) {
      case 'name':
        result.name = XmlParser.attr(e, 'val');
        break;
      case 'multiLevelType':
        result.multiLevelType = XmlParser.attr(e, 'val');
        break;
      case 'numStyleLink':
        result.numberingStyleLink = XmlParser.attr(e, 'val');
        break;
      case 'styleLink':
        result.styleLink = XmlParser.attr(e, 'val');
        break;
      case 'lvl':
        result.levels.push(parseNumberingLevel(e));
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}

export function parseNumberingLevel(elem: Element): NumberingLevel {
  const result = <NumberingLevel>{
    level: XmlParser.intAttr(elem, 'ilvl'),
  };
  const elements = XmlParser.elements(elem);
  for (let index = 0; index < elements.length;) {
    const e = elements[index];
    switch (e.localName) {
      case 'start':
        result.start = XmlParser.attr(e, 'val');
        break;
      case 'lvlRestart':
        result.restart = XmlParser.intAttr(e, 'val');
        break;
      case 'numFmt':
        result.format = XmlParser.attr(e, 'val');
        break;
      case 'lvlText':
        result.text = XmlParser.attr(e, 'val');
        break;
      case 'lvlJc':
        result.justification = XmlParser.attr(e, 'val');
        break;
      case 'lvlPicBulletId':
        result.bulletPictureId = XmlParser.attr(e, 'val');
        break;
      case 'pStyle':
        result.paragraphStyle = XmlParser.attr(e, 'val');
        break;
      case 'pPr':
        result.paragraphProps = parseParagraphProperties(e);
        break;
      case 'rPr':
        result.runProps = parseRunProperties(e);
        break;
      default: break;
    }
    index += 1;
  }

  return result;
}
