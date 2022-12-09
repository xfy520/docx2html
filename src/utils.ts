import { ParagraphTab } from './document/paragraph';
import { XmlParser } from './parser/xml-parser';
import {
  CommonProperties,
  Length,
  LengthUsage, LengthUsageType, ns, Relationship,
} from './types';

export function parseRelationships(root: Element): Relationship[] {
  return XmlParser.elements(root).map((e) => <Relationship>{
    id: XmlParser.attr(e, 'Id'),
    type: XmlParser.attr(e, 'Type'),
    target: XmlParser.attr(e, 'Target'),
    targetMode: XmlParser.attr(e, 'TargetMode'),
  });
}

export function splitPath(path: string): [string, string] {
  const si = path.lastIndexOf('/') + 1;
  const folder = si === 0 ? '' : path.substring(0, si);
  const fileName = si === 0 ? path : path.substring(si);
  return [folder, fileName];
}

export function resolvePath(path: string, base: string): string {
  try {
    const prefix = 'http://docx/';
    const url = new URL(path, prefix + base).toString();
    return url.substring(prefix.length);
  } catch {
    return `${base}${path}`;
  }
}

export function convertLength(val: string, usage: LengthUsageType = LengthUsage.Dxa): string {
  if (val == null || /.+(p[xt]|[%])$/.test(val)) {
    return val;
  }

  return `${(parseInt(val, 10) * usage.mul).toFixed(2)}${usage.unit}`;
}

export function convertBoolean(v: string, defaultValue = false): boolean {
  switch (v) {
    case '1': return true;
    case '0': return false;
    case 'on': return true;
    case 'off': return false;
    case 'true': return true;
    case 'false': return false;
    default: return defaultValue;
  }
}

export function convertPercentage(val: string): number {
  return val ? parseInt(val, 10) / 100 : null;
}

export function parseCommonProperty(elem: Element, props: CommonProperties): boolean {
  if (elem.namespaceURI !== ns.wordml) { return false; }

  switch (elem.localName) {
    case 'color':
      props.color = XmlParser.attr(elem, 'val');
      break;

    case 'sz':
      props.fontSize = XmlParser.lengthAttr(elem, 'val', LengthUsage.FontSize);
      break;

    default:
      return false;
  }

  return true;
}

const knownColors = ['black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen',
  'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'none', 'red', 'white', 'yellow'];

export class xmlUtil {
  // eslint-disable-next-line no-unused-vars
  static foreach(node: Element, cb: (n: Element) => void) {
    for (let i = 0; i < node.childNodes.length;) {
      const n = node.childNodes[i];
      if (n.nodeType === Node.ELEMENT_NODE) { cb(n as Element); }
      i += 1;
    }
  }

  static colorAttr(node: Element, attrName: string, defValue: string = null, autoColor = 'black') {
    const v = XmlParser.attr(node, attrName);

    if (v) {
      if (v === 'auto') {
        return autoColor;
      } if (knownColors.includes(v)) {
        return v;
      }

      return `#${v}`;
    }

    const themeColor = XmlParser.attr(node, 'themeColor');

    return themeColor ? `var(--docx-${themeColor}-color)` : defValue;
  }

  static sizeValue(node: Element, type: LengthUsageType = LengthUsage.Dxa) {
    return convertLength(node.textContent, type);
  }
}

export const autos = {
  shd: 'inherit',
  color: 'black',
  borderColor: 'black',
  highlight: 'transparent',
};

export class values {
  static themeValue(c: Element, attr: string) {
    const val = XmlParser.attr(c, attr);
    return val ? `var(--docx-${val}-font)` : null;
  }

  static valueOfSize(c: Element, attr: string) {
    let type = LengthUsage.Dxa;

    switch (XmlParser.attr(c, 'type')) {
      case 'dxa': break;
      case 'pct': type = LengthUsage.Percent; break;
      case 'auto': return 'auto';
      default: break;
    }

    return XmlParser.lengthAttr(c, attr, type);
  }

  static valueOfMargin(c: Element) {
    return XmlParser.lengthAttr(c, 'w');
  }

  static valueOfBorder(c: Element) {
    const type = XmlParser.attr(c, 'val');

    if (type === 'nil') { return 'none'; }

    const color = xmlUtil.colorAttr(c, 'color');
    const size = XmlParser.lengthAttr(c, 'sz', LengthUsage.Border);

    return `${size} solid ${color === 'auto' ? autos.borderColor : color}`;
  }

  static valueOfTblLayout(c: Element) {
    const type = XmlParser.attr(c, 'val');
    return type === 'fixed' ? 'fixed' : 'auto';
  }

  static classNameOfCnfStyle(c: Element) {
    const val = XmlParser.attr(c, 'val');
    const classes = [
      'first-row', 'last-row', 'first-col', 'last-col',
      'odd-col', 'even-col', 'odd-row', 'even-row',
      'ne-cell', 'nw-cell', 'se-cell', 'sw-cell',
    ];

    return classes.filter((_, i) => val[i] === '1').join(' ');
  }

  static valueOfJc(c: Element) {
    const type = XmlParser.attr(c, 'val');

    switch (type) {
      case 'start':
      case 'left': return 'left';
      case 'center': return 'center';
      case 'end':
      case 'right': return 'right';
      case 'both': return 'justify';
      default: break;
    }

    return type;
  }

  static valueOfVertAlign(c: Element, asTagName = false) {
    const type = XmlParser.attr(c, 'val');

    switch (type) {
      case 'subscript': return 'sub';
      case 'superscript': return asTagName ? 'sup' : 'super';
      default: break;
    }

    return asTagName ? null : type;
  }

  static valueOfTextAlignment(c: Element) {
    const type = XmlParser.attr(c, 'val');

    switch (type) {
      case 'auto':
      case 'baseline': return 'baseline';
      case 'top': return 'top';
      case 'center': return 'middle';
      case 'bottom': return 'bottom';
      default: break;
    }

    return type;
  }

  static addSize(a: string, b: string): string {
    if (a == null) return b;
    if (b == null) return a;

    return `calc(${a} + ${b})`;
  }

  static classNameOftblLook(c: Element) {
    const val = XmlParser.hexAttr(c, 'val', 0);
    let className = '';

    if (XmlParser.boolAttr(c, 'firstRow') || (val & 0x0020)) className += ' first-row';
    if (XmlParser.boolAttr(c, 'lastRow') || (val & 0x0040)) className += ' last-row';
    if (XmlParser.boolAttr(c, 'firstColumn') || (val & 0x0080)) className += ' first-col';
    if (XmlParser.boolAttr(c, 'lastColumn') || (val & 0x0100)) className += ' last-col';
    if (XmlParser.boolAttr(c, 'noHBand') || (val & 0x0200)) className += ' no-hband';
    if (XmlParser.boolAttr(c, 'noVBand') || (val & 0x0400)) className += ' no-vband';

    return className.trim();
  }
}

export function isObject(item) {
  return item && typeof item === 'object' && !Array.isArray(item);
}

export function isString(item: unknown): item is string {
  return item && typeof item === 'string' || item instanceof String;
}

export function keyBy<T = string>(array: T[], by: (x: T) => string): Record<string, T> {
  return array.reduce((props, x) => {
    props[by(x)] = x;
    return props;
  }, {});
}

export function mergeDeep(target, ...sources) {
  if (!sources.length) { return target; }

  const source = sources.shift();

  if (isObject(target) && isObject(source)) {
    for (const key in source) {
      if (isObject(source[key])) {
        const val = target[key] ?? (target[key] = {});
        mergeDeep(val, source[key]);
      } else {
        target[key] = source[key];
      }
    }
  }

  return mergeDeep(target, ...sources);
}

export function escapeClassName(className: string) {
  return className?.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase();
}

export function blobToBase64(blob: Blob): Promise<string | ArrayBuffer> {
  return new Promise((resolve, _) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.readAsDataURL(blob);
  });
}

export function deobfuscate(data: Uint8Array, guidKey: string): Uint8Array {
  const len = 16;
  const trimmed = guidKey.replace(/{|}|-/g, '');
  const numbers = new Array(len);

  for (let i = 0; i < len;) {
    numbers[len - i - 1] = parseInt(trimmed.substr(i * 2, 2), 16);
    i += 1;
  }

  for (let i = 0; i < 32;) {
    data[i] = data[i] ^ numbers[i % len];
    i += 1;
  }

  return data;
}

export interface TabStop {
  pos: number;
  leader: string;
  style: string;
}

const defaultTab: TabStop = { pos: 0, leader: 'none', style: 'left' };
const maxTabs = 50;

export function computePixelToPoint(container: HTMLElement = document.body) {
  const temp = document.createElement('div');
  temp.style.width = '100pt';

  container.appendChild(temp);
  const result = 100 / temp.offsetWidth;
  container.removeChild(temp);

  return result;
}

export function updateTabStop(elem: HTMLElement, tabs: ParagraphTab[], defaultTabSize: Length, pixelToPoint: number = 72 / 96) {
  const p = elem.closest('p');

  const ebb = elem.getBoundingClientRect();
  const pbb = p.getBoundingClientRect();
  const pcs = getComputedStyle(p);

  const tabStops = tabs?.length > 0 ? tabs.map((t) => ({
    pos: lengthToPoint(t.position),
    leader: t.leader,
    style: t.style,
  })).sort((a, b) => a.pos - b.pos) : [defaultTab];

  const lastTab = tabStops[tabStops.length - 1];
  const pWidthPt = pbb.width * pixelToPoint;
  const size = lengthToPoint(defaultTabSize);
  let pos = lastTab.pos + size;

  if (pos < pWidthPt) {
    for (; pos < pWidthPt && tabStops.length < maxTabs; pos += size) {
      tabStops.push({ ...defaultTab, pos });
    }
  }

  const marginLeft = parseFloat(pcs.marginLeft);
  const pOffset = pbb.left + marginLeft;
  const left = (ebb.left - pOffset) * pixelToPoint;
  const tab = tabStops.find((t) => t.style !== 'clear' && t.pos > left);

  if (tab == null) { return; }

  let width = 1;

  if (tab.style === 'right' || tab.style === 'center') {
    const tabStops = Array.from(p.querySelectorAll(`.${elem.className}`));
    const nextIdx = tabStops.indexOf(elem) + 1;
    const range = document.createRange();
    range.setStart(elem, 1);

    if (nextIdx < tabStops.length) {
      range.setEndBefore(tabStops[nextIdx]);
    } else {
      range.setEndAfter(p);
    }

    const mul = tab.style === 'center' ? 0.5 : 1;
    const nextBB = range.getBoundingClientRect();
    const offset = nextBB.left + mul * nextBB.width - (pbb.left - marginLeft);

    width = tab.pos - offset * pixelToPoint;
  } else {
    width = tab.pos - left;
  }

  elem.innerHTML = '&nbsp;';
  elem.style.textDecoration = 'inherit';
  elem.style.wordSpacing = `${width.toFixed(0)}pt`;

  switch (tab.leader) {
    case 'dot':
    case 'middleDot':
      elem.style.textDecoration = 'underline';
      elem.style.textDecorationStyle = 'dotted';
      break;

    case 'hyphen':
    case 'heavy':
    case 'underscore':
      elem.style.textDecoration = 'underline';
      break;
    default: break;
  }
}

function lengthToPoint(length: Length): number {
  return parseFloat(length);
}

export function asArray<T>(val: T | T[]): T[] {
  return Array.isArray(val) ? val : [val];
}
