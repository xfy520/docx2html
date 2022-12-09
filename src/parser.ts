import { parseBookmarkEnd, parseBookmarkStart } from './document/bookmarks';
import { DocumentElement } from './document/document';
import { WmlFieldChar, WmlFieldSimple, WmlInstructionText } from './document/fields';
import { ParagraphProperties, parseParagraphProperties, parseParagraphProperty } from './document/paragraph';
import { parseRunProperties, WmlRun } from './document/run';
import { parseSectionProperties } from './document/section';
import { WmlFootnote } from './notes/elements';
import { XmlParser } from './parser/xml-parser';
import { IDomStyle, IDomSubStyle } from './styles/styles-part';
import {
  DomType, XmlElement, ParserOptions, LengthUsage, WmlText, WmlBreak, WmlSymbol,
  WmlNoteReference, IDomImage, WmlHyperlink, WmlTable, WmlTableRow, WmlTableCell,
  WmlTableColumn, IDomNumbering, NumberingPicBullet,
} from './types';
import { autos, values, xmlUtil } from './utils';
import { parseVmlElement } from './vml/vml';

export interface WmlParagraph extends XmlElement, ParagraphProperties {
}

const supportedNamespaceURIs = [];

const mmlTagMap = {
  oMath: DomType.MmlMath,
  oMathPara: DomType.MmlMathParagraph,
  f: DomType.MmlFraction,
  num: DomType.MmlNumerator,
  den: DomType.MmlDenominator,
  rad: DomType.MmlRadical,
  deg: DomType.MmlDegree,
  e: DomType.MmlBase,
  sSup: DomType.MmlSuperscript,
  sSub: DomType.MmlSubscript,
  sup: DomType.MmlSuperArgument,
  sub: DomType.MmlSubArgument,
  d: DomType.MmlDelimiter,
  nary: DomType.MmlNary,
};

export default class Parser {
  options: ParserOptions;

  constructor(options?: Partial<ParserOptions>) {
    this.options = {
      ignoreWidth: false,
      debug: false,
      ...options,
    };
  }

  parseNotes(xmlDoc: Element, elemName: string, elemClass: unknown): WmlFootnote[] {
    const result = [];

    const elements = XmlParser.elements(xmlDoc, elemName);

    for (let index = 0; index < elements.length;) {
      const el = elements[index];
      // @ts-ignore
      const node = new elemClass();
      node.id = XmlParser.attr(el, 'id');
      node.noteType = XmlParser.attr(el, 'type');
      node.children = this.parseBodyElements(el);
      result.push(node);
      index += 1;
    }

    return result;
  }

  parseStylesFile(xstyles: Element): IDomStyle[] {
    const result = [];

    xmlUtil.foreach(xstyles, (n) => {
      switch (n.localName) {
        case 'style':
          result.push(this.parseStyle(n));
          break;

        case 'docDefaults':
          result.push(this.parseDefaultStyles(n));
          break;
        default: break;
      }
    });

    return result;
  }

  parseDefaultStyles(node: Element): IDomStyle {
    const result = <IDomStyle>{
      id: null,
      name: null,
      target: null,
      basedOn: null,
      styles: [],
    };

    xmlUtil.foreach(node, (c) => {
      let rPr = null;
      let pPr = null;
      switch (c.localName) {
        case 'rPrDefault':
          rPr = XmlParser.element(c, 'rPr');
          if (rPr) {
            result.styles.push({
              target: 'span',
              values: this.parseDefaultProperties(rPr, {}),
            });
          }
          break;

        case 'pPrDefault':
          pPr = XmlParser.element(c, 'pPr');
          if (pPr) {
            result.styles.push({
              target: 'p',
              values: this.parseDefaultProperties(pPr, {}),
            });
          }
          break;
        default: break;
      }
    });

    return result;
  }

  parseStyle(node: Element): IDomStyle {
    const result = <IDomStyle>{
      id: XmlParser.attr(node, 'styleId'),
      isDefault: XmlParser.boolAttr(node, 'default'),
      name: null,
      target: null,
      basedOn: null,
      styles: [],
      linked: null,
    };

    switch (XmlParser.attr(node, 'type')) {
      case 'paragraph': result.target = 'p'; break;
      case 'table': result.target = 'table'; break;
      case 'character': result.target = 'span'; break;
      case 'numbering': result.target = 'p'; break;
      default: break;
    }

    xmlUtil.foreach(node, (n) => {
      let styles = [];
      switch (n.localName) {
        case 'basedOn':
          result.basedOn = XmlParser.attr(n, 'val');
          break;

        case 'name':
          result.name = XmlParser.attr(n, 'val');
          break;

        case 'link':
          result.linked = XmlParser.attr(n, 'val');
          break;

        case 'next':
          result.next = XmlParser.attr(n, 'val');
          break;

        case 'aliases':
          result.aliases = XmlParser.attr(n, 'val').split(',');
          break;

        case 'pPr':
          result.styles.push({
            target: 'p',
            values: this.parseDefaultProperties(n, {}),
          });
          result.paragraphProps = parseParagraphProperties(n);
          break;

        case 'rPr':
          result.styles.push({
            target: 'span',
            values: this.parseDefaultProperties(n, {}),
          });
          result.runProps = parseRunProperties(n);
          break;

        case 'tblPr':
        case 'tcPr':
          result.styles.push({
            target: 'td',
            values: this.parseDefaultProperties(n, {}),
          });
          break;

        case 'tblStylePr':
          styles = this.parseTableStyle(n);
          for (let index = 0; index < styles.length;) {
            result.styles.push(styles[index]);
            index += 1;
          }
          break;

        case 'rsid':
        case 'qFormat':
        case 'hidden':
        case 'semiHidden':
        case 'unhideWhenUsed':
        case 'autoRedefine':
        case 'uiPriority':
          break;

        default:
          if (this.options.debug) {
            console.warn(`DOCX: Unknown style element: ${n.localName}`);
          }
      }
    });

    return result;
  }

  parseTableStyle(node: Element): IDomSubStyle[] {
    const result = [];

    const type = XmlParser.attr(node, 'type');
    let selector = '';
    let modificator = '';

    switch (type) {
      case 'firstRow':
        modificator = '.first-row';
        selector = 'tr.first-row td';
        break;
      case 'lastRow':
        modificator = '.last-row';
        selector = 'tr.last-row td';
        break;
      case 'firstCol':
        modificator = '.first-col';
        selector = 'td.first-col';
        break;
      case 'lastCol':
        modificator = '.last-col';
        selector = 'td.last-col';
        break;
      case 'band1Vert':
        modificator = ':not(.no-vband)';
        selector = 'td.odd-col';
        break;
      case 'band2Vert':
        modificator = ':not(.no-vband)';
        selector = 'td.even-col';
        break;
      case 'band1Horz':
        modificator = ':not(.no-hband)';
        selector = 'tr.odd-row';
        break;
      case 'band2Horz':
        modificator = ':not(.no-hband)';
        selector = 'tr.even-row';
        break;
      default: return [];
    }

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case 'pPr':
          result.push({
            target: `${selector} p`,
            mod: modificator,
            values: this.parseDefaultProperties(n, {}),
          });
          break;

        case 'rPr':
          result.push({
            target: `${selector} span`,
            mod: modificator,
            values: this.parseDefaultProperties(n, {}),
          });
          break;

        case 'tblPr':
        case 'tcPr':
          result.push({
            target: selector,
            mod: modificator,
            values: this.parseDefaultProperties(n, {}),
          });
          break;
        default: break;
      }
    });

    return result;
  }

  parseNumberingFile(xnums: Element): IDomNumbering[] {
    const result = [];
    const mapping = {};
    const bullets = [];

    xmlUtil.foreach(xnums, (n) => {
      let numId = null;
      let abstractNumId = null;
      switch (n.localName) {
        case 'abstractNum':
          this.parseAbstractNumbering(n, bullets)
            .forEach((x) => result.push(x));
          break;

        case 'numPicBullet':
          bullets.push(Parser.parseNumberingPicBullet(n));
          break;

        case 'num':
          numId = XmlParser.attr(n, 'numId');
          abstractNumId = XmlParser.elementAttr(n, 'abstractNumId', 'val');
          mapping[abstractNumId] = numId;
          break;
        default: break;
      }
    });

    result.forEach((props) => {
      props.id = mapping[props.id];
    });

    return result;
  }

  static parseNumberingPicBullet(elem: Element): NumberingPicBullet {
    const pict = XmlParser.element(elem, 'pict');
    const shape = pict && XmlParser.element(pict, 'shape');
    const imagedata = shape && XmlParser.element(shape, 'imagedata');

    return imagedata ? {
      id: XmlParser.intAttr(elem, 'numPicBulletId'),
      src: XmlParser.attr(imagedata, 'id'),
      style: XmlParser.attr(shape, 'style'),
    } : null;
  }

  parseAbstractNumbering(node: Element, bullets: NumberingPicBullet[]): IDomNumbering[] {
    const result = [];
    const id = XmlParser.attr(node, 'abstractNumId');

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case 'lvl':
          result.push(this.parseNumberingLevel(id, n, bullets));
          break;
        default: break;
      }
    });

    return result;
  }

  parseNumberingLevel(id: string, node: Element, bullets: NumberingPicBullet[]): IDomNumbering {
    const result: IDomNumbering = {
      id,
      level: XmlParser.intAttr(node, 'ilvl'),
      pStyleName: undefined,
      pStyle: {},
      rStyle: {},
      suff: 'tab',
    };

    xmlUtil.foreach(node, (n) => {
      let id = null;
      switch (n.localName) {
        case 'pPr':
          this.parseDefaultProperties(n, result.pStyle);
          break;

        case 'rPr':
          this.parseDefaultProperties(n, result.rStyle);
          break;

        case 'lvlPicBulletId':
          id = XmlParser.intAttr(n, 'val');
          result.bullet = bullets.find((x) => x.id === id);
          break;

        case 'lvlText':
          result.levelText = XmlParser.attr(n, 'val');
          break;

        case 'pStyle':
          result.pStyleName = XmlParser.attr(n, 'val');
          break;

        case 'numFmt':
          result.format = XmlParser.attr(n, 'val');
          break;

        case 'suff':
          result.suff = XmlParser.attr(n, 'val');
          break;
        default: break;
      }
    });

    return result;
  }

  parseDocumentFile(xmlDoc: Element): DocumentElement {
    const xbody = XmlParser.element(xmlDoc, 'body');
    const background = XmlParser.element(xmlDoc, 'background');
    const sectPr = XmlParser.element(xbody, 'sectPr');

    return {
      type: DomType.Document,
      children: this.parseBodyElements(xbody),
      props: sectPr ? parseSectionProperties(sectPr) : null,
      cssStyle: background ? Parser.parseBackground(background) : {},
    };
  }

  static parseBackground(elem: Element): Record<string, string> {
    const result = {};
    const color = xmlUtil.colorAttr(elem, 'color');

    if (color) {
      result['background-color'] = color;
    }

    return result;
  }

  parseBodyElements(elem: Element): XmlElement[] {
    const children: XmlElement[] = [];

    const elements = XmlParser.elements(elem);
    for (let index = 0; index < elements.length;) {
      const element = elements[index];
      switch (element.localName) {
        case 'p':
          children.push(this.parseParagraph(element));
          break;
        case 'tbl':
          children.push(this.parseTable(element));
          break;
        case 'sdt':
          children.push(...Parser.parseSdt(element, (e) => this.parseBodyElements(e)));
          break;
        default: break;
      }
      index += 1;
    }

    return children;
  }

  parseTable(node: Element): WmlTable {
    const result: WmlTable = { type: DomType.Table, children: [] };

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case 'tr':
          result.children.push(this.parseTableRow(c));
          break;

        case 'tblGrid':
          result.columns = Parser.parseTableColumns(c);
          break;

        case 'tblPr':
          this.parseTableProperties(c, result);
          break;
        default: break;
      }
    });

    return result;
  }

  parseTableProperties(elem: Element, table: WmlTable) {
    table.cssStyle = {};
    table.cellStyle = {};

    this.parseDefaultProperties(elem, table.cssStyle, table.cellStyle, (c) => {
      switch (c.localName) {
        case 'tblStyle':
          table.styleName = XmlParser.attr(c, 'val');
          break;

        case 'tblLook':
          table.className = values.classNameOftblLook(c);
          break;

        case 'tblpPr':
          Parser.parseTablePosition(c, table);
          break;

        case 'tblStyleColBandSize':
          table.colBandSize = XmlParser.intAttr(c, 'val');
          break;

        case 'tblStyleRowBandSize':
          table.rowBandSize = XmlParser.intAttr(c, 'val');
          break;

        default:
          return false;
      }

      return true;
    });

    switch (table.cssStyle['text-align']) {
      case 'center':
        delete table.cssStyle['text-align'];
        table.cssStyle['margin-left'] = 'auto';
        table.cssStyle['margin-right'] = 'auto';
        break;

      case 'right':
        delete table.cssStyle['text-align'];
        table.cssStyle['margin-left'] = 'auto';
        break;
      default: break;
    }
  }

  static parseTablePosition(node: Element, table: WmlTable) {
    const topFromText = XmlParser.lengthAttr(node, 'topFromText');
    const bottomFromText = XmlParser.lengthAttr(node, 'bottomFromText');
    const rightFromText = XmlParser.lengthAttr(node, 'rightFromText');
    const leftFromText = XmlParser.lengthAttr(node, 'leftFromText');

    table.cssStyle.float = 'left';
    table.cssStyle['margin-bottom'] = values.addSize(table.cssStyle['margin-bottom'], bottomFromText);
    table.cssStyle['margin-left'] = values.addSize(table.cssStyle['margin-left'], leftFromText);
    table.cssStyle['margin-right'] = values.addSize(table.cssStyle['margin-right'], rightFromText);
    table.cssStyle['margin-top'] = values.addSize(table.cssStyle['margin-top'], topFromText);
  }

  static parseTableColumns(node: Element): WmlTableColumn[] {
    const result = [];

    xmlUtil.foreach(node, (n) => {
      switch (n.localName) {
        case 'gridCol':
          result.push({ width: XmlParser.lengthAttr(n, 'w') });
          break;
        default: break;
      }
    });

    return result;
  }

  parseTableRow(node: Element): WmlTableRow {
    const result: WmlTableRow = { type: DomType.Row, children: [] };

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case 'tc':
          result.children.push(this.parseTableCell(c));
          break;

        case 'trPr':
          this.parseTableRowProperties(c, result);
          break;
        default: break;
      }
    });

    return result;
  }

  parseTableRowProperties(elem: Element, row: WmlTableRow) {
    row.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
      switch (c.localName) {
        case 'cnfStyle':
          row.className = values.classNameOfCnfStyle(c);
          break;

        case 'tblHeader':
          row.isHeader = XmlParser.boolAttr(c, 'val');
          break;

        default:
          return false;
      }

      return true;
    });
  }

  parseTableCell(node: Element): XmlElement {
    const result: WmlTableCell = { type: DomType.Cell, children: [] };

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case 'tbl':
          result.children.push(this.parseTable(c));
          break;

        case 'p':
          result.children.push(this.parseParagraph(c));
          break;

        case 'tcPr':
          this.parseTableCellProperties(c, result);
          break;
        default: break;
      }
    });

    return result;
  }

  parseTableCellProperties(elem: Element, cell: WmlTableCell) {
    cell.cssStyle = this.parseDefaultProperties(elem, {}, null, (c) => {
      switch (c.localName) {
        case 'gridSpan':
          cell.span = XmlParser.intAttr(c, 'val', null);
          break;

        case 'vMerge':
          cell.verticalMerge = XmlParser.attr(c, 'val') ?? 'continue';
          break;

        case 'cnfStyle':
          cell.className = values.classNameOfCnfStyle(c);
          break;

        default:
          return false;
      }

      return true;
    });
  }

  parseParagraph(node: Element): XmlElement {
    const result = <WmlParagraph>{ type: DomType.Paragraph, children: [] };

    const elements = XmlParser.elements(node);
    for (let index = 0; index < elements.length;) {
      const el = elements[index];
      switch (el.localName) {
        case 'pPr':
          this.parseParagraphProperties(el, result);
          break;
        case 'r':
          result.children.push(this.parseRun(el, result));
          break;
        case 'hyperlink':
          result.children.push(this.parseHyperlink(el, result));
          break;
        case 'bookmarkStart':
          result.children.push(parseBookmarkStart(el));
          break;
        case 'bookmarkEnd':
          result.children.push(parseBookmarkEnd(el));
          break;
        case 'oMath':
        case 'oMathPara':
          result.children.push(this.parseMathElement(el));
          break;
        case 'sdt':
          result.children.push(...Parser.parseSdt(el, (e) => this.parseParagraph(e).children));
          break;
        case 'ins':
          result.children.push(Parser.parseInserted(el, (e) => this.parseParagraph(e)));
          break;
        case 'del':
          result.children.push(Parser.parseDeleted(el, (e) => this.parseParagraph(e)));
          break;
        default: break;
      }
      index += 1;
    }

    return result;
  }

  static parseSdt(node: Element, parser: Function): XmlElement[] {
    const sdtContent = XmlParser.element(node, 'sdtContent');
    return sdtContent ? parser(sdtContent) : [];
  }

  static parseInserted(node: Element, parentParser: Function): XmlElement {
    return <XmlElement>{
      type: DomType.Inserted,
      children: parentParser(node)?.children ?? [],
    };
  }

  static parseDeleted(node: Element, parentParser: Function): XmlElement {
    return <XmlElement>{
      type: DomType.Deleted,
      children: parentParser(node)?.children ?? [],
    };
  }

  parseMathElement(elem: Element): XmlElement {
    const propsTag = `${elem.localName}Pr`;
    const result = { type: mmlTagMap[elem.localName], children: [] } as XmlElement;

    const elements = XmlParser.elements(elem);
    for (let index = 0; index < elements.length;) {
      const el = elements[index];
      const childType = mmlTagMap[el.localName];

      if (childType) {
        result.children.push(this.parseMathElement(el));
      } else if (el.localName === 'r') {
        result.children.push(this.parseRun(el));
      } else if (el.localName === propsTag) {
        result.props = Parser.parseMathProperies(el);
      }
      index += 1;
    }

    return result;
  }

  static parseMathProperies(elem: Element): Record<string, unknown> {
    const result: Record<string, unknown> = {};

    const elements = XmlParser.elements(elem);
    for (let index = 0; index < elements.length;) {
      const el = elements[index];
      switch (el.localName) {
        case 'chr': result.char = XmlParser.attr(el, 'val'); break;
        case 'degHide': result.hideDegree = XmlParser.boolAttr(el, 'val'); break;
        case 'begChr': result.beginChar = XmlParser.attr(el, 'val'); break;
        case 'endChr': result.endChar = XmlParser.attr(el, 'val'); break;
        default: break;
      }
      index += 1;
    }

    return result;
  }

  parseHyperlink(node: Element, parent?: XmlElement): WmlHyperlink {
    const result: WmlHyperlink = <WmlHyperlink>{ type: DomType.Hyperlink, parent, children: [] };
    const anchor = XmlParser.attr(node, 'anchor');
    const relId = XmlParser.attr(node, 'id');

    if (anchor) { result.href = `#${anchor}`; }

    if (relId) { result.id = relId; }

    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case 'r':
          result.children.push(this.parseRun(c, result));
          break;
        default: break;
      }
    });

    return result;
  }

  parseRun(node: Element, parent?: XmlElement): WmlRun {
    const result: WmlRun = <WmlRun>{ type: DomType.Run, parent, children: [] };

    xmlUtil.foreach(node, (c) => {
      const content = Parser.checkAlternateContent(c);
      let data = null;
      switch (content.localName) {
        case 't':
          result.children.push(<WmlText>{
            type: DomType.Text,
            text: content.textContent,
          });
          break;

        case 'delText':
          result.children.push(<WmlText>{
            type: DomType.DeletedText,
            text: content.textContent,
          });
          break;

        case 'fldSimple':
          result.children.push(<WmlFieldSimple>{
            type: DomType.SimpleField,
            instruction: XmlParser.attr(c, 'instr'),
            lock: XmlParser.boolAttr(c, 'lock', false),
            dirty: XmlParser.boolAttr(c, 'dirty', false),
          });
          break;

        case 'instrText':
          result.fieldRun = true;
          result.children.push(<WmlInstructionText>{
            type: DomType.Instruction,
            text: c.textContent,
          });
          break;

        case 'fldChar':
          result.fieldRun = true;
          result.children.push(<WmlFieldChar>{
            type: DomType.ComplexField,
            charType: XmlParser.attr(c, 'fldCharType'),
            lock: XmlParser.boolAttr(c, 'lock', false),
            dirty: XmlParser.boolAttr(c, 'dirty', false),
          });
          break;

        case 'noBreakHyphen':
          result.children.push({ type: DomType.NoBreakHyphen });
          break;

        case 'br':
          result.children.push(<WmlBreak>{
            type: DomType.Break,
            break: XmlParser.attr(c, 'type') || 'textWrapping',
          });
          break;

        case 'lastRenderedPageBreak':
          result.children.push(<WmlBreak>{
            type: DomType.Break,
            break: 'lastRenderedPageBreak',
          });
          break;

        case 'sym':
          result.children.push(<WmlSymbol>{
            type: DomType.Symbol,
            font: XmlParser.attr(c, 'font'),
            char: XmlParser.attr(c, 'char'),
          });
          break;

        case 'tab':
          result.children.push({ type: DomType.Tab });
          break;

        case 'footnoteReference':
          result.children.push(<WmlNoteReference>{
            type: DomType.FootnoteReference,
            id: XmlParser.attr(c, 'id'),
          });
          break;

        case 'endnoteReference':
          result.children.push(<WmlNoteReference>{
            type: DomType.EndnoteReference,
            id: XmlParser.attr(c, 'id'),
          });
          break;

        case 'drawing':
          data = Parser.parseDrawing(c);

          if (data) { result.children = [data]; }
          break;

        case 'pict':
          result.children.push(Parser.parseVmlPicture(c));
          break;

        case 'rPr':
          this.parseRunProperties(c, result);
          break;

        default: break;
      }
    });

    return result;
  }

  parseRunProperties(elem: Element, run: WmlRun) {
    this.parseDefaultProperties(elem, run.cssStyle = {}, null, (c) => {
      switch (c.localName) {
        case 'rStyle':
          run.styleName = XmlParser.attr(c, 'val');
          break;

        case 'vertAlign':
          run.verticalAlign = values.valueOfVertAlign(c, true);
          break;

        default:
          return false;
      }
      return true;
    });
  }

  static parseVmlPicture(elem: Element): XmlElement {
    const result = { type: DomType.VmlPicture, children: [] };

    const elements = XmlParser.elements(elem);
    for (let index = 0; index < elements.length;) {
      const el = elements[index];
      const child = parseVmlElement(el);
      if (child) {
        result.children.push(child);
      }
      index += 1;
    }

    return result;
  }

  static parseDrawing(node: Element): XmlElement {
    const elements = XmlParser.elements(node);
    for (let index = 0; index < elements.length;) {
      const element = elements[index];
      switch (element.localName) {
        case 'inline':
        case 'anchor':
          return Parser.parseDrawingWrapper(element);
        default: break;
      }
      index += 1;
    }
    return null;
  }

  static parseDrawingWrapper(node: Element): XmlElement {
    const result = <XmlElement>{ type: DomType.Drawing, children: [], cssStyle: {} };
    const isAnchor = node.localName === 'anchor';

    let wrapType: 'wrapTopAndBottom' | 'wrapNone' | null = null;
    const simplePos = XmlParser.boolAttr(node, 'simplePos');

    const posX = { relative: 'page', align: 'left', offset: '0' };
    const posY = { relative: 'page', align: 'top', offset: '0' };

    const elements = XmlParser.elements(node);

    for (let index = 0; index < elements.length;) {
      const n = elements[index];
      let g = null;
      switch (n.localName) {
        case 'simplePos':
          if (simplePos) {
            posX.offset = XmlParser.lengthAttr(n, 'x', LengthUsage.Emu);
            posY.offset = XmlParser.lengthAttr(n, 'y', LengthUsage.Emu);
          }
          break;

        case 'extent':
          result.cssStyle.width = XmlParser.lengthAttr(n, 'cx', LengthUsage.Emu);
          result.cssStyle.height = XmlParser.lengthAttr(n, 'cy', LengthUsage.Emu);
          break;

        case 'positionH':
        case 'positionV':
          if (!simplePos) {
            const pos = n.localName === 'positionH' ? posX : posY;
            const alignNode = XmlParser.element(n, 'align');
            const offsetNode = XmlParser.element(n, 'posOffset');

            pos.relative = XmlParser.attr(n, 'relativeFrom') ?? pos.relative;

            if (alignNode) { pos.align = alignNode.textContent; }

            if (offsetNode) { pos.offset = xmlUtil.sizeValue(offsetNode, LengthUsage.Emu); }
          }
          break;

        case 'wrapTopAndBottom':
          wrapType = 'wrapTopAndBottom';
          break;

        case 'wrapNone':
          wrapType = 'wrapNone';
          break;

        case 'graphic':
          g = Parser.parseGraphic(n);
          if (g) { result.children.push(g); }
          break;
        default: break;
      }
      index += 1;
    }

    if (wrapType === 'wrapTopAndBottom') {
      result.cssStyle.display = 'block';

      if (posX.align) {
        result.cssStyle['text-align'] = posX.align;
        result.cssStyle.width = '100%';
      }
    } else if (wrapType === 'wrapNone') {
      result.cssStyle.display = 'block';
      result.cssStyle.position = 'relative';
      result.cssStyle.width = '0px';
      result.cssStyle.height = '0px';

      if (posX.offset) { result.cssStyle.left = posX.offset; }
      if (posY.offset) { result.cssStyle.top = posY.offset; }
    } else if (isAnchor && (posX.align === 'left' || posX.align === 'right')) {
      result.cssStyle.float = posX.align;
    }

    return result;
  }

  static parseGraphic(elem: Element): XmlElement {
    const graphicData = XmlParser.element(elem, 'graphicData');
    const elements = XmlParser.elements(graphicData);

    for (let index = 0; index < elements.length;) {
      const n = elements[index];
      switch (n.localName) {
        case 'pic':
          return Parser.parsePicture(n);
        default: break;
      }
      index += 1;
    }

    return null;
  }

  static parsePicture(elem: Element): IDomImage {
    const result = <IDomImage>{ type: DomType.Image, src: '', cssStyle: {} };
    const blipFill = XmlParser.element(elem, 'blipFill');
    const blip = XmlParser.element(blipFill, 'blip');

    result.src = XmlParser.attr(blip, 'embed');

    const spPr = XmlParser.element(elem, 'spPr');
    const xfrm = XmlParser.element(spPr, 'xfrm');

    result.cssStyle.position = 'relative';

    const elements = XmlParser.elements(xfrm);

    for (let index = 0; index < elements.length;) {
      const n = elements[index];
      switch (n.localName) {
        case 'ext':
          result.cssStyle.width = XmlParser.lengthAttr(n, 'cx', LengthUsage.Emu);
          result.cssStyle.height = XmlParser.lengthAttr(n, 'cy', LengthUsage.Emu);
          break;

        case 'off':
          result.cssStyle.left = XmlParser.lengthAttr(n, 'x', LengthUsage.Emu);
          result.cssStyle.top = XmlParser.lengthAttr(n, 'y', LengthUsage.Emu);
          break;
        default: break;
      }
      index += 1;
    }
    return result;
  }

  static checkAlternateContent(elem: Element): Element {
    if (elem.localName !== 'AlternateContent') { return elem; }

    const choice = XmlParser.element(elem, 'Choice');

    if (choice) {
      const requires = XmlParser.attr(choice, 'Requires');
      const namespaceURI = elem.lookupNamespaceURI(requires);

      if (supportedNamespaceURIs.includes(namespaceURI)) { return choice.firstElementChild; }
    }

    return XmlParser.element(elem, 'Fallback')?.firstElementChild;
  }

  parseParagraphProperties(elem: Element, paragraph: WmlParagraph) {
    this.parseDefaultProperties(elem, paragraph.cssStyle = {}, null, (c) => {
      if (parseParagraphProperty(c, paragraph)) { return true; }
      switch (c.localName) {
        case 'pStyle':
          paragraph.styleName = XmlParser.attr(c, 'val');
          break;
        case 'cnfStyle':
          paragraph.className = values.classNameOfCnfStyle(c);
          break;
        case 'framePr':
          Parser.parseFrame(c, paragraph);
          break;
        case 'rPr':
          break;
        default:
          return false;
      }

      return true;
    });
  }

  static parseFrame(node: Element, paragraph: WmlParagraph) {
    const dropCap = XmlParser.attr(node, 'dropCap');

    if (dropCap === 'drop') { paragraph.cssStyle.float = 'left'; }
  }

  parseDefaultProperties(elem: Element, _style: Record<string, string> = null, childStyle: Record<string, string> = null, handler: (_arg: Element) => boolean = null): Record<string, string> {
    const style = _style || {};

    xmlUtil.foreach(elem, (c) => {
      if (handler?.(c)) { return; }
      switch (c.localName) {
        case 'jc':
          style['text-align'] = values.valueOfJc(c);
          break;

        case 'textAlignment':
          style['vertical-align'] = values.valueOfTextAlignment(c);
          break;

        case 'color':
          style.color = xmlUtil.colorAttr(c, 'val', null, autos.color);
          break;

        case 'sz':
          style['min-height'] = XmlParser.lengthAttr(c, 'val', LengthUsage.FontSize);
          style['font-size'] = style['min-height'];
          break;

        case 'shd':
          style['background-color'] = xmlUtil.colorAttr(c, 'fill', null, autos.shd);
          break;

        case 'highlight':
          style['background-color'] = xmlUtil.colorAttr(c, 'val', null, autos.highlight);
          break;

        case 'vertAlign':
          break;

        case 'position':
          style.verticalAlign = XmlParser.lengthAttr(c, 'val', LengthUsage.FontSize);
          break;

        case 'tcW':
          if (this.options.ignoreWidth) { break; }

        case 'tblW':
          style.width = values.valueOfSize(c, 'w');
          break;

        case 'trHeight':
          Parser.parseTrHeight(c, style);
          break;

        case 'strike':
          style['text-decoration'] = XmlParser.boolAttr(c, 'val', true) ? 'line-through' : 'none';
          break;

        case 'b':
          style['font-weight'] = XmlParser.boolAttr(c, 'val', true) ? 'bold' : 'normal';
          break;

        case 'i':
          style['font-style'] = XmlParser.boolAttr(c, 'val', true) ? 'italic' : 'normal';
          break;

        case 'caps':
          style['text-transform'] = XmlParser.boolAttr(c, 'val', true) ? 'uppercase' : 'none';
          break;

        case 'smallCaps':
          style['text-transform'] = XmlParser.boolAttr(c, 'val', true) ? 'lowercase' : 'none';
          break;

        case 'u':
          Parser.parseUnderline(c, style);
          break;

        case 'ind':
        case 'tblInd':
          Parser.parseIndentation(c, style);
          break;

        case 'rFonts':
          Parser.parseFont(c, style);
          break;

        case 'tblBorders':
          Parser.parseBorderProperties(c, childStyle || style);
          break;

        case 'tblCellSpacing':
          style['border-spacing'] = values.valueOfMargin(c);
          style['border-collapse'] = 'separate';
          break;

        case 'pBdr':
          Parser.parseBorderProperties(c, style);
          break;

        case 'bdr':
          style.border = values.valueOfBorder(c);
          break;

        case 'tcBorders':
          Parser.parseBorderProperties(c, style);
          break;

        case 'vanish':
          if (XmlParser.boolAttr(c, 'val', true)) { style.display = 'none'; }
          break;

        case 'kern':
          break;

        case 'noWrap':
          break;

        case 'tblCellMar':
        case 'tcMar':
          Parser.parseMarginProperties(c, childStyle || style);
          break;

        case 'tblLayout':
          style['table-layout'] = values.valueOfTblLayout(c);
          break;

        case 'vAlign':
          style['vertical-align'] = values.valueOfTextAlignment(c);
          break;

        case 'spacing':
          if (elem.localName === 'pPr') { Parser.parseSpacing(c, style); }
          break;

        case 'wordWrap':
          if (XmlParser.boolAttr(c, 'val')) {
            style['overflow-wrap'] = 'break-word';
          }
          break;

        case 'bCs':
        case 'iCs':
        case 'szCs':
        case 'tabs':
        case 'outlineLvl':
        case 'contextualSpacing':
        case 'tblStyleColBandSize':
        case 'tblStyleRowBandSize':
        case 'webHidden':
        case 'pageBreakBefore':
        case 'suppressLineNumbers':
        case 'keepLines':
        case 'keepNext':
        case 'lang':
        case 'widowControl':
        case 'bidi':
        case 'rtl':
        case 'noProof':
          break;
        default:
          if (this.options.debug) { console.warn(`DOCX: Unknown document element: ${elem.localName}.${c.localName}`); }
          break;
      }
    });

    return style;
  }

  static parseTrHeight(node: Element, output: Record<string, string>) {
    switch (XmlParser.attr(node, 'hRule')) {
      case 'exact':
        output.height = XmlParser.lengthAttr(node, 'val');
        break;

      case 'atLeast':
      default:
        output.height = XmlParser.lengthAttr(node, 'val');
        break;
    }
  }

  static parseUnderline(node: Element, style: Record<string, string>) {
    const val = XmlParser.attr(node, 'val');

    if (val == null) { return; }

    switch (val) {
      case 'dash':
      case 'dashDotDotHeavy':
      case 'dashDotHeavy':
      case 'dashedHeavy':
      case 'dashLong':
      case 'dashLongHeavy':
      case 'dotDash':
      case 'dotDotDash':
        style['text-decoration-style'] = 'dashed';
        break;

      case 'dotted':
      case 'dottedHeavy':
        style['text-decoration-style'] = 'dotted';
        break;

      case 'double':
        style['text-decoration-style'] = 'double';
        break;

      case 'single':
      case 'thick':
        style['text-decoration'] = 'underline';
        break;

      case 'wave':
      case 'wavyDouble':
      case 'wavyHeavy':
        style['text-decoration-style'] = 'wavy';
        break;

      case 'words':
        style['text-decoration'] = 'underline';
        break;

      case 'none':
        style['text-decoration'] = 'none';
        break;

      default: break;
    }

    const col = xmlUtil.colorAttr(node, 'color');

    if (col) { style['text-decoration-color'] = col; }
  }

  static parseIndentation(node: Element, style: Record<string, string>) {
    const firstLine = XmlParser.lengthAttr(node, 'firstLine');
    const hanging = XmlParser.lengthAttr(node, 'hanging');
    const left = XmlParser.lengthAttr(node, 'left');
    const start = XmlParser.lengthAttr(node, 'start');
    const right = XmlParser.lengthAttr(node, 'right');
    const end = XmlParser.lengthAttr(node, 'end');

    if (firstLine) style['text-indent'] = firstLine;
    if (hanging) style['text-indent'] = `-${hanging}`;
    if (left || start) style['margin-left'] = left || start;
    if (right || end) style['margin-right'] = right || end;
  }

  static parseFont(node: Element, style: Record<string, string>) {
    const ascii = XmlParser.attr(node, 'ascii');
    const asciiTheme = values.themeValue(node, 'asciiTheme');

    const fonts = [ascii, asciiTheme].filter((x) => x).join(', ');

    if (fonts.length > 0) { style['font-family'] = fonts; }
  }

  static parseBorderProperties(node: Element, output: Record<string, string>) {
    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case 'start':
        case 'left':
          output['border-left'] = values.valueOfBorder(c);
          break;

        case 'end':
        case 'right':
          output['border-right'] = values.valueOfBorder(c);
          break;

        case 'top':
          output['border-top'] = values.valueOfBorder(c);
          break;

        case 'bottom':
          output['border-bottom'] = values.valueOfBorder(c);
          break;
        default: break;
      }
    });
  }

  static parseMarginProperties(node: Element, output: Record<string, string>) {
    xmlUtil.foreach(node, (c) => {
      switch (c.localName) {
        case 'left':
          output['padding-left'] = values.valueOfMargin(c);
          break;

        case 'right':
          output['padding-right'] = values.valueOfMargin(c);
          break;

        case 'top':
          output['padding-top'] = values.valueOfMargin(c);
          break;

        case 'bottom':
          output['padding-bottom'] = values.valueOfMargin(c);
          break;
        default: break;
      }
    });
  }

  static parseSpacing(node: Element, style: Record<string, string>) {
    const before = XmlParser.lengthAttr(node, 'before');
    const after = XmlParser.lengthAttr(node, 'after');
    const line = XmlParser.intAttr(node, 'line', null);
    const lineRule = XmlParser.attr(node, 'lineRule');

    if (before) style['margin-top'] = before;
    if (after) style['margin-bottom'] = after;

    if (line !== null) {
      switch (lineRule) {
        case 'auto':
          style['line-height'] = `${(line / 240).toFixed(2)}`;
          break;

        case 'atLeast':
          style['line-height'] = `calc(100% + ${line / 20}pt)`;
          break;

        default:
          style['min-height'] = `${line / 20}pt`;
          style['line-height'] = style['min-height'];
          break;
      }
    }
  }
}
