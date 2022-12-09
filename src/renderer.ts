import Part from './base/part';
import { WmlBookmarkStart } from './document/bookmarks';
import { DocumentElement } from './document/document';
import { WmlParagraph } from './document/paragraph';
import { WmlRun } from './document/run';
import { FooterHeaderReference, SectionProperties } from './document/section';
import FontTablePart from './font-table/font-table';
import { BaseHeaderFooterPart } from './header-footer/header-footer-parts';
import { WmlBaseNote, WmlFootnote } from './notes/elements';
import { IDomStyle } from './styles/styles-part';
import ThemePart from './theme/theme-part';
import {
  CommonProperties,
  DomType, IDomImage, IDomNumbering, ns, Options, WmlBreak, WmlHyperlink, WmlNoteReference, WmlSymbol, WmlTable, WmlTableCell, WmlTableColumn, WmlText, XmlElement,
} from './types';
import {
  asArray,
  computePixelToPoint,
  escapeClassName, isString, keyBy, mergeDeep, updateTabStop,
} from './utils';
import { VmlElement } from './vml/vml';
import Word from './word';

interface CellPos {
  col: number;
  row: number;
}

type CellVerticalMergeType = Record<number, HTMLTableCellElement>;

type ChildType = Node | string;

function removeAllElements(elem: HTMLElement) {
  elem.innerHTML = '';
}

function appendComment(elem: HTMLElement, comment: string) {
  elem.appendChild(document.createComment(comment));
}

function appendChildren(elem: Element, children: (Node | string)[]) {
  children.forEach((c) => elem.appendChild(isString(c) ? document.createTextNode(c) : c));
}

function createElementNS(ns: string, tagName: string, props?: Partial<Record<string, unknown>>, children?: ChildType[]): any {
  const result = ns ? document.createElementNS(ns, tagName) : document.createElement(tagName);
  Object.assign(result, props);
  if (children) {
    appendChildren(result, children);
  }
  return result;
}

function createElement<T extends keyof HTMLElementTagNameMap>(
  tagName: T,
  props?: Partial<Record<keyof HTMLElementTagNameMap[T], any>>,
  children?: ChildType[],
): HTMLElementTagNameMap[T] {
  return createElementNS(undefined, tagName, props, children);
}

function createStyleElement(cssText: string) {
  return createElement('style', { innerHTML: cssText });
}

function createSvgElement<T extends keyof SVGElementTagNameMap>(
  tagName: T,
  props?: Partial<Record<keyof SVGElementTagNameMap[T], any>>,
  children?: ChildType[],
): SVGElementTagNameMap[T] {
  return createElementNS(ns.svg, tagName, props, children);
}

function findParent<T extends XmlElement>(elem: XmlElement, type: DomType): T {
  let { parent } = elem;

  while (parent != null && parent.type !== type) { parent = parent.parent; }

  return <T>parent;
}

export default class Renderer {
  public dom: Document;

  className = 'docx';

  rootSelector: string;

  word: Word;

  options: Options;

  styleMap: Record<string, IDomStyle> = {};

  currentPart: Part = null;

  tableVerticalMerges: CellVerticalMergeType[] = [];

  currentVerticalMerge: CellVerticalMergeType = null;

  tableCellPositions: CellPos[] = [];

  currentCellPosition: CellPos = null;

  footnoteMap: Record<string, WmlFootnote> = {};

  endnoteMap: Record<string, WmlFootnote> = {};

  currentFootnoteIds: string[];

  currentEndnoteIds: string[] = [];

  usedHederFooterParts: unknown[] = [];

  defaultTabSize: string;

  currentTabs: any[] = [];

  tabsTimeout = 0;

  constructor(dom: Document) {
    this.dom = dom;
  }

  render(word: Word, bodyDom: HTMLElement, _styleDom: HTMLElement, options: Options) {
    this.word = word;
    this.options = options;
    this.className = options.className;
    this.rootSelector = options.inWrapper ? `.${this.className}-wrapper` : ':root';
    this.styleMap = null;

    const styleDom = _styleDom || bodyDom;

    removeAllElements(styleDom);
    removeAllElements(bodyDom);

    appendComment(styleDom, 'docxjs library predefined styles');
    styleDom.appendChild(this.renderDefaultStyle());

    if (!window.MathMLElement && options.useMathMLPolyfill) {
      appendComment(styleDom, 'docxjs mathml polyfill styles');
      styleDom.appendChild(createStyleElement(''));
    }

    if (word.themePart) {
      appendComment(styleDom, 'docxjs document theme values');
      this.renderTheme(word.themePart, styleDom);
    }

    if (word.stylesPart != null) {
      this.styleMap = this.processStyles(word.stylesPart.styles);

      appendComment(styleDom, 'docxjs document styles');
      styleDom.appendChild(this.renderStyles(word.stylesPart.styles));
    }

    if (word.numberingPart) {
      this.prodessNumberings(word.numberingPart.domNumberings);

      appendComment(styleDom, 'docxjs document numbering styles');
      styleDom.appendChild(this.renderNumbering(word.numberingPart.domNumberings, styleDom));
      // styleContainer.appendChild(this.renderNumbering2(document.numberingPart, styleContainer));
    }

    if (word.footnotesPart) {
      this.footnoteMap = keyBy(word.footnotesPart.notes, (x) => x.id);
    }

    if (word.endnotesPart) {
      this.endnoteMap = keyBy(word.endnotesPart.notes, (x) => x.id);
    }

    if (word.settingsPart) {
      this.defaultTabSize = word.settingsPart.settings?.defaultTabStop;
    }

    if (!options.ignoreFonts && word.fontTablePart) { this.renderFontTable(word.fontTablePart, styleDom); }

    const sectionElements = this.renderSections(word.documentPart.body);

    if (this.options.inWrapper) {
      bodyDom.appendChild(this.renderWrapper(sectionElements));
    } else {
      appendChildren(bodyDom, sectionElements);
    }

    this.refreshTabStops();
  }

  renderWrapper(children: HTMLElement[]) {
    return this.createElement('div', { className: `${this.className}-wrapper` }, children);
  }

  renderSections(document: DocumentElement): HTMLElement[] {
    const result = [];

    this.processElement(document);
    const sections = this.splitBySection(document.children);
    let prevProps = null;

    for (let i = 0, l = sections.length; i < l;) {
      this.currentFootnoteIds = [];

      const section = sections[i];
      const props = section.sectProps || document.props;
      const sectionElement = this.createSection(this.className, props);
      Renderer.renderStyleValues(document.cssStyle, sectionElement);

      this.options.renderHeaders && this.renderHeaderFooter(
        props.headerRefs,
        props,
        result.length,
        prevProps !== props,
        sectionElement,
      );

      const contentElement = this.createElement('article');
      this.renderElements(section.elements, contentElement);
      sectionElement.appendChild(contentElement);

      if (this.options.renderFootnotes) {
        this.renderNotes(this.currentFootnoteIds, this.footnoteMap, sectionElement);
      }

      if (this.options.renderEndnotes && i === l - 1) {
        this.renderNotes(this.currentEndnoteIds, this.endnoteMap, sectionElement);
      }

      this.options.renderFooters && this.renderHeaderFooter(
        props.footerRefs,
        props,
        result.length,
        prevProps !== props,
        sectionElement,
      );

      result.push(sectionElement);
      prevProps = props;
      i += 1;
    }

    return result;
  }

  renderNotes(noteIds: string[], notesMap: Record<string, WmlBaseNote>, into: HTMLElement) {
    const notes = noteIds.map((id) => notesMap[id]).filter((x) => x);

    if (notes.length > 0) {
      const result = this.createElement('ol', null, this.renderElements(notes));
      into.appendChild(result);
    }
  }

  renderHeaderFooter(refs: FooterHeaderReference[], props: SectionProperties, page: number, firstOfSection: boolean, into: HTMLElement) {
    if (!refs) return;

    const ref = (props.titlePage && firstOfSection ? refs.find((x) => x.type === 'first') : null)
      ?? (page % 2 === 1 ? refs.find((x) => x.type === 'even') : null)
      ?? refs.find((x) => x.type === 'default');

    const part = ref && this.word.findPartByRelId(ref.id, this.word.documentPart) as BaseHeaderFooterPart;

    if (part) {
      this.currentPart = part;
      if (!this.usedHederFooterParts.includes(part.path)) {
        this.processElement(part.rootElement);
        this.usedHederFooterParts.push(part.path);
      }
      this.renderElements([part.rootElement], into);
      this.currentPart = null;
    }
  }

  renderElements(elems: XmlElement[], into?: Element): Node[] {
    if (elems == null) { return null; }

    const result = elems.flatMap((e) => this.renderElement(e)).filter((e) => e != null);

    if (into) { appendChildren(into, result); }

    return result;
  }

  renderEndnoteReference(elem: WmlNoteReference) {
    const result = this.createElement('sup');
    this.currentEndnoteIds.push(elem.id);
    result.textContent = `${this.currentEndnoteIds.length}`;
    return result;
  }

  renderRun(elem: WmlRun) {
    if (elem.fieldRun) { return null; }

    const result = this.createElement('span');

    if (elem.id) { result.id = elem.id; }

    this.renderClass(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);

    if (elem.verticalAlign) {
      const wrapper = this.createElement(elem.verticalAlign as any);
      this.renderChildren(elem, wrapper);
      result.appendChild(wrapper);
    } else {
      this.renderChildren(elem, result);
    }

    return result;
  }

  renderTableColumns(columns: WmlTableColumn[]) {
    const result = this.createElement('colgroup');

    for (const col of columns) {
      const colElem = this.createElement('col');

      if (col.width) { colElem.style.width = col.width; }

      result.appendChild(colElem);
    }

    return result;
  }

  renderTableRow(elem: XmlElement) {
    const result = this.createElement('tr');

    this.currentCellPosition.col = 0;

    this.renderClass(elem, result);
    this.renderChildren(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);

    this.currentCellPosition.row += 1;

    return result;
  }

  renderTable(elem: WmlTable) {
    const result = this.createElement('table');

    this.tableCellPositions.push(this.currentCellPosition);
    this.tableVerticalMerges.push(this.currentVerticalMerge);
    this.currentVerticalMerge = {};
    this.currentCellPosition = { col: 0, row: 0 };

    if (elem.columns) { result.appendChild(this.renderTableColumns(elem.columns)); }

    this.renderClass(elem, result);
    this.renderChildren(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);

    this.currentVerticalMerge = this.tableVerticalMerges.pop();
    this.currentCellPosition = this.tableCellPositions.pop();

    return result;
  }

  renderTableCell(elem: WmlTableCell) {
    const result = this.createElement('td');

    const key = this.currentCellPosition.col;

    if (elem.verticalMerge) {
      if (elem.verticalMerge === 'restart') {
        this.currentVerticalMerge[key] = result;
        result.rowSpan = 1;
      } else if (this.currentVerticalMerge[key]) {
        this.currentVerticalMerge[key].rowSpan += 1;
        result.style.display = 'none';
      }
    } else {
      this.currentVerticalMerge[key] = null;
    }

    this.renderClass(elem, result);
    this.renderChildren(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);

    if (elem.span) { result.colSpan = elem.span; }

    this.currentCellPosition.col += result.colSpan;

    return result;
  }

  renderHyperlink(elem: WmlHyperlink) {
    const result = this.createElement('a');

    this.renderChildren(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);

    if (elem.href) {
      result.href = elem.href;
    } else if (elem.id) {
      const rel = this.word.documentPart.rels
        .find((it) => it.id === elem.id && it.targetMode === 'External');
      result.href = rel?.target;
    }

    return result;
  }

  renderDrawing(elem: XmlElement) {
    const result = this.createElement('div');

    result.style.display = 'inline-block';
    result.style.position = 'relative';
    result.style.textIndent = '0px';

    this.renderChildren(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);

    return result;
  }

  renderImage(elem: IDomImage) {
    const result = this.createElement('img');

    Renderer.renderStyleValues(elem.cssStyle, result);

    if (this.word) {
      this.word.loadDocumentImage(elem.src, this.currentPart).then((x) => {
        result.src = x as string;
      });
    }

    return result;
  }

  renderText(elem: WmlText) {
    return this.dom.createTextNode(elem.text);
  }

  renderDeletedText(elem: WmlText) {
    return this.options.renderEndnotes ? this.dom.createTextNode(elem.text) : null;
  }

  renderBreak(elem: WmlBreak) {
    if (elem.break === 'textWrapping') {
      return this.createElement('br');
    }

    return null;
  }

  renderSymbol(elem: WmlSymbol) {
    const span = this.createElement('span');
    span.style.fontFamily = elem.font;
    span.innerHTML = `&#x${elem.char};`;
    return span;
  }

  tabStopClass() {
    return `${this.className}-tab-stop`;
  }

  renderTab(elem: XmlElement) {
    const tabSpan = this.createElement('span');

    tabSpan.innerHTML = '&emsp;';// "&nbsp;";

    if (this.options.experimental) {
      tabSpan.className = this.tabStopClass();
      const stops = findParent<WmlParagraph>(elem, DomType.Paragraph)?.tabs;
      this.currentTabs.push({ stops, span: tabSpan });
    }

    return tabSpan;
  }

  renderFootnoteReference(elem: WmlNoteReference) {
    const result = this.createElement('sup');
    this.currentFootnoteIds.push(elem.id);
    result.textContent = `${this.currentFootnoteIds.length}`;
    return result;
  }

  renderElement(elem: XmlElement): Node | Node[] {
    switch (elem.type) {
      case DomType.Paragraph:
        return this.renderParagraph(elem as WmlParagraph);

      case DomType.BookmarkStart:
        return this.renderBookmarkStart(elem as WmlBookmarkStart);

      case DomType.BookmarkEnd:
        return null; // ignore bookmark end

      case DomType.Run:
        return this.renderRun(elem as WmlRun);

      case DomType.Table:
        return this.renderTable(elem);

      case DomType.Row:
        return this.renderTableRow(elem);

      case DomType.Cell:
        return this.renderTableCell(elem);

      case DomType.Hyperlink:
        return this.renderHyperlink(elem);

      case DomType.Drawing:
        return this.renderDrawing(elem);

      case DomType.Image:
        return this.renderImage(elem as IDomImage);

      case DomType.Text:
        return this.renderText(elem as WmlText);

      case DomType.DeletedText:
        return this.renderDeletedText(elem as WmlText);

      case DomType.Tab:
        return this.renderTab(elem);

      case DomType.Symbol:
        return this.renderSymbol(elem as WmlSymbol);

      case DomType.Break:
        return this.renderBreak(elem as WmlBreak);

      case DomType.Footer:
        return this.renderContainer(elem, 'footer');

      case DomType.Header:
        return this.renderContainer(elem, 'header');

      case DomType.Footnote:
      case DomType.Endnote:
        return this.renderContainer(elem, 'li');

      case DomType.FootnoteReference:
        return this.renderFootnoteReference(elem as WmlNoteReference);

      case DomType.EndnoteReference:
        return this.renderEndnoteReference(elem as WmlNoteReference);

      case DomType.NoBreakHyphen:
        return this.createElement('wbr');

      case DomType.VmlPicture:
        return this.renderVmlPicture(elem);

      case DomType.VmlElement:
        return this.renderVmlElement(elem as VmlElement);

      case DomType.MmlMath:
        return this.renderContainerNS(elem, ns.mathML, 'math', { xmlns: ns.mathML });

      case DomType.MmlMathParagraph:
        return this.renderContainer(elem, 'span');

      case DomType.MmlFraction:
        return this.renderContainerNS(elem, ns.mathML, 'mfrac');

      case DomType.MmlNumerator:
      case DomType.MmlDenominator:
        return this.renderContainerNS(elem, ns.mathML, 'mrow');

      case DomType.MmlRadical:
        return this.renderMmlRadical(elem);

      case DomType.MmlDegree:
        return this.renderContainerNS(elem, ns.mathML, 'mn');

      case DomType.MmlSuperscript:
        return this.renderContainerNS(elem, ns.mathML, 'msup');

      case DomType.MmlSubscript:
        return this.renderContainerNS(elem, ns.mathML, 'msub');

      case DomType.MmlBase:
        return this.renderContainerNS(elem, ns.mathML, 'mrow');

      case DomType.MmlSuperArgument:
        return this.renderContainerNS(elem, ns.mathML, 'mn');

      case DomType.MmlSubArgument:
        return this.renderContainerNS(elem, ns.mathML, 'mn');

      case DomType.MmlDelimiter:
        return this.renderMmlDelimiter(elem);

      case DomType.MmlNary:
        return this.renderMmlNary(elem);

      case DomType.Inserted:
        return this.renderInserted(elem);

      case DomType.Deleted:
        return this.renderDeleted(elem);
      default: break;
    }

    return null;
  }

  renderVmlPicture(elem: XmlElement) {
    const result = createElement('div');
    this.renderChildren(elem, result);
    return result;
  }

  renderVmlElement(elem: VmlElement): SVGElement {
    const container = createSvgElement('svg');

    container.setAttribute('style', elem.cssStyleText);

    const result = createSvgElement(elem.tagName as any);
    Object.entries(elem.attrs).forEach(([k, v]) => result.setAttribute(k, v));

    if (elem.imageHref?.id) {
      this.word?.loadDocumentImage(elem.imageHref.id, this.currentPart)
        .then((x) => result.setAttribute('href', x));
    }

    container.appendChild(result);

    setTimeout(() => {
      const bb = (container.firstElementChild as any).getBBox();

      container.setAttribute('width', `${Math.ceil(bb.x + bb.width)}`);
      container.setAttribute('height', `${Math.ceil(bb.y + bb.height)}`);
    }, 0);

    return container;
  }

  renderMmlRadical(elem: XmlElement): HTMLElement {
    const base = elem.children.find((el) => el.type === DomType.MmlBase);

    if (elem.props?.hideDegree) {
      return createElementNS(ns.mathML, 'msqrt', null, this.renderElements([base]));
    }

    const degree = elem.children.find((el) => el.type === DomType.MmlDegree);
    return createElementNS(ns.mathML, 'mroot', null, this.renderElements([base, degree]));
  }

  renderContainerNS(elem: XmlElement, ns: string, tagName: string, props?: Record<string, any>) {
    return createElementNS(ns, tagName, props, this.renderChildren(elem));
  }

  renderMmlDelimiter(elem: XmlElement): HTMLElement {
    const children = [];

    children.push(createElementNS(ns.mathML, 'mo', null, [elem.props.beginChar ?? '(']));
    children.push(...this.renderElements(elem.children));
    children.push(createElementNS(ns.mathML, 'mo', null, [elem.props.endChar ?? ')']));

    return createElementNS(ns.mathML, 'mrow', null, children);
  }

  renderMmlNary(elem: XmlElement): HTMLElement {
    const children = [];
    const grouped = keyBy(elem.children, (x) => x.type);

    const sup = grouped[DomType.MmlSuperArgument];
    const sub = grouped[DomType.MmlSubArgument];
    const supElem = sup ? createElementNS(ns.mathML, 'mo', null, asArray(this.renderElement(sup))) : null;
    const subElem = sub ? createElementNS(ns.mathML, 'mo', null, asArray(this.renderElement(sub))) : null;

    if (elem.props?.char) {
      const charElem = createElementNS(ns.mathML, 'mo', null, [elem.props.char]);

      if (supElem || subElem) {
        children.push(createElementNS(ns.mathML, 'munderover', null, [charElem, subElem, supElem]));
      } else if (supElem) {
        children.push(createElementNS(ns.mathML, 'mover', null, [charElem, supElem]));
      } else if (subElem) {
        children.push(createElementNS(ns.mathML, 'munder', null, [charElem, subElem]));
      } else {
        children.push(charElem);
      }
    }

    children.push(...this.renderElements(grouped[DomType.MmlBase].children));

    return createElementNS(ns.mathML, 'mrow', null, children);
  }

  renderInserted(elem: XmlElement): Node | Node[] {
    if (this.options.renderChanges) { return this.renderContainer(elem, 'ins'); }

    return this.renderChildren(elem);
  }

  renderDeleted(elem: XmlElement): Node {
    if (this.options.renderChanges) { return this.renderContainer(elem, 'del'); }

    return null;
  }

  renderContainer(elem: XmlElement, tagName: keyof HTMLElementTagNameMap, props?: Record<string, any>) {
    return this.createElement(tagName, props, this.renderChildren(elem));
  }

  renderBookmarkStart(elem: WmlBookmarkStart): HTMLElement {
    const result = this.createElement('span');
    result.id = elem.name;
    return result;
  }

  renderParagraph(elem: WmlParagraph) {
    const result = this.createElement('p');

    const style = this.findStyle(elem.styleName);
    elem.tabs ??= style?.paragraphProps?.tabs; // TODO

    this.renderClass(elem, result);
    this.renderChildren(elem, result);
    Renderer.renderStyleValues(elem.cssStyle, result);
    Renderer.renderCommonProperties(result.style, elem);

    const numbering = elem.numbering ?? style?.paragraphProps?.numbering;

    if (numbering) {
      result.classList.add(this.numberingClass(numbering.id, numbering.level));
    }

    return result;
  }

  renderChildren(elem: XmlElement, into?: Element): Node[] {
    return this.renderElements(elem.children, into);
  }

  static renderCommonProperties(style: any, props: CommonProperties) {
    if (props == null) { return; }

    if (props.color) {
      style.color = props.color;
    }

    if (props.fontSize) {
      style['font-size'] = props.fontSize;
    }
  }

  renderClass(input: XmlElement, ouput: HTMLElement) {
    if (input.className) { ouput.className = input.className; }

    if (input.styleName) { ouput.classList.add(this.processStyleName(input.styleName)); }
  }

  static renderStyleValues(style: Record<string, string>, ouput: HTMLElement) {
    Object.assign(ouput.style, style);
  }

  createSection(className: string, props: SectionProperties) {
    const elem = this.createElement('section', { className });

    if (props) {
      if (props.pageMargins) {
        elem.style.paddingLeft = props.pageMargins.left;
        elem.style.paddingRight = props.pageMargins.right;
        elem.style.paddingTop = props.pageMargins.top;
        elem.style.paddingBottom = props.pageMargins.bottom;
      }

      if (props.pageSize) {
        if (!this.options.ignoreWidth) { elem.style.width = props.pageSize.width; }
        if (!this.options.ignoreHeight) { elem.style.minHeight = props.pageSize.height; }
      }

      if (props.columns && props.columns.numberOfColumns) {
        elem.style.columnCount = `${props.columns.numberOfColumns}`;
        elem.style.columnGap = props.columns.space;

        if (props.columns.separator) {
          elem.style.columnRule = '1px solid black';
        }
      }
    }

    return elem;
  }

  splitBySection(elements: XmlElement[]): { sectProps: SectionProperties, elements: XmlElement[] }[] {
    let current = { sectProps: null, elements: [] };
    const result = [current];

    for (const elem of elements) {
      if (elem.type === DomType.Paragraph) {
        const s = this.findStyle((elem as WmlParagraph).styleName);

        if (s?.paragraphProps?.pageBreakBefore) {
          // @ts-ignore
          current.sectProps = sectProps;
          current = { sectProps: null, elements: [] };
          result.push(current);
        }
      }

      current.elements.push(elem);

      if (elem.type === DomType.Paragraph) {
        const p = elem as WmlParagraph;

        const sectProps = p.sectionProps;
        let pBreakIndex = -1;
        let rBreakIndex = -1;

        if (this.options.breakPages && p.children) {
          pBreakIndex = p.children.findIndex((r) => {
            rBreakIndex = r.children?.findIndex(this.isPageBreakElement.bind(this)) ?? -1;
            return rBreakIndex !== -1;
          });
        }

        if (sectProps || pBreakIndex !== -1) {
          current.sectProps = sectProps;
          current = { sectProps: null, elements: [] };
          result.push(current);
        }

        if (pBreakIndex !== -1) {
          const breakRun = p.children[pBreakIndex];
          const splitRun = rBreakIndex < breakRun.children.length - 1;

          if (pBreakIndex < p.children.length - 1 || splitRun) {
            const { children } = elem;
            const newParagraph = { ...elem, children: children.slice(pBreakIndex) };
            elem.children = children.slice(0, pBreakIndex);
            current.elements.push(newParagraph);

            if (splitRun) {
              const runChildren = breakRun.children;
              const newRun = { ...breakRun, children: runChildren.slice(0, rBreakIndex) };
              elem.children.push(newRun);
              breakRun.children = runChildren.slice(rBreakIndex);
            }
          }
        }
      }
    }

    let currentSectProps = null;

    for (let i = result.length - 1; i >= 0;) {
      if (result[i].sectProps == null) {
        result[i].sectProps = currentSectProps;
      } else {
        currentSectProps = result[i].sectProps;
      }
      i -= 1;
    }

    return result;
  }

  isPageBreakElement(elem: XmlElement): boolean {
    if (elem.type !== DomType.Break) { return false; }

    if ((elem as WmlBreak).break === 'lastRenderedPageBreak') { return !this.options.ignoreLastRenderedPageBreak; }

    return (elem as WmlBreak).break === 'page';
  }

  processElement(element: XmlElement) {
    if (element.children) {
      for (const e of element.children) {
        e.parent = element;

        if (e.type === DomType.Table) {
          this.processTable(e);
        } else {
          this.processElement(e);
        }
      }
    }
  }

  processTable(table: WmlTable) {
    for (const r of table.children) {
      for (const c of r.children) {
        c.cssStyle = Renderer.copyStyleProperties(table.cellStyle, c.cssStyle, [
          'border-left', 'border-right', 'border-top', 'border-bottom',
          'padding-left', 'padding-right', 'padding-top', 'padding-bottom',
        ]);

        this.processElement(c);
      }
    }
  }

  renderFontTable(fontsPart: FontTablePart, styleContainer: HTMLElement) {
    for (const f of fontsPart.fonts) {
      for (const ref of f.embedFontRefs) {
        this.word.loadFont(ref.id, ref.key).then((fontData) => {
          const cssValues = {
            'font-family': f.name,
            src: `url(${fontData})`,
          };

          if (ref.type === 'bold' || ref.type === 'boldItalic') {
            cssValues['font-weight'] = 'bold';
          }

          if (ref.type === 'italic' || ref.type === 'boldItalic') {
            cssValues['font-style'] = 'italic';
          }

          appendComment(styleContainer, `docxjs ${f.name} font`);
          const cssText = Renderer.styleToString('@font-face', cssValues);
          styleContainer.appendChild(createStyleElement(cssText));
          this.refreshTabStops();
        });
      }
    }
  }

  refreshTabStops() {
    if (!this.options.experimental) { return; }

    clearTimeout(this.tabsTimeout);

    this.tabsTimeout = setTimeout(() => {
      const pixelToPoint = computePixelToPoint();

      for (const tab of this.currentTabs) {
        updateTabStop(tab.span, tab.stops, this.defaultTabSize, pixelToPoint);
      }
    }, 500);
  }

  renderNumbering(numberings: IDomNumbering[], styleContainer: HTMLElement) {
    let styleText = '';
    const rootCounters = [];

    for (const num of numberings) {
      const selector = `p.${this.numberingClass(num.id, num.level)}`;
      let listStyleType = 'none';

      if (num.bullet) {
        const valiable = `--${this.className}-${num.bullet.src}`.toLowerCase();

        styleText += Renderer.styleToString(`${selector}:before`, {
          content: "' '",
          display: 'inline-block',
          background: `var(${valiable})`,
        }, num.bullet.style);

        this.word.loadNumberingImage(num.bullet.src).then((data) => {
          const text = `${this.rootSelector} { ${valiable}: url(${data}) }`;
          styleContainer.appendChild(createStyleElement(text));
        });
      } else if (num.levelText) {
        const counter = this.numberingCounter(num.id, num.level);

        if (num.level > 0) {
          styleText += Renderer.styleToString(`p.${this.numberingClass(num.id, num.level - 1)}`, {
            'counter-reset': counter,
          });
        } else {
          rootCounters.push(counter);
        }

        styleText += Renderer.styleToString(`${selector}:before`, {
          content: this.levelTextToContent(num.levelText, num.suff, num.id, Renderer.numFormatToCssValue(num.format)),
          'counter-increment': counter,
          ...num.rStyle,
        });
      } else {
        listStyleType = Renderer.numFormatToCssValue(num.format);
      }

      styleText += Renderer.styleToString(selector, {
        display: 'list-item',
        'list-style-position': 'inside',
        'list-style-type': listStyleType,
        ...num.pStyle,
      });
    }

    if (rootCounters.length > 0) {
      styleText += Renderer.styleToString(this.rootSelector, {
        'counter-reset': rootCounters.join(' '),
      });
    }

    return createStyleElement(styleText);
  }

  static numFormatToCssValue(format: string) {
    const mapping = {
      none: 'none',
      bullet: 'disc',
      decimal: 'decimal',
      lowerLetter: 'lower-alpha',
      upperLetter: 'upper-alpha',
      lowerRoman: 'lower-roman',
      upperRoman: 'upper-roman',
    };

    return mapping[format] || format;
  }

  levelTextToContent(text: string, suff: string, id: string, numformat: string) {
    const suffMap = {
      tab: '\\9',
      space: '\\a0',
    };

    const result = text.replace(/%\d*/g, (s) => {
      const lvl = parseInt(s.substring(1), 10) - 1;
      return `"counter(${this.numberingCounter(id, lvl)}, ${numformat})"`;
    });

    return `"${result}${suffMap[suff] ?? ''}"`;
  }

  numberingCounter(id: string, lvl: number) {
    return `${this.className}-num-${id}-${lvl}`;
  }

  numberingClass(id: string, lvl: number) {
    return `${this.className}-num-${id}-${lvl}`;
  }

  prodessNumberings(numberings: IDomNumbering[]) {
    for (const num of numberings.filter((n) => n.pStyleName)) {
      const style = this.findStyle(num.pStyleName);

      if (style?.paragraphProps?.numbering) {
        style.paragraphProps.numbering.level = num.level;
      }
    }
  }

  findStyle(styleName: string) {
    return styleName && this.styleMap?.[styleName];
  }

  renderStyles(styles: IDomStyle[]): HTMLElement {
    let styleText = '';
    const stylesMap = this.styleMap;
    const defautStyles = keyBy(styles.filter((s) => s.isDefault), (s) => s.target);

    for (const style of styles) {
      let subStyles = style.styles;

      if (style.linked) {
        const linkedStyle = style.linked && stylesMap[style.linked];

        if (linkedStyle) { subStyles = subStyles.concat(linkedStyle.styles); } else if (this.options.debug) { console.warn(`Can't find linked style ${style.linked}`); }
      }

      for (const subStyle of subStyles) {
        let selector = `${style.target ?? ''}.${style.cssName}`; // ${subStyle.mod ?? ''}

        if (style.target !== subStyle.target) { selector += ` ${subStyle.target}`; }

        if (defautStyles[style.target] === style) { selector = `.${this.className} ${style.target}, ${selector}`; }

        styleText += Renderer.styleToString(selector, subStyle.values);
      }
    }

    return createStyleElement(styleText);
  }

  processStyles(styles: IDomStyle[]) {
    const stylesMap = keyBy(styles.filter((x) => x.id != null), (x) => x.id);

    for (const style of styles.filter((x) => x.basedOn)) {
      const baseStyle = stylesMap[style.basedOn];

      if (baseStyle) {
        style.paragraphProps = mergeDeep(style.paragraphProps, baseStyle.paragraphProps);
        style.runProps = mergeDeep(style.runProps, baseStyle.runProps);

        for (const baseValues of baseStyle.styles) {
          const styleValues = style.styles.find((x) => x.target === baseValues.target);

          if (styleValues) {
            Renderer.copyStyleProperties(baseValues.values, styleValues.values);
          } else {
            style.styles.push({ ...baseValues, values: { ...baseValues.values } });
          }
        }
      } else if (this.options.debug) { console.warn(`Can't find base style ${style.basedOn}`); }
    }

    for (const style of styles) {
      style.cssName = this.processStyleName(style.id);
    }

    return stylesMap;
  }

  processStyleName(className: string): string {
    return className ? `${this.className}_${escapeClassName(className)}` : this.className;
  }

  static copyStyleProperties(input: Record<string, string>, output: Record<string, string>, attrs: string[] = null): Record<string, string> {
    if (!input) { return output; }

    if (output == null) output = {};
    if (attrs == null) attrs = Object.getOwnPropertyNames(input);

    for (const key of attrs) {
      if (input.hasOwnProperty(key) && !output.hasOwnProperty(key)) { output[key] = input[key]; }
    }

    return output;
  }

  renderDefaultStyle() {
    const c = this.className;
    const styleText = `
.${c}-wrapper { background: gray; padding: 30px; padding-bottom: 0px; display: flex; flex-flow: column; align-items: center; }
.${c}-wrapper>section.${c} { background: white; box-shadow: 0 0 10px rgba(0, 0, 0, 0.5); margin-bottom: 30px; }
.${c} { color: black; }
section.${c} { box-sizing: border-box; display: flex; flex-flow: column nowrap; position: relative; overflow: hidden; }
section.${c}>article { margin-bottom: auto; }
.${c} table { border-collapse: collapse; }
.${c} table td, .${c} table th { vertical-align: top; }
.${c} p { margin: 0pt; min-height: 1em; }
.${c} span { white-space: pre-wrap; overflow-wrap: break-word; }
.${c} a { color: inherit; text-decoration: inherit; }
`;

    return createStyleElement(styleText);
  }

  renderTheme(themePart: ThemePart, styleContainer: HTMLElement) {
    const variables = {};
    const fontScheme = themePart.theme?.fontScheme;

    if (fontScheme) {
      if (fontScheme.majorFont) {
        variables['--docx-majorHAnsi-font'] = fontScheme.majorFont.latinTypeface;
      }

      if (fontScheme.minorFont) {
        variables['--docx-minorHAnsi-font'] = fontScheme.minorFont.latinTypeface;
      }
    }

    const colorScheme = themePart.theme?.colorScheme;

    if (colorScheme) {
      for (const [k, v] of Object.entries(colorScheme.colors)) {
        variables[`--docx-${k}-color`] = `#${v}`;
      }
    }

    const cssText = Renderer.styleToString(`.${this.className}`, variables);
    styleContainer.appendChild(createStyleElement(cssText));
  }

  static styleToString(selectors: string, values: Record<string, string>, cssText: string = null) {
    let result = `${selectors} {\r\n`;

    for (const key in values) {
      result += `  ${key}: ${values[key]};\r\n`;
    }

    if (cssText) { result += cssText; }

    return `${result}}\r\n`;
  }

  createElement = createElement;
}
