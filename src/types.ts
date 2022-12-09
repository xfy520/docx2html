export interface Options {
  inWrapper: boolean;
  ignoreWidth: boolean;
  ignoreHeight: boolean;
  ignoreFonts: boolean;
  breakPages: boolean;
  debug: boolean;
  experimental: boolean;
  className: string;
  renderHeaders: boolean;
  renderFooters: boolean;
  renderFootnotes: boolean;
  renderEndnotes: boolean;
  ignoreLastRenderedPageBreak: boolean;
  useBase64URL: boolean;
  useMathMLPolyfill: boolean;
  renderChanges: boolean;
  trimXmlDeclaration: boolean,
  keepOrigin?: boolean,
}

export interface XmlOptions {
  trimXmlDeclaration: boolean;
  keepOrigin?: boolean;
}

export interface ParserOptions {
  ignoreWidth: boolean;
  debug: boolean;
}

export enum RelationshipTypes {
  OfficeDocument = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  FontTable = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  Image = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  Numbering = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  Styles = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  StylesWithEffects = 'http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects',
  Theme = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  Settings = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  WebSettings = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  Hyperlink = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  Footnotes = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  Endnotes = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  Footer = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  Header = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  ExtendedProperties = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  CoreProperties = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  CustomProperties = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/custom-properties',
}

export interface Relationship {
  id: string,
  type: RelationshipTypes | string,
  target: string
  targetMode: '' | 'External' | string
}

export enum DomType {
  Document = 'document',
  Paragraph = 'paragraph',
  Run = 'run',
  Break = 'break',
  NoBreakHyphen = 'noBreakHyphen',
  Table = 'table',
  Row = 'row',
  Cell = 'cell',
  Hyperlink = 'hyperlink',
  Drawing = 'drawing',
  Image = 'image',
  Text = 'text',
  Tab = 'tab',
  Symbol = 'symbol',
  BookmarkStart = 'bookmarkStart',
  BookmarkEnd = 'bookmarkEnd',
  Footer = 'footer',
  Header = 'header',
  FootnoteReference = 'footnoteReference',
  EndnoteReference = 'endnoteReference',
  Footnote = 'footnote',
  Endnote = 'endnote',
  SimpleField = 'simpleField',
  ComplexField = 'complexField',
  Instruction = 'instruction',
  VmlPicture = 'vmlPicture',
  MmlMath = 'mmlMath',
  MmlMathParagraph = 'mmlMathParagraph',
  MmlFraction = 'mmlFraction',
  MmlNumerator = 'mmlNumerator',
  MmlDenominator = 'mmlDenominator',
  MmlRadical = 'mmlRadical',
  MmlBase = 'mmlBase',
  MmlDegree = 'mmlDegree',
  MmlSuperscript = 'mmlSuperscript',
  MmlSubscript = 'mmlSubscript',
  MmlSubArgument = 'mmlSubArgument',
  MmlSuperArgument = 'mmlSuperArgument',
  MmlNary = 'mmlNary',
  MmlDelimiter = 'mmlDelimiter',
  VmlElement = 'vmlElement',
  Inserted = 'inserted',
  Deleted = 'deleted',
  DeletedText = 'deletedText'
}

export interface XmlElement {
  type: DomType;
  children?: XmlElement[];
  cssStyle?: Record<string, string>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  props?: Record<string, any>;

  styleName?: string;
  className?: string;

  parent?: XmlElement;
}

export interface WmlHyperlink extends XmlElement {
  id?: string;
  href?: string;
}

export interface WmlNoteReference extends XmlElement {
  id: string;
}

export interface WmlBreak extends XmlElement {
  break: 'page' | 'lastRenderedPageBreak' | 'textWrapping';
}

export interface WmlText extends XmlElement {
  text: string;
}

export interface WmlSymbol extends XmlElement {
  font: string;
  char: string;
}

export interface WmlTableColumn {
  width?: string;
}

export interface WmlTable extends XmlElement {
  columns?: WmlTableColumn[];
  cellStyle?: Record<string, string>;

  colBandSize?: number;
  rowBandSize?: number;
}

export interface WmlTableRow extends XmlElement {
  isHeader?: boolean;
}

export interface WmlTableCell extends XmlElement {
  verticalMerge?: 'restart' | 'continue' | string;
  span?: number;
}

export interface IDomImage extends XmlElement {
  src: string;
}

export interface NumberingPicBullet {
  id: number;
  src: string;
  style?: string;
}

export interface IDomNumbering {
  id: string;
  level: number;
  pStyleName: string;
  pStyle: Record<string, string>;
  rStyle: Record<string, string>;
  levelText?: string;
  suff: string;
  format?: string;
  bullet?: NumberingPicBullet;
}

export const ns = {
  wordml: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  drawingml: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  picture: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  compatibility: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  math: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  svg: 'http://www.w3.org/2000/svg',
  mathML: 'http://www.w3.org/1998/Math/MathML',
};

export type LengthType = 'px' | 'pt' | '%' | '';
export type Length = string;

export interface Font {
  name: string;
  family: string;
}

export interface CommonProperties {
  fontSize: Length;
  color: string;
}

export type LengthUsageType = { mul: number, unit: LengthType };

export const LengthUsage: Record<string, LengthUsageType> = {
  Dxa: { mul: 0.05, unit: 'pt' }, // twips
  Emu: { mul: 1 / 12700, unit: 'pt' },
  FontSize: { mul: 0.5, unit: 'pt' },
  Border: { mul: 0.125, unit: 'pt' },
  Point: { mul: 1, unit: 'pt' },
  Percent: { mul: 0.02, unit: '%' },
  LineHeight: { mul: 1 / 240, unit: '' },
  VmlEmu: { mul: 1 / 12700, unit: '' },
};
