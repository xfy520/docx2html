import * as JSZip from 'jszip';
import { Relationship, XmlOptions } from '../types';
import { parseXmlString, XmlParser } from '../parser/xml-parser';
import { parseRelationships, splitPath } from '../utils';

function normalizePath(path: string) {
  return path.startsWith('/') ? path.substr(1) : path;
}

export default class Xml {
  xmlParser: XmlParser = new XmlParser();

  private _zip: JSZip;

  private _text: string;

  public options: XmlOptions;

  constructor(_zip: JSZip, _text: string, options: XmlOptions) {
    this._zip = _zip;
    this._text = _text;
    this.options = options;
  }

  get(path: string): JSZip.JSZipObject {
    return this._zip.files[normalizePath(path)];
  }

  update(path: string, content: unknown, options?: JSZip.JSZipFileOptions) {
    this._zip.file(path, content, options);
  }

  save(type: JSZip.OutputType = 'blob'): Promise<unknown> {
    return this._zip.generateAsync({ type });
  }

  static load(input: Blob | string, options: XmlOptions): Promise<Xml> {
    if (typeof input === 'string') {
      return Promise.resolve(new Xml(null, input, options));
    }
    return JSZip.loadAsync(input).then((zip) => new Xml(zip, '', options));
  }

  load(path: string, type: JSZip.OutputType = 'string'): Promise<string> {
    if (this._zip) {
      return this.get(path)?.async<JSZip.OutputType>(type) ?? Promise.resolve(null);
    }
    return Promise.resolve(this._text);
  }

  loadRelationships(path: string = null): Promise<Relationship[]> {
    let relsPath = '_rels/.rels';

    if (path != null) {
      const [f, fn] = splitPath(path);
      relsPath = `${f}_rels/${fn}.rels`;
    }

    return this.load(relsPath)
      .then((txt) => (txt ? parseRelationships(this.parseXmlDocument(txt).firstElementChild) : null));
  }

  parseXmlDocument(txt: string): Document {
    return parseXmlString(txt, this.options.trimXmlDeclaration);
  }
}
