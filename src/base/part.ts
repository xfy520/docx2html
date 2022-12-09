import { serializeXmlString } from '../parser/xml-parser';
import { Relationship } from '../types';
import Xml from './xml';

export default class Part {
  protected _xmlDocument: Document;

  rels: Relationship[];

  protected _xml: Xml;

  public path: string;

  constructor(_xml: Xml, path: string) {
    this._xml = _xml;
    this.path = path;
  }

  load(): Promise<unknown> {
    return Promise.all([
      this._xml.loadRelationships(this.path).then((rels) => {
        this.rels = rels;
      }),
      this._xml.load(this.path).then((text) => {
        const xmlDoc = this._xml.parseXmlDocument(text);

        if (this._xml.options.keepOrigin) {
          this._xmlDocument = xmlDoc;
        }
        this.parseXml(xmlDoc.firstElementChild);
      }),
    ]);
  }

  save() {
    this._xml.update(this.path, serializeXmlString(this._xmlDocument));
  }

  protected parseXml(root: Element) {
    console.log(this.path);
    console.log(root);
  }
}
