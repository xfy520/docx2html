import * as JSZip from 'jszip';
import { OutputType } from 'jszip';
import Part from './base/part';
import Xml from './base/xml';
import { topLevelRels } from './constant';
import CorePropsPart from './document-props/core-props-part';
import CustomPropsPart from './document-props/custom-props-part';
import ExtendedPropsPart from './document-props/extended-props-part';
import DocumentPart from './document/document-part';
import FontTablePart from './font-table/font-table';
import { FooterPart, HeaderPart } from './header-footer/header-footer-parts';
import { EndnotesPart, FootnotesPart } from './notes/notes-parts';
import NumberingPart from './numbering/numbering-part';
import Parser from './parser';
import SettingsPart from './settings/settings-part';
import StylesPart from './styles/styles-part';
import ThemePart from './theme/theme-part';
import { Options, Relationship, RelationshipTypes } from './types';
import {
  blobToBase64, deobfuscate, resolvePath, splitPath,
} from './utils';

export default class Word {
  private _xml: Xml;

  private _parser: Parser;

  private _options: Options;

  rels: Relationship[];

  parts: Part[] = [];

  partsMap: Record<string, Part> = {};

  documentPart: DocumentPart;

  fontTablePart: FontTablePart;

  numberingPart: NumberingPart;

  stylesPart: StylesPart;

  footnotesPart: FootnotesPart;

  endnotesPart: EndnotesPart;

  themePart: ThemePart;

  corePropsPart: CorePropsPart;

  extendedPropsPart: ExtendedPropsPart;

  settingsPart: SettingsPart;

  static load(blob: Blob | string, parser: Parser, options: Options) {
    const word = new Word();
    word._options = options;
    word._parser = parser;

    return Xml.load(blob, options)
      .then((xml) => {
        word._xml = xml;
        return word._xml.loadRelationships();
      }).then((rels) => {
        word.rels = rels;

        const tasks = topLevelRels.map((rel) => {
          const r = rels.find((x) => x.type === rel.type) ?? rel;
          return word.loadRelationshipPart(r.target, r.type);
        });

        return Promise.all(tasks);
      }).then(() => word);
  }

  save(type: JSZip.OutputType = 'blob'): Promise<unknown> {
    return this._xml.save(type);
  }

  private loadRelationshipPart(path: string, type: string): Promise<Part> {
    if (this.partsMap[path]) { return Promise.resolve(this.partsMap[path]); }

    if (!this._xml.get(path)) { return Promise.resolve(null); }

    let part: Part = null;

    switch (type) {
      case RelationshipTypes.OfficeDocument:
        part = new DocumentPart(this._xml, path, this._parser);
        this.documentPart = part as DocumentPart;
        break;

      case RelationshipTypes.FontTable:
        part = new FontTablePart(this._xml, path);
        this.fontTablePart = part as FontTablePart;
        break;

      case RelationshipTypes.Numbering:
        part = new NumberingPart(this._xml, path, this._parser);
        this.numberingPart = part as NumberingPart;
        break;

      case RelationshipTypes.Styles:
        part = new StylesPart(this._xml, path, this._parser);
        this.stylesPart = part as StylesPart;
        break;

      case RelationshipTypes.Theme:
        part = new ThemePart(this._xml, path);
        this.themePart = part as ThemePart;
        break;

      case RelationshipTypes.Footnotes:
        part = new FootnotesPart(this._xml, path, this._parser);
        this.footnotesPart = part as FootnotesPart;
        break;

      case RelationshipTypes.Endnotes:
        part = new EndnotesPart(this._xml, path, this._parser);
        this.endnotesPart = part as EndnotesPart;
        break;

      case RelationshipTypes.Footer:
        part = new FooterPart(this._xml, path, this._parser);
        break;

      case RelationshipTypes.Header:
        part = new HeaderPart(this._xml, path, this._parser);
        break;

      case RelationshipTypes.CoreProperties:
        part = new CorePropsPart(this._xml, path);
        this.corePropsPart = part as CorePropsPart;
        break;

      case RelationshipTypes.ExtendedProperties:
        part = new ExtendedPropsPart(this._xml, path);
        this.extendedPropsPart = part as ExtendedPropsPart;
        break;

      case RelationshipTypes.CustomProperties:
        part = new CustomPropsPart(this._xml, path);
        break;

      case RelationshipTypes.Settings:
        part = new SettingsPart(this._xml, path);
        this.settingsPart = part as SettingsPart;
        break;
      default:
        break;
    }

    if (part == null) { return Promise.resolve(null); }

    this.partsMap[path] = part;
    this.parts.push(part);

    return part.load().then(() => {
      if (part.rels == null || part.rels.length === 0) { return part; }

      const [folder] = splitPath(part.path);
      const rels = part.rels.map((rel) => this.loadRelationshipPart(resolvePath(rel.target, folder), rel.type));

      return Promise.all(rels).then(() => part);
    });
  }

  loadFont(id: string, key: string): PromiseLike<string> {
    return this.loadResource(this.fontTablePart, id, 'uint8array')
      .then((x) => (x ? this.blobToURL(new Blob([deobfuscate(x, key)])) : x));
  }

  loadNumberingImage(id: string): PromiseLike<string | ArrayBuffer> {
    return this.loadResource(this.numberingPart, id, 'blob')
      .then((x) => this.blobToURL(x));
  }

  private blobToURL(blob: Blob): string | PromiseLike<string | ArrayBuffer> {
    if (!blob) { return null; }

    if (this._options.useBase64URL) {
      return blobToBase64(blob);
    }

    return URL.createObjectURL(blob);
  }

  static getPathById(part: Part, id: string): string {
    const rel = part.rels.find((x) => x.id === id);
    const [folder] = splitPath(part.path);
    return rel ? resolvePath(rel.target, folder) : null;
  }

  private loadResource(part: Part, id: string, outputType: OutputType) {
    const path = Word.getPathById(part, id);
    return path ? this._xml.load(path, outputType) : Promise.resolve(null);
  }

  findPartByRelId(id: string, basePart: Part = null) {
    const rel = (basePart.rels ?? this.rels).find((r) => r.id === id);
    const folder = basePart ? splitPath(basePart.path)[0] : '';
    return rel ? this.partsMap[resolvePath(rel.target, folder)] : null;
  }

  loadDocumentImage(id: string, part?: Part): PromiseLike<string | ArrayBuffer> {
    return this.loadResource(part ?? this.documentPart, id, 'blob')
      .then((x) => this.blobToURL(x));
  }
}
