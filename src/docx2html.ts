import Word from './word';
import Parser from './parser';
import Renderer from './renderer';
import { Options } from './types';

const defaultOptions: Options = {
  debug: true,
  className: 'word',
  trimXmlDeclaration: true,
  ignoreHeight: false,
  ignoreWidth: false,
  ignoreFonts: false,
  breakPages: true,
  experimental: false,
  inWrapper: true,
  ignoreLastRenderedPageBreak: true,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  useBase64URL: false,
  useMathMLPolyfill: false,
  renderChanges: false,
};

export function prase(data: Blob | string, options: Partial<Options> = null): Promise<unknown> {
  const opts = { ...defaultOptions, ...options };
  return Word.load(data, new Parser(opts), opts);
}

export function render(data: Blob | string, bodyDom: HTMLElement, styleDom: HTMLElement = null, options: Partial<Options> = null): Promise<unknown> {
  const opts = { ...defaultOptions, ...options };
  if (!bodyDom) {
    return Promise.reject(new Error('容器不能为空'));
  }
  const renderer = new Renderer(window.document);
  return Word
    .load(data, new Parser(opts), opts)
    .then((doc) => {
      renderer.render(doc, bodyDom, styleDom, opts);
      return doc;
    });
}

export function renderDocument(data: Blob | string, options: Partial<Options> = null): Promise<unknown> {
  const opts = { ...defaultOptions, ...options };
  const renderer = new Renderer(window.document);
  return Word
    .load(data, new Parser(opts), opts)
    .then((doc) => {
      renderer.render(doc, null, null, opts);
      return doc;
    });
}
