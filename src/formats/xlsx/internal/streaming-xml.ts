import {
  SaxesParser,
  type SaxesAttributeNS,
  type SaxesTagNS,
  type XMLDecl,
} from 'saxes';

import {
  XmlComplexityLimitError,
  XmlStructureError,
} from '../../../common/xml/types';
import { ZipEntrySizeLimitError } from '../../../common/archive/read-entry';

export interface XlsxXmlElement {
  attributes: ReadonlyMap<string, string>;
  localName: string;
  namespace: string;
}

export interface XlsxXmlEventSink {
  closeElement?(element: XlsxXmlElement): void;
  openElement?(element: XlsxXmlElement): void;
  text?(value: string): void;
}

export interface XlsxStreamingXmlLimits {
  consumeBytes?: (byteLength: number) => void;
  consumeNodes?: (nodeCount: number) => void;
  maxBytes: number;
  maxDepth: number;
  maxNodes: number;
}

type XmlEncoding = 'utf-16be' | 'utf-16le' | 'utf-8';

function concatenate(left: Uint8Array, right: Uint8Array): Uint8Array {
  const output = new Uint8Array(left.byteLength + right.byteLength);
  output.set(left);
  output.set(right, left.byteLength);
  return output;
}

function expandedAttributeName(attribute: SaxesAttributeNS): string {
  return `{${attribute.uri}}${attribute.local}`;
}

function element(tag: SaxesTagNS): XlsxXmlElement {
  const attributes = new Map<string, string>();
  for (const attribute of Object.values(tag.attributes)) {
    attributes.set(expandedAttributeName(attribute), attribute.value);
  }
  return {
    attributes,
    localName: tag.local,
    namespace: tag.uri,
  };
}

function declaredEncodingMatches(
  actual: XmlEncoding,
  declared: string | undefined,
): boolean {
  if (declared === undefined) return true;
  const normalized = declared.toLowerCase();
  if (actual === 'utf-8') return normalized === 'utf-8';
  return normalized === 'utf-16' || normalized === actual;
}

export class XlsxStreamingXmlParser {
  private byteLength = 0;
  private closed = false;
  private decoder: TextDecoder | undefined;
  private depth = 0;
  private encoding: XmlEncoding | undefined;
  private nodeCount = 0;
  private pending: Uint8Array = new Uint8Array();
  private readonly parser = new SaxesParser({
    xmlns: true,
  });

  constructor(
    private readonly limits: XlsxStreamingXmlLimits,
    private readonly sink: XlsxXmlEventSink,
  ) {
    this.parser.on('doctype', () => {
      throw new XmlStructureError(
        'XML document type declarations are not allowed',
      );
    });
    this.parser.on('xmldecl', (declaration: XMLDecl) => {
      if (declaration.version !== '1.0') {
        throw new XmlStructureError('Only XML 1.0 documents are supported');
      }
      if (!declaredEncodingMatches(this.encoding!, declaration.encoding)) {
        throw new XmlStructureError(
          'XML declaration encoding does not match source bytes',
        );
      }
    });
    this.parser.on('opentag', (tag: SaxesTagNS) => {
      this.nodeCount += 1;
      this.depth += 1;
      if (this.nodeCount > this.limits.maxNodes) {
        throw new XmlComplexityLimitError(
          'maxXmlNodes',
          this.nodeCount,
          this.limits.maxNodes,
        );
      }
      if (this.depth > this.limits.maxDepth) {
        throw new XmlComplexityLimitError(
          'maxXmlDepth',
          this.depth,
          this.limits.maxDepth,
        );
      }
      this.limits.consumeNodes?.(1);
      this.sink.openElement?.(element(tag));
    });
    this.parser.on('text', (value: string) => this.sink.text?.(value));
    this.parser.on('cdata', (value: string) => this.sink.text?.(value));
    this.parser.on('closetag', (tag: SaxesTagNS) => {
      this.sink.closeElement?.(element(tag));
      this.depth -= 1;
    });
    this.parser.on('error', (cause: Error) => {
      throw new XmlStructureError('Invalid XML structure', { cause });
    });
  }

  write(chunk: Uint8Array): void {
    if (this.closed) throw new Error('XLSX streaming XML parser is closed');
    this.byteLength += chunk.byteLength;
    if (this.byteLength > this.limits.maxBytes) {
      throw new ZipEntrySizeLimitError(this.byteLength, this.limits.maxBytes);
    }
    this.limits.consumeBytes?.(chunk.byteLength);
    if (!this.decoder) {
      this.pending = concatenate(this.pending, chunk);
      this.initializeDecoder(false);
      return;
    }
    this.writeDecoded(chunk, true);
  }

  close(): void {
    if (this.closed) throw new Error('XLSX streaming XML parser is closed');
    this.closed = true;
    if (!this.decoder) this.initializeDecoder(true);
    this.writeDecoded(new Uint8Array(), false);
    this.parser.close();
  }

  private initializeDecoder(final: boolean): void {
    if (!final && this.pending.byteLength < 3) return;
    let offset = 0;
    const first = this.pending[0];
    const second = this.pending[1];
    if (first === 0xff && second === 0xfe) {
      this.encoding = 'utf-16le';
      offset = 2;
    } else if (first === 0xfe && second === 0xff) {
      this.encoding = 'utf-16be';
      offset = 2;
    } else if (first === 0x3c && second === 0x00) {
      this.encoding = 'utf-16le';
    } else if (first === 0x00 && second === 0x3c) {
      this.encoding = 'utf-16be';
    } else {
      this.encoding = 'utf-8';
    }
    this.decoder = new TextDecoder(this.encoding, { fatal: true });
    const pending = this.pending.subarray(offset);
    this.pending = new Uint8Array();
    this.writeDecoded(pending, true);
  }

  private writeDecoded(bytes: Uint8Array, stream: boolean): void {
    let text: string;
    try {
      text = this.decoder!.decode(bytes, { stream });
    } catch (cause) {
      throw new XmlStructureError(
        `Invalid ${this.encoding!.toUpperCase()} XML`,
        { cause },
      );
    }
    this.parser.write(text);
  }
}
