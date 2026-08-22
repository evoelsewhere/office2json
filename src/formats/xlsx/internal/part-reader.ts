import type JSZip from 'jszip';

import {
  readZipEntryBytes,
  ZipExpansionBudgetLimitError,
  ZipEntrySizeLimitError,
} from '../../../common/archive/read-entry';
import {
  readXmlFileResult,
  XmlComplexityLimitError,
  XmlStructureError,
  type XmlReadResult,
} from '../../../common/xml/read-xml';
import type { XmlLookupValue } from '../../../common/xml/tree';
import { XlsxParseError } from '../errors';
import type { XlsxDiagnostic } from '../types';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
  resourceLimitDiagnostic,
  XlsxResourceLimitError,
} from './resource-limits';
import { XlsxStreamingXmlParser, type XlsxXmlEventSink } from './streaming-xml';

interface ReadXmlOptions {
  required?: boolean;
}

type ByteLimitName = 'maxMediaBytes' | 'maxPartBytes';

type StreamXmlOutcome =
  | { status: 'ok' }
  | { error: unknown; phase: 'parse' | 'read'; status: 'error' };

interface StreamableXmlEntry {
  _data?: { uncompressedSize?: unknown };
  internalStream(type: 'uint8array'): XmlEntryStream;
}

interface XmlEntryStream {
  on(event: 'data', listener: (chunk: Uint8Array) => void): XmlEntryStream;
  on(event: 'end', listener: () => void): XmlEntryStream;
  on(event: 'error', listener: (error: unknown) => void): XmlEntryStream;
  pause(): XmlEntryStream;
  resume(): XmlEntryStream;
}

function asError(value: unknown): Error {
  return value instanceof Error ? value : new Error(String(value));
}

function declaredEntrySize(entry: JSZip.JSZipObject): number | null {
  const value = (entry as unknown as StreamableXmlEntry)._data
    ?.uncompressedSize;
  return typeof value === 'number' && Number.isSafeInteger(value)
    ? value
    : null;
}

function streamXmlEntry(
  entry: JSZip.JSZipObject,
  parser: XlsxStreamingXmlParser,
): Promise<StreamXmlOutcome> {
  return new Promise((resolve) => {
    let settled = false;
    let stream: ReturnType<StreamableXmlEntry['internalStream']>;
    try {
      stream = (entry as unknown as StreamableXmlEntry).internalStream(
        'uint8array',
      );
    } catch (error) {
      resolve({ error: asError(error), phase: 'read', status: 'error' });
      return;
    }

    const fail = (error: unknown, phase: 'parse' | 'read'): void => {
      if (settled) return;
      settled = true;
      stream.pause();
      resolve({ error, phase, status: 'error' });
    };

    try {
      stream
        .on('data', (chunk) => {
          try {
            parser.write(chunk);
          } catch (error) {
            fail(error, 'parse');
          }
        })
        .on('error', (error) => fail(asError(error), 'read'))
        .on('end', () => {
          try {
            parser.close();
          } catch (error) {
            fail(error, 'parse');
            return;
          }
          settled = true;
          resolve({ status: 'ok' });
        })
        .resume();
    } catch (error) {
      fail(asError(error), 'read');
    }
  });
}

export class XlsxPartReader {
  private readonly xmlCache = new Map<string, XmlReadResult<XmlLookupValue>>();
  private totalExpandedBytes = 0;
  private totalXmlNodes = 0;

  constructor(
    private readonly zip: JSZip,
    private readonly diagnostics: XlsxDiagnostic[],
    private readonly limits: ResolvedXlsxResourceLimits = defaultXlsxResourceLimits(),
  ) {}

  hasPart(part: string): boolean {
    return this.zip.file(part) !== null;
  }

  async readXml(
    part: string,
    options: { required: true },
  ): Promise<XmlLookupValue>;
  async readXml(
    part: string,
    options?: ReadXmlOptions,
  ): Promise<XmlLookupValue | null>;
  async readXml(
    part: string,
    options: ReadXmlOptions = {},
  ): Promise<XmlLookupValue | null> {
    let result = this.xmlCache.get(part);
    if (!result) {
      result = await readXmlFileResult<XmlLookupValue>(this.zip, part, {
        consumeBytes: (byteLength) => this.consumeExpandedBytes(byteLength),
        consumeNodes: (nodeCount) => this.consumeXmlNodes(nodeCount),
        maxBytes: this.limits.maxXmlBytes,
        maxDepth: this.limits.maxXmlDepth,
        maxNodes: this.limits.maxXmlNodes,
      });
      this.xmlCache.set(part, result);
    }

    if (result.status === 'ok') return result.value;
    if (result.status === 'missing') {
      if (options.required) this.failMissing(part);
      return null;
    }
    return this.failXml(part, result.error, result.phase);
  }

  async readBytes(
    part: string,
    limitName: ByteLimitName,
  ): Promise<Uint8Array | null> {
    const entry = this.zip.file(part);
    if (!entry) return null;
    try {
      return await readZipEntryBytes(
        entry,
        this.limits[limitName],
        (byteLength) => this.consumeExpandedBytes(byteLength),
      );
    } catch (error) {
      if (error instanceof ZipEntrySizeLimitError) {
        this.failResource(
          new XlsxResourceLimitError(
            limitName,
            error.actual,
            error.limit,
            part,
          ),
        );
      }
      if (error instanceof ZipExpansionBudgetLimitError) {
        this.failResource(
          new XlsxResourceLimitError(
            'maxTotalUncompressedBytes',
            error.actual,
            error.limit,
            part,
          ),
        );
      }
      throw error;
    }
  }

  async streamXml(
    part: string,
    sink: XlsxXmlEventSink,
    options: ReadXmlOptions = {},
  ): Promise<boolean> {
    const entry = this.zip.file(part);
    if (!entry) {
      if (options.required) this.failMissing(part);
      return false;
    }

    const expectedSize = declaredEntrySize(entry);
    if (expectedSize !== null && expectedSize > this.limits.maxXmlBytes) {
      this.failResource(
        new XlsxResourceLimitError(
          'maxXmlBytes',
          expectedSize,
          this.limits.maxXmlBytes,
          part,
        ),
      );
    }

    const parser = new XlsxStreamingXmlParser(
      {
        consumeBytes: (byteLength) => this.consumeExpandedBytes(byteLength),
        consumeNodes: (nodeCount) => this.consumeXmlNodes(nodeCount),
        maxBytes: this.limits.maxXmlBytes,
        maxDepth: this.limits.maxXmlDepth,
        maxNodes: this.limits.maxXmlNodes,
      },
      sink,
    );
    const outcome = await streamXmlEntry(entry, parser);
    if (outcome.status === 'ok') return true;
    if (outcome.phase === 'parse') {
      if (outcome.error instanceof XlsxParseError) throw outcome.error;
      if (
        !(outcome.error instanceof ZipEntrySizeLimitError) &&
        !(outcome.error instanceof XmlComplexityLimitError) &&
        !(outcome.error instanceof ZipExpansionBudgetLimitError) &&
        !(outcome.error instanceof XmlStructureError)
      ) {
        throw outcome.error;
      }
    }
    return this.failXml(part, outcome.error, outcome.phase);
  }

  private failMissing(part: string): never {
    const diagnostic: XlsxDiagnostic = {
      code: 'missing-required-part',
      message: part
        ? `Required XLSX part is missing: ${part}`
        : 'Required XLSX part name is empty',
      ...(part ? { part } : {}),
      severity: 'error',
    };
    this.diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic);
  }

  private failXml(
    part: string,
    error: unknown,
    phase: 'limit' | 'parse' | 'read',
  ): never {
    if (error instanceof ZipEntrySizeLimitError) {
      this.failResource(
        new XlsxResourceLimitError(
          'maxXmlBytes',
          error.actual,
          error.limit,
          part,
        ),
      );
    }
    if (error instanceof XmlComplexityLimitError) {
      this.failResource(
        new XlsxResourceLimitError(
          error.limitName,
          error.actual,
          error.limit,
          part,
        ),
      );
    }
    if (error instanceof ZipExpansionBudgetLimitError) {
      this.failResource(
        new XlsxResourceLimitError(
          'maxTotalUncompressedBytes',
          error.actual,
          error.limit,
          part,
        ),
      );
    }

    const diagnostic: XlsxDiagnostic = {
      code: phase === 'parse' ? 'xml-parse-failed' : 'xml-read-failed',
      message: `Failed to ${phase} XLSX part ${part}`,
      part,
      severity: 'error',
    };
    this.diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic, { cause: error });
  }

  private consumeExpandedBytes(byteLength: number): void {
    const next = this.totalExpandedBytes + byteLength;
    if (
      !Number.isSafeInteger(next) ||
      next > this.limits.maxTotalUncompressedBytes
    ) {
      throw new ZipExpansionBudgetLimitError(
        next,
        this.limits.maxTotalUncompressedBytes,
      );
    }
    this.totalExpandedBytes = next;
  }

  private consumeXmlNodes(nodeCount: number): void {
    const next = this.totalXmlNodes + nodeCount;
    if (!Number.isSafeInteger(next) || next > this.limits.maxTotalXmlNodes) {
      throw new XmlComplexityLimitError(
        'maxTotalXmlNodes',
        next,
        this.limits.maxTotalXmlNodes,
      );
    }
    this.totalXmlNodes = next;
  }

  private failResource(error: XlsxResourceLimitError): never {
    const diagnostic = resourceLimitDiagnostic(error);
    this.diagnostics.push(diagnostic);
    throw new XlsxParseError(diagnostic, { cause: error });
  }
}
