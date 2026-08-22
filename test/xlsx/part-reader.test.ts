import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import type {
  StreamableZipObject,
  ZipEntryStream,
} from '../../src/common/archive/read-entry';
import { XlsxParseError } from '../../src/formats/xlsx/errors';
import { XlsxPartReader } from '../../src/formats/xlsx/internal/part-reader';
import {
  defaultXlsxResourceLimits,
  type ResolvedXlsxResourceLimits,
} from '../../src/formats/xlsx/internal/resource-limits';
import type { XlsxDiagnostic } from '../../src/formats/xlsx/types';

function archive(): JSZip {
  const zip = new JSZip();
  zip.file('first.xml', '<root><child id="first"/></root>');
  zip.file('second.xml', '<root><child id="second"/></root>');
  zip.file('invalid.xml', '<root><child></root>');
  zip.file(
    'doctype.xml',
    '<!DOCTYPE root [<!ENTITY payload "secret">]><root>&payload;</root>',
  );
  zip.file('invalid-utf8.xml', Uint8Array.from([0x3c, 0x72, 0x3e, 0xff]));
  zip.file('media.bin', Uint8Array.from([0, 127, 128, 255]));
  return zip;
}

function limits(
  overrides: Partial<ResolvedXlsxResourceLimits>,
): ResolvedXlsxResourceLimits {
  return { ...defaultXlsxResourceLimits(), ...overrides };
}

function readFailureArchive(error: unknown): JSZip {
  const listeners = new Map<string, (value?: unknown) => void>();
  const streamImplementation = {
    on(event: string, listener: (value?: unknown) => void) {
      listeners.set(event, listener);
      return streamImplementation;
    },
    pause() {
      return streamImplementation;
    },
    resume() {
      listeners.get('error')?.(error);
      return streamImplementation;
    },
  };
  const entry: StreamableZipObject = {
    internalStream() {
      return streamImplementation as unknown as ZipEntryStream;
    },
    name: 'unreadable.xml',
  };
  return {
    file(part: string) {
      return part === entry.name ? entry : null;
    },
  } as unknown as JSZip;
}

interface ScriptedArchiveOptions {
  afterEndError?: unknown;
  chunks?: readonly Uint8Array[];
  declaredSize?: unknown;
  error?: unknown;
  ignorePause?: boolean;
  internalStreamError?: unknown;
  throwOnEvent?: string;
}

interface ScriptedArchiveState {
  chunksEmitted: number;
  pauses: number;
  resumes: number;
}

function scriptedArchive(options: ScriptedArchiveOptions): {
  state: ScriptedArchiveState;
  zip: JSZip;
} {
  const listeners = new Map<string, (value?: unknown) => void>();
  const state: ScriptedArchiveState = {
    chunksEmitted: 0,
    pauses: 0,
    resumes: 0,
  };
  let paused = false;
  const streamImplementation = {
    on(event: string, listener: (value?: unknown) => void) {
      if (options.throwOnEvent === event)
        throw new Error(`cannot bind ${event}`);
      listeners.set(event, listener);
      return streamImplementation;
    },
    pause() {
      paused = true;
      state.pauses += 1;
      return streamImplementation;
    },
    resume() {
      state.resumes += 1;
      for (const chunk of options.chunks ?? []) {
        if (paused && !options.ignorePause) break;
        state.chunksEmitted += 1;
        listeners.get('data')?.(chunk);
      }
      if ('error' in options) {
        listeners.get('error')?.(options.error);
      } else {
        listeners.get('end')?.();
        if ('afterEndError' in options) {
          listeners.get('error')?.(options.afterEndError);
        }
      }
      return streamImplementation;
    },
  };
  const entry: StreamableZipObject = {
    _data: { uncompressedSize: options.declaredSize },
    internalStream() {
      if ('internalStreamError' in options) throw options.internalStreamError;
      return streamImplementation as unknown as ZipEntryStream;
    },
    name: 'stream.xml',
  };
  return {
    state,
    zip: {
      file(part: string) {
        return part === entry.name ? entry : null;
      },
    } as unknown as JSZip,
  };
}

function captureParseError(action: Promise<unknown>): Promise<XlsxParseError> {
  return action.then(
    () => {
      throw new Error('Expected XLSX part read to fail');
    },
    (error: unknown) => {
      expect(error).toBeInstanceOf(XlsxParseError);
      return error as XlsxParseError;
    },
  );
}

describe('XLSX bounded part reader', () => {
  it('returns a cached XML tree without charging budgets twice', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(
      archive(),
      diagnostics,
      limits({ maxTotalXmlNodes: 2 }),
    );

    const first = await reader.readXml('first.xml', { required: true });
    const cached = await reader.readXml('first.xml', { required: true });

    expect(cached).toBe(first);
    expect(first).toMatchObject({
      root: { child: { attrs: { id: 'first' } } },
    });
    expect(diagnostics).toEqual([]);
  });

  it('accepts expanded XML exactly at the cumulative byte limit', async () => {
    const reader = new XlsxPartReader(
      archive(),
      [],
      limits({ maxTotalUncompressedBytes: 32 }),
    );

    await expect(reader.readXml('first.xml')).resolves.toMatchObject({
      root: { child: {} },
    });
  });

  it('returns null for an absent optional part', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(archive(), diagnostics);

    await expect(reader.readXml('missing.xml')).resolves.toBeNull();
    expect(diagnostics).toEqual([]);
  });

  it.each(['missing.xml', ''])(
    'throws a structured missing-required-part error for %s',
    async (part) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(archive(), diagnostics);
      const error = await captureParseError(
        reader.readXml(part, { required: true }),
      );

      expect(error.diagnostic).toEqual({
        code: 'missing-required-part',
        message: part
          ? `Required XLSX part is missing: ${part}`
          : 'Required XLSX part name is empty',
        ...(part ? { part } : {}),
        severity: 'error',
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );

  it.each([
    ['invalid.xml', 'xml-parse-failed'],
    ['doctype.xml', 'xml-parse-failed'],
    ['invalid-utf8.xml', 'xml-parse-failed'],
  ] as const)('rejects unsafe XML for %s', async (part, code) => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(archive(), diagnostics);
    const error = await captureParseError(reader.readXml(part));

    expect(error.diagnostic).toMatchObject({
      code,
      part,
      severity: 'error',
    });
    expect(error.cause).toBeInstanceOf(Error);
    expect(diagnostics).toEqual([error.diagnostic]);
  });

  it('distinguishes archive read failures from XML failures', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(
      readFailureArchive(new Error('storage unavailable')),
      diagnostics,
    );
    const error = await captureParseError(reader.readXml('unreadable.xml'));

    expect(error.diagnostic).toEqual({
      code: 'xml-read-failed',
      message: 'Failed to read XLSX part unreadable.xml',
      part: 'unreadable.xml',
      severity: 'error',
    });
    expect((error.cause as Error).message).toBe('storage unavailable');
  });

  it.each([
    [{ maxXmlBytes: 31 }, 'first.xml', 'maxXmlBytes', 32, 31],
    [{ maxXmlDepth: 1 }, 'first.xml', 'maxXmlDepth', 2, 1],
    [{ maxXmlNodes: 1 }, 'first.xml', 'maxXmlNodes', 2, 1],
    [{ maxTotalXmlNodes: 3 }, 'second.xml', 'maxTotalXmlNodes', 4, 3],
    [
      { maxTotalUncompressedBytes: 63 },
      'second.xml',
      'maxTotalUncompressedBytes',
      65,
      63,
    ],
  ] as const)(
    'maps XML budget failures to %s',
    async (overrides, failingPart, limitName, actual, limit) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(
        archive(),
        diagnostics,
        limits(overrides),
      );
      if (failingPart === 'second.xml') {
        await reader.readXml('first.xml');
      }
      const error = await captureParseError(reader.readXml(failingPart));

      expect(error.diagnostic).toMatchObject({
        actual,
        code: 'resource-limit-exceeded',
        limit,
        limitName,
        part: failingPart,
        severity: 'error',
      });
      expect(error.cause).toMatchObject({
        actual,
        limit,
        limitName,
        name: 'XlsxResourceLimitError',
        part: failingPart,
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );

  it('returns exact binary bytes and null for a missing part', async () => {
    const reader = new XlsxPartReader(archive(), []);

    await expect(
      reader.readBytes('media.bin', 'maxMediaBytes'),
    ).resolves.toEqual(Uint8Array.from([0, 127, 128, 255]));
    await expect(
      reader.readBytes('missing.bin', 'maxMediaBytes'),
    ).resolves.toBeNull();
  });

  it('streams namespace-aware XML events and reports optional presence', async () => {
    const reader = new XlsxPartReader(archive(), []);
    const events: string[] = [];

    await expect(
      reader.streamXml('first.xml', {
        closeElement: ({ localName }) => events.push(`/${localName}`),
        openElement: ({ attributes, localName }) =>
          events.push(`${localName}:${attributes.get('{}id') ?? ''}`),
      }),
    ).resolves.toBe(true);
    await expect(reader.streamXml('missing.xml', {})).resolves.toBe(false);
    expect(events).toEqual(['root:', 'child:first', '/child', '/root']);
  });

  it('reports a missing required streamed part once', async () => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(archive(), diagnostics);
    const error = await captureParseError(
      reader.streamXml('missing.xml', {}, { required: true }),
    );

    expect(error.diagnostic).toEqual({
      code: 'missing-required-part',
      message: 'Required XLSX part is missing: missing.xml',
      part: 'missing.xml',
      severity: 'error',
    });
    expect(diagnostics).toEqual([error.diagnostic]);
  });

  it.each([
    ['invalid.xml', 'xml-parse-failed'],
    ['invalid-utf8.xml', 'xml-parse-failed'],
  ] as const)('maps streamed XML failure for %s', async (part, code) => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(archive(), diagnostics);
    const error = await captureParseError(reader.streamXml(part, {}));

    expect(error.diagnostic).toEqual({
      code,
      message: `Failed to parse XLSX part ${part}`,
      part,
      severity: 'error',
    });
    expect(error.cause).toBeInstanceOf(Error);
    expect(diagnostics).toEqual([error.diagnostic]);
  });

  it.each([
    [new Error('storage unavailable'), 'storage unavailable'],
    ['non-error failure', 'non-error failure'],
  ])('maps streamed archive read failure %#', async (failure, message) => {
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(readFailureArchive(failure), diagnostics);
    const error = await captureParseError(
      reader.streamXml('unreadable.xml', {}),
    );

    expect(error.diagnostic).toEqual({
      code: 'xml-read-failed',
      message: 'Failed to read XLSX part unreadable.xml',
      part: 'unreadable.xml',
      severity: 'error',
    });
    expect((error.cause as Error).message).toBe(message);
    expect(diagnostics).toEqual([error.diagnostic]);
  });

  it('rejects declared XML size before starting decompression', async () => {
    const { state, zip } = scriptedArchive({ declaredSize: 33 });
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(
      zip,
      diagnostics,
      limits({ maxXmlBytes: 32 }),
    );
    const error = await captureParseError(reader.streamXml('stream.xml', {}));

    expect(error.diagnostic).toMatchObject({
      actual: 33,
      code: 'resource-limit-exceeded',
      limit: 32,
      limitName: 'maxXmlBytes',
      part: 'stream.xml',
    });
    expect(state).toEqual({ chunksEmitted: 0, pauses: 0, resumes: 0 });
  });

  it('accepts a declared XML size exactly at the limit', async () => {
    const bytes = new TextEncoder().encode('<r/>');
    const { state, zip } = scriptedArchive({
      chunks: [bytes],
      declaredSize: bytes.byteLength,
    });
    const reader = new XlsxPartReader(
      zip,
      [],
      limits({ maxXmlBytes: bytes.byteLength }),
    );

    await expect(reader.streamXml('stream.xml', {})).resolves.toBe(true);
    expect(state).toEqual({ chunksEmitted: 1, pauses: 0, resumes: 1 });
  });

  it.each([
    [{ maxXmlBytes: 31 }, 'first.xml', 'maxXmlBytes', 32, 31],
    [{ maxXmlDepth: 1 }, 'first.xml', 'maxXmlDepth', 2, 1],
    [{ maxXmlNodes: 1 }, 'first.xml', 'maxXmlNodes', 2, 1],
    [{ maxTotalXmlNodes: 3 }, 'second.xml', 'maxTotalXmlNodes', 4, 3],
    [
      { maxTotalUncompressedBytes: 64 },
      'second.xml',
      'maxTotalUncompressedBytes',
      65,
      64,
    ],
  ] as const)(
    'maps streamed XML budget failures to %s',
    async (overrides, failingPart, limitName, actual, limit) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(
        archive(),
        diagnostics,
        limits(overrides),
      );
      if (failingPart === 'second.xml') {
        await reader.streamXml('first.xml', {});
      }
      const error = await captureParseError(reader.streamXml(failingPart, {}));

      expect(error.diagnostic).toMatchObject({
        actual,
        code: 'resource-limit-exceeded',
        limit,
        limitName,
        part: failingPart,
        severity: 'error',
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );

  it('accepts exact cumulative streamed XML budgets', async () => {
    const reader = new XlsxPartReader(
      archive(),
      [],
      limits({ maxTotalUncompressedBytes: 65, maxTotalXmlNodes: 4 }),
    );

    await expect(reader.streamXml('first.xml', {})).resolves.toBe(true);
    await expect(reader.streamXml('second.xml', {})).resolves.toBe(true);
  });

  it('stops decompression and ignores callbacks after a parse failure', async () => {
    const { state, zip } = scriptedArchive({
      chunks: [
        new TextEncoder().encode('<r><x></r>'),
        new TextEncoder().encode('<ignored/>'),
      ],
      ignorePause: true,
    });
    const diagnostics: XlsxDiagnostic[] = [];
    const reader = new XlsxPartReader(zip, diagnostics);
    const error = await captureParseError(reader.streamXml('stream.xml', {}));

    expect(error.diagnostic.code).toBe('xml-parse-failed');
    expect(state).toEqual({ chunksEmitted: 2, pauses: 1, resumes: 1 });
    expect(diagnostics).toHaveLength(1);
  });

  it('maps an XML error raised only when the stream closes', async () => {
    const { state, zip } = scriptedArchive({
      chunks: [new TextEncoder().encode('<r>')],
    });
    const diagnostics: XlsxDiagnostic[] = [];
    const error = await captureParseError(
      new XlsxPartReader(zip, diagnostics).streamXml('stream.xml', {}),
    );

    expect(error.diagnostic.code).toBe('xml-parse-failed');
    expect(state).toEqual({ chunksEmitted: 1, pauses: 1, resumes: 1 });
    expect(diagnostics).toEqual([error.diagnostic]);
  });

  it('ignores data and errors delivered after a consumer abort', async () => {
    const sinkError = new Error('stop consumer');
    const opened: string[] = [];
    const closed: string[] = [];
    const { state, zip } = scriptedArchive({
      chunks: [
        new TextEncoder().encode('<r>'),
        new TextEncoder().encode('<ignored/></r>'),
      ],
      error: new Error('late stream error'),
      ignorePause: true,
    });

    await expect(
      new XlsxPartReader(zip, []).streamXml('stream.xml', {
        closeElement: ({ localName }) => closed.push(localName),
        openElement: ({ localName }) => {
          opened.push(localName);
          throw sinkError;
        },
      }),
    ).rejects.toBe(sinkError);
    expect(opened).toEqual(['r']);
    expect(closed).toEqual([]);
    expect(state).toEqual({ chunksEmitted: 2, pauses: 1, resumes: 1 });
  });

  it('ignores an error delivered after a successful end event', async () => {
    const bytes = new TextEncoder().encode('<r/>');
    const { state, zip } = scriptedArchive({
      afterEndError: new Error('late error'),
      chunks: [bytes],
    });

    await expect(
      new XlsxPartReader(zip, []).streamXml('stream.xml', {}),
    ).resolves.toBe(true);
    expect(state).toEqual({ chunksEmitted: 1, pauses: 0, resumes: 1 });
  });

  it.each([
    [
      { internalStreamError: new Error('cannot create stream') },
      'cannot create stream',
    ],
    [{ throwOnEvent: 'end' }, 'cannot bind end'],
  ])('maps stream setup failure %#', async (options, message) => {
    const { state, zip } = scriptedArchive(options);
    const error = await captureParseError(
      new XlsxPartReader(zip, []).streamXml('stream.xml', {}),
    );

    expect(error.diagnostic.code).toBe('xml-read-failed');
    expect((error.cause as Error).message).toBe(message);
    expect(state.resumes).toBe(0);
  });

  it('propagates consumer errors without relabeling them as XML failures', async () => {
    const sinkError = new Error('consumer failed');
    const diagnostic: XlsxDiagnostic = {
      code: 'invalid-document-value',
      message: 'consumer parse failure',
      severity: 'error',
    };
    const parseError = new XlsxParseError(diagnostic);
    const firstReader = new XlsxPartReader(archive(), []);
    const secondReader = new XlsxPartReader(archive(), []);

    await expect(
      firstReader.streamXml('first.xml', {
        openElement: () => {
          throw sinkError;
        },
      }),
    ).rejects.toBe(sinkError);
    await expect(
      secondReader.streamXml('first.xml', {
        openElement: () => {
          throw parseError;
        },
      }),
    ).rejects.toBe(parseError);
  });

  it.each([
    [{ maxMediaBytes: 3 }, 'maxMediaBytes', 4, 3],
    [{ maxTotalUncompressedBytes: 3 }, 'maxTotalUncompressedBytes', 4, 3],
  ] as const)(
    'enforces binary byte budget %s',
    async (overrides, limitName, actual, limit) => {
      const diagnostics: XlsxDiagnostic[] = [];
      const reader = new XlsxPartReader(
        archive(),
        diagnostics,
        limits(overrides),
      );
      const error = await captureParseError(
        reader.readBytes('media.bin', 'maxMediaBytes'),
      );

      expect(error.diagnostic).toMatchObject({
        actual,
        code: 'resource-limit-exceeded',
        limit,
        limitName,
        part: 'media.bin',
        severity: 'error',
      });
      expect(diagnostics).toEqual([error.diagnostic]);
    },
  );
});
