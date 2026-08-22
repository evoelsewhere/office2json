import { XlsxParseError } from '../errors';
import type { XlsxRichTextRun } from '../types';
import { XlsxPartReader } from './part-reader';
import {
  type ResolvedXlsxResourceLimits,
  XlsxResourceLimitError,
} from './resource-limits';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';

const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const RUN_PROPERTY_ELEMENTS = new Set([
  'b',
  'charset',
  'color',
  'condense',
  'extend',
  'family',
  'i',
  'outline',
  'rFont',
  'scheme',
  'shadow',
  'strike',
  'sz',
  'u',
  'vertAlign',
]);

export interface XlsxPhoneticProperties {
  alignment?: 'center' | 'distributed' | 'left' | 'no-control';
  fontId?: number;
  type?:
    | 'full-width-katakana'
    | 'half-width-katakana'
    | 'hiragana'
    | 'no-conversion';
}

export interface XlsxPhoneticRun {
  end: number;
  start: number;
  text: string;
}

export interface XlsxSharedString {
  phoneticProperties?: XlsxPhoneticProperties;
  phoneticRuns?: readonly XlsxPhoneticRun[];
  runs?: readonly XlsxRichTextRun[];
  text: string;
}

export interface XlsxSharedStringTable {
  part: string | null;
  values: readonly XlsxSharedString[];
}

interface PendingPhoneticRun {
  end: number;
  hasText: boolean;
  start: number;
  text: string;
}

interface PendingRichRun {
  hasText: boolean;
  text: string;
}

type TextTarget = 'phonetic' | 'plain' | 'rich' | null;

function structureFailure(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message,
    part,
    severity: 'error',
  });
}

function valueFailure(part: string, message: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-value',
    message,
    part,
    severity: 'error',
  });
}

function attribute(
  element: XlsxXmlElement,
  localName: string,
): string | undefined {
  return element.attributes.get(`{}${localName}`);
}

function unsignedInteger(
  value: string | undefined,
  part: string,
  message: string,
): number {
  if (value === undefined || !/^(?:0|[1-9]\d*)$/u.test(value)) {
    valueFailure(part, message);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > 0xffff_ffff) {
    valueFailure(part, message);
  }
  return parsed;
}

function optionalUnsignedInteger(
  value: string | undefined,
  part: string,
  message: string,
): number | undefined {
  return value === undefined
    ? undefined
    : unsignedInteger(value, part, message);
}

function phoneticAlignment(
  value: string | undefined,
  part: string,
): XlsxPhoneticProperties['alignment'] {
  if (value === undefined) return undefined;
  if (value === 'center' || value === 'distributed' || value === 'left') {
    return value;
  }
  if (value === 'noControl') return 'no-control';
  valueFailure(part, 'Shared-string phonetic alignment is invalid');
}

function phoneticType(
  value: string | undefined,
  part: string,
): XlsxPhoneticProperties['type'] {
  if (value === undefined) return undefined;
  if (value === 'fullwidthKatakana') return 'full-width-katakana';
  if (value === 'halfwidthKatakana') return 'half-width-katakana';
  if (value === 'Hiragana') return 'hiragana';
  if (value === 'noConversion') return 'no-conversion';
  valueFailure(part, 'Shared-string phonetic type is invalid');
}

class SharedStringSink implements XlsxXmlEventSink {
  private capture: TextTarget = null;
  private declaredCount: number | undefined;
  private declaredUniqueCount: number | undefined;
  private mode: 'plain' | 'rich' | 'unset' = 'unset';
  private phoneticProperties: XlsxPhoneticProperties | undefined;
  private readonly phoneticRuns: XlsxPhoneticRun[] = [];
  private phoneticSeen!: boolean;
  private plainText!: string;
  private readonly richRuns: XlsxRichTextRun[] = [];
  private currentPhoneticRun: PendingPhoneticRun | undefined;
  private currentRichRun: PendingRichRun | undefined;
  private readonly stack: XlsxXmlElement[] = [];
  private textCharacters = 0;
  private textRuns = 0;
  private readonly values: XlsxSharedString[] = [];

  constructor(
    private readonly part: string,
    private readonly namespace: string,
    private readonly limits: ResolvedXlsxResourceLimits,
  ) {}

  openElement(element: XlsxXmlElement): void {
    if (element.namespace !== this.namespace) {
      structureFailure(
        this.part,
        'Shared-string element has an unsupported namespace',
      );
    }
    const parent = this.stack.at(-1);
    if (!parent) {
      if (element.localName !== 'sst') {
        structureFailure(this.part, 'Shared-string root is missing');
      }
      this.declaredCount = optionalUnsignedInteger(
        attribute(element, 'count'),
        this.part,
        'Shared-string count is invalid',
      );
      this.declaredUniqueCount = optionalUnsignedInteger(
        attribute(element, 'uniqueCount'),
        this.part,
        'Shared-string unique count is invalid',
      );
      this.stack.push(element);
      return;
    }

    this.openChild(parent.localName, element);
    this.stack.push(element);
  }

  closeElement(element: XlsxXmlElement): void {
    this.capture = null;
    if (element.localName === 'r') this.closeRichRun();
    if (element.localName === 'rPh') this.closePhoneticRun();
    if (element.localName === 'si') this.closeString();
    this.stack.pop();
  }

  text(value: string): void {
    if (this.capture === null) {
      if (value.trim().length > 0) {
        structureFailure(
          this.part,
          'Shared-string text must be contained by a text element',
        );
      }
      return;
    }
    this.consumeText(value.length);
    if (this.capture === 'plain') this.plainText += value;
    if (this.capture === 'rich') this.currentRichRun!.text += value;
    if (this.capture === 'phonetic') this.currentPhoneticRun!.text += value;
  }

  result(): XlsxSharedStringTable {
    if (
      this.declaredUniqueCount !== undefined &&
      this.declaredUniqueCount !== this.values.length
    ) {
      valueFailure(
        this.part,
        'Shared-string unique count does not match entries',
      );
    }
    if (
      this.declaredCount !== undefined &&
      this.declaredCount < this.values.length
    ) {
      valueFailure(this.part, 'Shared-string count is smaller than its table');
    }
    return Object.freeze({
      part: this.part,
      values: Object.freeze([...this.values]),
    });
  }

  private openChild(parent: string, element: XlsxXmlElement): void {
    if (parent === 'sst' && element.localName === 'si') {
      this.openString();
      return;
    }
    if (parent === 'si') {
      if (element.localName === 't') {
        this.openPlainText(element);
        return;
      }
      if (element.localName === 'r') {
        this.openRichRun();
        return;
      }
      if (element.localName === 'rPh') {
        this.openPhoneticRun(element);
        return;
      }
      if (element.localName === 'phoneticPr') {
        this.openPhoneticProperties(element);
        return;
      }
    }
    if (parent === 'r') {
      if (element.localName === 'rPr') return;
      if (element.localName === 't') {
        this.openRunText(element, 'rich');
        return;
      }
    }
    if (element.localName === 't') {
      this.openRunText(element, 'phonetic');
      return;
    }
    if (parent === 'rPr' && RUN_PROPERTY_ELEMENTS.has(element.localName)) {
      return;
    }
    structureFailure(this.part, 'Shared-string element nesting is invalid');
  }

  private openString(): void {
    const actual = this.values.length + 1;
    if (actual > this.limits.maxSharedStrings) {
      throw new XlsxResourceLimitError(
        'maxSharedStrings',
        actual,
        this.limits.maxSharedStrings,
        this.part,
      );
    }
    this.mode = 'unset';
    this.plainText = '';
    this.richRuns.length = 0;
    this.phoneticRuns.length = 0;
    this.phoneticProperties = undefined;
    this.phoneticSeen = false;
  }

  private openPlainText(element: XlsxXmlElement): void {
    if (this.mode !== 'unset' || this.phoneticSeen) {
      structureFailure(this.part, 'Shared-string plain text is out of order');
    }
    this.mode = 'plain';
    this.validateXmlSpace(element);
    this.capture = 'plain';
  }

  private openRichRun(): void {
    if (this.mode === 'plain' || this.phoneticSeen) {
      structureFailure(this.part, 'Shared-string rich text is out of order');
    }
    this.mode = 'rich';
    this.consumeRun();
    this.currentRichRun = { hasText: false, text: '' };
  }

  private openPhoneticRun(element: XlsxXmlElement): void {
    if (this.phoneticProperties !== undefined) {
      structureFailure(this.part, 'Shared-string phonetic run is out of order');
    }
    this.phoneticSeen = true;
    this.consumeRun();
    this.currentPhoneticRun = {
      end: unsignedInteger(
        attribute(element, 'eb'),
        this.part,
        'Shared-string phonetic end index is invalid',
      ),
      hasText: false,
      start: unsignedInteger(
        attribute(element, 'sb'),
        this.part,
        'Shared-string phonetic start index is invalid',
      ),
      text: '',
    };
  }

  private openPhoneticProperties(element: XlsxXmlElement): void {
    if (this.phoneticProperties !== undefined) {
      structureFailure(
        this.part,
        'Shared string has duplicate phonetic properties',
      );
    }
    const fontId = optionalUnsignedInteger(
      attribute(element, 'fontId'),
      this.part,
      'Shared-string phonetic font ID is invalid',
    );
    const alignment = phoneticAlignment(
      attribute(element, 'alignment'),
      this.part,
    );
    const type = phoneticType(attribute(element, 'type'), this.part);
    this.phoneticSeen = true;
    this.phoneticProperties = Object.freeze({
      ...(alignment === undefined ? {} : { alignment }),
      ...(fontId === undefined ? {} : { fontId }),
      ...(type === undefined ? {} : { type }),
    });
  }

  private openRunText(
    element: XlsxXmlElement,
    target: Exclude<TextTarget, 'plain' | null>,
  ): void {
    const run =
      target === 'rich' ? this.currentRichRun : this.currentPhoneticRun;
    if (!run || run.hasText) {
      structureFailure(this.part, 'Shared-string run text is invalid');
    }
    run.hasText = true;
    this.validateXmlSpace(element);
    this.capture = target;
  }

  private closeRichRun(): void {
    if (!this.currentRichRun?.hasText) {
      structureFailure(this.part, 'Shared-string rich run has no text');
    }
    this.richRuns.push(Object.freeze({ text: this.currentRichRun.text }));
    this.currentRichRun = undefined;
  }

  private closePhoneticRun(): void {
    if (!this.currentPhoneticRun?.hasText) {
      structureFailure(this.part, 'Shared-string phonetic run has no text');
    }
    const { end, start, text } = this.currentPhoneticRun;
    this.phoneticRuns.push(Object.freeze({ end, start, text }));
    this.currentPhoneticRun = undefined;
  }

  private closeString(): void {
    const text =
      this.mode === 'rich'
        ? this.richRuns.map((run) => run.text).join('')
        : this.plainText;
    for (const run of this.phoneticRuns) {
      if (run.start >= run.end || run.end > text.length) {
        valueFailure(this.part, 'Shared-string phonetic range is invalid');
      }
    }
    const value: XlsxSharedString = Object.freeze({
      ...(this.phoneticProperties === undefined
        ? {}
        : { phoneticProperties: this.phoneticProperties }),
      ...(this.phoneticRuns.length === 0
        ? {}
        : { phoneticRuns: Object.freeze([...this.phoneticRuns]) }),
      ...(this.mode === 'rich'
        ? { runs: Object.freeze([...this.richRuns]) }
        : {}),
      text,
    });
    this.values.push(value);
  }

  private consumeRun(): void {
    const actual = this.textRuns + 1;
    if (actual > this.limits.maxRichTextRuns) {
      throw new XlsxResourceLimitError(
        'maxRichTextRuns',
        actual,
        this.limits.maxRichTextRuns,
        this.part,
      );
    }
    this.textRuns = actual;
  }

  private consumeText(length: number): void {
    const actual = this.textCharacters + length;
    if (
      !Number.isSafeInteger(actual) ||
      actual > this.limits.maxTextCharacters
    ) {
      throw new XlsxResourceLimitError(
        'maxTextCharacters',
        actual,
        this.limits.maxTextCharacters,
        this.part,
      );
    }
    this.textCharacters = actual;
  }

  private validateXmlSpace(element: XlsxXmlElement): void {
    const value = element.attributes.get(`{${XML_NAMESPACE}}space`);
    if (value !== undefined && value !== 'default' && value !== 'preserve') {
      valueFailure(this.part, 'Shared-string xml:space value is invalid');
    }
  }
}

export async function parseXlsxSharedStringPart(
  part: string,
  dialect: XlsxWorkbookDiscovery['dialect'],
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<XlsxSharedStringTable> {
  const sink = new SharedStringSink(
    part,
    XLSX_SPREADSHEET_NAMESPACES[dialect],
    limits,
  );
  await reader.streamXml(part, sink, { required: true });
  return sink.result();
}
