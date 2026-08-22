import { XlsxParseError } from '../errors';
import { XlsxPartReader } from './part-reader';
import type { XlsxXmlElement, XlsxXmlEventSink } from './streaming-xml';
import {
  type XlsxWorkbookDiscovery,
  XLSX_SPREADSHEET_NAMESPACES,
} from './workbook-discovery';

function rootFailure(part: string): never {
  throw new XlsxParseError({
    code: 'invalid-document-structure',
    message: 'chartsheet root is missing or has the wrong namespace',
    part,
    severity: 'error',
  });
}

class ChartSheetSink implements XlsxXmlEventSink {
  private rootSeen = false;

  constructor(
    private readonly part: string,
    private readonly namespace: string,
  ) {}

  openElement(element: XlsxXmlElement): void {
    if (this.rootSeen) return;
    if (
      element.localName !== 'chartsheet' ||
      element.namespace !== this.namespace
    ) {
      rootFailure(this.part);
    }
    this.rootSeen = true;
  }

  closeElement(): void {}

  text(): void {}
}

export async function validateXlsxChartSheetPart(
  part: string,
  dialect: XlsxWorkbookDiscovery['dialect'],
  reader: XlsxPartReader,
): Promise<void> {
  const sink = new ChartSheetSink(part, XLSX_SPREADSHEET_NAMESPACES[dialect]);
  await reader.streamXml(part, sink, { required: true });
}
