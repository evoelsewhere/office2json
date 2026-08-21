import { parse } from './parser';
import type {
  PptxDocument,
  PptxInput,
  PptxParseOptions,
  PptxParseResult,
} from './types';

export { validatePptxScene } from './scene-validation';
export { createPptx } from './creator';
export { renderPptxDocumentToSvg, renderPptxToSvg } from './render-svg';
export {
  defaultPptxRoundTripPortableLimits,
  parsePptxRoundTripJson,
  PptxRoundTripPortableLimitError,
  serializePptxRoundTripJson,
} from './roundtrip/portable';
export { readPptxRoundTrip } from './roundtrip/read';
export {
  replacePptxRoundTripText,
  setPptxRoundTripChartTransform,
  setPptxRoundTripGroupTransform,
  setPptxRoundTripImageCrop,
  setPptxRoundTripImageTransform,
  setPptxRoundTripShapeTransform,
  setPptxRoundTripTableTransform,
  setPptxRoundTripTextTransform,
} from './roundtrip/edit';
export { writePptxRoundTrip } from './roundtrip/write';

export { PptxParseError } from './errors';
export { PptxRenderError } from './render-error';
export { PptxWriteError } from './write-error';

/** Parse a PowerPoint Open XML package into the current structured JSON model. */
export async function parsePptx(
  input: PptxInput,
  options: PptxParseOptions = {},
): Promise<PptxDocument> {
  return parse(input, options);
}

/** Parse a PowerPoint package and return recoverable diagnostics. */
export async function parsePptxWithDiagnostics(
  input: PptxInput,
  options: PptxParseOptions = {},
): Promise<PptxParseResult> {
  const diagnostics: PptxParseResult['diagnostics'] = [];
  const document = await parse(input, options, diagnostics);
  return { document, diagnostics };
}

export type * from './types';
export type * from './scene-types';
export type * from './write-types';
export type * from './roundtrip/types';
export type * from './render-types';
