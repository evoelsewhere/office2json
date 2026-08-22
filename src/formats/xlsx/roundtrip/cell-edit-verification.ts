import type JSZip from 'jszip';

import { canonicalXlsxJson } from './canonical-json';
import { XlsxWriteError } from './errors';
import type {
  XlsxPackageGraph,
  XlsxPackageGraphPart,
} from './internal/package-graph';
import type { XlsxPartFidelity } from './types';

interface XlsxZipStream {
  on(event: 'data', listener: (chunk: Uint8Array) => void): XlsxZipStream;
  on(event: 'end', listener: () => void): XlsxZipStream;
  on(event: 'error', listener: (error: unknown) => void): XlsxZipStream;
  pause(): XlsxZipStream;
  resume(): XlsxZipStream;
}

export function generateBoundedXlsxZip(
  archive: Pick<JSZip, 'generateInternalStream'>,
  maxOutputBytes: number,
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let length = 0;
    const stream = archive.generateInternalStream({
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      platform: 'DOS',
      streamFiles: true,
      type: 'uint8array',
    }) as unknown as XlsxZipStream;
    stream
      .on('data', (chunk) => {
        length += chunk.byteLength;
        if (length > maxOutputBytes) {
          chunks.length = 0;
          stream.pause();
          reject(
            new XlsxWriteError(
              'resource-limit-exceeded',
              'XLSX generated package exceeds its output byte limit',
              {
                actual: length,
                limit: maxOutputBytes,
                limitName: 'maxOutputBytes',
              },
            ),
          );
          return;
        }
        chunks.push(chunk);
      })
      .on('error', () => {
        reject(
          new XlsxWriteError(
            'generated-package-invalid',
            'Failed to generate the XLSX output package',
          ),
        );
      })
      .on('end', () => {
        const output = new Uint8Array(length);
        let offset = 0;
        for (const chunk of chunks) {
          output.set(chunk, offset);
          offset += chunk.byteLength;
        }
        resolve(output);
      })
      .resume();
  });
}

export function xlsxCellEditPartTopologyEqual(
  source: XlsxPackageGraphPart,
  output: XlsxPackageGraphPart,
): boolean {
  return (
    source.contentType === output.contentType &&
    source.name === output.name &&
    source.relationshipPart === output.relationshipPart
  );
}

export function verifyXlsxCellEditR1Parts(
  source: XlsxPackageGraph,
  output: XlsxPackageGraph,
  dirtyParts: ReadonlySet<string>,
  addedParts: ReadonlySet<string> = new Set(),
  changedRelationshipOwners: ReadonlySet<string> = new Set(),
): XlsxPartFidelity[] {
  const sourceRelationships = source.relationships.filter(
    (relationship) =>
      relationship.owner === null ||
      !changedRelationshipOwners.has(relationship.owner),
  );
  const outputRelationships = output.relationships.filter(
    (relationship) =>
      relationship.owner === null ||
      !changedRelationshipOwners.has(relationship.owner),
  );
  if (
    source.conformance !== output.conformance ||
    source.parts.length + addedParts.size !== output.parts.length ||
    canonicalXlsxJson(sourceRelationships) !==
      canonicalXlsxJson(outputRelationships)
  ) {
    throw new XlsxWriteError(
      'generated-package-invalid',
      'XLSX edited package topology differs from its source',
    );
  }
  const sourceByName = new Map(source.parts.map((part) => [part.name, part]));
  const outputNames = new Set(output.parts.map((part) => part.name));
  for (const sourcePart of source.parts) {
    if (!outputNames.has(sourcePart.name)) {
      throw new XlsxWriteError(
        'generated-package-invalid',
        'XLSX edited package removed a source part unexpectedly',
        { part: sourcePart.name },
      );
    }
  }
  const parts: XlsxPartFidelity[] = [];
  for (const outputPart of output.parts) {
    const sourcePart = sourceByName.get(outputPart.name);
    if (!sourcePart) {
      if (!addedParts.has(outputPart.name)) {
        throw new XlsxWriteError(
          'generated-package-invalid',
          'XLSX edited package contains an unexpected part',
          { part: outputPart.name },
        );
      }
      parts.push({
        byteLength: outputPart.byteLength,
        disposition: 'add',
        name: outputPart.name,
        sha256: outputPart.sha256,
      });
      continue;
    }
    if (!xlsxCellEditPartTopologyEqual(sourcePart, outputPart)) {
      throw new XlsxWriteError(
        'generated-package-invalid',
        'XLSX edited package contains an unexpected part',
        { part: outputPart.name },
      );
    }
    const disposition = dirtyParts.has(outputPart.name) ? 'patch' : 'copy';
    if (
      disposition === 'copy' &&
      (sourcePart.byteLength !== outputPart.byteLength ||
        sourcePart.sha256 !== outputPart.sha256)
    ) {
      throw new XlsxWriteError(
        'generated-package-invalid',
        'XLSX copied part bytes changed unexpectedly',
        { part: outputPart.name },
      );
    }
    parts.push({
      byteLength: outputPart.byteLength,
      disposition,
      name: outputPart.name,
      sha256: outputPart.sha256,
      sourceByteLength: sourcePart.byteLength,
      sourceSha256: sourcePart.sha256,
    });
  }
  return parts;
}
