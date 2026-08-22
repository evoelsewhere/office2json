import { getXlsxRelationshipPartName } from './package-identity';
import { XlsxPartReader } from './part-reader';
import { parseXlsxRelationships, type XlsxRelationship } from './relationships';
import type { ResolvedXlsxResourceLimits } from './resource-limits';

export async function loadXlsxWorksheetRelationships(
  worksheetPart: string,
  reader: XlsxPartReader,
  limits: ResolvedXlsxResourceLimits,
): Promise<ReadonlyMap<string, XlsxRelationship>> {
  const relationshipPart = getXlsxRelationshipPartName(worksheetPart);
  const xml = await reader.readXml(relationshipPart);
  return xml === null
    ? new Map()
    : parseXlsxRelationships(xml, worksheetPart, limits.maxRelationships);
}
