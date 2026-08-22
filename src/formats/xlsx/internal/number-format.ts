import type { XlsxCellValue } from '../types';

export type XlsxDatePrecision = Extract<
  XlsxCellValue,
  { kind: 'date' }
>['precision'];

function builtinFormats(): Readonly<Record<number, string>> {
  return {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
  };
}

export function xlsxBuiltinNumberFormatCode(id: number): string | undefined {
  return builtinFormats()[id];
}

function splitSections(code: string): string[] | undefined {
  const sections: string[] = [];
  let current = '';
  let bracketDepth = 0;
  let quoted = false;
  let skipUntil = 0;
  for (const [index, character] of code.split('').entries()) {
    if (index < skipUntil) continue;
    if (character === '"') {
      quoted = !quoted;
      current += character;
      continue;
    }
    if (
      !quoted &&
      (character === '\\' || character === '_' || character === '*')
    ) {
      current += character;
      if (index === code.length - 1) return undefined;
      current += '0';
      skipUntil = index + 2;
      continue;
    }
    if (!quoted && character === '[') bracketDepth += 1;
    if (!quoted && character === ']') {
      bracketDepth -= 1;
    }
    if (!quoted && bracketDepth === 0 && character === ';') {
      sections.push(current);
      current = '';
      continue;
    }
    current += character;
  }
  if (quoted || bracketDepth !== 0) return undefined;
  sections.push(current);
  return sections.length <= 4 ? sections : undefined;
}

function containsConditionalSection(section: string): boolean {
  return /\[(?:<=|>=|<>|=|<|>)/u.test(section);
}

function selectNumericSection(code: string, value: number): string | undefined {
  const sections = splitSections(code);
  if (!sections || sections.some(containsConditionalSection)) return undefined;
  if (sections.length === 1) return sections[0];
  if (sections.length === 2) return sections[value < 0 ? 1 : 0];
  if (value > 0) return sections[0];
  if (value < 0) return sections[1];
  return sections[2];
}

function elapsedToken(value: string): boolean {
  const normalized = value.toLowerCase();
  return /^(?:h+|m+|s+)$/u.test(normalized);
}

function classifySection(section: string): XlsxDatePrecision | undefined {
  let hasDate = false;
  let hasMinutesOrMonths = false;
  let hasTime = false;
  let skipUntil = 0;
  for (const [index, character] of section.split('').entries()) {
    if (index < skipUntil) continue;
    if (character === '"') {
      const end = section.indexOf('"', index + 1);
      skipUntil = end + 1;
      continue;
    }
    if (character === '[') {
      const end = section.indexOf(']', index + 1);
      if (elapsedToken(section.slice(index + 1, end))) return 'duration';
      skipUntil = end + 1;
      continue;
    }
    const remainder = section.slice(index).toUpperCase();
    if (remainder.startsWith('AM/PM')) {
      hasTime = true;
      continue;
    }
    if (remainder.startsWith('A/P')) {
      hasTime = true;
      continue;
    }
    const lower = character.toLowerCase();
    if (lower === 'y' || lower === 'd') hasDate = true;
    else if (lower === 'h' || lower === 's') hasTime = true;
    else if (lower === 'm') hasMinutesOrMonths = true;
  }
  if (hasMinutesOrMonths && !hasTime) hasDate = true;
  if (hasDate && hasTime) return 'date-time';
  if (hasDate) return 'date';
  return hasTime ? 'time' : undefined;
}

export function xlsxNumberFormatDatePrecision(
  code: string,
  value: number,
): XlsxDatePrecision | undefined {
  if (!Number.isFinite(value)) return undefined;
  const section = selectNumericSection(code, value);
  return section === undefined ? undefined : classifySection(section);
}
