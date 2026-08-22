const URI_SCHEME = /^[A-Za-z][A-Za-z\d+.-]*:/;

function hasControlCharacter(value: string): boolean {
  for (const character of value) {
    const codePoint = character.codePointAt(0) ?? 0;
    if (codePoint <= 0x1f || codePoint === 0x7f) {
      return true;
    }
  }
  return false;
}

function invalidPartName(message: string): never {
  throw new TypeError(`Invalid XLSX package part name: ${message}`);
}

function decodedSegment(rawSegment: string): string {
  let decoded: string;
  try {
    decoded = decodeURIComponent(rawSegment);
  } catch (cause) {
    throw new TypeError('Invalid XLSX package part name: percent encoding', {
      cause,
    });
  }

  if (!decoded) invalidPartName('empty path segment');
  if (decoded === '.' || decoded === '..') {
    invalidPartName('dot segment');
  }
  if (
    decoded.includes('/') ||
    decoded.includes('\\') ||
    decoded.includes('?') ||
    decoded.includes('#')
  ) {
    invalidPartName('encoded path delimiter');
  }
  if (hasControlCharacter(decoded)) {
    invalidPartName('control character');
  }
  return decoded;
}

function rawPartPath(value: string): string {
  if (value.includes('\\')) invalidPartName('backslash delimiter');
  if (value.includes('?') || value.includes('#')) {
    invalidPartName('query or fragment');
  }
  const withoutRootSlash = value.startsWith('/') ? value.slice(1) : value;
  if (!withoutRootSlash || withoutRootSlash.startsWith('/')) {
    invalidPartName('empty rooted name');
  }
  return withoutRootSlash;
}

/** Normalize an OPC package part to its unambiguous case-sensitive identity. */
export function canonicalizeXlsxPartName(value: string): string {
  const canonical = rawPartPath(value).split('/').map(decodedSegment).join('/');
  if (URI_SCHEME.test(canonical)) {
    invalidPartName('URI scheme');
  }
  return canonical;
}

/** Resolve an internal relationship target relative to the part that owns it. */
export function resolveXlsxPartTarget(
  ownerPart: string,
  target: string,
): string {
  const ownerSegments = canonicalizeXlsxPartName(ownerPart).split('/');
  ownerSegments.pop();

  if (!target || target !== target.trim()) {
    invalidPartName('empty or padded relationship target');
  }
  if (URI_SCHEME.test(target)) invalidPartName('external URI scheme');
  if (target.startsWith('//')) invalidPartName('network-path reference');

  const rooted = target.startsWith('/');
  const targetPath = rawPartPath(target);
  const resolved = rooted ? [] : ownerSegments;

  for (const rawSegment of targetPath.split('/')) {
    if (rawSegment === '..') {
      if (resolved.length === 0) {
        invalidPartName('relationship target escapes package root');
      }
      resolved.pop();
    } else {
      if (rawSegment === '.') invalidPartName('dot target segment');
      resolved.push(decodedSegment(rawSegment));
    }
  }

  if (resolved.length === 0) invalidPartName('empty relationship target');
  return resolved.join('/');
}

/** Resolve an internal package relationship target against the package root. */
export function resolveXlsxRootTarget(target: string): string {
  if (!target || target !== target.trim()) {
    invalidPartName('empty or padded relationship target');
  }
  if (URI_SCHEME.test(target)) invalidPartName('external URI scheme');
  if (target.startsWith('//')) invalidPartName('network-path reference');
  return canonicalizeXlsxPartName(target);
}

/** Return the relationship part owned by an XLSX package part. */
export function getXlsxRelationshipPartName(ownerPart: string): string {
  const canonicalOwner = canonicalizeXlsxPartName(ownerPart);
  const separatorIndex = canonicalOwner.lastIndexOf('/');
  const directory = canonicalOwner.slice(0, separatorIndex + 1);
  const filename = canonicalOwner.slice(separatorIndex + 1);
  return `${directory}_rels/${filename}.rels`;
}
