const MAX_CANONICAL_DEPTH = 128;

function canonicalValue(
  value: unknown,
  active: WeakSet<object>,
  depth: number,
): string {
  if (depth > MAX_CANONICAL_DEPTH) {
    throw new TypeError('Canonical XLSX JSON exceeds maximum depth');
  }
  if (value === null) return 'null';
  if (typeof value === 'boolean' || typeof value === 'string') {
    return JSON.stringify(value);
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new TypeError('Canonical XLSX JSON requires finite numbers');
    }
    return JSON.stringify(value);
  }
  if (typeof value !== 'object') {
    throw new TypeError(`Canonical XLSX JSON does not support ${typeof value}`);
  }
  if (active.has(value)) {
    throw new TypeError('Canonical XLSX JSON does not support cycles');
  }
  active.add(value);
  try {
    if (Array.isArray(value)) {
      if (Object.getPrototypeOf(value) !== Array.prototype) {
        throw new TypeError('Canonical XLSX JSON requires plain arrays');
      }
      return `[${value
        .map((child) => canonicalValue(child, active, depth + 1))
        .join(',')}]`;
    }
    if (Object.getPrototypeOf(value) !== Object.prototype) {
      throw new TypeError('Canonical XLSX JSON requires plain objects');
    }
    const record = value as Record<string, unknown>;
    return `{${Object.keys(record)
      .sort()
      .map(
        (key) =>
          `${JSON.stringify(key)}:${canonicalValue(record[key], active, depth + 1)}`,
      )
      .join(',')}}`;
  } finally {
    active.delete(value);
  }
}

export function canonicalXlsxJson(value: unknown): string {
  return canonicalValue(value, new WeakSet<object>(), 0);
}
