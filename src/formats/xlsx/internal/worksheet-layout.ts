import type { XlsxColumnRange, XlsxRange } from '../types';

export interface XlsxAuthoredColumnRange extends XlsxColumnRange {
  order: number;
}

interface HeapEntry {
  end: number;
  order: number;
  value: XlsxAuthoredColumnRange;
}

function higherPriority(left: HeapEntry, right: HeapEntry): boolean {
  const difference = left.order - right.order || left.end - right.end;
  return Math.sign(difference) === 1;
}

function heapPush(heap: HeapEntry[], value: HeapEntry): void {
  heap.push(value);
  let index = heap.length - 1;
  Array.from({ length: Math.ceil(Math.log2(heap.length + 1)) }).forEach(() => {
    if (index === 0) return;
    const parent = Math.floor((index - 1) / 2);
    if (higherPriority(heap[parent]!, value)) return;
    heap[index] = heap[parent]!;
    index = parent;
  });
  heap[index] = value;
}

function heapPop(heap: HeapEntry[]): HeapEntry {
  const first = heap[0]!;
  const last = heap.pop()!;
  if (heap.length === 0) return first;
  let index = 0;
  Array.from({ length: Math.ceil(Math.log2(heap.length + 1)) }).forEach(() => {
    const left = index * 2 + 1;
    if (left >= heap.length) return;
    const right = left + 1;
    const child =
      right < heap.length && higherPriority(heap[right]!, heap[left]!)
        ? right
        : left;
    if (higherPriority(last, heap[child]!)) return;
    heap[index] = heap[child]!;
    index = child;
  });
  heap[index] = last;
  return first;
}

function propertyKey(value: XlsxAuthoredColumnRange): string {
  return JSON.stringify([
    value.collapsed,
    value.hidden,
    value.outlineLevel,
    value.style,
    value.width,
  ]);
}

function outputRange(
  value: XlsxAuthoredColumnRange,
  start: number,
  end: number,
): XlsxColumnRange {
  return {
    end,
    ...(value.collapsed === undefined ? {} : { collapsed: value.collapsed }),
    ...(value.hidden === undefined ? {} : { hidden: value.hidden }),
    ...(value.outlineLevel === undefined
      ? {}
      : { outlineLevel: value.outlineLevel }),
    start,
    ...(value.style === undefined ? {} : { style: value.style }),
    ...(value.width === undefined ? {} : { width: value.width }),
  };
}

export function normalizeXlsxColumnRanges(
  values: readonly XlsxAuthoredColumnRange[],
): XlsxColumnRange[] {
  const starts = new Map<number, XlsxAuthoredColumnRange[]>();
  const boundaries = new Set<number>();
  for (const value of values) {
    const entries = starts.get(value.start);
    if (entries) entries.push(value);
    else starts.set(value.start, [value]);
    boundaries.add(value.start);
    boundaries.add(value.end + 1);
  }
  const positions = [...boundaries].sort((left, right) => left - right);
  const intervals = positions
    .slice(1)
    .map((nextPosition, index) => [positions[index]!, nextPosition] as const);
  const heap: HeapEntry[] = [];
  const output: XlsxColumnRange[] = [];
  let previousKey: string | undefined;
  for (const [position, nextPosition] of intervals) {
    for (const value of starts.get(position) ?? []) {
      heapPush(heap, { end: value.end, order: value.order, value });
    }
    Array.from({ length: heap.length }).forEach(() => {
      if (heap[0] && heap[0].end < position) heapPop(heap);
    });
    const active = heap[0]?.value;
    if (!active) {
      previousKey = undefined;
      continue;
    }
    const end = nextPosition - 1;
    const key = propertyKey(active);
    const previous = output.at(-1);
    if (previous && previousKey === key) {
      previous.end = end;
    } else {
      output.push(outputRange(active, position, end));
    }
    previousKey = key;
  }
  return output;
}

class ColumnOccupancy {
  private readonly words = new Map<number, number>();

  overlaps(start: number, end: number): boolean {
    let overlap = false;
    this.eachWord(start, end, (index, mask) => {
      if (((this.words.get(index) ?? 0) & mask) !== 0) overlap = true;
    });
    return overlap;
  }

  set(start: number, end: number, occupied: boolean): void {
    this.eachWord(start, end, (index, mask) => {
      const current = this.words.get(index) ?? 0;
      const next = occupied ? current | mask : current & ~mask;
      this.words.set(index, next);
    });
  }

  private eachWord(
    start: number,
    end: number,
    callback: (index: number, mask: number) => void,
  ): void {
    const firstWord = Math.floor((start - 1) / 32);
    const lastWord = Math.floor((end - 1) / 32);
    Array.from({ length: lastWord - firstWord + 1 }).forEach((_, offset) => {
      const index = firstWord + offset;
      const firstBit = index === firstWord ? (start - 1) % 32 : 0;
      const lastBit = index === lastWord ? (end - 1) % 32 : 31;
      const leadingMask = 0xffff_ffff << firstBit;
      const trailingMask = 2 ** (lastBit + 1) - 1;
      callback(index, leadingMask & trailingMask);
    });
  }
}

export function xlsxMergedRangesOverlap(ranges: readonly XlsxRange[]): boolean {
  const additions = new Map<number, XlsxRange[]>();
  const removals = new Map<number, XlsxRange[]>();
  const rows = new Set<number>();
  for (const range of ranges) {
    const startRow = range.start.row;
    const endRow = range.end.row + 1;
    const starting = additions.get(startRow);
    if (starting) starting.push(range);
    else additions.set(startRow, [range]);
    const ending = removals.get(endRow);
    if (ending) ending.push(range);
    else removals.set(endRow, [range]);
    rows.add(startRow);
    rows.add(endRow);
  }
  const occupancy = new ColumnOccupancy();
  for (const row of [...rows].sort((left, right) => left - right)) {
    for (const range of removals.get(row) ?? []) {
      occupancy.set(range.start.column, range.end.column, false);
    }
    for (const range of additions.get(row) ?? []) {
      if (occupancy.overlaps(range.start.column, range.end.column)) return true;
      occupancy.set(range.start.column, range.end.column, true);
    }
  }
  return false;
}
