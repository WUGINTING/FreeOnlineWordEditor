// Bounded caches for pure functions of kept XML.
//
// A document repeats the same run / paragraph / cell properties thousands of times (a 63-page
// document: 52 distinct w:rPr over 2,657 runs), and patching one means parsing it into a whole
// XML document. The results are strings or frozen data, so they can be shared safely; never
// cache a DOM element that a caller may change.

/** A small least-recently-used map. */
export class Lru<V> {
  private map = new Map<string, V>();
  constructor(private readonly max: number) {}

  get(key: string): V | undefined {
    const v = this.map.get(key);
    if (v !== undefined) {
      // Most recently used last.
      this.map.delete(key);
      this.map.set(key, v);
    }
    return v;
  }

  set(key: string, value: V): void {
    this.map.set(key, value);
    if (this.map.size > this.max) this.map.delete(this.map.keys().next().value as string);
  }

  clear(): void {
    this.map.clear();
  }

  get size(): number {
    return this.map.size;
  }
}

/** Non-finite numbers would all become null in JSON; keep them apart. */
function replacer(_key: string, value: unknown): unknown {
  return typeof value === 'number' && !Number.isFinite(value) ? { '\u0000nonFinite': String(value) } : value;
}

/**
 * A cache key for the given inputs: distinct for inputs a pure function could tell apart
 * (null vs. "null", NaN vs. null ...). Objects are compared by their JSON, so give them in a
 * fixed key order (as the writer's model builders do); a different order only misses the cache.
 */
export function memoKey(...inputs: unknown[]): string {
  return JSON.stringify(inputs, replacer);
}

/** `f()` cached in `cache` under `key`. */
export function memo<V>(cache: Lru<V>, key: string, f: () => V): V {
  let v = cache.get(key);
  if (v === undefined) {
    v = f();
    cache.set(key, v);
  }
  return v;
}

/** Freeze an object and everything in it (cached data must not be changed by callers). */
export function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const v of Object.values(value as object)) deepFreeze(v);
  }
  return value;
}
