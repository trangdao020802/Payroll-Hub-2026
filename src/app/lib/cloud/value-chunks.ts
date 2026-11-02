type Path = (string | number)[];
export type CloudValueChunk =
  | { kind: 'value'; path: Path; value: unknown }
  | { kind: 'object'; path: Path }
  | { kind: 'array'; path: Path; length: number }
  | { kind: 'items'; path: Path; offset: number; values: unknown[] };

// Count cells rather than rows: a wide imported row must not turn a supposedly
// small message into a clone of a complete nested table.
function fits(value: unknown, budget: { left: number }): boolean {
  if (--budget.left < 0) return false;
  if (!value || typeof value !== 'object' || value instanceof Blob || value instanceof Date) return true;
  for (const key of Object.keys(value)) {
    if (!fits((value as Record<string, unknown>)[key], budget)) return false;
  }
  return true;
}

/** Used only for worker-owned decoded values. Release sent data after each ack. */
export function* streamCloudValue(value: unknown, path: Path = []): Generator<CloudValueChunk> {
  if (Array.isArray(value)) {
    yield { kind: 'array', path, length: value.length };
    for (let offset = 0; offset < value.length;) {
      const values: unknown[] = [];
      const budget = { left: 8_000 };
      while (values.length < 500 && offset + values.length < value.length && fits(value[offset + values.length], budget)) {
        values.push(value[offset + values.length]);
      }
      if (values.length) {
        yield { kind: 'items', path, offset, values };
        for (let i = 0; i < values.length; i++) delete value[offset + i];
        offset += values.length;
      } else {
        yield* streamCloudValue(value[offset], [...path, offset]);
        delete value[offset++];
      }
    }
  } else if (value && typeof value === 'object' && !fits(value, { left: 1_000 })) {
    yield { kind: 'object', path };
    for (const key of Object.keys(value)) {
      yield* streamCloudValue((value as Record<string, unknown>)[key], [...path, key]);
      delete (value as Record<string, unknown>)[key];
    }
  } else {
    yield { kind: 'value', path, value };
  }
}

/** Reconstruct the original shape without a single giant structured clone. */
export function createCloudValueReceiver() {
  let root: unknown;
  let started = false;
  const arrays: { value: unknown[]; next: number }[] = [];
  const positions = new WeakMap<unknown[], { value: unknown[]; next: number }>();
  const at = (path: Path): unknown => {
    let value = root;
    for (const key of path) {
      if (!value || typeof value !== 'object' || !Object.hasOwn(value, key)) throw new Error('Thiếu phần dữ liệu cloud.');
      value = (value as Record<string | number, unknown>)[key];
    }
    return value;
  };
  const assign = (path: Path, value: unknown) => {
    if (!path.length) {
      if (started) throw new Error('Dữ liệu cloud bị lặp.');
      root = value; started = true; return;
    }
    const parent = at(path.slice(0, -1));
    if (!parent || typeof parent !== 'object') throw new Error('Sai cấu trúc dữ liệu cloud.');
    const key = path[path.length - 1];
    if (Array.isArray(parent)) {
      const position = positions.get(parent);
      if (!position || key !== position.next || position.next >= parent.length) throw new Error('Sai thứ tự dữ liệu cloud.');
      position.next++;
    } else if (Object.hasOwn(parent, key)) throw new Error('Dữ liệu cloud bị lặp.');
    // Imported column names such as __proto__ are data, never setters.
    Object.defineProperty(parent, key, { value, enumerable: true, writable: true, configurable: true });
  };
  return {
    accept(part: CloudValueChunk) {
      if (part.kind === 'items') {
        const target = at(part.path);
        if (!Array.isArray(target)) throw new Error('Sai cấu trúc bảng cloud.');
        for (let i = 0; i < part.values.length; i++) assign([...part.path, part.offset + i], part.values[i]);
      } else if (part.kind === 'array') {
        const value = new Array(part.length);
        const position = { value, next: 0 };
        positions.set(value, position); arrays.push(position); assign(part.path, value);
      } else assign(part.path, part.kind === 'object' ? {} : part.value);
    },
    finish() {
      if (!started || arrays.some(item => item.next !== item.value.length)) throw new Error('Dữ liệu cloud chưa đầy đủ.');
      return root;
    },
  };
}
