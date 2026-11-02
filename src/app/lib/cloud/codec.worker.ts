import { encode, decode, digest } from "./codec";
import { streamCloudValue } from './value-chunks';
import localforage from 'localforage';
localforage.config({ name: 'PayrollApp', storeName: 'app_data' });
const outputs = new Map<number, { chunks: ReturnType<typeof streamCloudValue>; sequence: number }>();
const sendNext = (id: number) => {
  const output = outputs.get(id);
  if (!output) return;
  const next = output.chunks.next();
  if (next.done) {
    outputs.delete(id);
    self.postMessage({ id, type: 'codec-complete' });
  } else self.postMessage({ id, type: 'codec-chunk', sequence: ++output.sequence, part: next.value });
};
self.onmessage = async ({ data }) => {
  if (data.type === 'codec-ack') {
    if (outputs.get(data.id)?.sequence === data.sequence) sendNext(data.id);
    return;
  }
  try {
    const result =
      data.operation === "encode"
        ? await encode(data.value)
        : data.operation === 'cache-read' ? await localforage.getItem(String(data.value))
        : await decode(data.value);
    if (data.operation !== 'encode' && data.chunked) {
      outputs.set(data.id, { chunks: streamCloudValue(result), sequence: 0 });
      sendNext(data.id);
      return;
    }
    self.postMessage({
      id: data.id,
      result,
      hash: result instanceof Blob ? await digest(result) : undefined,
    });
  } catch (error) {
    self.postMessage({
      id: data.id,
      error: error instanceof Error ? error.message : "Không thể đọc dữ liệu",
      code: data.operation === 'cache-read' && error instanceof Error && /No available storage method|not allowed|denied/i.test(error.message)
        ? 'CACHE_WORKER_UNAVAILABLE' : undefined,
    });
  }
};
