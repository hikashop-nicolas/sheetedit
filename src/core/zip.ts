// Promise wrappers over fflate's asynchronous zip/unzip. The async functions run the
// deflate/inflate on a Web Worker (fflate spins one up internally), so compressing a
// large workbook on save no longer blocks the main thread. The synchronous variants
// stay available for tests and tiny fixed-size payloads where a worker is not worth it.
import { unzip, zip, type AsyncZippable, type Unzipped } from "fflate";

export function zipAsync(files: AsyncZippable): Promise<Uint8Array> {
  return new Promise((resolve, reject) => zip(files, (err, data) => (err ? reject(err) : resolve(data))));
}

export function unzipAsync(bytes: Uint8Array): Promise<Unzipped> {
  return new Promise((resolve, reject) => unzip(bytes, (err, data) => (err ? reject(err) : resolve(data))));
}
