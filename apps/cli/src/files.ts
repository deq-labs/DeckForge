import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

/** An error whose message is already written for the user; the CLI prints it without a stack. */
export class UserError extends Error {}

export function readInput(path: string, what: string): Uint8Array {
  try {
    return new Uint8Array(readFileSync(path));
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (code === 'ENOENT') throw new UserError(`Cannot find the ${what} "${path}". Check the path and try again.`);
    if (code === 'EISDIR') throw new UserError(`"${path}" is a folder; give the path to the ${what} file.`);
    if (code === 'EBUSY' || code === 'EPERM') throw new UserError(`Cannot read "${path}". If it is open in another program, close it and try again.`);
    throw e;
  }
}

export function writeOutput(path: string, data: Uint8Array, opts: { overwrite: boolean; what: string }): void {
  if (!opts.overwrite && existsSync(path)) {
    throw new UserError(`"${path}" already exists. Use --force to replace it, or choose another name with -o.`);
  }
  try {
    mkdirSync(dirname(resolve(path)), { recursive: true });
    writeFileSync(path, data);
  } catch (e) {
    const code = (e as NodeJS.ErrnoException).code;
    if (code === 'EBUSY' || code === 'EPERM' || code === 'EACCES') {
      throw new UserError(`Cannot write the ${opts.what} "${path}". If it is open in PowerPoint or Excel, close it and try again.`);
    }
    throw e;
  }
}

export function samePath(a: string, b: string): boolean {
  const norm = (p: string): string => resolve(p).toLowerCase(); // case-insensitive on Windows/macOS defaults
  return norm(a) === norm(b);
}
