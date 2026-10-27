import { randomUUID } from "node:crypto"
import { link, mkdir, rm, writeFile } from "node:fs/promises"
import { basename, dirname, extname, join, resolve } from "node:path"

import type { Either } from "functype/either"
import { Left, Right } from "functype/either"

// The one write path for tools that save to the caller's disk (save_attachment, export_contacts).
// A file is written in full to a private temp file beside the target, then hard-linked into place:
//   - "wx" on the temp file never follows or reuses something already there.
//   - 0600, because the default folder is tmpdir(), which other users on the machine can list.
//   - link rather than rename: rename silently replaces a file at the target, link fails with
//     EEXIST and leaves it alone. So a save never destroys anything already on disk, and a failed
//     write (a full disk, say) never leaves a truncated file where a good one is expected.
// The trade-off is that a filesystem without hard links (FAT/exFAT, some network shares) fails the
// write instead of risking an overwrite.

const toError = (err: unknown): Error => (err instanceof Error ? err : new Error(String(err)))

const numbered = (target: string, attempt: number): string => {
  if (attempt === 0) return target
  const ext = extname(target)
  return join(resolve(target, ".."), `${basename(target, ext)} (${attempt})${ext}`)
}

const linkUnderFreeName = async (
  temp: string,
  target: string,
  maxNames: number,
  attempt = 0,
): Promise<Either<Error, string>> => {
  if (attempt >= maxNames) {
    return Left(
      new Error(
        maxNames === 1
          ? `${basename(target)} already exists`
          : `${maxNames} files named like ${basename(target)} already exist`,
      ),
    )
  }
  const candidate = numbered(target, attempt)
  try {
    await link(temp, candidate)
    return Right(candidate)
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code === "EEXIST") return linkUnderFreeName(temp, target, maxNames, attempt + 1)
    return Left(toError(err))
  }
}

/**
 * Writes `data` to `dir/name` without replacing anything, and returns the path written.
 * With `maxNames` above 1, a taken name falls back to "name (1).ext", "name (2).ext" and so on.
 */
export const writeNewFile = async (
  dir: string,
  name: string,
  data: string | Buffer,
  maxNames = 1,
): Promise<Either<Error, string>> => {
  const target = join(dir, name)
  // join() on a sanitised name cannot escape dir, but assert it rather than trust the reasoning:
  // save_attachment's names come from whoever sent the email.
  // A prefix test would pass "/out" for "/outside/x", so the parent folder is compared exactly.
  if (dirname(resolve(target)) !== resolve(dir)) return Left(new Error(`Refusing to write outside ${dir}.`))

  const temp = join(dir, `.${name}.${randomUUID()}.tmp`)
  try {
    await mkdir(dir, { recursive: true })
    await writeFile(temp, data, { flag: "wx", mode: 0o600 })
    return await linkUnderFreeName(temp, target, maxNames)
  } catch (err) {
    return Left(toError(err))
  } finally {
    // After a successful link this only drops the second name; a stray temp file is not worth
    // failing a save that has already landed.
    await rm(temp, { force: true }).catch(() => undefined)
  }
}
