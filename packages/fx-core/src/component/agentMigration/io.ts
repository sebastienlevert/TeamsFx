// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import { open } from "fs/promises";
import { constants } from "fs";
import path from "path";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { Artifacts } from "./model";
import { cancelled, isErrno, migrationError } from "./errors";
import { packageLimits } from "./paths";

/** Fault-injection seam at actual I/O boundaries; algorithms always use real artifacts. */
export const agentMigrationIo = {
  promote: (source: string, destination: string): Promise<void> => fs.rename(source, destination),
  replace: (source: string, destination: string): Promise<void> => fs.rename(source, destination),
  beforeCommit: (): Promise<void> => Promise.resolve(),
};

export async function safeAncestors(target: string): Promise<Result<undefined, FxError>> {
  const full = path.resolve(target);
  const root = path.parse(full).root;
  let current = root;
  try {
    for (const part of full.slice(root.length).split(path.sep).filter(Boolean)) {
      current = path.join(current, part);
      const stat = await fs.lstat(current);
      if (stat.isSymbolicLink()) return err(migrationError("AgentPackagePathInvalid"));
    }
    return ok(undefined);
  } catch (error) {
    return err(
      migrationError(
        isErrno(error, "ENOENT") ? "AgentPackageSourceInvalid" : "AgentMigrationIoError",
        error
      )
    );
  }
}

export async function readFile(
  file: string,
  limit = packageLimits.fileBytes
): Promise<Result<Buffer, FxError>> {
  const safe = await safeAncestors(file);
  if (safe.isErr()) return err(safe.error);
  try {
    const handle = await open(file, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0));
    try {
      const stat = await handle.stat();
      if (!stat.isFile() || stat.nlink !== 1) return err(migrationError("AgentPackagePathInvalid"));
      if (stat.size > limit) return err(migrationError("AgentPackageLimitExceeded"));
      const bytes = Buffer.alloc(stat.size + 1);
      let count = 0;
      while (count < bytes.length) {
        const read = await handle.read(bytes, count, bytes.length - count, count);
        if (read.bytesRead === 0) break;
        count += read.bytesRead;
      }
      const after = await handle.stat();
      if (count > limit) return err(migrationError("AgentPackageLimitExceeded"));
      if (
        stat.size !== count ||
        stat.size !== after.size ||
        stat.mtimeMs !== after.mtimeMs ||
        stat.ctimeMs !== after.ctimeMs
      )
        return err(migrationError("AgentEditsConflict"));
      return ok(bytes.subarray(0, count));
    } finally {
      await handle.close();
    }
  } catch (error) {
    return err(
      migrationError(
        isErrno(error, "ENOENT") ? "AgentPackageReferenceMissing" : "AgentMigrationIoError",
        error
      )
    );
  }
}

export async function writeArtifacts(
  root: string,
  files: Artifacts,
  signal?: AbortSignal
): Promise<Result<undefined, FxError>> {
  try {
    for (const [name, bytes] of files) {
      const cancel = cancelled(signal);
      if (cancel) return err(cancel);
      const file = path.join(root, ...name.split("/"));
      await fs.mkdir(path.dirname(file), { recursive: true });
      await fs.writeFile(file, bytes, { flag: "wx" });
    }
    return ok(undefined);
  } catch (error) {
    return err(migrationError("AgentMigrationIoError", error));
  }
}

export async function requireAbsent(file: string): Promise<Result<undefined, FxError>> {
  try {
    await fs.lstat(file);
    return err(migrationError("AgentPackageDestinationExists"));
  } catch (error) {
    return isErrno(error, "ENOENT")
      ? ok(undefined)
      : err(migrationError("AgentMigrationIoError", error));
  }
}

export async function prepareParent(parent: string): Promise<Result<string[], FxError>> {
  const missing: string[] = [];
  const created: string[] = [];
  let current = parent;
  try {
    for (;;) {
      try {
        const stat = await fs.lstat(current);
        if (!stat.isDirectory() || stat.isSymbolicLink())
          return err(migrationError("AgentPackagePathInvalid"));
        break;
      } catch (error) {
        if (!isErrno(error, "ENOENT")) return err(migrationError("AgentMigrationIoError", error));
        missing.push(current);
        const ancestor = path.dirname(current);
        if (ancestor === current) return err(migrationError("AgentPackagePathInvalid"));
        current = ancestor;
      }
    }
    const safe = await safeAncestors(current);
    if (safe.isErr()) return err(safe.error);
    for (const directory of missing.reverse()) {
      await fs.mkdir(directory);
      created.push(directory);
    }
    return ok(created);
  } catch (error) {
    try {
      for (const directory of created.reverse()) await fs.rmdir(directory);
    } catch (cleanupError) {
      return err(migrationError("AgentMigrationRecoveryRequired", cleanupError));
    }
    return err(migrationError("AgentMigrationIoError", error));
  }
}
