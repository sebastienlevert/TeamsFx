// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import * as lockfile from "proper-lockfile";
import { ConfigFolderName, err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { getLockFolder } from "../../core/middleware/concurrentLocker";
import { cancelled, isErrno, migrationError } from "./errors";
import { readDirectory } from "./intake";
import { agentMigrationIo, readFile, safeAncestors, writeArtifacts } from "./io";
import { artifactDigest, Artifacts, digest, jsonBytes } from "./model";

export function recoveryPath(project: string): string {
  return path.join(path.dirname(project), `.${path.basename(project)}.agent-edit`);
}

export async function checkRecovery(project: string): Promise<Result<undefined, FxError>> {
  try {
    await fs.lstat(recoveryPath(project));
    return err(migrationError("AgentMigrationRecoveryRequired"));
  } catch (error) {
    return isErrno(error, "ENOENT")
      ? ok(undefined)
      : err(migrationError("AgentMigrationIoError", error));
  }
}

async function currentBytes(file: string): Promise<Result<Buffer | undefined, FxError>> {
  try {
    await fs.lstat(file);
    return readFile(file);
  } catch (error) {
    return isErrno(error, "ENOENT")
      ? ok(undefined)
      : err(migrationError("AgentMigrationIoError", error));
  }
}

function equal(left: Buffer | undefined, right: Buffer | undefined): boolean {
  return left === undefined ? right === undefined : right !== undefined && left.equals(right);
}

export async function commitEdits(
  project: string,
  before: Artifacts,
  after: Artifacts,
  signal?: AbortSignal
): Promise<Result<undefined, FxError>> {
  const changes = [...after].filter(([name, bytes]) => !before.get(name)?.equals(bytes));
  const journal = recoveryPath(project);
  let ownedJournal = false;
  let retainJournal = false;
  let release: (() => Promise<void>) | undefined;
  const attempted: string[] = [];
  const createdDirectories: string[] = [];
  let result: Result<undefined, FxError>;

  async function ensureParent(file: string): Promise<Result<undefined, FxError>> {
    const relative = path.relative(project, path.dirname(file));
    let current = project;
    for (const part of relative.split(path.sep).filter(Boolean)) {
      current = path.join(current, part);
      try {
        await fs.mkdir(current);
        createdDirectories.push(current);
      } catch (error) {
        if (!isErrno(error, "EEXIST")) return err(migrationError("AgentMigrationIoError", error));
      }
      const safe = await safeAncestors(current);
      if (safe.isErr()) return err(safe.error);
    }
    return ok(undefined);
  }

  async function commit(): Promise<Result<undefined, FxError>> {
    await agentMigrationIo.beforeCommit();
    const current = await readDirectory(project, signal, true);
    if (current.isErr()) return err(current.error);
    if (artifactDigest(current.value) !== artifactDigest(before))
      return err(migrationError("AgentEditsConflict"));
    for (const [name] of changes) {
      const cancel = cancelled(signal);
      if (cancel) return err(cancel);
      const target = path.join(project, ...name.split("/"));
      const current = await currentBytes(target);
      if (current.isErr()) return err(current.error);
      if (!equal(current.value, before.get(name))) return err(migrationError("AgentEditsConflict"));
      const directory = await ensureParent(target);
      if (directory.isErr()) return err(directory.error);
      attempted.push(name);
      await agentMigrationIo.replace(path.join(journal, "after", ...name.split("/")), target);
    }
    const committed = await readDirectory(project, undefined, true);
    if (committed.isErr()) return err(committed.error);
    if (artifactDigest(committed.value) !== artifactDigest(after)) {
      return err(migrationError("AgentEditsConflict"));
    }
    return ok(undefined);
  }

  async function rollback(): Promise<Result<undefined, FxError>> {
    for (const name of [...attempted].reverse()) {
      const target = path.join(project, ...name.split("/"));
      const current = await currentBytes(target);
      if (current.isErr()) return err(current.error);
      if (equal(current.value, before.get(name))) continue;
      if (!equal(current.value, after.get(name)))
        return err(migrationError("AgentMigrationRecoveryRequired"));
      if (before.has(name))
        await agentMigrationIo.replace(path.join(journal, "before", ...name.split("/")), target);
      else await fs.unlink(target);
    }
    for (const directory of [...createdDirectories].reverse()) await fs.rmdir(directory);
    return ok(undefined);
  }

  try {
    const lockDirectory = getLockFolder(project);
    await fs.mkdir(lockDirectory, { recursive: true });
    release = await lockfile.lock(project, {
      retries: 0,
      lockfilePath: path.join(lockDirectory, `${ConfigFolderName}.lock`),
    });
    const recovery = await checkRecovery(project);
    if (recovery.isErr()) result = err(recovery.error);
    else {
      await fs.mkdir(journal);
      ownedJournal = true;
      const prepared: Artifacts = new Map();
      for (const [name, bytes] of changes) {
        prepared.set(`after/${name}`, bytes);
        const old = before.get(name);
        if (old) prepared.set(`before/${name}`, old);
      }
      prepared.set(
        "journal.json",
        jsonBytes({
          version: 1,
          projectPath: project,
          files: changes.map(([name, bytes]) => ({
            path: name,
            before: before.has(name) ? digest(before.get(name)!) : null,
            after: digest(bytes),
          })),
        })
      );
      const preparation = await writeArtifacts(journal, prepared, signal);
      result = preparation.isErr() ? err(preparation.error) : await commit();
    }
  } catch (error) {
    result = err(
      migrationError(
        isErrno(error, "ELOCKED") ? "AgentEditsConflict" : "AgentMigrationIoError",
        error
      )
    );
  }
  if (result.isErr() && attempted.length > 0) {
    try {
      const restored = await rollback();
      if (restored.isErr()) {
        retainJournal = true;
        result = err(migrationError("AgentMigrationRecoveryRequired", restored.error));
      }
    } catch (error) {
      retainJournal = true;
      result = err(migrationError("AgentMigrationRecoveryRequired", error));
    }
  }
  try {
    if (ownedJournal && !retainJournal) await fs.remove(journal);
    if (release) await release();
  } catch (error) {
    result = err(migrationError("AgentMigrationRecoveryRequired", error));
  }
  return result;
}
