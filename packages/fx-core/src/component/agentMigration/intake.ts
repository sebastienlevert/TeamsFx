// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { inflateRawSync } from "zlib";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { Artifacts, crc32, text } from "./model";
import { cancelled, isErrno, migrationError } from "./errors";
import { NameIndex, packageLimits } from "./paths";
import { readFile, safeAncestors } from "./io";

export function readArchive(bytes: Buffer, signal?: AbortSignal): Result<Artifacts, FxError> {
  if (bytes.length > packageLimits.archiveBytes)
    return err(migrationError("AgentPackageLimitExceeded"));
  const files: Artifacts = new Map();
  const names = new NameIndex();
  let expanded = 0;
  try {
    const zip = new AdmZip(bytes);
    if (zip.getEntryCount() > packageLimits.entries)
      return err(migrationError("AgentPackageLimitExceeded"));
    const entries = zip.getEntries();
    for (const entry of entries) {
      const cancel = cancelled(signal);
      if (cancel) return err(cancel);
      const decoded = text(entry.rawEntryName);
      if (decoded.isErr()) return err(decoded.error);
      const name = names.add(decoded.value, entry.isDirectory);
      if (name.isErr()) return err(name.error);
      const mode = (entry.attr >>> 16) & 0xf000;
      if (
        (mode !== 0 && mode !== 0x8000 && mode !== 0x4000) ||
        (mode === 0x4000 && !entry.isDirectory)
      )
        return err(migrationError("AgentPackagePathInvalid"));
      if (
        (entry.header.flags & 0x41) !== 0 ||
        ![0, 8].includes(entry.header.method) ||
        entry.header.diskNumStart !== 0
      )
        return err(migrationError("AgentPackageUnsupported"));
      if (
        entry.header.size > packageLimits.fileBytes ||
        expanded + entry.header.size > packageLimits.expandedBytes
      )
        return err(migrationError("AgentPackageLimitExceeded"));
      const compressed = entry.getCompressedData();
      if (compressed.length !== entry.header.compressedSize)
        return err(migrationError("AgentPackageIntegrityInvalid"));
      // Do not trust a zero/forged declared size to bound the ZIP dependency's inflater.
      const data =
        entry.header.method === 0
          ? compressed
          : inflateRawSync(compressed, {
              maxOutputLength:
                Math.min(packageLimits.fileBytes, packageLimits.expandedBytes - expanded) + 1,
            });
      if (data.length > packageLimits.fileBytes)
        return err(migrationError("AgentPackageLimitExceeded"));
      if (
        data.length !== entry.header.size ||
        crc32(data) !== entry.header.crc ||
        (entry.isDirectory && data.length !== 0)
      )
        return err(migrationError("AgentPackageIntegrityInvalid"));
      expanded += data.length;
      if (!entry.isDirectory) files.set(name.value, data);
    }
    return ok(files);
  } catch (error) {
    return err(
      migrationError(
        isErrno(error, "ERR_BUFFER_TOO_LARGE")
          ? "AgentPackageLimitExceeded"
          : "AgentPackageIntegrityInvalid",
        error
      )
    );
  }
}

export async function readDirectory(
  root: string,
  signal?: AbortSignal,
  project = false
): Promise<Result<Artifacts, FxError>> {
  const safe = await safeAncestors(root);
  if (safe.isErr()) return err(safe.error);
  const names = new NameIndex();
  const files: Artifacts = new Map();
  let count = 0;
  let bytes = 0;
  async function walk(relative: string): Promise<Result<undefined, FxError>> {
    const entries = await fs.readdir(path.join(root, ...relative.split("/")), {
      withFileTypes: true,
    });
    for (const entry of entries.sort((a, b) => (a.name < b.name ? -1 : 1))) {
      const cancel = cancelled(signal);
      if (cancel) return err(cancel);
      const name = relative ? `${relative}/${entry.name}` : entry.name;
      if (project && (name === ".git" || name === "node_modules" || name === "appPackage/build"))
        continue;
      if (++count > packageLimits.entries) return err(migrationError("AgentPackageLimitExceeded"));
      if (entry.isSymbolicLink() || (!entry.isFile() && !entry.isDirectory()))
        return err(migrationError("AgentPackagePathInvalid"));
      const registered = names.add(name, entry.isDirectory());
      if (registered.isErr()) return err(registered.error);
      if (entry.isDirectory()) {
        const walked = await walk(name);
        if (walked.isErr()) return err(walked.error);
      } else {
        const data = await readFile(path.join(root, ...name.split("/")));
        if (data.isErr()) return err(data.error);
        bytes += data.value.length;
        if (bytes > packageLimits.expandedBytes)
          return err(migrationError("AgentPackageLimitExceeded"));
        files.set(name, data.value);
      }
    }
    return ok(undefined);
  }
  try {
    const walked = await walk("");
    return walked.isErr() ? err(walked.error) : ok(files);
  } catch (error) {
    return err(migrationError("AgentMigrationIoError", error));
  }
}

export async function intake(
  source: string,
  signal?: AbortSignal
): Promise<Result<{ files: Artifacts; kind: "zip" | "directory" }, FxError>> {
  const cancel = cancelled(signal);
  if (cancel) return err(cancel);
  try {
    const safe = await safeAncestors(source);
    if (safe.isErr()) return err(safe.error);
    const stat = await fs.lstat(source);
    if (stat.isDirectory()) {
      const files = await readDirectory(source, signal);
      return files.isErr() ? err(files.error) : ok({ files: files.value, kind: "directory" });
    }
    if (!stat.isFile() || path.extname(source).toLowerCase() !== ".zip")
      return err(migrationError("AgentPackageSourceInvalid"));
    const bytes = await readFile(source, packageLimits.archiveBytes);
    if (bytes.isErr()) return err(bytes.error);
    const files = readArchive(bytes.value, signal);
    return files.isErr() ? err(files.error) : ok({ files: files.value, kind: "zip" });
  } catch (error) {
    return err(migrationError("AgentPackageSourceInvalid", error));
  }
}
