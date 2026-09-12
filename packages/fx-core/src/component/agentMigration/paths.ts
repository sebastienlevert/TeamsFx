// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { migrationError } from "./errors";

export const packageLimits = Object.freeze({
  archiveBytes: 10 * 1024 * 1024,
  expandedBytes: 64 * 1024 * 1024,
  fileBytes: 10 * 1024 * 1024,
  entries: 2048,
  pathLength: 240,
  depth: 16,
  changesBytes: 1024 * 1024,
});

export function safeName(input: string, directory = false): Result<string, FxError> {
  const name = input.replace(/\\/g, "/").replace(directory ? /\/$/ : /$^/, "");
  if (name.length > packageLimits.pathLength || name.split("/").length > packageLimits.depth) {
    return err(migrationError("AgentPackageLimitExceeded"));
  }
  if (
    !name ||
    name
      .split("/")
      .some(
        (part) =>
          !part ||
          part === "." ||
          part === ".." ||
          /[<>:"|?*\u0000-\u001f\u007f]/.test(part) ||
          /[. ]$/.test(part) ||
          /^(con|prn|aux|nul|com[0-9\u00b9\u00b2\u00b3]|lpt[0-9\u00b9\u00b2\u00b3])(?:\.|$)/i.test(
            part
          )
      )
  ) {
    return err(migrationError("AgentPackagePathInvalid"));
  }
  return ok(name);
}

export function contained(root: string, file: string): boolean {
  const relative = path.relative(root, file);
  return (
    relative === "" ||
    (!path.isAbsolute(relative) && relative !== ".." && !relative.startsWith(`..${path.sep}`))
  );
}

export function localReference(owner: string, reference: string): Result<string, FxError> {
  const portable = reference.replace(/\\/g, "/");
  if (/^(?:[a-zA-Z][a-zA-Z0-9+.-]*:|\/)/.test(portable) || /[%?#]/.test(portable)) {
    return err(migrationError("AgentPackagePathInvalid"));
  }
  return safeName(path.posix.normalize(path.posix.join(path.posix.dirname(owner), portable)));
}

/** Case-folded names include implicit directories, so A/file and a/other collide too. */
export class NameIndex {
  private readonly names = new Map<
    string,
    { name: string; directory: boolean; explicit: boolean }
  >();

  add(input: string, directory = false): Result<string, FxError> {
    const safe = safeName(input, directory);
    if (safe.isErr()) return err(safe.error);
    const parts = safe.value.split("/");
    for (let index = 0; index < parts.length; index++) {
      const name = parts.slice(0, index + 1).join("/");
      const key = name.normalize("NFC").toLowerCase();
      const last = index === parts.length - 1;
      const isDirectory = !last || directory;
      const previous = this.names.get(key);
      if (
        previous &&
        (previous.name !== name ||
          !previous.directory ||
          !isDirectory ||
          (last && previous.explicit))
      ) {
        return err(migrationError("AgentPackageCollision"));
      }
      this.names.set(key, {
        name,
        directory: isDirectory,
        explicit: last || previous?.explicit === true,
      });
    }
    return ok(safe.value);
  }
}
