// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { Artifacts, JsonObject } from "./model";
import { NameIndex } from "./paths";
import { migrationError } from "./errors";

export function instructionFile(
  files: Artifacts,
  agentPath: string,
  document: JsonObject,
  content: string,
  preferred?: string
): Result<string, FxError> {
  const index = new NameIndex();
  for (const name of files.keys()) {
    const added = index.add(name);
    if (added.isErr()) return err(added.error);
  }
  let file = preferred;
  if (
    file &&
    path.posix.dirname(file) !== path.posix.dirname(agentPath) &&
    !file.startsWith(`${path.posix.dirname(agentPath)}/`)
  )
    return err(migrationError("AgentPackagePathInvalid"));
  if (!file) {
    const base = path.posix.join(path.posix.dirname(agentPath), "instructions");
    for (let suffix = 0; suffix <= files.size; suffix++) {
      const candidate = `${base}${suffix ? `-${suffix}` : ""}.txt`;
      if (index.add(candidate).isOk()) {
        file = candidate;
        break;
      }
    }
  }
  if (!file) return err(migrationError("AgentPackageCollision"));
  const relative = path.posix.relative(path.posix.dirname(agentPath), file);
  if (relative.includes("'")) return err(migrationError("AgentPackagePathInvalid"));
  document.instructions = `$[file('${relative}', 'raw')]`;
  files.set(file, Buffer.from(content));
  return ok(file);
}
