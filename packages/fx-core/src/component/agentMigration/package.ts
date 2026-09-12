// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { err, FxError, ManifestType, ok, resolveManifest, Result } from "@microsoft/teamsfx-api";
import { isErrno, migrationError } from "./errors";
import { inspectGraph } from "./graph";
import { readDirectory } from "./intake";
import { readFile } from "./io";
import { parseJson } from "./model";

/**
 * Retain the validated local closure of an imported package. Legacy packaging
 * already resolves primary manifests, but does not enumerate local workers or
 * transitive OpenAPI/card/tool files. Ordinary projects keep their existing path.
 */
export async function appendImportedPackageFiles(
  zip: AdmZip,
  project: string,
  packageDirectory: string,
  envs: Record<string, string>
): Promise<Result<undefined, FxError>> {
  const provenance = path.join(project, ".atk", "import.json");
  try {
    await fs.lstat(provenance);
  } catch (error) {
    return isErrno(error, "ENOENT")
      ? ok(undefined)
      : err(migrationError("AgentMigrationIoError", error));
  }
  const record = await readFile(provenance);
  if (record.isErr()) return err(record.error);
  const metadata = parseJson(record.value);
  if (metadata.isErr() || metadata.value.ruleVersion !== "agent-package/1")
    return err(migrationError("AgentPackageUnsupported"));
  if (path.resolve(packageDirectory) !== path.join(path.resolve(project), "appPackage")) {
    return err(migrationError("AgentPackageUnsupported"));
  }
  const snapshot = await readDirectory(project, undefined, true);
  if (snapshot.isErr()) return err(snapshot.error);
  const files = new Map(
    [...snapshot.value]
      .filter(([name]) => name.startsWith("appPackage/"))
      .map(([name, bytes]) => [name.slice(11), bytes])
  );
  const graph = await inspectGraph(files);
  if (graph.isErr()) return err(graph.error);
  try {
    for (const name of graph.value.files) {
      const document = graph.value.documents.get(name);
      if (zip.getEntry(name) && !document) continue;
      let bytes = files.get(name)!;
      if (document) {
        const manifestType =
          document.kind === "agent"
            ? ManifestType.DeclarativeCopilotManifest
            : document.kind === "plugin"
              ? ManifestType.PluginManifest
              : ManifestType.TeamsManifest;
        const resolved = await resolveManifest(bytes.toString("utf8"), {
          fromPath: path.join(packageDirectory, ...name.split("/")),
          envs,
          manifestType,
        });
        bytes = Buffer.from(resolved.content);
      }
      zip.addFile(name, bytes);
    }
    return ok(undefined);
  } catch (error) {
    return err(migrationError("AgentPackageSchemaInvalid", error));
  }
}
