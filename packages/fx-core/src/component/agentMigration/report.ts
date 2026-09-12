// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import YAML from "yaml";
import {
  AgentMigrationFile,
  AgentMigrationIdentity,
  AgentMigrationReport,
  AgentTemplateIdentity,
  err,
  FxError,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../common/localizeUtils";
import { PackageGraph } from "./graph";
import { artifactDigest, Artifacts, digest, object, parseJson } from "./model";
import { knowledgeExtensions } from "./validation";
import { migrationError } from "./errors";

export function projectTrackingId(files: Artifacts): string | null {
  for (const name of ["m365agents.yml", "teamsapp.yml"]) {
    const bytes = files.get(name);
    if (bytes) {
      const doc = YAML.parseDocument(bytes.toString("utf8"));
      const id: unknown = doc.get("projectId");
      return typeof id === "string" ? id : null;
    }
  }
  return null;
}

export function importedTemplate(files: Artifacts): Result<AgentTemplateIdentity | null, FxError> {
  const bytes = files.get(".atk/import.json");
  if (!bytes) return ok(null);
  const data = parseJson(bytes);
  if (data.isErr()) return err(data.error);
  const template = object(data.value.template);
  return template &&
    typeof template.id === "string" &&
    typeof template.version === "string" &&
    typeof template.digest === "string"
    ? ok({ id: template.id, version: template.version, digest: template.digest })
    : err(migrationError("AgentPackageSchemaInvalid"));
}

export function makeReport(
  projectPath: string,
  graph: PackageGraph,
  before: Artifacts,
  after: Artifacts,
  dryRun: boolean,
  importing: boolean,
  sourceKind: "zip" | "directory" | "project",
  identity: AgentMigrationIdentity,
  template: AgentTemplateIdentity | null,
  transformations: AgentMigrationReport["transformations"]
): AgentMigrationReport {
  const changed = importing || artifactDigest(before) !== artifactDigest(after);
  const files: AgentMigrationFile[] = [];
  for (const [name, bytes] of after) {
    const old =
      importing && name.startsWith("appPackage/") ? before.get(name.slice(11)) : before.get(name);
    files.push({
      path: name,
      action: old?.equals(bytes) ? "preserved" : old ? "modified" : "created",
      digest: digest(bytes),
      ...(old && importing ? { sourcePath: name.slice(11) } : {}),
    });
  }
  if (importing) {
    for (const [name, bytes] of before) {
      if (!graph.files.has(name))
        files.push({
          path: name,
          sourcePath: name,
          action: knowledgeExtensions.has(path.posix.extname(name).toLowerCase())
            ? "candidate"
            : "skipped",
          digest: digest(bytes),
        });
    }
  }
  const diagnostics: AgentMigrationReport["diagnostics"] = graph.external.map((reference) => ({
    code: "ExternalConfiguration",
    severity: "info",
    path: reference.path,
    pointer: reference.pointer,
    ruleId: "agent-package/1",
    message: getLocalizedString("agentMigration.diagnostic.external"),
  }));
  for (const file of files.filter((file) => file.action === "candidate"))
    diagnostics.push({
      code: "UnreferencedCandidate",
      severity: "info",
      path: file.path,
      ruleId: "agent-package/1",
      message: getLocalizedString("agentMigration.diagnostic.candidate"),
    });
  if (!importing) {
    for (const name of after.keys()) {
      if (name.startsWith("appPackage/") && !graph.files.has(name.slice(11)))
        diagnostics.push({
          code: "UnreferencedFile",
          severity: "warning",
          path: name,
          ruleId: "agent-package/1",
          message: getLocalizedString("agentMigration.diagnostic.unreferenced"),
        });
    }
  }
  return {
    reportVersion: 1,
    projectPath,
    dryRun,
    changed,
    operationMode: dryRun ? "dry-run" : !changed ? "no-op" : importing ? "imported" : "edited",
    source: { kind: sourceKind, digest: artifactDigest(before) },
    projectDigest: artifactDigest(after),
    identity,
    template,
    transformations,
    files,
    referenceMap: Object.fromEntries(
      [...graph.files]
        .filter((name) => before.has(importing ? name : `appPackage/${name}`))
        .map((name) => [name, `appPackage/${name}`])
    ),
    diagnostics,
    configuration: {
      structurallyValid: true,
      readyToProvision: "not-evaluated",
      readyToPublish: "not-evaluated",
      requirements: [
        ...(importing ? [getLocalizedString("agentMigration.configuration.identity")] : []),
        ...(graph.external.length
          ? [getLocalizedString("agentMigration.configuration.resources")]
          : []),
      ],
    },
    recoveryRequired: false,
  };
}
