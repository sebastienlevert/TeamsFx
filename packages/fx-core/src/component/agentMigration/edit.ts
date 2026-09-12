// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import {
  AgentEditRequest,
  AgentMigrationReport,
  err,
  FxError,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../common/localizeUtils";
import { readEdits, planEdits } from "./editPlan";
import { cancelled, migrationError } from "./errors";
import { inspectGraph } from "./graph";
import { readDirectory } from "./intake";
import { artifactDigest } from "./model";
import { importedTemplate, makeReport, projectTrackingId } from "./report";
import { checkRecovery, commitEdits } from "./transaction";

export async function applyAgentEdits(
  request: AgentEditRequest,
  signal?: AbortSignal
): Promise<Result<AgentMigrationReport, FxError>> {
  const cancel = cancelled(signal);
  if (cancel) return err(cancel);
  if (
    typeof request?.projectPath !== "string" ||
    !request.projectPath ||
    typeof request.changesFile !== "string" ||
    !request.changesFile ||
    (request.expectedDigest !== undefined &&
      (typeof request.expectedDigest !== "string" ||
        !/^sha256:[a-f0-9]{64}$/.test(request.expectedDigest))) ||
    (request.dryRun !== undefined && typeof request.dryRun !== "boolean")
  )
    return err(migrationError("AgentEditsInvalid"));
  const project = path.resolve(request.projectPath);
  try {
    const recovery = await checkRecovery(project);
    if (recovery.isErr()) return err(recovery.error);
    const before = await readDirectory(project, signal, true);
    if (before.isErr()) return err(before.error);
    if (request.expectedDigest && request.expectedDigest !== artifactDigest(before.value))
      return err(migrationError("AgentEditsStale"));
    const changesFile = path.resolve(request.changesFile);
    const changes = await readEdits(changesFile);
    if (changes.isErr()) return err(changes.error);
    const packageFiles = new Map(
      [...before.value]
        .filter(([name]) => name.startsWith("appPackage/"))
        .map(([name, data]) => [name.slice(11), data])
    );
    const graph = await inspectGraph(packageFiles, signal);
    if (graph.isErr()) return err(graph.error);
    const planned = await planEdits(changes.value, changesFile, graph.value, packageFiles, signal);
    if (planned.isErr()) return err(planned.error);
    const finalGraph = await inspectGraph(planned.value, signal);
    if (finalGraph.isErr()) return err(finalGraph.error);
    const after = new Map([
      ...before.value,
      ...[...planned.value].map(([name, data]): [string, Buffer] => [`appPackage/${name}`, data]),
    ]);
    const appId = graph.value.manifest.id;
    if (typeof appId !== "string") return err(migrationError("AgentPackageSchemaInvalid"));
    const template = importedTemplate(before.value);
    if (template.isErr()) return err(template.error);
    const result = makeReport(
      project,
      finalGraph.value,
      before.value,
      after,
      request.dryRun === true,
      false,
      "project",
      {
        policy: "preserved",
        provisionedByOperation: false,
        agentId: graph.value.agentId,
        projectId: projectTrackingId(before.value),
        appId,
      },
      template.value,
      changes.value.operations.map((operation) => ({
        ruleId: `edit/${operation.kind}/1`,
        path: `appPackage/${graph.value.agentPath}`,
        description: getLocalizedString("agentMigration.rule.edit"),
      }))
    );
    if (result.changed && !request.dryRun) {
      const committed = await commitEdits(project, before.value, after, signal);
      if (committed.isErr()) return err(committed.error);
    }
    const finalCancel = cancelled(signal);
    if (finalCancel && (request.dryRun || !result.changed)) return err(finalCancel);
    return ok(result);
  } catch (error) {
    return err(migrationError("AgentMigrationIoError", error));
  }
}
