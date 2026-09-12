// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import * as lockfile from "proper-lockfile";
import YAML from "yaml";
import {
  AgentImportRequest,
  AgentMigrationReport,
  Context,
  err,
  FxError,
  ManifestType,
  ok,
  Platform,
  resolveManifest,
  Result,
} from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../common/localizeUtils";
import { cancelled, isErrno, migrationError } from "./errors";
import { AgentPackageGenerator } from "./generator";
import { inspectGraph, packageRoot } from "./graph";
import { instructionFile } from "./instructions";
import { intake, readDirectory } from "./intake";
import { agentMigrationIo, prepareParent, requireAbsent, writeArtifacts } from "./io";
import { artifactDigest, Artifacts, digest, jsonBytes, object, parseJson } from "./model";
import { contained, safeName } from "./paths";
import { makeReport } from "./report";

export async function importAgentPackage(
  request: AgentImportRequest,
  context: Context,
  signal?: AbortSignal
): Promise<Result<AgentMigrationReport, FxError>> {
  try {
    return await runImport(request, context, signal);
  } catch (error) {
    return err(migrationError("AgentMigrationIoError", error));
  }
}

async function runImport(
  request: AgentImportRequest,
  context: Context,
  signal?: AbortSignal
): Promise<Result<AgentMigrationReport, FxError>> {
  const cancel = cancelled(signal);
  if (cancel) return err(cancel);
  if (
    typeof request?.sourcePath !== "string" ||
    !request.sourcePath ||
    (request.outputPath !== undefined &&
      (typeof request.outputPath !== "string" || !request.outputPath)) ||
    (request.dryRun !== undefined && typeof request.dryRun !== "boolean")
  )
    return err(migrationError("AgentPackageSourceInvalid"));
  const source = path.resolve(request.sourcePath);
  const input = await intake(source, signal);
  if (input.isErr()) return err(input.error);
  const root = packageRoot(input.value.files);
  if (root.isErr()) return err(root.error);
  const graph = await inspectGraph(root.value, signal, true);
  if (graph.isErr()) return err(graph.error);
  const sourceBase =
    path.basename(source) === "appPackage"
      ? path.basename(path.dirname(source))
      : path.basename(source, input.value.kind === "zip" ? path.extname(source) : "");
  const defaultName = safeName(`${sourceBase}-imported`);
  if (!request.outputPath && defaultName.isErr()) return err(defaultName.error);
  const destination = path.resolve(request.outputPath ?? `${sourceBase}-imported`);
  const destinationName = safeName(path.basename(destination));
  if (destinationName.isErr()) return err(destinationName.error);
  if (contained(source, destination) || contained(destination, source))
    return err(migrationError("AgentPackagePathInvalid"));
  const absent = await requireAbsent(destination);
  if (absent.isErr()) return err(absent.error);
  const parent = path.dirname(destination);
  const appName = object(graph.value.manifest.name)?.short;
  if (typeof appName !== "string") return err(migrationError("AgentPackageSchemaInvalid"));
  let stage: string | undefined;
  let release: (() => Promise<void>) | undefined;
  let createdParents: string[] = [];
  let promoted = false;
  let result: Result<AgentMigrationReport, FxError>;
  try {
    const parents = await prepareParent(parent);
    if (parents.isErr()) return err(parents.error);
    createdParents = parents.value;
    release = await lockfile.lock(destination, {
      realpath: false,
      retries: 0,
      lockfilePath: path.join(parent, `.${path.basename(destination)}.agent-import.lock`),
    });
    stage = await fs.mkdtemp(path.join(parent, `.${path.basename(destination)}.agent-import-`));
    result = await stageImport(
      stage,
      destination,
      request,
      input.value.kind,
      root.value,
      input.value.files,
      graph.value,
      context,
      signal
    );
    if (result.isOk() && !request.dryRun) {
      await agentMigrationIo.beforeCommit();
      const cancel = cancelled(signal);
      const stillAbsent = await requireAbsent(destination);
      if (cancel) result = err(cancel);
      else if (stillAbsent.isErr()) result = err(stillAbsent.error);
      else {
        await agentMigrationIo.promote(stage, destination);
        stage = undefined;
        promoted = true;
      }
    }
  } catch (error) {
    result = err(
      migrationError(
        isErrno(error, "ELOCKED") ? "AgentEditsConflict" : "AgentMigrationIoError",
        error
      )
    );
  }
  try {
    if (stage) await fs.remove(stage);
    if (release) await release();
    if (!promoted) for (const directory of [...createdParents].reverse()) await fs.rmdir(directory);
  } catch (error) {
    return err(migrationError("AgentMigrationRecoveryRequired", error));
  }
  return result;
}

async function stageImport(
  stage: string,
  destination: string,
  request: AgentImportRequest,
  kind: "zip" | "directory",
  source: Artifacts,
  allSourceFiles: Artifacts,
  graph: import("./graph").PackageGraph,
  context: Context,
  signal?: AbortSignal
): Promise<Result<AgentMigrationReport, FxError>> {
  const appName = object(graph.manifest.name)?.short;
  if (typeof appName !== "string") return err(migrationError("AgentPackageSchemaInvalid"));
  const generator = new AgentPackageGenerator(appName, signal);
  const generated = await generator.run(context, { platform: Platform.CLI }, stage);
  if (generated.isErr()) return err(generated.error);
  if (!generator.projectId || !generator.template)
    return err(migrationError("AgentMigrationIoError"));
  const packageFiles = new Map([...source].filter(([name]) => graph.files.has(name)));
  const sourceAppId = graph.manifest.id;
  const primary = graph.documents.get(graph.agentPath)!.value;
  const sourceAgentId = primary.id;
  const sourceAgents = Object.fromEntries(
    [...graph.documents]
      .filter(([, document]) => document.kind === "agent" && typeof document.value.id === "string")
      .map(([name, document]) => [name, document.value.id])
  );
  graph.manifest.id = "${{TEAMS_APP_ID}}";
  const transformations: AgentMigrationReport["transformations"] = [
    {
      ruleId: "new-identity/1",
      path: "appPackage/manifest.json",
      description: getLocalizedString("agentMigration.rule.identity"),
    },
  ];
  for (const name of graph.normalized) {
    const document = graph.documents.get(name)?.value ?? graph.dataDocuments.get(name);
    if (!document) return err(migrationError("AgentPackageSchemaInvalid"));
    packageFiles.set(
      name,
      /\.ya?ml$/i.test(name) ? Buffer.from(YAML.stringify(document)) : jsonBytes(document)
    );
    transformations.push({
      ruleId: "portable-references/1",
      path: `appPackage/${name}`,
      description: getLocalizedString("agentMigration.rule.paths"),
    });
  }
  for (const [file, document] of graph.documents) {
    if (document.kind !== "agent") continue;
    delete document.value.id;
    const instructions = graph.instructions.get(file);
    if (instructions) {
      // Reserve every source name, including skipped candidates, before choosing an instruction file.
      const reserved = new Map([...source, ...packageFiles]);
      const target = instructionFile(
        reserved,
        file,
        document.value,
        instructions.content,
        instructions.sourcePath
      );
      if (target.isErr()) return err(target.error);
      packageFiles.set(target.value, reserved.get(target.value)!);
      transformations.push({
        ruleId: "literal-instructions/1",
        path: `appPackage/${target.value}`,
        description: getLocalizedString("agentMigration.rule.instructions"),
      });
    }
    packageFiles.set(file, jsonBytes(document.value));
  }
  packageFiles.set("manifest.json", jsonBytes(graph.manifest));
  const identity: AgentMigrationReport["identity"] = {
    policy: "new",
    provisionedByOperation: false,
    agentId: graph.agentId,
    projectId: generator.projectId,
    appId: "${{TEAMS_APP_ID}}",
    ...(typeof sourceAppId === "string" ? { sourceAppId } : {}),
    ...(typeof sourceAgentId === "string" ? { sourceAgentId } : {}),
  };
  const output = new Map([...packageFiles].map(([name, data]) => [`appPackage/${name}`, data]));
  output.set(
    ".atk/import.json",
    jsonBytes({
      reportVersion: 1,
      ruleVersion: "agent-package/1",
      sourceDigest: artifactDigest(allSourceFiles),
      identity,
      sourceAgents,
      template: generator.template,
    })
  );
  const written = await writeArtifacts(stage, output, signal);
  if (written.isErr()) return err(written.error);
  const finalGraph = await inspectGraph(packageFiles, signal);
  if (finalGraph.isErr()) return err(finalGraph.error);
  for (const [file, original] of graph.instructions) {
    const resolved = await resolveManifest(packageFiles.get(file)!.toString("utf8"), {
      fromPath: path.join(stage, "appPackage", ...file.split("/")),
      envs: {},
      manifestType: ManifestType.DeclarativeCopilotManifest,
    });
    const parsed = parseJson(Buffer.from(resolved.content));
    if (parsed.isErr()) return err(parsed.error);
    if (parsed.value.instructions !== original.content)
      return err(migrationError("AgentPackageSchemaInvalid"));
  }
  const snapshot = await readDirectory(stage, signal, true);
  if (snapshot.isErr()) return err(snapshot.error);
  const report = makeReport(
    destination,
    finalGraph.value,
    source,
    snapshot.value,
    request.dryRun === true,
    true,
    kind,
    identity,
    generator.template,
    transformations
  );
  report.source.digest = artifactDigest(allSourceFiles);
  const wrapped = !allSourceFiles.has("manifest.json");
  if (wrapped) {
    for (const file of report.files)
      if (file.sourcePath) file.sourcePath = `appPackage/${file.sourcePath}`;
    report.referenceMap = Object.fromEntries(
      Object.entries(report.referenceMap).map(([source, destination]) => [
        `appPackage/${source}`,
        destination,
      ])
    );
    for (const [name, data] of allSourceFiles) {
      if (!name.startsWith("appPackage/"))
        report.files.push({
          path: name,
          sourcePath: name,
          action: "skipped",
          digest: digest(data),
        });
    }
  }
  return ok(report);
}
