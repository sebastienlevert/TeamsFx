// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Ajv from "ajv";
import path from "path";
import {
  AgentEditDocument,
  agentEditDocumentSchema,
  AgentEditOperation,
  err,
  FxError,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import { cancelled, migrationError } from "./errors";
import { PackageGraph } from "./graph";
import { instructionFile } from "./instructions";
import { readFile } from "./io";
import { Artifacts, jsonBytes, JsonObject, object, objects, parseJson, text } from "./model";
import { localReference, NameIndex, packageLimits, safeName } from "./paths";
import { schemaUrl, validatePng } from "./validation";

export async function readEdits(file: string): Promise<Result<AgentEditDocument, FxError>> {
  const bytes = await readFile(file, packageLimits.changesBytes);
  if (bytes.isErr()) return err(bytes.error);
  const parsed = parseJson(bytes.value);
  if (parsed.isErr()) return err(migrationError("AgentEditsInvalid", parsed.error));
  const validate = new Ajv().compile<AgentEditDocument>(agentEditDocumentSchema);
  return validate(parsed.value) ? ok(parsed.value) : err(migrationError("AgentEditsInvalid"));
}

export async function planEdits(
  changes: AgentEditDocument,
  changesFile: string,
  graph: PackageGraph,
  files: Artifacts,
  signal?: AbortSignal
): Promise<Result<Artifacts, FxError>> {
  if (changes.agentId !== graph.agentId) return err(migrationError("AgentEditsAgentNotFound"));
  const result = new Map(files);
  const agent = graph.documents.get(graph.agentPath)!.value;

  async function asset(sourceFile: string): Promise<Result<Buffer, FxError>> {
    const safe = safeName(sourceFile);
    if (safe.isErr()) return err(safe.error);
    return readFile(path.join(path.dirname(changesFile), ...safe.value.split("/")));
  }

  async function apply(operation: AgentEditOperation): Promise<Result<undefined, FxError>> {
    const encoded = parseJson(jsonBytes(operation));
    if (encoded.isErr()) return err(encoded.error);
    const value = encoded.value.value;
    switch (operation.kind) {
      case "setAppMetadata":
      case "setAgentMetadata": {
        const destination = operation.kind === "setAppMetadata" ? graph.manifest : agent;
        const metadata = object(value);
        if (!metadata) return err(migrationError("AgentEditsInvalid"));
        for (const [key, value] of Object.entries(metadata)) {
          const previousMembers = object(destination[key]);
          const suppliedMembers = object(value);
          destination[key] =
            previousMembers && suppliedMembers ? { ...previousMembers, ...suppliedMembers } : value;
        }
        return ok(undefined);
      }
      case "replaceConversationStarters":
        if (!Array.isArray(value)) return err(migrationError("AgentEditsInvalid"));
        agent.conversation_starters = value;
        return ok(undefined);
      case "setBehaviorOverrides":
        if (!object(value)) return err(migrationError("AgentEditsInvalid"));
        agent.behavior_overrides = value!;
        return ok(undefined);
      case "setSchemaVersion":
        agent.version = operation.value;
        agent.$schema = schemaUrl("agent", operation.value);
        return ok(undefined);
      case "upsertCapability": {
        const capability = object(value);
        if (!capability) return err(migrationError("AgentEditsInvalid"));
        const capabilities = objects(agent.capabilities);
        const index = capabilities.findIndex((item) => item.name === capability.name);
        if (index < 0) capabilities.push(capability);
        else capabilities[index] = capability;
        agent.capabilities = capabilities;
        return ok(undefined);
      }
      case "removeCapability":
        if (Array.isArray(agent.capabilities))
          agent.capabilities = objects(agent.capabilities).filter(
            (item) => item.name !== operation.name
          );
        return ok(undefined);
      case "replaceInstructions": {
        const bytes = await asset(operation.sourceFile);
        if (bytes.isErr()) return err(bytes.error);
        const value = text(bytes.value);
        if (value.isErr()) return err(value.error);
        const currentFile = graph.instructions.get(graph.agentPath)?.sourcePath;
        const shared =
          currentFile &&
          [...graph.instructions].some(
            ([name, reference]) => name !== graph.agentPath && reference.sourcePath === currentFile
          );
        const instructions = instructionFile(
          result,
          graph.agentPath,
          agent,
          value.value,
          shared ? undefined : currentFile
        );
        if (instructions.isErr()) return err(instructions.error);
        graph.instructions.set(graph.agentPath, {
          content: value.value,
          sourcePath: instructions.value,
        });
        return ok(undefined);
      }
      case "replaceIcon": {
        const bytes = await asset(operation.sourceFile);
        if (bytes.isErr()) return err(bytes.error);
        const valid = validatePng(bytes.value, operation.icon === "color" ? 192 : 32);
        if (valid.isErr()) return err(valid.error);
        const file = object(graph.manifest.icons)?.[operation.icon];
        if (typeof file !== "string") return err(migrationError("AgentPackageSchemaInvalid"));
        const reference = localReference("manifest.json", file);
        if (reference.isErr()) return err(reference.error);
        result.set(reference.value, bytes.value);
        return ok(undefined);
      }
      case "attachEmbeddedKnowledge": {
        const capabilities = objects(agent.capabilities);
        const existing = capabilities.find((item) => item.name === "EmbeddedKnowledge");
        if (existing?.embedded_resource_snapshot_id !== undefined)
          return err(migrationError("AgentEditsInvalid"));
        const capability: JsonObject = existing ?? { name: "EmbeddedKnowledge", files: [] };
        const resources = objects(capability.files);
        const index = new NameIndex();
        for (const name of result.keys()) {
          const valid = index.add(name);
          if (valid.isErr()) return err(valid.error);
        }
        for (const file of operation.files) {
          const safe = safeName(file.targetPath);
          if (safe.isErr()) return err(safe.error);
          const destination = `knowledge/${safe.value}`;
          const bytes = await asset(file.sourceFile);
          if (bytes.isErr()) return err(bytes.error);
          const old = result.get(destination);
          if (old && !old.equals(bytes.value)) return err(migrationError("AgentPackageCollision"));
          if (!old) {
            const valid = index.add(destination);
            if (valid.isErr()) return err(valid.error);
            result.set(destination, bytes.value);
          }
          if (!resources.some((resource) => resource.file === destination))
            resources.push({ file: destination });
        }
        capability.files = resources;
        if (!existing) capabilities.push(capability);
        agent.capabilities = capabilities;
        return ok(undefined);
      }
    }
  }

  for (const operation of changes.operations) {
    const cancel = cancelled(signal);
    if (cancel) return err(cancel);
    const applied = await apply(operation);
    if (applied.isErr()) return err(applied.error);
  }
  for (const [name, document] of graph.documents) {
    const before = parseJson(files.get(name)!);
    if (before.isErr()) return err(before.error);
    if (JSON.stringify(before.value) !== JSON.stringify(document.value))
      result.set(name, jsonBytes(document.value));
  }
  return ok(result);
}
