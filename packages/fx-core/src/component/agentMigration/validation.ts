// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ApiPluginManifestConverter,
  AppManifestUtils,
  DeclarativeAgentManifestConverter,
  err,
  FxError,
  ok,
  Result,
  TeamsManifestConverter,
} from "@microsoft/teamsfx-api";
import { crc32, JsonObject, jsonBytes, parseJson } from "./model";
import { inflateSync } from "zlib";
import { migrationError } from "./errors";

export type DocumentKind = "teams" | "agent" | "plugin";
export const fixtureAppId = "00000000-0000-4000-8000-000000000001";
export const knowledgeExtensions = new Set([
  ".txt",
  ".md",
  ".pdf",
  ".docx",
  ".pptx",
  ".xlsx",
  ".csv",
]);

export function schemaUrl(kind: DocumentKind, version: string): string {
  const base = "https://developer.microsoft.com/json-schemas";
  return kind === "teams"
    ? `${base}/teams/v${version}/MicrosoftTeams.schema.json`
    : `${base}/copilot/${kind === "agent" ? "declarative-agent" : "plugin"}/${version}/schema.json`;
}

export async function validateDocument(
  input: JsonObject,
  kind: DocumentKind
): Promise<Result<undefined, FxError>> {
  const version =
    input[kind === "teams" ? "manifestVersion" : kind === "agent" ? "version" : "schema_version"];
  if (typeof version !== "string" || !/^(?:v)?\d+\.\d+$/.test(version)) {
    return err(migrationError("AgentPackageUnsupported"));
  }
  const url = schemaUrl(kind, version);
  if (
    input.$schema !== undefined &&
    (typeof input.$schema !== "string" ||
      input.$schema.replace(/\/[a-z]{2}-[a-z]{2}\/json-schemas\//i, "/json-schemas/") !== url)
  ) {
    return err(migrationError("AgentPackageSchemaInvalid"));
  }
  let schema;
  try {
    schema = await AppManifestUtils.fetchSchema(url, { localOnly: true });
  } catch (error) {
    return err(migrationError("AgentPackageUnsupported", error));
  }
  try {
    const clone = parseJson(jsonBytes(input));
    if (clone.isErr()) return err(clone.error);
    if (kind === "teams" && clone.value.id === "${{TEAMS_APP_ID}}") clone.value.id = fixtureAppId;
    const content = JSON.stringify(clone.value);
    const manifest =
      kind === "teams"
        ? TeamsManifestConverter.jsonToManifest(content)
        : kind === "agent"
          ? DeclarativeAgentManifestConverter.jsonToManifest(content)
          : ApiPluginManifestConverter.jsonToManifest(content);
    const issues = await AppManifestUtils.validateAgainstSchema(manifest, schema);
    if (issues.length > 0) {
      const error = migrationError("AgentPackageSchemaInvalid");
      error.userData = JSON.stringify({ issues });
      return err(error);
    }
    return ok(undefined);
  } catch (error) {
    return err(migrationError("AgentPackageSchemaInvalid", error));
  }
}

export function validatePng(bytes: Buffer, size: number): Result<undefined, FxError> {
  if (
    bytes.length < 45 ||
    !bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])) ||
    bytes.toString("ascii", 12, 16) !== "IHDR" ||
    bytes.readUInt32BE(16) !== size ||
    bytes.readUInt32BE(20) !== size ||
    bytes.toString("ascii", bytes.length - 8, bytes.length - 4) !== "IEND"
  ) {
    return err(migrationError("AgentPackageSchemaInvalid"));
  }
  const compressed: Buffer[] = [];
  let offset = 8;
  while (offset + 12 <= bytes.length) {
    const length = bytes.readUInt32BE(offset);
    if (length > bytes.length - offset - 12)
      return err(migrationError("AgentPackageSchemaInvalid"));
    const type = bytes.toString("ascii", offset + 4, offset + 8);
    const payload = bytes.subarray(offset + 8, offset + 8 + length);
    if (
      crc32(bytes.subarray(offset + 4, offset + 8 + length)) !==
      bytes.readUInt32BE(offset + 8 + length)
    ) {
      return err(migrationError("AgentPackageSchemaInvalid"));
    }
    if (type === "IHDR" && (offset !== 8 || length !== 13))
      return err(migrationError("AgentPackageSchemaInvalid"));
    if (type === "IDAT") compressed.push(payload);
    offset += length + 12;
    if (type === "IEND" && (length !== 0 || offset !== bytes.length))
      return err(migrationError("AgentPackageSchemaInvalid"));
  }
  if (offset !== bytes.length || compressed.length === 0)
    return err(migrationError("AgentPackageSchemaInvalid"));
  try {
    inflateSync(Buffer.concat(compressed), { maxOutputLength: 2 * 1024 * 1024 });
  } catch (error) {
    return err(migrationError("AgentPackageSchemaInvalid", error));
  }
  return ok(undefined);
}
