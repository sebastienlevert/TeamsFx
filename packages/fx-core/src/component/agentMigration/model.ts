// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { createHash } from "crypto";
import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { migrationError } from "./errors";

export type JsonValue = string | number | boolean | null | JsonValue[] | JsonObject;
export interface JsonObject {
  [key: string]: JsonValue;
}
export type Artifacts = Map<string, Buffer>;

export function isObject(value: unknown): value is JsonObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function object(value: JsonValue | undefined): JsonObject | undefined {
  return isObject(value) ? value : undefined;
}

export function objects(value: JsonValue | undefined): JsonObject[] {
  return Array.isArray(value) ? value.filter(isObject) : [];
}

export function text(data: Buffer): Result<string, FxError> {
  try {
    return ok(new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(data));
  } catch (error) {
    return err(migrationError("AgentPackageSourceInvalid", error));
  }
}

export function parseJson(data: Buffer): Result<JsonObject, FxError> {
  const decoded = text(data);
  if (decoded.isErr()) return err(decoded.error);
  try {
    const value: unknown = JSON.parse(decoded.value.replace(/^\uFEFF/, ""));
    if (!isObject(value)) return err(migrationError("AgentPackageSchemaInvalid"));
    return ok(value);
  } catch (error) {
    return err(migrationError("AgentPackageSchemaInvalid", error));
  }
}

export function jsonBytes(value: unknown): Buffer {
  return Buffer.from(`${JSON.stringify(value, undefined, 2)}\n`);
}

export function digest(data: Buffer): string {
  return `sha256:${createHash("sha256").update(data).digest("hex")}`;
}

export function decodeProfileArchive(
  data: Buffer,
  expectedSha256: string
): Result<Buffer, FxError> {
  const decoded = text(data);
  if (decoded.isErr()) return err(migrationError("AgentPackageIntegrityInvalid", decoded.error));
  const encoded = decoded.value.replace(/\r?\n$/, "");
  if (
    !encoded ||
    !/^[a-f0-9]{64}$/.test(expectedSha256) ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)
  ) {
    return err(migrationError("AgentPackageIntegrityInvalid"));
  }
  const archive = Buffer.from(encoded, "base64");
  if (archive.toString("base64") !== encoded || digest(archive) !== `sha256:${expectedSha256}`) {
    return err(migrationError("AgentPackageIntegrityInvalid"));
  }
  return ok(archive);
}

export function artifactDigest(files: Artifacts): string {
  const hash = createHash("sha256");
  for (const name of [...files.keys()].sort()) {
    const data = files.get(name)!;
    hash.update(`${name}\0${data.length}\0`).update(data);
  }
  return `sha256:${hash.digest("hex")}`;
}

export function crc32(data: Buffer): number {
  const table = Array.from({ length: 256 }, (_, index) => {
    let crc = index;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    return crc >>> 0;
  });
  let crc = 0xffffffff;
  for (const byte of data) crc = (crc >>> 8) ^ table[(crc ^ byte) & 0xff];
  return (crc ^ 0xffffffff) >>> 0;
}
