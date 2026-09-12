// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import YAML from "yaml";
import { err, FxError, getStaticManifestFileReference, ok, Result } from "@microsoft/teamsfx-api";
import { cancelled, migrationError } from "./errors";
import {
  Artifacts,
  isObject,
  JsonObject,
  JsonValue,
  object,
  objects,
  parseJson,
  text,
} from "./model";
import { localReference } from "./paths";
import { DocumentKind, knowledgeExtensions, validateDocument, validatePng } from "./validation";

export interface PackageGraph {
  manifest: JsonObject;
  agentId: string;
  agentPath: string;
  documents: Map<string, { kind: DocumentKind; value: JsonObject }>;
  instructions: Map<string, { content: string; sourcePath?: string }>;
  files: Set<string>;
  external: Array<{ path: string; pointer: string }>;
  normalized: Set<string>;
  dataDocuments: Map<string, JsonObject>;
}

export function packageRoot(files: Artifacts): Result<Artifacts, FxError> {
  const roots = ["manifest.json", "appPackage/manifest.json"].filter((file) => files.has(file));
  if (roots.length !== 1) return err(migrationError("AgentPackageUnsupported"));
  if (roots[0] === "manifest.json") return ok(files);
  return ok(
    new Map(
      [...files]
        .filter(([name]) => name.startsWith("appPackage/"))
        .map(([name, data]) => [name.slice("appPackage/".length), data])
    )
  );
}

export async function inspectGraph(
  files: Artifacts,
  signal?: AbortSignal,
  normalizePaths = false
): Promise<Result<PackageGraph, FxError>> {
  const bytes = files.get("manifest.json");
  if (!bytes) return err(migrationError("AgentPackageReferenceMissing"));
  const parsed = parseJson(bytes);
  if (parsed.isErr()) return err(parsed.error);
  const manifest = parsed.value;
  const agents = objects(object(manifest.copilotAgents)?.declarativeAgents);
  if (
    agents.length !== 1 ||
    typeof agents[0].id !== "string" ||
    typeof agents[0].file !== "string" ||
    ["bots", "staticTabs", "configurableTabs", "composeExtensions"].some(
      (field) => objects(manifest[field]).length > 0
    ) ||
    objects(object(manifest.copilotAgents)?.customEngineAgents).length > 0
  ) {
    return err(migrationError("AgentPackageUnsupported"));
  }
  const agentPath = localReference("manifest.json", agents[0].file);
  if (agentPath.isErr()) return err(agentPath.error);
  const graph: PackageGraph = {
    manifest,
    agentId: agents[0].id,
    agentPath: agentPath.value,
    documents: new Map(),
    instructions: new Map(),
    files: new Set(),
    external: [],
    normalized: new Set(),
    dataDocuments: new Map(),
  };
  if (normalizePaths && agents[0].file.includes("\\")) {
    agents[0].file = agents[0].file.replace(/\\/g, "/");
    graph.normalized.add("manifest.json");
  }
  const visiting = new Set<string>();
  const dataDocuments = graph.dataDocuments;

  function reference(
    owner: string,
    value: JsonValue | undefined,
    pointer: string,
    binding?: { parent: JsonObject; key: string }
  ): Result<string | undefined, FxError> {
    if (typeof value !== "string") return err(migrationError("AgentPackageSchemaInvalid"));
    if (/^https?:\/\//i.test(value)) {
      graph.external.push({ path: owner, pointer });
      return ok(undefined);
    }
    const resolved = localReference(owner, value);
    if (resolved.isErr()) return err(resolved.error);
    if (!files.has(resolved.value)) return err(migrationError("AgentPackageReferenceMissing"));
    graph.files.add(resolved.value);
    if (normalizePaths && binding && value.includes("\\")) {
      binding.parent[binding.key] = value.replace(/\\/g, "/");
      graph.normalized.add(owner);
    }
    return ok(resolved.value);
  }

  async function dataFile(
    owner: string,
    value: JsonValue | undefined,
    pointer: string,
    openapi = false,
    binding?: { parent: JsonObject; key: string }
  ): Promise<Result<undefined, FxError>> {
    const resolved = reference(owner, value, pointer, binding);
    if (resolved.isErr()) return err(resolved.error);
    if (!resolved.value || dataDocuments.has(resolved.value)) return ok(undefined);
    const content = text(files.get(resolved.value)!);
    if (content.isErr()) return err(content.error);
    let data: unknown;
    try {
      data = /\.ya?ml$/i.test(resolved.value)
        ? YAML.parse(content.value, { maxAliasCount: 100 })
        : JSON.parse(content.value.replace(/^\uFEFF/, ""));
    } catch (error) {
      return err(migrationError("AgentPackageSchemaInvalid", error));
    }
    if (!isObject(data)) return err(migrationError("AgentPackageSchemaInvalid"));
    if (
      openapi &&
      (!(
        (typeof data.openapi === "string" && /^3\./.test(data.openapi)) ||
        data.swagger === "2.0"
      ) ||
        !object(data.info) ||
        !object(data.paths))
    )
      return err(migrationError("AgentPackageUnsupported"));
    dataDocuments.set(resolved.value, data);
    return walkReferences(data, resolved.value);
  }

  async function walkReferences(
    value: JsonValue,
    owner: string
  ): Promise<Result<undefined, FxError>> {
    if (Array.isArray(value)) {
      for (const item of value) {
        const walked = await walkReferences(item, owner);
        if (walked.isErr()) return err(walked.error);
      }
    } else if (isObject(value)) {
      const imageReferences: Array<{ parent: JsonObject; key: string }> = [];
      if (value.type === "Image") imageReferences.push({ parent: value, key: "url" });
      if (typeof value.iconUrl === "string")
        imageReferences.push({ parent: value, key: "iconUrl" });
      if (typeof value.backgroundImage === "string")
        imageReferences.push({ parent: value, key: "backgroundImage" });
      const background = object(value.backgroundImage);
      if (background && typeof background.url === "string")
        imageReferences.push({ parent: background, key: "url" });
      for (const binding of imageReferences) {
        const image = binding.parent[binding.key];
        if (typeof image !== "string" || /^data:image\//i.test(image)) continue;
        const resolved = reference(owner, image, "/image", binding);
        if (resolved.isErr()) return err(resolved.error);
      }
      if (typeof value.$ref === "string") {
        if (normalizePaths && !/^https?:\/\//i.test(value.$ref) && value.$ref.includes("\\")) {
          value.$ref = value.$ref.replace(/\\/g, "/");
          graph.normalized.add(owner);
        }
        const [file, fragment] = value.$ref.split("#");
        if (file) {
          const read = await dataFile(owner, file, "$ref");
          if (read.isErr()) return err(read.error);
        }
        if (fragment && !/^https?:\/\//i.test(file)) {
          const target = file ? localReference(owner, file) : ok(owner);
          if (target.isErr()) return err(target.error);
          let selected: JsonValue | undefined = dataDocuments.get(target.value);
          if (!fragment.startsWith("/")) return err(migrationError("AgentPackageUnsupported"));
          for (const segment of fragment.slice(1).split("/")) {
            const parent = object(selected);
            const key = segment.replace(/~1/g, "/").replace(/~0/g, "~");
            selected = parent && Object.hasOwn(parent, key) ? parent[key] : undefined;
          }
          if (selected === undefined) return err(migrationError("AgentPackageReferenceMissing"));
        }
      }
      for (const child of Object.values(value)) {
        const walked = await walkReferences(child, owner);
        if (walked.isErr()) return err(walked.error);
      }
    }
    return ok(undefined);
  }

  async function visit(file: string, kind: DocumentKind): Promise<Result<undefined, FxError>> {
    const cancel = cancelled(signal);
    if (cancel) return err(cancel);
    if (visiting.has(file)) return err(migrationError("AgentPackageUnsupported"));
    if (graph.documents.has(file))
      return graph.documents.get(file)?.kind === kind
        ? ok(undefined)
        : err(migrationError("AgentPackageSchemaInvalid"));
    const bytes = files.get(file);
    if (!bytes) return err(migrationError("AgentPackageReferenceMissing"));
    const parsed = parseJson(bytes);
    if (parsed.isErr()) return err(parsed.error);
    const document = kind === "teams" ? graph.manifest : parsed.value;
    const resolved = { ...document };
    visiting.add(file);
    graph.files.add(file);
    if (kind === "agent" && typeof document.instructions === "string") {
      const ref = getStaticManifestFileReference(document.instructions);
      if (!ref && /^\$\[\s*file\([\s\S]*\)\s*\]$/.test(document.instructions)) {
        return err(migrationError("AgentPackageUnsupported"));
      }
      let content = document.instructions;
      let sourcePath: string | undefined;
      if (ref) {
        if (!/\.(txt|md)$/i.test(ref.path)) return err(migrationError("AgentPackageUnsupported"));
        const local = reference(file, ref.path, "/instructions");
        if (local.isErr()) return err(local.error);
        if (!local.value) return err(migrationError("AgentPackageUnsupported"));
        const value = text(files.get(local.value)!);
        if (value.isErr()) return err(value.error);
        content = value.value;
        sourcePath = local.value;
      }
      resolved.instructions = content;
      graph.instructions.set(file, { content, sourcePath });
    }
    const valid = await validateDocument(resolved, kind);
    if (valid.isErr()) return err(valid.error);
    graph.documents.set(file, { kind, value: document });
    if (kind === "agent") {
      for (const [field, childKind] of [
        ["actions", "plugin"],
        ["worker_agents", "agent"],
      ]) {
        for (const item of objects(document[field])) {
          if (item.file !== undefined) {
            const child = reference(file, item.file, `/${field}`, { parent: item, key: "file" });
            if (child.isErr()) return err(child.error);
            if (!child.value) return err(migrationError("AgentPackageUnsupported"));
            const result = await visit(child.value, childKind === "agent" ? "agent" : "plugin");
            if (result.isErr()) return err(result.error);
          } else if (item.id !== undefined)
            graph.external.push({ path: file, pointer: `/${field}/id` });
        }
      }
      const capabilities = objects(document.capabilities);
      if (new Set(capabilities.map((capability) => capability.name)).size !== capabilities.length)
        return err(migrationError("AgentPackageSchemaInvalid"));
      for (const capability of capabilities) {
        if (capability.name === "EmbeddedKnowledge") {
          for (const resource of objects(capability.files)) {
            // Embedded-knowledge paths are package-relative, as in the native package builder.
            const child = reference(
              "manifest.json",
              resource.file,
              "/capabilities/EmbeddedKnowledge/files"
            );
            if (child.isErr()) return err(child.error);
            if (
              normalizePaths &&
              typeof resource.file === "string" &&
              resource.file.includes("\\")
            ) {
              resource.file = resource.file.replace(/\\/g, "/");
              graph.normalized.add(file);
            }
            const editorialAnswers = object(document.editorial_answers);
            if (editorialAnswers?.url !== undefined) {
              const result = await dataFile(
                file,
                editorialAnswers.url,
                "/editorial_answers/url",
                false,
                { parent: editorialAnswers, key: "url" }
              );
              if (result.isErr()) return err(result.error);
            }
            if (
              !child.value ||
              files.get(child.value)!.length > 1024 * 1024 ||
              !knowledgeExtensions.has(path.posix.extname(child.value).toLowerCase())
            )
              return err(migrationError("AgentPackageLimitExceeded"));
          }
        } else if (
          capability.name !== "WebSearch" &&
          capability.name !== "CodeInterpreter" &&
          capability.name !== "GraphicArt"
        ) {
          graph.external.push({ path: file, pointer: "/capabilities" });
        }
      }
    } else if (kind === "plugin") {
      for (const runtime of objects(document.runtimes)) {
        const spec = object(runtime.spec);
        if (runtime.type === "LocalPlugin") return err(migrationError("AgentPackageUnsupported"));
        if (runtime.type === "OpenApi") {
          if (spec?.api_description !== undefined)
            return err(migrationError("AgentPackageUnsupported"));
          const result = await dataFile(
            file,
            spec?.url,
            "/runtimes/spec/url",
            true,
            spec ? { parent: spec, key: "url" } : undefined
          );
          if (result.isErr()) return err(result.error);
        } else if (runtime.type === "RemoteMCPServer") {
          const remote = reference(file, spec?.url, "/runtimes/spec/url");
          if (remote.isErr()) return err(remote.error);
          if (remote.value) return err(migrationError("AgentPackageUnsupported"));
          const tools = object(spec?.mcp_tool_description);
          if (tools?.file !== undefined) {
            const result = await dataFile(
              file,
              tools.file,
              "/runtimes/spec/mcp_tool_description/file",
              false,
              { parent: tools, key: "file" }
            );
            if (result.isErr()) return err(result.error);
          }
        }
      }
      for (const fn of objects(document.functions)) {
        const card = object(object(object(fn.capabilities)?.response_semantics)?.static_template);
        if (card?.file !== undefined) {
          const result = await dataFile(
            file,
            card.file,
            "/functions/capabilities/response_semantics/static_template/file",
            false,
            { parent: card, key: "file" }
          );
          if (result.isErr()) return err(result.error);
        }
      }
    }
    visiting.delete(file);
    return ok(undefined);
  }

  const container = await visit("manifest.json", "teams");
  if (container.isErr()) return err(container.error);
  const icons = object(manifest.icons);
  for (const [icon, size] of [
    ["color", 192],
    ["outline", 32],
    ["color32x32", 32],
  ]) {
    if (typeof icon !== "string" || typeof size !== "number" || icons?.[icon] === undefined)
      continue;
    const resolved = reference("manifest.json", icons[icon], `/icons/${icon}`, {
      parent: icons,
      key: icon,
    });
    if (resolved.isErr()) return err(resolved.error);
    if (!resolved.value) return err(migrationError("AgentPackageUnsupported"));
    const png = validatePng(files.get(resolved.value)!, size);
    if (png.isErr()) return err(png.error);
  }
  const localization = object(manifest.localizationInfo);
  for (const locale of objects(localization?.additionalLanguages)) {
    const result = await dataFile(
      "manifest.json",
      locale.file,
      "/localizationInfo/additionalLanguages/file",
      false,
      { parent: locale, key: "file" }
    );
    if (result.isErr()) return err(result.error);
  }
  if (localization?.defaultLanguageFile !== undefined) {
    const result = await dataFile(
      "manifest.json",
      localization.defaultLanguageFile,
      "/localizationInfo/defaultLanguageFile",
      false,
      { parent: localization, key: "defaultLanguageFile" }
    );
    if (result.isErr()) return err(result.error);
  }
  const agent = await visit(graph.agentPath, "agent");
  return agent.isErr() ? err(agent.error) : ok(graph);
}
