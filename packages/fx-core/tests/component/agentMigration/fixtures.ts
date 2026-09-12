// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import { createHash } from "crypto";

export const sourceAppId = "11111111-2222-4333-8444-555555555555";
export const sourceAgentId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";
export const instructionText =
  "\uFEFFUse the source exactly.\r\nCaf\u00e9 \u{1F680}\nLiteral ${{NOT_AN_ENV}} and $[file('not-a-file.txt')]. $& $$ $` $'\n";

export async function sourceFixture(root: string, referenced = false): Promise<string> {
  const source = path.join(root, "source");
  const template = path.resolve(
    __dirname,
    "../../../../..",
    "templates",
    "vsc",
    "common",
    "declarative-agent-basic",
    "appPackage"
  );
  await fs.outputJSON(path.join(source, "manifest.json"), {
    $schema: "https://developer.microsoft.com/json-schemas/teams/v1.24/MicrosoftTeams.schema.json",
    manifestVersion: "1.24",
    version: "2.3.4",
    id: sourceAppId,
    developer: {
      name: "Original maker",
      websiteUrl: "https://example.com",
      privacyUrl: "https://example.com/privacy",
      termsOfUseUrl: "https://example.com/terms",
    },
    name: { short: "Original agent", full: "Original agent full name" },
    description: { short: "Original short description", full: "Original full description" },
    icons: { color: "branding/original.png", outline: "branding/line.png" },
    accentColor: "#FFFFFF",
    copilotAgents: {
      declarativeAgents: [{ id: "logical-agent", file: "agents/declarativeAgent_7.json" }],
    },
    localizationInfo: {
      defaultLanguageTag: "en-us",
      additionalLanguages: [{ languageTag: "fr-fr", file: "locales/fr.json" }],
    },
  });
  await fs.outputJSON(path.join(source, "agents", "declarativeAgent_7.json"), {
    $schema:
      "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.8/schema.json",
    version: "v1.8",
    id: sourceAgentId,
    name: "Original agent",
    description: "Source behavior",
    instructions: referenced ? "$[file('original.txt')]" : instructionText,
    disclaimer: { text: "Original disclaimer" },
    sensitivity_label: { id: "label-resource-id" },
    worker_agents: [{ id: "external-worker-id" }],
    user_overrides: [{ path: "$.capabilities[0]", allowed_actions: ["remove"] }],
    behavior_overrides: { special_instructions: { discourage_model_knowledge: false } },
    conversation_starters: [{ title: "Start", text: "Original question" }],
    capabilities: [
      { name: "WebSearch" },
      {
        name: "OneDriveAndSharePoint",
        items_by_url: [{ url: "https://example.sharepoint.com/sites/original" }],
      },
      { name: "EmbeddedKnowledge", files: [{ file: "agents/docs/handbook.txt" }] },
    ],
    actions: [{ id: "original-action", file: "actions/plugin.json" }],
  });
  if (referenced) await fs.outputFile(path.join(source, "agents", "original.txt"), instructionText);
  await fs.copy(path.join(template, "color.png"), path.join(source, "branding", "original.png"));
  await fs.copy(path.join(template, "outline.png"), path.join(source, "branding", "line.png"));
  await fs.outputJSON(path.join(source, "locales", "fr.json"), {
    name: { short: "Agent original" },
    description: { short: "Description originale" },
  });
  await fs.outputFile(
    path.join(source, "agents", "docs", "handbook.txt"),
    "Approved source knowledge"
  );
  await fs.outputJSON(path.join(source, "agents", "actions", "plugin.json"), {
    $schema: "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.4/schema.json",
    schema_version: "v2.4",
    name_for_human: "Original tool",
    namespace: "OriginalTool",
    description_for_human: "Preserved tool",
    functions: [
      {
        name: "lookup",
        description: "Read approved data",
        capabilities: {
          response_semantics: {
            data_path: "$",
            static_template: { file: "cards/result.json" },
          },
        },
      },
    ],
    runtimes: [
      {
        type: "OpenApi",
        auth: { type: "None" },
        spec: { url: "api/openapi.yaml" },
        run_for_functions: ["lookup"],
      },
    ],
  });
  await fs.outputFile(
    path.join(source, "agents", "actions", "api", "openapi.yaml"),
    [
      "openapi: 3.0.1",
      "info: { title: Original, version: 1.0.0 }",
      "servers: [{ url: 'https://example.com' }]",
      "paths:",
      "  /lookup:",
      "    get:",
      "      operationId: lookup",
      "      responses:",
      "        '200':",
      "          description: result",
      "          content:",
      "            application/json:",
      "              schema:",
      "                $ref: './schemas/result.yaml'",
      "",
    ].join("\n")
  );
  await fs.outputFile(
    path.join(source, "agents", "actions", "api", "schemas", "result.yaml"),
    "type: object\nproperties:\n  value:\n    type: string\n"
  );
  await fs.outputJSON(path.join(source, "agents", "actions", "cards", "result.json"), {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: "Original" },
      { type: "Image", url: "images/original.png" },
    ],
  });
  await fs.copy(
    path.join(template, "color.png"),
    path.join(source, "agents", "actions", "cards", "images", "original.png")
  );
  for (let i = 0; i < 12; i++) {
    await fs.outputFile(path.join(source, "candidates", String(i), "same.txt"), `candidate ${i}`);
  }
  await fs.outputFile(path.join(source, "instruction.txt"), "NOT the active instructions");
  await fs.outputFile(path.join(source, "untrusted.js"), "throw new Error('never execute');");
  return source;
}

export async function snapshot(root: string): Promise<Record<string, string>> {
  const entries: Record<string, string> = {};
  async function walk(directory: string): Promise<void> {
    for (const entry of (await fs.readdir(directory, { withFileTypes: true })).sort((a, b) =>
      a.name.localeCompare(b.name)
    )) {
      const full = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(full);
      else
        entries[path.relative(root, full).replace(/\\/g, "/")] = createHash("sha256")
          .update(await fs.readFile(full))
          .digest("hex");
    }
  }
  await walk(root);
  return entries;
}

export async function zipFixture(source: string, target: string, wrapped = false): Promise<string> {
  const zip = new AdmZip();
  zip.addLocalFolder(source, wrapped ? "appPackage" : "");
  await fs.writeFile(target, zip.toBuffer());
  return target;
}

export async function changeFile(
  root: string,
  operations: unknown[],
  agentId = "logical-agent"
): Promise<string> {
  const target = path.join(root, "approved", "changes.json");
  await fs.outputJSON(target, { schemaVersion: 1, agentId, operations });
  return target;
}
