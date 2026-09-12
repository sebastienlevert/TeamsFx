// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import os from "os";
import path from "path";
import AdmZip from "adm-zip";
import Ajv from "ajv";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  AgentMigrationReport,
  agentMigrationReportSchema,
  FxError,
  Result,
  ManifestType,
  Platform,
  resolveManifest,
} from "@microsoft/teamsfx-api";
import { FxCoreClient } from "../../../src/core/FxCoreClient";
import { MockTools } from "../../core/utils";
import { agentMigrationIo } from "../../../src/component/agentMigration/io";
import {
  changeFile,
  instructionText,
  snapshot,
  sourceAgentId,
  sourceAppId,
  sourceFixture,
  zipFixture,
} from "./fixtures";

function report(result: Result<AgentMigrationReport, FxError>): AgentMigrationReport {
  if (result.isErr()) throw result.error;
  return result.value;
}

describe("native agent package operations", () => {
  let root: string;
  let client: FxCoreClient;
  let source: string;
  let output: string;
  beforeEach(async () => {
    root = await fs.mkdtemp(path.join(os.tmpdir(), "atk-agent-import-"));
    source = await sourceFixture(root);
    output = path.join(root, "new project");
    client = new FxCoreClient(new MockTools());
  });
  afterEach(async () => {
    vi.restoreAllMocks();
    await fs.remove(root);
  });
  const agentPath = (project: string) =>
    path.join(project, "appPackage", "agents", "declarativeAgent_7.json");
  async function imported(): Promise<AgentMigrationReport> {
    return report(await client.importAgentPackage({ sourcePath: source, outputPath: output }));
  }
  async function effective(project: string): Promise<string> {
    const fromPath = agentPath(project);
    const result = await resolveManifest(await fs.readFile(fromPath, "utf8"), {
      fromPath,
      envs: {},
      manifestType: ManifestType.DeclarativeCopilotManifest,
    });
    return JSON.parse(result.content).instructions;
  }

  it.each(["directory", "zip", "wrapped"])(
    "IMP-01: imports real %s with a unique graph and immutable source",
    async (kind) => {
      const before = await snapshot(source);
      const sourcePath =
        kind === "directory"
          ? source
          : await zipFixture(source, path.join(root, "source.zip"), kind === "wrapped");
      const result = report(await client.importAgentPackage({ sourcePath, outputPath: output }));
      expect(result.operationMode).toBe("imported");
      expect(result.projectPath).toBe(output);
      expect(result.projectDigest).toMatch(/^sha256:[a-f0-9]{64}$/);
      expect(await snapshot(source)).toEqual(before);
      expect(await fs.readFile(path.join(output, "m365agents.yml"), "utf8")).toContain(
        "teamsApp/zipAppPackage"
      );
      expect(result.template?.id).toBe("declarative-agent-basic");
      expect(await fs.pathExists(path.join(output, "appPackage", "declarativeAgent.json"))).toBe(
        false
      );
    }
  );

  it.each([false, true])(
    "IMP-02: preserves exact effective instructions and subsequent edits (file=%s)",
    async (referenced) => {
      if (referenced) source = await sourceFixture(root, true);
      await imported();
      expect(await effective(output)).toBe(instructionText);
      const agent = await fs.readJSON(agentPath(output));
      const match = /file\('([^']+)', 'raw'\)/.exec(agent.instructions);
      expect(match).not.toBeNull();
      await fs.writeFile(
        path.join(path.dirname(agentPath(output)), match![1]),
        "Approved later\r\nchange"
      );
      expect(await effective(output)).toBe("Approved later\r\nchange");
    }
  );

  it("IMP-03: preserves modern fields and resource scopes without source deployment takeover", async () => {
    const before = await fs.readJSON(path.join(source, "agents", "declarativeAgent_7.json"));
    const result = await imported();
    const after = await fs.readJSON(agentPath(output));
    for (const field of [
      "name",
      "description",
      "disclaimer",
      "sensitivity_label",
      "worker_agents",
      "user_overrides",
      "behavior_overrides",
      "capabilities",
      "conversation_starters",
      "version",
    ]) {
      expect(after[field]).toEqual(before[field]);
    }
    expect(after.id).toBeUndefined();
    expect(result.identity.sourceAppId).toBe(sourceAppId);
    expect(result.identity.sourceAgentId).toBe(sourceAgentId);
    expect(result.identity.agentId).toBe("logical-agent");
    expect((await fs.readJSON(path.join(output, "appPackage", "manifest.json"))).id).toBe(
      "${{TEAMS_APP_ID}}"
    );
  });

  it("IMP-04: preserves nested action/OpenAPI/card/localization/knowledge and image bytes", async () => {
    const before = await snapshot(source);
    await imported();
    const after = await snapshot(path.join(output, "appPackage"));
    for (const file of [
      "branding/original.png",
      "branding/line.png",
      "locales/fr.json",
      "agents/actions/plugin.json",
      "agents/actions/cards/result.json",
      "agents/actions/cards/images/original.png",
      "agents/actions/api/openapi.yaml",
      "agents/actions/api/schemas/result.yaml",
      "agents/docs/handbook.txt",
    ]) {
      expect(after[file]).toBe(before[file]);
    }
  });

  it("IMP-05: inventories all candidates without attaching or flattening them", async () => {
    const result = await imported();
    expect(result.files.filter((file) => file.action === "candidate")).toHaveLength(13);
    expect(await fs.pathExists(path.join(output, "appPackage", "instruction.txt"))).toBe(false);
    expect(await fs.pathExists(path.join(output, "appPackage", "untrusted.js"))).toBe(false);
    expect((await fs.readJSON(agentPath(output))).capabilities[2].files).toHaveLength(1);
  });

  it.each(["ambiguous", "multiple", "missing", "json", "schema", "utf8"])(
    "IMP-06: rejects %s source before destination writes",
    async (kind) => {
      const manifest = path.join(source, "manifest.json");
      if (kind === "ambiguous")
        await fs.copy(manifest, path.join(source, "appPackage", "manifest.json"));
      if (kind === "multiple") {
        const value = await fs.readJSON(manifest);
        value.copilotAgents.declarativeAgents.push({
          id: "other",
          file: "agents/declarativeAgent_7.json",
        });
        await fs.writeJSON(manifest, value);
      }
      if (kind === "missing") await fs.remove(path.join(source, "branding", "original.png"));
      if (kind === "json") await fs.writeFile(manifest, "{ bad");
      if (kind === "utf8") await fs.writeFile(manifest, Buffer.from([0xff, 0xfe, 0x41]));
      if (kind === "schema") {
        const value = await fs.readJSON(manifest);
        value.manifestVersion = "99.0";
        await fs.writeJSON(manifest, value);
      }
      const result = await client.importAgentPackage({ sourcePath: source, outputPath: output });
      expect(result.isErr()).toBe(true);
      expect(await fs.pathExists(output)).toBe(false);
    }
  );

  it("IMP-07: bounds actual expanded bytes and verifies archive integrity", async () => {
    const zip = new AdmZip();
    zip.addFile("bomb.txt", Buffer.alloc(10 * 1024 * 1024 + 1, 65));
    const filename = path.join(root, "bomb.zip");
    await fs.writeFile(filename, zip.toBuffer());
    const result = await client.importAgentPackage({ sourcePath: filename, outputPath: output });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("AgentPackageLimitExceeded");
    await fs.writeFile(filename, Buffer.from("not a zip"));
    expect(
      (await client.importAgentPackage({ sourcePath: filename, outputPath: output })).isErr()
    ).toBe(true);
  });

  it.each(["../escape", "C:\\escape", "\\\\server\\share", "/absolute", "CON.txt", "trailing."])(
    "IMP-08: rejects unsafe package names %s",
    async (name) => {
      const rawName = Buffer.from(name, "utf8");
      const zip = new AdmZip();
      zip.addFile("x".repeat(rawName.length), Buffer.from("invalid"));
      const archive = zip.toBuffer();
      const central = archive.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
      // The writer sanitizes names; inject the actual unsafe bytes into both real ZIP headers.
      rawName.copy(archive, 30);
      rawName.copy(archive, central + 46);
      const filename = path.join(root, "unsafe.zip");
      await fs.writeFile(filename, archive);
      const result = await client.importAgentPackage({ sourcePath: filename, outputPath: output });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("AgentPackagePathInvalid");
      expect(await fs.pathExists(output)).toBe(false);
    }
  );

  it("IMP-08: rejects case collisions and file/directory prefix collisions", async () => {
    const zip = new AdmZip();
    zip.addFile("A.txt", Buffer.from("a"));
    zip.addFile("a.txt", Buffer.from("b"));
    const filename = path.join(root, "collision.zip");
    await fs.writeFile(filename, zip.toBuffer());
    const result = await client.importAgentPackage({ sourcePath: filename, outputPath: output });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("AgentPackageCollision");
  });

  it("IMP-09: native generation is offline, noninteractive, and uses a fresh tracking ID", async () => {
    const tools = new MockTools();
    for (const key of ["selectOption", "inputText", "showMessage"] as const) {
      vi.spyOn(tools.ui, key).mockRejectedValue(new Error("UI forbidden"));
    }
    vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("Network forbidden"));
    client = new FxCoreClient(tools);
    const first = await imported();
    const second = report(
      await client.importAgentPackage({ sourcePath: source, outputPath: path.join(root, "other") })
    );
    expect(first.identity.projectId).not.toBe(second.identity.projectId);
    expect(first.identity.provisionedByOperation).toBe(false);
    expect(first.configuration.readyToPublish).toBe("not-evaluated");
    expect(await fs.readFile(path.join(output, "env", ".env.dev"), "utf8")).not.toContain(
      sourceAppId
    );
    expect(globalThis.fetch).not.toHaveBeenCalled();
  });

  it("IMP-10: dry run, destination guard, cancellation and promotion faults leave no partial project", async () => {
    expect(
      report(
        await client.importAgentPackage({ sourcePath: source, outputPath: output, dryRun: true })
      ).operationMode
    ).toBe("dry-run");
    expect(await fs.pathExists(output)).toBe(false);
    const controller = new AbortController();
    controller.abort();
    expect(
      (
        await client.importAgentPackage(
          { sourcePath: source, outputPath: output },
          { signal: controller.signal }
        )
      ).isErr()
    ).toBe(true);
    vi.spyOn(agentMigrationIo, "promote").mockRejectedValueOnce(
      new Error("injected promotion failure")
    );
    expect(
      (await client.importAgentPackage({ sourcePath: source, outputPath: output })).isErr()
    ).toBe(true);
    expect(await fs.pathExists(output)).toBe(false);
    await fs.ensureDir(output);
    await fs.writeFile(path.join(output, "mine.txt"), "keep");
    const result = await client.importAgentPackage({ sourcePath: source, outputPath: output });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("AgentPackageDestinationExists");
    expect(await fs.readFile(path.join(output, "mine.txt"), "utf8")).toBe("keep");
  });

  it("IMP-11: resolves relative paths without changing CWD", async () => {
    const cwd = process.cwd();
    const result = report(
      await client.importAgentPackage({
        sourcePath: path.relative(cwd, source),
        outputPath: path.relative(cwd, path.join(root, "sp ace-\u00e9")),
      })
    );
    expect(result.projectPath).toBe(path.join(root, "sp ace-\u00e9"));
    expect(process.cwd()).toBe(cwd);
  });

  it("EDT-01: applies exact replacements and removals retaining omitted fields", async () => {
    await imported();
    const changes = await changeFile(root, [
      { kind: "setAppMetadata", value: { version: "2.3.5" } },
      { kind: "setAgentMetadata", value: { description: "Approved description" } },
      { kind: "replaceConversationStarters", value: [{ title: "New", text: "Approved question" }] },
      { kind: "upsertCapability", value: { name: "OneDriveAndSharePoint" } },
      { kind: "removeCapability", name: "WebSearch" },
      {
        kind: "setBehaviorOverrides",
        value: { special_instructions: { discourage_model_knowledge: true } },
      },
      { kind: "setSchemaVersion", value: "v1.8" },
    ]);
    const result = report(
      await client.applyAgentEdits({ projectPath: output, changesFile: changes })
    );
    const agent = await fs.readJSON(agentPath(output));
    expect(result.operationMode).toBe("edited");
    expect((await fs.readJSON(path.join(output, "appPackage", "manifest.json"))).version).toBe(
      "2.3.5"
    );
    expect(agent.name).toBe("Original agent");
    expect(agent.conversation_starters).toEqual([{ title: "New", text: "Approved question" }]);
    expect(agent.capabilities).toEqual([
      { name: "OneDriveAndSharePoint" },
      { name: "EmbeddedKnowledge", files: [{ file: "agents/docs/handbook.txt" }] },
    ]);
    expect(agent.behavior_overrides.special_instructions.discourage_model_knowledge).toBe(true);
  });

  it("EDT-01: descriptive metadata edits preserve omitted members of name and developer", async () => {
    await imported();
    const file = path.join(output, "appPackage", "manifest.json");
    const before = await fs.readJSON(file);
    const changes = await changeFile(root, [
      {
        kind: "setAppMetadata",
        value: {
          name: { short: "Approved short name" },
          developer: { websiteUrl: "https://example.com/new" },
        },
      },
    ]);
    report(await client.applyAgentEdits({ projectPath: output, changesFile: changes }));
    const after = await fs.readJSON(file);
    expect(after.name).toEqual({ ...before.name, short: "Approved short name" });
    expect(after.developer).toEqual({ ...before.developer, websiteUrl: "https://example.com/new" });
  });

  it.each([
    { kind: "unknown" },
    { kind: "setAppMetadata", value: { id: "forbidden" } },
    { kind: "replaceInstructions", sourceFile: "a.txt", extra: true },
    { kind: "setSchemaVersion", value: "v99.0" },
    { kind: "replaceConversationStarters", value: [] },
  ])("EDT-02: invalid operation fails without writes: %j", async (operation) => {
    await imported();
    const before = await snapshot(output);
    const changes = await changeFile(root, [operation]);
    expect(
      (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
    ).toBe(true);
    expect(await snapshot(output)).toEqual(before);
  });

  it("EDT-03: explicit instruction/icon/knowledge assets retain bytes and nested names", async () => {
    await imported();
    await fs.outputFile(path.join(root, "approved", "text.txt"), "Approved\r\n${{literal}}");
    await fs.outputFile(path.join(root, "approved", "one.txt"), "one");
    await fs.copy(
      path.join(source, "branding", "line.png"),
      path.join(root, "approved", "line.png")
    );
    const changes = await changeFile(root, [
      { kind: "replaceInstructions", sourceFile: "text.txt" },
      { kind: "replaceIcon", icon: "outline", sourceFile: "line.png" },
      {
        kind: "attachEmbeddedKnowledge",
        files: [{ sourceFile: "one.txt", targetPath: "nested/one.txt" }],
      },
    ]);
    report(await client.applyAgentEdits({ projectPath: output, changesFile: changes }));
    expect(await effective(output)).toBe("Approved\r\n${{literal}}");
    expect(
      await fs.readFile(path.join(output, "appPackage", "knowledge", "nested", "one.txt"), "utf8")
    ).toBe("one");
  });

  it("EDT-04: dry-run/no-op and stale digest preserve bytes and mtimes", async () => {
    const initial = await imported();
    const changes = await changeFile(root, [
      { kind: "setAgentMetadata", value: { name: "Approved" } },
    ]);
    const before = await snapshot(output);
    report(
      await client.applyAgentEdits({ projectPath: output, changesFile: changes, dryRun: true })
    );
    expect(await snapshot(output)).toEqual(before);
    const applied = report(
      await client.applyAgentEdits({ projectPath: output, changesFile: changes })
    );
    const stat = await fs.stat(agentPath(output));
    const noop = report(
      await client.applyAgentEdits({
        projectPath: output,
        changesFile: changes,
        expectedDigest: applied.projectDigest,
      })
    );
    expect(noop.operationMode).toBe("no-op");
    expect((await fs.stat(agentPath(output))).mtimeMs).toBe(stat.mtimeMs);
    const stale = await client.applyAgentEdits({
      projectPath: output,
      changesFile: changes,
      expectedDigest: initial.projectDigest,
    });
    expect(stale.isErr()).toBe(true);
    if (stale.isErr()) expect(stale.error.name).toBe("AgentEditsStale");
  });

  it("EDT-05: concurrent mutation before commit is never overwritten", async () => {
    await imported();
    const changes = await changeFile(root, [
      { kind: "setAgentMetadata", value: { name: "Approved" } },
    ]);
    const original = agentMigrationIo.beforeCommit;
    vi.spyOn(agentMigrationIo, "beforeCommit").mockImplementationOnce(async () => {
      await fs.appendFile(agentPath(output), "\n ");
      await original();
    });
    const result = await client.applyAgentEdits({ projectPath: output, changesFile: changes });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("AgentEditsConflict");
    expect((await fs.readFile(agentPath(output), "utf8")).endsWith("\n ")).toBe(true);
  });

  it("EDT-06: a mid-commit fault rolls back only owned writes", async () => {
    await imported();
    const before = await snapshot(output);
    const changes = await changeFile(root, [
      { kind: "setAppMetadata", value: { version: "3.0.0" } },
      { kind: "setAgentMetadata", value: { name: "Approved" } },
    ]);
    const replace = agentMigrationIo.replace;
    let count = 0;
    vi.spyOn(agentMigrationIo, "replace").mockImplementation(async (...args) => {
      if (++count === 2) throw new Error("injected replacement failure");
      return replace(...args);
    });
    expect(
      (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
    ).toBe(true);
    expect(await snapshot(output)).toEqual(before);
  });

  it("EDT-07: edits preserve existing deployment/env and report unknown template truthfully", async () => {
    await imported();
    await fs.remove(path.join(output, ".atk", "import.json"));
    const file = path.join(output, "appPackage", "manifest.json");
    const manifest = await fs.readJSON(file);
    manifest.id = sourceAppId;
    await fs.writeJSON(file, manifest);
    await fs.appendFile(path.join(output, "env", ".env.dev"), `\nM365_APP_ID=${sourceAppId}\n`);
    const changes = await changeFile(root, []);
    const before = await snapshot(output);
    const result = report(
      await client.applyAgentEdits({ projectPath: output, changesFile: changes })
    );
    expect(result.template).toBeNull();
    expect(result.identity.policy).toBe("preserved");
    expect(result.identity.appId).toBe(sourceAppId);
    expect(result.configuration.readyToPublish).toBe("not-evaluated");
    expect(await snapshot(output)).toEqual(before);
  });

  it("EDT-08: explicit asset traversal and knowledge count overflow fail as a whole", async () => {
    await imported();
    const before = await snapshot(output);
    const changes = await changeFile(root, [
      { kind: "replaceInstructions", sourceFile: "..\\outside.txt" },
    ]);
    expect(
      (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
    ).toBe(true);
    expect(await snapshot(output)).toEqual(before);
  });

  it("SCN-AGENT-PACKAGE-01: native public import/edit/package preserves effective text and full asset closure", async () => {
    await imported();
    await fs.appendFile(path.join(output, "env", ".env.dev"), `\nTEAMS_APP_ID=${sourceAppId}\n`);
    async function packaged(): Promise<AdmZip> {
      const result = await client.package({
        platform: Platform.CLI,
        projectPath: output,
        env: "dev",
        "output-package-file": path.join(root, "compiled.zip"),
        "output-folder": path.join(output, "appPackage", "build"),
      });
      if (result.isErr()) throw result.error;
      return new AdmZip(await fs.readFile(result.value.packagePath));
    }
    const first = await packaged();
    expect(JSON.parse(first.readAsText("agents/declarativeAgent_7.json")).instructions).toBe(
      instructionText
    );
    expect(
      JSON.parse(first.readAsText("agents/actions/plugin.json")).functions[0].capabilities
        .response_semantics.static_template
    ).toEqual({ file: "cards/result.json" });
    const sourceFiles = await snapshot(source);
    for (const name of [
      "branding/original.png",
      "branding/line.png",
      "locales/fr.json",
      "agents/actions/cards/result.json",
      "agents/actions/api/openapi.yaml",
      "agents/actions/cards/images/original.png",
      "agents/actions/api/schemas/result.yaml",
      "agents/docs/handbook.txt",
    ]) {
      const entry = first.getEntry(name);
      expect(entry, name).not.toBeNull();
      expect(entry?.getData()).toEqual(await fs.readFile(path.join(source, ...name.split("/"))));
    }
    await fs.outputFile(
      path.join(root, "approved", "text.txt"),
      "Reviewed exact instructions\r\n${{literal}}"
    );
    const changes = await changeFile(root, [
      { kind: "replaceInstructions", sourceFile: "text.txt" },
    ]);
    report(await client.applyAgentEdits({ projectPath: output, changesFile: changes }));
    const second = await packaged();
    expect(JSON.parse(second.readAsText("agents/declarativeAgent_7.json")).instructions).toBe(
      "Reviewed exact instructions\r\n${{literal}}"
    );
    expect(await snapshot(source)).toEqual(sourceFiles);
  });

  it("SCN-AGENT-PACKAGE-03: the exported report schema validates real import and edit reports", async () => {
    const validate = new Ajv().compile(agentMigrationReportSchema);
    const importedReport = await imported();
    expect(validate(importedReport), JSON.stringify(validate.errors)).toBe(true);
    const changes = await changeFile(root, []);
    const edited = report(
      await client.applyAgentEdits({ projectPath: output, changesFile: changes })
    );
    expect(validate(edited), JSON.stringify(validate.errors)).toBe(true);
  });
});
