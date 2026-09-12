// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLIContext } from "@microsoft/teamsfx-api";
import { FeatureFlags, UserCancelError, featureFlagManager } from "@microsoft/teamsfx-core";
import fs from "fs-extra";
import { cloneDeep } from "lodash";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import * as activate from "../../../src/activate";
import { engine } from "../../../src/commands/engine";
import { createLocalAgentClient } from "../../../src/commands/models/agentPackage";
import { importCommand } from "../../../src/commands/models/import";
import { rootCommand } from "../../../src/commands/models/root";
import CliTelemetry from "../../../src/telemetry/cliTelemetry";
import UI from "../../../src/userInteraction";
import { createSource, createWorkspace, snapshot, writeEdits } from "./agentPackage.fixtures";

describe("SCN-AGENT-PACKAGE-02: native CLI command harness", () => {
  let workspace: string;
  let stdout: string[];
  let stderr: string[];
  let originalExitCode: typeof process.exitCode;
  let originalInteractive: boolean;

  beforeEach(async () => {
    workspace = await createWorkspace();
    stdout = [];
    stderr = [];
    originalExitCode = process.exitCode;
    originalInteractive = UI.interactive;
    engine.debugLogs = [];
    vi.spyOn(process.stdout, "write").mockImplementation((chunk, encoding, callback) => {
      stdout.push(String(chunk));
      const done = typeof encoding === "function" ? encoding : callback;
      done?.();
      return true;
    });
    vi.spyOn(process.stderr, "write").mockImplementation((chunk, encoding, callback) => {
      stderr.push(String(chunk));
      const done = typeof encoding === "function" ? encoding : callback;
      done?.();
      return true;
    });
    vi.spyOn(activate, "getFxCore").mockImplementation(() => {
      throw new Error("Local commands must not activate online CLI tools.");
    });
    vi.spyOn(UI, "inputText").mockImplementation(() => {
      throw new Error("Local commands must never ask questions.");
    });
    vi.spyOn(CliTelemetry, "sendTelemetryEvent");
    vi.spyOn(CliTelemetry, "sendTelemetryErrorEvent");
    vi.spyOn(CliTelemetry, "withRootFolder");
  });

  afterEach(async () => {
    try {
      expect(activate.getFxCore).not.toHaveBeenCalled();
      expect(UI.inputText).not.toHaveBeenCalled();
      expect(CliTelemetry.sendTelemetryEvent).not.toHaveBeenCalled();
      expect(CliTelemetry.sendTelemetryErrorEvent).not.toHaveBeenCalled();
      expect(CliTelemetry.withRootFolder).not.toHaveBeenCalled();
    } finally {
      vi.restoreAllMocks();
      process.exitCode = originalExitCode;
      UI.interactive = originalInteractive;
      await fs.remove(workspace);
    }
  });

  function commandContext(args: string[]) {
    const root = cloneDeep(rootCommand);
    const found = engine.findCommand(root, args);
    const context: CLIContext = {
      command: found.cmd,
      optionValues: {},
      globalOptionValues: {},
      argumentValues: [],
      telemetryProperties: {},
    };
    return { root, context, remaining: found.remainingArgs };
  }

  async function execute(args: string[]) {
    const { root, context, remaining } = commandContext(args);
    const cwd = process.cwd();
    const result = await engine.execute(context, root, remaining);
    expect(process.cwd()).toBe(cwd);
    return { context, result };
  }

  it("SCN-AGENT-PACKAGE-02: registers import/edit agent independently of plugin feature flags", () => {
    for (const verb of ["import", "edit"]) {
      const { context, remaining } = commandContext([verb, "agent"]);
      expect(context.command.fullName).toBe(`atk ${verb} agent`);
      expect(context.command.handler).toBeTypeOf("function");
      expect(context.command.defaultInteractiveOption).toBe(false);
      expect(remaining).toEqual([]);
    }
    const getFlag = featureFlagManager.getBooleanValue.bind(featureFlagManager);
    vi.spyOn(featureFlagManager, "getBooleanValue").mockImplementation((flag) =>
      flag === FeatureFlags.OpenPluginImportExport ? false : getFlag(flag)
    );
    expect(importCommand().commands?.map((command) => command.name)).toEqual(["agent"]);
    vi.mocked(featureFlagManager.getBooleanValue).mockImplementation((flag) =>
      flag === FeatureFlags.OpenPluginImportExport ? true : getFlag(flag)
    );
    const plugin = importCommand().commands?.find((command) => command.name === "openplugin");
    expect(plugin?.aliases).toContain("agentplugin");
  });

  it("SCN-AGENT-PACKAGE-02: maps native options without platform inputs or interactive trimming", () => {
    const { root, context, remaining } = commandContext([
      "import",
      "agent",
      "--source=export=name.zip",
      "--output",
      "New café project",
      "--dry-run",
      "--format=json",
      "-i",
      "true",
      "--telemetry=true",
    ]);
    expect(engine.parseArgs(context, root, remaining).isOk()).toBe(true);
    expect(context.optionValues).toMatchObject({
      sourcePath: "export=name.zip",
      outputPath: "New café project",
      dryRun: true,
      format: "json",
      nonInteractive: true,
    });
    expect(context.optionValues.platform).toBeUndefined();
    expect(context.globalOptionValues.interactive).toBe(false);
    expect(context.globalOptionValues.telemetry).toBe(false);
    expect(stdout).toEqual([]);
  });

  it.each(["import", "edit"])(
    "SCN-AGENT-PACKAGE-02: %s agent help accurately describes local-only global options",
    async (verb) => {
      const { result } = await execute([verb, "agent", "--help"]);
      expect(result.isOk()).toBe(true);
      expect(stdout.join("")).toContain("Telemetry is disabled for local agent commands.");
      expect(stdout.join("")).toContain("always noninteractive");
      expect(stderr).toEqual([]);
    }
  );

  it.each([
    { verb: "import", values: {}, option: "source" },
    { verb: "edit", values: {}, option: "folder" },
    { verb: "edit", values: { projectPath: "." }, option: "changes" },
  ])(
    "SCN-AGENT-PACKAGE-02: direct $verb handler rejects missing --$option",
    async ({ verb, values, option }) => {
      const { context } = commandContext([verb, "agent"]);
      context.optionValues = values;
      const handler = context.command.handler;
      expect(handler).toBeTypeOf("function");
      if (!handler) throw new Error("Native command handler is missing.");
      const result = await handler(context);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.name).toBe("MissingRequiredOptionError");
        expect(result.error.message).toContain(option);
      }
    }
  );

  it.each([
    { args: ["import", "agent"], name: "MissingRequiredOptionError", option: "--source" },
    { args: ["edit", "agent"], name: "MissingRequiredOptionError", option: "--folder" },
    {
      args: ["edit", "agent", "--folder", "."],
      name: "MissingRequiredOptionError",
      option: "--changes",
    },
    {
      args: ["import", "agent", "--source", ".", "--output"],
      name: "MissingRequiredOptionError",
      option: "--output",
    },
    {
      args: ["import", "agent", "--source", "--unknown"],
      name: "MissingRequiredOptionError",
      option: "--source",
    },
    {
      args: ["import", "agent", "--source", ".", "--wrong"],
      name: "UnknownOptionError",
      option: "--wrong",
    },
    {
      args: ["import", "agent", "--source", ".", "--format", "xml"],
      name: "InvalidChoiceError",
      option: "--format",
    },
  ])("SCN-AGENT-PACKAGE-02: rejects $args without prompts", async ({ args, name, option }) => {
    const { result } = await execute([...args, "-i", "true"]);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe(name);
      expect(result.error.message).toContain(option);
    }
    expect(stdout).toEqual([]);
  });

  it("SCN-AGENT-PACKAGE-02 / IMP-01: imports through real core and preserves the entire report", async () => {
    const source = await createSource(workspace);
    const sourceBefore = await snapshot(source);
    const output = path.join(workspace, "Native project");
    const { result } = await execute([
      "import",
      "agent",
      "--source",
      source,
      "--output",
      output,
      "--format",
      "json",
      "-i",
      "false",
      "--debug",
    ]);
    expect(result.isOk(), result.isErr() ? result.error.message : "").toBe(true);
    expect(stdout).toHaveLength(1);
    const envelope = JSON.parse(stdout.join(""));
    expect(envelope).toMatchObject({
      success: true,
      result: {
        operationMode: "imported",
        projectPath: output,
        identity: { policy: "new", provisionedByOperation: false },
        configuration: {
          structurallyValid: true,
          readyToProvision: "not-evaluated",
          readyToPublish: "not-evaluated",
        },
        recoveryRequired: false,
      },
    });
    expect(envelope.result.diagnostics).toBeInstanceOf(Array);
    expect(envelope.result.files).toBeInstanceOf(Array);
    expect(envelope.result.transformations).toBeInstanceOf(Array);
    expect(envelope.result.template).not.toBeNull();
    expect(await snapshot(source)).toEqual(sourceBefore);
    expect(await fs.readFile(path.join(output, "m365agents.yml"), "utf8")).toContain("provision");
  });

  it("SCN-AGENT-PACKAGE-02 / IMP-10: forwards SIGINT and removes its handler without creating output", async () => {
    const source = await createSource(workspace);
    const output = path.join(workspace, "Canceled project");
    const listeners = process.listenerCount("SIGINT");
    const on = process.on.bind(process);
    vi.spyOn(process, "on").mockImplementation((event, listener) => {
      const result = on(event, listener);
      if (event === "SIGINT") queueMicrotask(() => process.emit("SIGINT"));
      return result;
    });
    const { context, result } = await execute([
      "import",
      "agent",
      "--source",
      source,
      "--output",
      output,
      "--format",
      "json",
    ]);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe(new UserCancelError().name);
      await engine.processResult(context, result.error);
    }
    expect(process.exitCode).toBe(130);
    expect(JSON.parse(stdout.join(""))).toMatchObject({
      success: false,
      error: { name: new UserCancelError().name },
    });
    expect(process.listenerCount("SIGINT")).toBe(listeners);
    expect(await fs.pathExists(output)).toBe(false);
  });

  it("SCN-AGENT-PACKAGE-02 / EDT-04: runs real dry-run/edit/no-op and preserves API report fields", async () => {
    const source = await createSource(workspace);
    const project = path.join(workspace, "Editable project");
    const imported = await execute(["import", "agent", "--source", source, "--output", project]);
    expect(
      imported.result.isOk(),
      imported.result.isErr() ? imported.result.error.message : ""
    ).toBe(true);
    const changes = await writeEdits(workspace, [
      { kind: "setAgentMetadata", value: { description: "Approved replacement description" } },
    ]);
    const before = await snapshot(project);
    for (const { flags, operationMode } of [
      { flags: ["--dry-run"], operationMode: "dry-run" },
      { flags: [], operationMode: "edited" },
      { flags: [], operationMode: "no-op" },
    ]) {
      stdout.length = 0;
      const { result } = await execute([
        "edit",
        "agent",
        "--folder",
        project,
        "--changes",
        changes,
        "--format",
        "json",
        ...flags,
      ]);
      expect(result.isOk(), result.isErr() ? result.error.message : "").toBe(true);
      const envelope = JSON.parse(stdout.join(""));
      expect(envelope.result.operationMode).toBe(operationMode);
      expect(envelope.result.identity.policy).toBe("preserved");
      if (operationMode === "dry-run") expect(await snapshot(project)).toEqual(before);
      if (operationMode === "no-op") {
        const direct = await createLocalAgentClient().applyAgentEdits({
          projectPath: project,
          changesFile: changes,
        });
        expect(direct.isOk(), direct.isErr() ? direct.error.message : "").toBe(true);
        if (direct.isOk()) expect(envelope.result).toEqual(direct.value);
      }
    }
    expect(UI.interactive).toBe(originalInteractive);
    expect(stderr.join("")).not.toContain("interactive");
  });
});
