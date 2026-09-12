// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserCancelError } from "@microsoft/teamsfx-core";
import fs from "fs-extra";
import path from "path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  CliResult,
  createSource,
  createSourceZip,
  createWorkspace,
  runCli,
  snapshot,
  writeEdits,
} from "./agentPackage.fixtures";

describe("SCN-AGENT-PACKAGE-02: real offline CLI with closed stdin", () => {
  let workspace: string;

  beforeEach(async () => {
    workspace = await createWorkspace();
  });

  afterEach(async () => {
    await fs.remove(workspace);
  });

  function envelope(result: CliResult, exitCode: number) {
    expect(result.signal, result.stderr).toBeNull();
    expect(result.code, result.stderr).toBe(exitCode);
    expect(result.stderr).not.toContain("LOCAL_AGENT_OFFLINE_BOUNDARY_VIOLATION");
    expect(result.stdout.trim().split("\n")).toHaveLength(1);
    const parsed = JSON.parse(result.stdout);
    expect(Object.keys(parsed).sort()).toEqual(
      exitCode === 0 ? ["result", "success"] : ["error", "success"]
    );
    if (exitCode !== 0) {
      expect(Object.keys(parsed.error).sort()).toEqual(["message", "name", "source"]);
    }
    return parsed;
  }

  it.each(["import", "edit"])(
    "SCN-AGENT-PACKAGE-02: exposes %s agent help and version without auth/network",
    async (verb) => {
      const help = await runCli(workspace, [verb, "agent", "--help"]);
      expect(help.code, help.stderr).toBe(0);
      expect(help.stdout).toContain(`Usage: atk ${verb} agent`);
      expect(help.stdout).toContain("--dry-run");
      expect(help.stdout).toContain("--format");
      expect(help.stdout).toContain(verb === "import" ? "--source" : "--changes");
      const version = await runCli(workspace, [verb, "agent", "--version"]);
      expect(version.code, version.stderr).toBe(0);
      expect(version.stdout.trim()).toMatch(/^\d+\.\d+\.\d+/);
    }
  );

  it.each([
    { args: ["import", "agent"], name: "MissingRequiredOptionError" },
    { args: ["edit", "agent"], name: "MissingRequiredOptionError" },
    { args: ["edit", "agent", "--folder", "."], name: "MissingRequiredOptionError" },
    { args: ["import", "agent", "--unknown"], name: "UnknownOptionError" },
    { args: ["import", "agent", "--format=json", "--unknown"], name: "UnknownOptionError" },
    { args: ["edit", "agent", "--unknown=value"], name: "UnknownOptionError" },
    { args: ["import", "agent", "--source"], name: "MissingRequiredOptionError" },
    {
      args: ["edit", "agent", "--folder", ".", "--changes"],
      name: "MissingRequiredOptionError",
    },
    {
      args: ["edit", "agent", "--folder", ".", "--changes", "changes.json", "--expected-digest"],
      name: "MissingRequiredOptionError",
    },
    {
      args: ["import", "agent", "--source", "missing.zip"],
      name: "AgentPackageSourceInvalid",
    },
    {
      args: ["edit", "agent", "--folder", "missing", "--changes", "missing.json"],
      name: "AgentPackageSourceInvalid",
    },
  ])(
    "SCN-AGENT-PACKAGE-02 / IMP-06: returns one JSON error for $args before or after --format",
    async ({ args, name }) => {
      const result = envelope(
        await runCli(workspace, [...args, "--format=json", "-i", "false"]),
        1
      );
      expect(result.success).toBe(false);
      expect(result.error.name).toBe(name);
    }
  );

  it("SCN-AGENT-PACKAGE-02 / IMP-01 / IMP-11: imports a real ZIP using the default output path", async () => {
    const source = await createSource(workspace);
    const zip = await createSourceZip(source);
    const before = await fs.readFile(zip);
    const imported = envelope(
      await runCli(workspace, ["import", "agent", "--source", zip, "--format=json"]),
      0
    );
    expect(imported.result.operationMode).toBe("imported");
    expect(imported.result.source.kind).toBe("zip");
    expect(imported.result.projectPath).toBe(path.join(workspace, "Export café-imported"));
    expect(await fs.readFile(zip)).toEqual(before);
  });

  it.each(["export root", "appPackage"])(
    "SCN-AGENT-PACKAGE-02 / IMP-11: defaults a relative %s source to its basename in a Unicode caller CWD",
    async (sourceKind) => {
      const source = await createSource(workspace);
      const sourceBefore = await snapshot(source);
      const caller = path.join(workspace, "Caller Ω");
      await fs.mkdir(caller);
      const packagePath = sourceKind === "appPackage" ? path.join(source, "appPackage") : source;
      const project = path.join(caller, "Export café-imported");
      const imported = envelope(
        await runCli(caller, [
          "import",
          "agent",
          "--source",
          path.relative(caller, packagePath),
          "--format",
          "json",
          "-i",
          "false",
        ]),
        0
      );
      expect(imported.result.projectPath).toBe(project);
      expect(imported.result.source.kind).toBe("directory");
      expect(imported.result.operationMode).toBe("imported");
      expect(await fs.readFile(path.join(project, "m365agents.yml"), "utf8")).toContain(
        "provision"
      );
      expect(await fs.pathExists(path.join(workspace, "Export café-imported"))).toBe(false);
      expect(await snapshot(source)).toEqual(sourceBefore);
    }
  );

  it("SCN-AGENT-PACKAGE-02 / IMP-10 / EDT-04: imports, previews, edits, no-ops and rejects stale digests", async () => {
    const source = await createSource(workspace);
    const sourceBefore = await snapshot(source);
    const project = path.join(workspace, "Native project café");
    const importArgs = [
      "import",
      "agent",
      "--source",
      path.relative(workspace, source),
      "--output",
      path.relative(workspace, project),
      "--format",
      "json",
      "-i",
      "false",
    ];
    const preview = envelope(await runCli(workspace, [...importArgs, "--dry-run"]), 0);
    expect(preview.result.operationMode).toBe("dry-run");
    expect(await fs.pathExists(project)).toBe(false);
    const imported = envelope(await runCli(workspace, importArgs), 0);
    expect(imported.result.operationMode).toBe("imported");
    expect(imported.result.projectPath).toBe(project);
    expect(imported.result.identity).toMatchObject({
      policy: "new",
      provisionedByOperation: false,
    });
    expect(imported.result.configuration).toMatchObject({
      structurallyValid: true,
      readyToProvision: "not-evaluated",
      readyToPublish: "not-evaluated",
    });
    expect(imported.result.diagnostics).toBeInstanceOf(Array);
    expect(imported.result.template).not.toBeNull();
    expect(await snapshot(source)).toEqual(sourceBefore);
    const duplicate = envelope(await runCli(workspace, importArgs), 1);
    expect(duplicate.error.name).toBe("AgentPackageDestinationExists");

    const changes = await writeEdits(workspace, [
      { kind: "setAgentMetadata", value: { description: "Approved local description" } },
    ]);
    const editArgs = [
      "edit",
      "agent",
      "--folder",
      project,
      "--changes",
      changes,
      "--format",
      "json",
      "-i",
      "false",
    ];
    const beforeEdit = await snapshot(project);
    const editPreview = envelope(await runCli(workspace, [...editArgs, "--dry-run"]), 0);
    expect(editPreview.result.operationMode).toBe("dry-run");
    expect(await snapshot(project)).toEqual(beforeEdit);
    const edited = envelope(await runCli(workspace, editArgs), 0);
    expect(edited.result.operationMode).toBe("edited");
    expect(edited.result.identity.policy).toBe("preserved");
    const afterEdit = await snapshot(project);
    const noOp = envelope(
      await runCli(workspace, [...editArgs, "--expected-digest", edited.result.projectDigest]),
      0
    );
    expect(noOp.result.operationMode).toBe("no-op");
    expect(noOp.result.changed).toBe(false);
    expect(await snapshot(project)).toEqual(afterEdit);
    const stale = envelope(
      await runCli(workspace, [...editArgs, "--expected-digest", `sha256:${"0".repeat(64)}`]),
      1
    );
    expect(stale.error.name).toBe("AgentEditsStale");
    expect(await snapshot(project)).toEqual(afterEdit);
  }, 120000);

  it("SCN-AGENT-PACKAGE-02 / EDT-02: reports invalid edit documents without partial writes", async () => {
    const source = await createSource(workspace);
    const project = path.join(workspace, "Edit validation");
    envelope(
      await runCli(workspace, [
        "import",
        "agent",
        "--source",
        source,
        "--output",
        project,
        "--format",
        "json",
      ]),
      0
    );
    const before = await snapshot(project);
    const changes = await writeEdits(workspace, [{ kind: "runScript", command: "not-executed" }]);
    const failure = envelope(
      await runCli(workspace, [
        "edit",
        "agent",
        "--folder",
        project,
        "--changes",
        changes,
        "--format",
        "json",
      ]),
      1
    );
    expect(failure.error.name).toBe("AgentEditsInvalid");
    expect(await snapshot(project)).toEqual(before);
  });

  it("SCN-AGENT-PACKAGE-02 / IMP-10: cancellation returns exit 130 and cleans owned output", async () => {
    const source = await createSource(workspace);
    const project = path.join(workspace, "Canceled project");
    const failure = envelope(
      await runCli(
        workspace,
        ["import", "agent", "--source", source, "--output", project, "--format", "json"],
        true
      ),
      130
    );
    expect(failure.error.name).toBe(new UserCancelError().name);
    expect(await fs.pathExists(project)).toBe(false);
  });
});
