// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import os from "os";
import path from "path";
import http from "http";
import https from "https";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { FxCoreClient } from "../../../src/core/FxCoreClient";
import { agentMigrationIo } from "../../../src/component/agentMigration/io";
import { readArchive } from "../../../src/component/agentMigration/intake";
import { packageLimits } from "../../../src/component/agentMigration/paths";
import { MockTools } from "../../core/utils";
import { changeFile, snapshot, sourceFixture, zipFixture } from "./fixtures";

describe("bounded local agent operations", () => {
  let root: string;
  let source: string;
  let output: string;
  let client: FxCoreClient;
  beforeEach(async () => {
    root = await fs.mkdtemp(path.join(os.tmpdir(), "atk-agent-safety-"));
    source = await sourceFixture(root);
    output = path.join(root, "project");
    client = new FxCoreClient(new MockTools());
  });
  afterEach(async () => {
    vi.restoreAllMocks();
    await fs.remove(root);
  });
  async function imported(): Promise<void> {
    const result = await client.importAgentPackage({ sourcePath: source, outputPath: output });
    if (result.isErr()) throw result.error;
  }

  it.each(["count", "aggregate", "path", "depth", "compressed"])(
    "IMP-07: enforces actual %s bounds with real archive data",
    (kind) => {
      const zip = new AdmZip();
      if (kind === "count") {
        for (let i = 0; i <= packageLimits.entries; i++)
          zip.addFile(`file-${i}.txt`, Buffer.alloc(0));
      } else if (kind === "aggregate") {
        for (let i = 0; i < 7; i++)
          zip.addFile(`file-${i}.txt`, Buffer.alloc(packageLimits.fileBytes));
      } else if (kind === "path")
        zip.addFile(`${"a".repeat(packageLimits.pathLength)}.txt`, Buffer.alloc(0));
      else if (kind === "depth")
        zip.addFile(`${"a/".repeat(packageLimits.depth)}file.txt`, Buffer.alloc(0));
      const bytes =
        kind === "compressed" ? Buffer.alloc(packageLimits.archiveBytes + 1) : zip.toBuffer();
      const result = readArchive(bytes);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("AgentPackageLimitExceeded");
    }
  );

  it.each(["crc", "encryption", "method", "forged-zero", "symlink", "prefix"])(
    "IMP-07 IMP-08: rejects crafted %s archive metadata",
    (kind) => {
      const zip = new AdmZip();
      zip.addFile(
        "payload.txt",
        kind === "forged-zero" ? Buffer.alloc(packageLimits.fileBytes + 2) : Buffer.from("payload")
      );
      if (kind === "symlink") zip.getEntry("payload.txt")!.attr = 0xa1ff0000;
      if (kind === "prefix") zip.addFile("payload.txt/child", Buffer.alloc(0));
      const bytes = zip.toBuffer();
      const central = bytes.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
      if (kind === "crc") bytes.writeUInt32LE(bytes.readUInt32LE(central + 16) ^ 1, central + 16);
      if (kind === "encryption")
        bytes.writeUInt16LE(bytes.readUInt16LE(central + 8) | 1, central + 8);
      if (kind === "method") bytes.writeUInt16LE(99, central + 10);
      if (kind === "forged-zero") {
        bytes.writeUInt32LE(0, central + 24);
        bytes.writeUInt32LE(0, 22);
      }
      expect(readArchive(bytes).isErr()).toBe(true);
    }
  );

  it("IMP-08: refuses directory junctions and hard links", async () => {
    const external = path.join(root, "external");
    await fs.ensureDir(external);
    await fs.symlink(external, path.join(source, "link"), "junction");
    expect(
      (await client.importAgentPackage({ sourcePath: source, outputPath: output })).isErr()
    ).toBe(true);
    await fs.unlink(path.join(source, "link"));
    await fs.writeFile(path.join(external, "data.txt"), "outside");
    await fs.link(path.join(external, "data.txt"), path.join(source, "hardlink.txt"));
    expect(
      (await client.importAgentPackage({ sourcePath: source, outputPath: output })).isErr()
    ).toBe(true);
    expect(await fs.pathExists(output)).toBe(false);
  });

  it("IMP-06: rejects an icon with an invalid PNG chunk checksum", async () => {
    const icon = path.join(source, "branding", "original.png");
    const bytes = await fs.readFile(icon);
    bytes[29] ^= 1;
    await fs.writeFile(icon, bytes);
    expect(
      (await client.importAgentPackage({ sourcePath: source, outputPath: output })).isErr()
    ).toBe(true);
    expect(await fs.pathExists(output)).toBe(false);
  });

  it("IMP-08: rejects canonically equivalent Unicode names", () => {
    const zip = new AdmZip();
    zip.addFile("caf\u00e9.txt", Buffer.from("one"));
    zip.addFile("cafe\u0301.txt", Buffer.from("two"));
    const result = readArchive(zip.toBuffer());
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("AgentPackageCollision");
  });

  it("IMP-05: an instruction filename collision cannot overwrite an unreferenced candidate", async () => {
    await fs.outputFile(path.join(source, "agents", "instructions.txt"), "unapproved");
    const result = await client.importAgentPackage({ sourcePath: source, outputPath: output });
    if (result.isErr()) throw result.error;
    const agent = await fs.readJSON(
      path.join(output, "appPackage", "agents", "declarativeAgent_7.json")
    );
    expect(agent.instructions).toContain("instructions-1.txt");
    expect(
      result.value.files.some(
        (file) => file.sourcePath === "agents/instructions.txt" && file.action === "candidate"
      )
    ).toBe(true);
    expect(await fs.pathExists(path.join(output, "appPackage", "agents", "instructions.txt"))).toBe(
      false
    );
  });

  it("IMP-03: supported empty scopes stay empty and unsupported nulls fail rather than being stripped", async () => {
    const agentFile = path.join(source, "agents", "declarativeAgent_7.json");
    const agent = await fs.readJSON(agentFile);
    agent.capabilities = [{ name: "OneDriveAndSharePoint" }];
    agent.user_overrides = [];
    agent.worker_agents = [];
    await fs.writeJSON(agentFile, agent);
    await imported();
    const after = await fs.readJSON(
      path.join(output, "appPackage", "agents", "declarativeAgent_7.json")
    );
    expect(after.capabilities).toEqual(agent.capabilities);
    expect(after.user_overrides).toEqual([]);
    expect(after.worker_agents).toEqual([]);
    agent.capabilities = null;
    await fs.writeJSON(agentFile, agent);
    expect(
      (
        await client.importAgentPackage({
          sourcePath: source,
          outputPath: path.join(root, "invalid"),
        })
      ).isErr()
    ).toBe(true);
  });

  it("IMP-04: preserves a local worker's graph and retains its deployment ID only in provenance", async () => {
    const file = path.join(source, "agents", "declarativeAgent_7.json");
    const agent = await fs.readJSON(file);
    agent.worker_agents.push({ file: "workers/helper.json" });
    await fs.writeJSON(file, agent);
    await fs.outputJSON(path.join(source, "agents", "workers", "helper.json"), {
      version: "v1.8",
      id: "old-worker-deployment",
      name: "Worker",
      description: "Original worker",
      instructions: "Original worker text",
    });
    await imported();
    const worker = await fs.readJSON(
      path.join(output, "appPackage", "agents", "workers", "helper.json")
    );
    expect(worker.id).toBeUndefined();
    expect(worker.name).toBe("Worker");
    const provenance = await fs.readJSON(path.join(output, ".atk", "import.json"));
    expect(provenance.sourceAgents["agents/workers/helper.json"]).toBe("old-worker-deployment");
  });

  it("IMP-09: local APIs never request HTTP, authentication, or AI", async () => {
    const httpRequest = vi.spyOn(http, "request");
    const httpsRequest = vi.spyOn(https, "request");
    const tools = new MockTools();
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const azure = vi.spyOn(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync");
    client = new FxCoreClient(tools);
    await imported();
    const changes = await changeFile(root, []);
    const result = await client.applyAgentEdits({ projectPath: output, changesFile: changes });
    expect(result.isOk()).toBe(true);
    for (const spy of [httpRequest, httpsRequest, token, azure]) expect(spy).not.toHaveBeenCalled();
  });

  it("IMP-09: the pinned migration profile does not read ordinary creation archives", async () => {
    const original = fs.readFile;
    vi.spyOn(fs, "readFile").mockImplementation((...args) => {
      if (String(args[0]).replace(/\\/g, "/").endsWith("/templates/fallback/common.zip")) {
        throw new Error("Ordinary creation templates must not be used for migration");
      }
      return original(...args);
    });
    const result = await client.importAgentPackage({ sourcePath: source, outputPath: output });
    if (result.isErr()) throw result.error;
    expect(result.value.template).toEqual({
      id: "declarative-agent-basic",
      version: "6.16.0",
      digest: "sha256:f32852fdf0fad7ead0d77443918171d627b2660ff8e670affbb9b4c7ee86f9a2",
    });
  });

  it.each(["!!!!", "AAAA", "ZE=="])(
    "IMP-09: rejects malformed or wrong-hash encoded profiles before ZIP parsing (%s)",
    async (encoded) => {
      const original = fs.readFile;
      vi.spyOn(fs, "readFile").mockImplementation((...args) => {
        if (String(args[0]).endsWith("template.zip.b64")) {
          return Promise.resolve(Buffer.from(`${encoded}\n`));
        }
        return original(...args);
      });
      const result = await client.importAgentPackage({ sourcePath: source, outputPath: output });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("AgentPackageIntegrityInvalid");
      expect(await fs.pathExists(output)).toBe(false);
    }
  );

  it("IMP-11: portable source references are normalized without changing prose or asset bytes", async () => {
    const file = path.join(source, "manifest.json");
    const manifest = await fs.readJSON(file);
    manifest.copilotAgents.declarativeAgents[0].file = "agents\\declarativeAgent_7.json";
    manifest.icons.color = "branding\\original.png";
    await fs.writeJSON(file, manifest);
    await imported();
    const after = await fs.readJSON(path.join(output, "appPackage", "manifest.json"));
    expect(after.copilotAgents.declarativeAgents[0].file).toBe("agents/declarativeAgent_7.json");
    expect(after.icons.color).toBe("branding/original.png");
    expect(after.description).toEqual(manifest.description);
  });

  it("IMP-11: a missing output parent is created only for a successful committed import", async () => {
    const nested = path.join(root, "new parent", "new child", "project");
    const preview = await client.importAgentPackage({
      sourcePath: source,
      outputPath: nested,
      dryRun: true,
    });
    expect(preview.isOk()).toBe(true);
    expect(await fs.pathExists(path.join(root, "new parent"))).toBe(false);
    const committed = await client.importAgentPackage({ sourcePath: source, outputPath: nested });
    expect(committed.isOk()).toBe(true);
    expect(await fs.pathExists(path.join(nested, "m365agents.yml"))).toBe(true);
  });

  it("IMP-01: wrapper-external files are accounted for and never copied into the project", async () => {
    const file = await zipFixture(source, path.join(root, "wrapped.zip"), true);
    const zip = new AdmZip(await fs.readFile(file));
    zip.addFile("outside.txt", Buffer.from("not package content"));
    await fs.writeFile(file, zip.toBuffer());
    const result = await client.importAgentPackage({ sourcePath: file, outputPath: output });
    if (result.isErr()) throw result.error;
    expect(
      result.value.files.some(
        (item) => item.sourcePath === "outside.txt" && item.action === "skipped"
      )
    ).toBe(true);
    expect(await fs.pathExists(path.join(output, "outside.txt"))).toBe(false);
  });

  it.each(["version", "agent", "field", "provenance"])(
    "EDT-02: rejects invalid %s before all writes",
    async (kind) => {
      await imported();
      const changes = await changeFile(root, []);
      const doc = await fs.readJSON(changes);
      if (kind === "version") doc.schemaVersion = 2;
      if (kind === "agent") doc.agentId = "unknown";
      if (kind === "field") doc.hook = "evil.js";
      if (kind === "provenance") await fs.writeFile(path.join(output, ".atk", "import.json"), "{}");
      await fs.writeJSON(changes, doc);
      const before = await snapshot(output);
      expect(
        (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
      ).toBe(true);
      expect(await snapshot(output)).toEqual(before);
    }
  );

  it("EDT-06: mid-commit cancellation rolls back the entire owned change set", async () => {
    await imported();
    const before = await snapshot(output);
    const changes = await changeFile(root, [
      { kind: "setAppMetadata", value: { version: "4.0.0" } },
      { kind: "setAgentMetadata", value: { name: "Approved" } },
    ]);
    const controller = new AbortController();
    const original = agentMigrationIo.replace;
    vi.spyOn(agentMigrationIo, "replace").mockImplementationOnce(async (...args) => {
      await original(...args);
      controller.abort();
    });
    const result = await client.applyAgentEdits(
      { projectPath: output, changesFile: changes },
      { signal: controller.signal }
    );
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("UserCancel");
    expect(await snapshot(output)).toEqual(before);
  });

  it("EDT-05 EDT-06: rollback retains recovery evidence rather than overwrite an independent edit", async () => {
    await imported();
    const changes = await changeFile(root, [
      { kind: "setAppMetadata", value: { version: "4.0.0" } },
      { kind: "setAgentMetadata", value: { name: "Approved" } },
    ]);
    const original = agentMigrationIo.replace;
    let count = 0;
    let written = "";
    vi.spyOn(agentMigrationIo, "replace").mockImplementation(async (from, to) => {
      if (++count === 2) {
        await fs.appendFile(written, "\n ");
        throw new Error("injected failure with concurrent write");
      }
      written = to;
      await original(from, to);
    });
    const result = await client.applyAgentEdits({ projectPath: output, changesFile: changes });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("AgentMigrationRecoveryRequired");
    expect((await fs.readFile(written, "utf8")).endsWith("\n ")).toBe(true);
    expect(await fs.pathExists(path.join(root, ".project.agent-edit", "journal.json"))).toBe(true);
    expect(
      (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
    ).toBe(true);
  });

  it("EDT-08: explicit knowledge count, collisions and bounded change files fail without truncation", async () => {
    await imported();
    const before = await snapshot(output);
    const files = [];
    for (let i = 0; i < 11; i++) {
      await fs.outputFile(path.join(root, "approved", `${i}.txt`), `approved ${i}`);
      files.push({ sourceFile: `${i}.txt`, targetPath: `${i}.txt` });
    }
    const changes = await changeFile(root, [{ kind: "attachEmbeddedKnowledge", files }]);
    expect(
      (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
    ).toBe(true);
    await fs.writeFile(changes, " ".repeat(packageLimits.changesBytes + 1));
    expect(
      (await client.applyAgentEdits({ projectPath: output, changesFile: changes })).isErr()
    ).toBe(true);
    expect(await snapshot(output)).toEqual(before);
  });
});
