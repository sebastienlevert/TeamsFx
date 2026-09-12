// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { spawn } from "child_process";
import { randomUUID } from "crypto";
import fs from "fs-extra";
import path from "path";

export const agentId = "localAgent";
export const originalInstructions = "Exact source text\r\nCafé Ω\n$1 ${{LITERAL}} $[not-a-file]";

export async function createWorkspace(): Promise<string> {
  const workspace = path.resolve(__dirname, "..", "..", `.agent-package-${randomUUID()}`);
  await fs.mkdir(workspace);
  return workspace;
}

export async function createSource(workspace: string): Promise<string> {
  const source = path.join(workspace, "Export café");
  const appPackage = path.join(source, "appPackage");
  await fs.ensureDir(path.join(appPackage, "agents"));
  await fs.writeJson(path.join(appPackage, "manifest.json"), {
    $schema:
      "https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
    manifestVersion: "1.19",
    version: "1.0.0",
    id: "11111111-2222-4333-8444-555555555555",
    developer: {
      name: "Local author",
      websiteUrl: "https://example.com",
      privacyUrl: "https://example.com/privacy",
      termsOfUseUrl: "https://example.com/terms",
    },
    name: { short: "Local agent", full: "Local exported agent" },
    description: { short: "Local agent", full: "A local exported declarative agent." },
    icons: { color: "color.png", outline: "outline.png" },
    accentColor: "#FFFFFF",
    copilotAgents: {
      declarativeAgents: [{ id: agentId, file: "agents/agent.json" }],
    },
  });
  await fs.writeJson(path.join(appPackage, "agents", "agent.json"), {
    $schema:
      "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.5/schema.json",
    version: "v1.5",
    name: "Local agent",
    description: "A local exported declarative agent.",
    instructions: originalInstructions,
  });
  const colorIcon = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAAApklEQVR42u3BMQEAAADCoPVPbQlPoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPgZA3gABFrSBrAAAAABJRU5ErkJggg==",
    "base64"
  );
  const outlineIcon = Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAGklEQVR42u3BAQEAAACCIP+vbkhAAQAAAO8GECAAAcm1w7EAAAAASUVORK5CYII=",
    "base64"
  );
  await fs.writeFile(path.join(appPackage, "color.png"), colorIcon);
  await fs.writeFile(path.join(appPackage, "outline.png"), outlineIcon);
  await fs.ensureDir(path.join(appPackage, "candidates"));
  for (let index = 0; index < 12; index++) {
    await fs.writeFile(path.join(appPackage, "candidates", `notes-${index}.txt`), `Notes ${index}`);
  }
  return source;
}

export async function snapshot(folder: string): Promise<Record<string, string>> {
  const files: Record<string, string> = {};
  async function walk(relative: string): Promise<void> {
    for (const entry of (await fs.readdir(path.join(folder, relative))).sort()) {
      const child = path.join(relative, entry);
      if ((await fs.lstat(path.join(folder, child))).isDirectory()) {
        await walk(child);
      } else {
        files[child] = (await fs.readFile(path.join(folder, child))).toString("base64");
      }
    }
  }
  await walk("");
  return files;
}

export async function createSourceZip(source: string): Promise<string> {
  const files = await snapshot(source);
  const localEntries: Buffer[] = [];
  const directory: Buffer[] = [];
  let offset = 0;
  for (const [relative, base64] of Object.entries(files)) {
    const name = Buffer.from(relative.split(path.sep).join("/"));
    const contents = Buffer.from(base64, "base64");
    let crc = 0xffffffff;
    for (const byte of contents) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) {
        crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
      }
    }
    crc = (crc ^ 0xffffffff) >>> 0;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x800, 6);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(contents.length, 18);
    local.writeUInt32LE(contents.length, 22);
    local.writeUInt16LE(name.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x800, 8);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(contents.length, 20);
    central.writeUInt32LE(contents.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    localEntries.push(local, name, contents);
    directory.push(central, name);
    offset += local.length + name.length + contents.length;
  }
  const central = Buffer.concat(directory);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(Object.keys(files).length, 8);
  end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(offset, 16);
  const zip = source + ".zip";
  await fs.writeFile(zip, Buffer.concat([...localEntries, central, end]));
  return zip;
}

export async function writeEdits(
  workspace: string,
  operations: unknown[],
  name = "changes.json"
): Promise<string> {
  const changes = path.join(workspace, name);
  await fs.writeJson(changes, { schemaVersion: 1, agentId, operations });
  return changes;
}

export interface CliResult {
  code: number | null;
  signal: NodeJS.Signals | null;
  stdout: string;
  stderr: string;
}

export async function runCli(cwd: string, args: string[], cancel = false): Promise<CliResult> {
  const cliRoot = path.resolve(__dirname, "..", "..", "..");
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [
        "--require",
        path.join(__dirname, "agentPackage.offline.cjs"),
        path.join(cliRoot, "cli.js"),
        ...args,
      ],
      {
        cwd,
        shell: false,
        stdio: ["pipe", "pipe", "pipe"],
        env: {
          ...process.env,
          FORCE_COLOR: "0",
          ATK_TEST_CANCEL_LOCAL_AGENT: cancel ? "true" : "false",
        },
      }
    );
    let stdout = "";
    let stderr = "";
    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error(`CLI did not exit with closed stdin.\n${stdout}\n${stderr}`));
    }, 25000);
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (data: string) => {
      stdout += data;
    });
    child.stderr.on("data", (data: string) => {
      stderr += data;
    });
    child.on("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });
    child.on("close", (code, signal) => {
      clearTimeout(timeout);
      resolve({ code, signal, stdout, stderr });
    });
    child.stdin.end();
  });
}
