// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const fs = require("fs/promises");
const path = require("path");
const AdmZip = require("adm-zip");
const {
  artifactDigest,
  decodeProfileArchive,
  digest,
} = require("../build/component/agentMigration/model");
const templateConfig = require("../src/common/templates-config.json");

async function main() {
  const version = process.argv[2] || "6.16.0";
  const templateId = "declarative-agent-basic";
  if (!/^\d+\.\d+\.\d+$/.test(version) || templateConfig.localVersion !== version) {
    throw new Error(
      "Build the matching native template release before freezing an import profile."
    );
  }
  const native = new AdmZip(
    await fs.readFile(path.join(__dirname, "..", "templates", "fallback", "common.zip"))
  );
  const files = new Map(
    native
      .getEntries()
      .filter((entry) => !entry.isDirectory && entry.entryName.startsWith(`${templateId}/`))
      .map((entry) => [entry.entryName, entry.getData()])
  );
  if (!files.size) throw new Error("The native declarative-agent-basic profile is missing.");
  const contentDigest = artifactDigest(files);
  const directory = path.join(__dirname, "..", "resource", "agent-import", version);
  const profilePath = path.join(directory, "profile.json");
  let archiveBytes;
  let existingProfile = false;
  let legacyBinary = false;
  try {
    const existing = JSON.parse(await fs.readFile(profilePath, "utf8"));
    existingProfile = true;
    if (existing.archiveEncoding === "base64") {
      const decoded = decodeProfileArchive(
        await fs.readFile(path.join(directory, "template.zip.b64")),
        existing.archiveSha256
      );
      if (decoded.isErr()) throw decoded.error;
      archiveBytes = decoded.value;
    } else {
      archiveBytes = await fs.readFile(path.join(directory, "template.zip"));
      legacyBinary = true;
    }
    const archive = new AdmZip(archiveBytes);
    const contents = new Map(
      archive
        .getEntries()
        .filter((entry) => !entry.isDirectory)
        .map((entry) => [entry.entryName, entry.getData()])
    );
    if (existing.contentDigest !== contentDigest || artifactDigest(contents) !== contentDigest) {
      throw new Error(
        "Import profiles are immutable. Freeze a new version instead of overwriting this one."
      );
    }
    if (!legacyBinary) {
      process.stdout.write(`Verified import profile ${version}: ${contentDigest}\n`);
      return;
    }
  } catch (error) {
    if (error.code !== "ENOENT" || existingProfile) throw error;
  }
  if (!archiveBytes) {
    const archive = new AdmZip();
    for (const [name, bytes] of files) archive.addFile(name, bytes);
    archiveBytes = archive.toBuffer();
  }
  await fs.mkdir(directory, { recursive: true });
  await fs.writeFile(
    path.join(directory, "template.zip.b64"),
    archiveBytes.toString("base64") + "\n",
    { flag: "wx" }
  );
  await fs.writeFile(
    profilePath,
    JSON.stringify(
      {
        profileVersion: 1,
        templateId,
        templateVersion: version,
        archiveEncoding: "base64",
        archiveSha256: digest(archiveBytes).slice("sha256:".length),
        contentDigest,
        source: "templates/vsc/common/declarative-agent-basic",
      },
      null,
      2
    ) + "\n",
    { flag: existingProfile ? "w" : "wx" }
  );
  if (legacyBinary) await fs.unlink(path.join(directory, "template.zip"));
  process.stdout.write(`Froze import profile ${version}: ${contentDigest}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error.message}\n`);
  process.exitCode = 1;
});
