// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import fs from "fs-extra";
import path from "path";
import YAML from "yaml";
import {
  AgentTemplateIdentity,
  Context,
  err,
  FxError,
  GeneratorResult,
  IGenerator,
  Inputs,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import { getResourceFolder } from "../../folder";
import { coordinator } from "../coordinator";
import { Generator } from "../generator/generator";
import { renderTemplateFileData, renderTemplateFileName, unzip } from "../generator/utils";
import { artifactDigest, Artifacts, decodeProfileArchive, parseJson } from "./model";
import { cancelled, migrationError } from "./errors";

/**
 * The native no-action profile, deliberately without new-agent enrichment or
 * template-channel selection. No source files are passed to the renderer.
 */
export class AgentPackageGenerator implements IGenerator {
  readonly componentName = "agent-package-generator";
  template: AgentTemplateIdentity | undefined;
  projectId: string | undefined;

  constructor(
    private readonly appName: string,
    private readonly signal?: AbortSignal
  ) {}

  async run(
    _context: Context,
    _inputs: Inputs,
    destinationPath: string
  ): Promise<Result<GeneratorResult, FxError>> {
    const cancel = cancelled(this.signal);
    if (cancel) return err(cancel);
    try {
      const templateName = "declarative-agent-basic";
      const version = "6.16.0";
      const directory = path.join(getResourceFolder(), "agent-import", version);
      const profile = parseJson(await fs.readFile(path.join(directory, "profile.json")));
      if (profile.isErr()) return err(profile.error);
      if (
        profile.value.profileVersion !== 1 ||
        profile.value.templateId !== templateName ||
        profile.value.templateVersion !== version ||
        profile.value.archiveEncoding !== "base64" ||
        typeof profile.value.archiveSha256 !== "string" ||
        typeof profile.value.contentDigest !== "string"
      ) {
        return err(migrationError("AgentPackageUnsupported"));
      }
      const archive = decodeProfileArchive(
        await fs.readFile(path.join(directory, "template.zip.b64")),
        profile.value.archiveSha256
      );
      if (archive.isErr()) return err(archive.error);
      const zip = new AdmZip(archive.value);
      const entries: Artifacts = new Map(
        zip
          .getEntries()
          .filter((entry) => !entry.isDirectory && entry.entryName.startsWith(`${templateName}/`))
          .map((entry) => [entry.entryName, entry.getData()])
      );
      if (entries.size === 0 || artifactDigest(entries) !== profile.value.contentDigest) {
        return err(migrationError("AgentPackageUnsupported"));
      }
      this.template = {
        id: templateName,
        version,
        digest: artifactDigest(entries),
      };
      const variables = Generator.getDefaultVariables("ImportedAgent");
      await unzip(
        zip,
        destinationPath,
        (name, data) =>
          renderTemplateFileName(name, data, variables).slice(templateName.length + 1),
        (name, data) => renderTemplateFileData(name, data, variables),
        (name) =>
          name.startsWith(`${templateName}/`) && !name.startsWith(`${templateName}/appPackage/`)
      );
      for (const name of ["m365agents.yml", "m365agents.local.yml"]) {
        const file = path.join(destinationPath, name);
        const document = YAML.parseDocument(await fs.readFile(file, "utf8"));
        if (document.errors.length > 0) return err(migrationError("AgentPackageSchemaInvalid"));
        document.setIn(["provision", 0, "with", "name"], `${this.appName}\${{APP_NAME_SUFFIX}}`);
        await fs.writeFile(file, document.toString());
      }
      const tracking = await coordinator.ensureTrackingId(destinationPath);
      if (tracking.isErr()) return err(tracking.error);
      this.projectId = tracking.value;
      return ok({});
    } catch (error) {
      return err(migrationError("AgentMigrationIoError", error));
    }
  }
}
