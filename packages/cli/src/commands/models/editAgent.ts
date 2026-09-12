// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AgentEditRequest, CLICommand, err } from "@microsoft/teamsfx-api";
import { MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { agentPackageOptions, runAgentPackageOperation } from "./agentPackage";

export const editAgentCommand: CLICommand = {
  name: "agent",
  description: commands["edit.agent"].description,
  defaultInteractiveOption: false,
  options: [
    {
      name: "folder",
      questionName: "projectPath",
      type: "string",
      required: true,
      description: commands["edit.agent"].options.folder,
    },
    {
      name: "changes",
      questionName: "changesFile",
      type: "string",
      required: true,
      description: commands["edit.agent"].options.changes,
    },
    {
      name: "expected-digest",
      questionName: "expectedDigest",
      type: "string",
      description: commands["edit.agent"].options.expectedDigest,
    },
    ...agentPackageOptions,
  ],
  examples: [
    {
      command: "atk edit agent --folder new-agent --changes changes.json --format json -i false",
      description: commands["edit.agent"].example,
    },
  ],
  handler: async (context) => {
    const { projectPath, changesFile, expectedDigest, dryRun } = context.optionValues;
    if (typeof projectPath !== "string" || projectPath.length === 0) {
      return err(new MissingRequiredOptionError(context.command.fullName, "folder"));
    }
    if (typeof changesFile !== "string" || changesFile.length === 0) {
      return err(new MissingRequiredOptionError(context.command.fullName, "changes"));
    }
    const request: AgentEditRequest = {
      projectPath,
      changesFile,
      expectedDigest: typeof expectedDigest === "string" ? expectedDigest : undefined,
      dryRun: dryRun === true,
    };
    return runAgentPackageOperation(context, (client, options) =>
      client.applyAgentEdits(request, options)
    );
  },
};
