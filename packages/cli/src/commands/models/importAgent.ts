// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AgentImportRequest, CLICommand, err } from "@microsoft/teamsfx-api";
import { MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { agentPackageOptions, runAgentPackageOperation } from "./agentPackage";

export const importAgentCommand: CLICommand = {
  name: "agent",
  description: commands["import.agent"].description,
  defaultInteractiveOption: false,
  options: [
    {
      name: "source",
      questionName: "sourcePath",
      type: "string",
      required: true,
      description: commands["import.agent"].options.source,
    },
    {
      name: "output",
      questionName: "outputPath",
      type: "string",
      description: commands["import.agent"].options.output,
    },
    ...agentPackageOptions,
  ],
  examples: [
    {
      command:
        "atk import agent --source exported-agent.zip --output new-agent --format json -i false",
      description: commands["import.agent"].example,
    },
  ],
  handler: async (context) => {
    const { sourcePath, outputPath, dryRun } = context.optionValues;
    if (typeof sourcePath !== "string" || sourcePath.length === 0) {
      return err(new MissingRequiredOptionError(context.command.fullName, "source"));
    }
    const request: AgentImportRequest = {
      sourcePath,
      outputPath: typeof outputPath === "string" ? outputPath : undefined,
      dryRun: dryRun === true,
    };
    return runAgentPackageOperation(context, (client, options) =>
      client.importAgentPackage(request, options)
    );
  },
};
