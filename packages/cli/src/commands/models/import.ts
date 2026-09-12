// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand } from "@microsoft/teamsfx-api";
import { FeatureFlags, featureFlagManager } from "@microsoft/teamsfx-core";
import { commands } from "../../resource";
import { importAgentCommand } from "./importAgent";
import { importOpenPluginCommand } from "./importOpenPlugin";

export function importCommand(): CLICommand {
  return {
    name: "import",
    description: commands.import.description,
    commands: [
      importAgentCommand,
      ...(featureFlagManager.getBooleanValue(FeatureFlags.OpenPluginImportExport)
        ? [importOpenPluginCommand]
        : []),
    ],
  };
}
