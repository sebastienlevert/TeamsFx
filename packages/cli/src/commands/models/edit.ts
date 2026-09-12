// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommand } from "@microsoft/teamsfx-api";
import { commands } from "../../resource";
import { editAgentCommand } from "./editAgent";

export const editCommand: CLICommand = {
  name: "edit",
  description: commands.edit.description,
  commands: [editAgentCommand],
};
