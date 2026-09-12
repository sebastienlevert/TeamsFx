// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { CLICommand } from "@microsoft/teamsfx-api";

export function isLocalAgentInvocation(args: string[]): boolean {
  return (args[0] === "import" || args[0] === "edit") && args[1] === "agent";
}

export function isLocalAgentCommand(command: CLICommand): boolean {
  return isLocalAgentInvocation(command.fullName?.split(" ").slice(-2) ?? []);
}

export function requestsLocalAgentJson(args: string[]): boolean {
  let json = false;
  for (let index = 0; index < args.length; index++) {
    if (args[index] === "--format") {
      json = args[index + 1] === "json";
    } else if (args[index].startsWith("--format=")) {
      json = args[index] === "--format=json";
    }
  }
  return json;
}

export async function writeLocalAgentOutput(message: string): Promise<void> {
  await writeStream(process.stdout, message + "\n");
}

export async function flushLocalAgentOutput(): Promise<void> {
  await Promise.all([writeStream(process.stdout, ""), writeStream(process.stderr, "")]);
}

async function writeStream(stream: NodeJS.WriteStream, message: string): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    stream.write(message, (error) => {
      if (error) reject(error);
      else resolve();
    });
  });
}
