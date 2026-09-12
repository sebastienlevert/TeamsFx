// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DeclarativeAgentManifestLatest, TeamsManifestLatest } from "@microsoft/app-manifest";
export { default as agentEditDocumentSchema } from "./schemas/agentEdits.json";
export { default as agentMigrationReportSchema } from "./schemas/agentMigrationReport.json";

export interface AgentImportRequest {
  sourcePath: string;
  outputPath?: string;
  dryRun?: boolean;
}

export interface AgentEditRequest {
  projectPath: string;
  changesFile: string;
  dryRun?: boolean;
  /** Strict precondition, including for dry runs and no-ops. */
  expectedDigest?: string;
}

export interface AgentAppMetadata {
  name?: Partial<TeamsManifestLatest["name"]>;
  description?: Partial<TeamsManifestLatest["description"]>;
  developer?: Partial<TeamsManifestLatest["developer"]>;
  version?: TeamsManifestLatest["version"];
  accentColor?: TeamsManifestLatest["accentColor"];
  validDomains?: TeamsManifestLatest["validDomains"];
}

export type AgentEditOperation =
  | { kind: "replaceInstructions"; sourceFile: string }
  | {
      kind: "setAppMetadata";
      value: AgentAppMetadata;
    }
  | {
      kind: "setAgentMetadata";
      value: Partial<
        Pick<
          DeclarativeAgentManifestLatest,
          "name" | "description" | "disclaimer" | "sensitivity_label"
        >
      >;
    }
  | {
      kind: "replaceConversationStarters";
      value: NonNullable<DeclarativeAgentManifestLatest["conversation_starters"]>;
    }
  | {
      kind: "upsertCapability";
      value: NonNullable<DeclarativeAgentManifestLatest["capabilities"]>[number];
    }
  | { kind: "removeCapability"; name: string }
  | {
      kind: "setBehaviorOverrides";
      value: NonNullable<DeclarativeAgentManifestLatest["behavior_overrides"]>;
    }
  | { kind: "setSchemaVersion"; value: string }
  | { kind: "replaceIcon"; icon: "color" | "outline"; sourceFile: string }
  | {
      kind: "attachEmbeddedKnowledge";
      files: Array<{ sourceFile: string; targetPath: string }>;
    };

export interface AgentEditDocument {
  schemaVersion: 1;
  agentId: string;
  operations: AgentEditOperation[];
}

export interface AgentMigrationDiagnostic {
  code: string;
  severity: "info" | "warning";
  message: string;
  path?: string;
  pointer?: string;
  ruleId?: string;
  remediation?: string;
}

export interface AgentMigrationFile {
  path: string;
  action: "created" | "modified" | "preserved" | "skipped" | "candidate";
  sourcePath?: string;
  digest?: string;
}

export interface AgentTemplateIdentity {
  id: string;
  version: string;
  digest: string;
}

export interface AgentMigrationIdentity {
  policy: "new" | "preserved";
  provisionedByOperation: false;
  agentId: string;
  projectId: string | null;
  appId: string;
  sourceAppId?: string;
  sourceAgentId?: string;
}

export interface AgentMigrationReport {
  reportVersion: 1;
  projectPath: string;
  operationMode: "imported" | "edited" | "no-op" | "dry-run";
  dryRun: boolean;
  changed: boolean;
  source: { kind: "zip" | "directory" | "project"; digest: string };
  projectDigest: string;
  identity: AgentMigrationIdentity;
  template: AgentTemplateIdentity | null;
  transformations: Array<{ ruleId: string; path: string; description: string }>;
  files: AgentMigrationFile[];
  referenceMap: Record<string, string>;
  diagnostics: AgentMigrationDiagnostic[];
  configuration: {
    structurallyValid: true;
    readyToProvision: "not-evaluated";
    readyToPublish: "not-evaluated";
    requirements: string[];
  };
  recoveryRequired: false;
}
