// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../common/localizeUtils";
import { UserCancelError } from "../../error";

export type MigrationErrorCode =
  | "AgentPackageSourceInvalid"
  | "AgentPackageUnsupported"
  | "AgentPackageLimitExceeded"
  | "AgentPackagePathInvalid"
  | "AgentPackageCollision"
  | "AgentPackageReferenceMissing"
  | "AgentPackageSchemaInvalid"
  | "AgentPackageIntegrityInvalid"
  | "AgentPackageDestinationExists"
  | "AgentEditsInvalid"
  | "AgentEditsAgentNotFound"
  | "AgentEditsStale"
  | "AgentEditsConflict"
  | "AgentMigrationIoError"
  | "AgentMigrationRecoveryRequired";

export function migrationError(code: MigrationErrorCode, cause?: unknown): FxError {
  const options = {
    source: "agent-migration",
    name: code,
    message: getLocalizedString(`error.agentMigration.${code}`),
    error: cause instanceof Error ? cause : undefined,
  };
  return code === "AgentMigrationIoError" || code === "AgentMigrationRecoveryRequired"
    ? new SystemError(options)
    : new UserError(options);
}

export function cancelled(signal?: AbortSignal): FxError | undefined {
  return signal?.aborted ? new UserCancelError("agent-migration") : undefined;
}

export function isErrno(error: unknown, code: string): boolean {
  return typeof error === "object" && error !== null && "code" in error && error.code === code;
}
