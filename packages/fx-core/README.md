# TeamsFx Core

The TeamsFx Core package implements shared capabilities for Microsoft 365 Agents Toolkit IDE Extensions and the CLI through API contracts defined in the [api](/packages/api).

## Local agent package import and edits

The supported `FxCoreClient` / `IFxCoreClient` boundary exposes two local,
noninteractive operations:

```typescript
const imported = await client.importAgentPackage({
  sourcePath: "exports/agent.zip",
  outputPath: "projects/imported-agent",
  dryRun: true,
});

const edited = await client.applyAgentEdits({
  projectPath: "projects/imported-agent",
  changesFile: "approved/changes.json",
  expectedDigest: reviewedProjectDigest,
});
```

Both return `Result<AgentMigrationReport, FxError>` and accept an optional second
argument `{ signal: AbortSignal }`. Requests, the edit-document schema, and report
types are exported from `@microsoft/teamsfx-api`. ZIP and extracted-folder imports
use the bundled native template and create a **new** project; they never overwrite
an existing directory or take over the source deployment identity.

Edits preserve the current project's deployment/environment identity and apply
only the approved version-1 JSON operations. A dry run changes no project files;
a no-op does not rewrite them. Import and edits do not authenticate, provision,
share, publish, run source code, fetch remote source references, or call AI.
Local structural validity is reported separately from cloud readiness, which is
not evaluated.

See the [import contract](../../docs/03-specs/operations/scaffolding/import-agent-package.md),
[edit contract](../../docs/03-specs/operations/scaffolding/apply-agent-edits.md),
and [workflow](../../docs/03-specs/scenarios/agent-package/import-and-edit.md) for
supported dialects, safety limits, exact collection semantics, and transaction
recovery. Imported instructions use the shared manifest resolver's opt-in
`$[file('instructions.txt', 'raw')]` syntax so literal template-like prose remains
literal when packaged.

The migration profile is independently pinned at
`resource/agent-import/6.16.0/`; it does not read or modify normal creation's
template archive or version configuration. Maintainers can verify the frozen
profile with `pnpm run bundle:agent-import-profile` after building the matching
native template release and core. A content change requires a new profile
version; the command refuses to overwrite an existing different profile.

## Data Collection.

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at https://go.microsoft.com/fwlink/?LinkID=824704. You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)

If you are interested in fixing issues and contributing directly to the code base, please see the [Contributing Guide](./CONTRIBUTING.md).

## Reporting Security Issues

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please report them to the Microsoft Security Response Center (MSRC) at [https://msrc.microsoft.com/create-report](https://msrc.microsoft.com/create-report).

If you prefer to submit without logging in, send email to [secure@microsoft.com](mailto:secure@microsoft.com). If possible, encrypt your message with our PGP key; please download it from the the [Microsoft Security Response Center PGP Key page](https://www.microsoft.com/en-us/msrc/pgp-key-msrc).

You should receive a response within 24 hours. If for some reason you do not, please follow up via email to ensure we received your original message. Additional information can be found at [microsoft.com/msrc](https://www.microsoft.com/msrc).

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE.txt) license.
