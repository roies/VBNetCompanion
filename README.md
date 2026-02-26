# VSExtensionForVB

VSExtensionForVB is a TypeScript-based VS Code extension scaffold focused on improving .NET editing parity for VB.NET with C#.

## Objective

Deliver a Visual Studio-like experience in VS Code for both C# and VB.NET, with priority on:

- F12 navigation (Go to Definition/Implementation)
- IntelliSense and semantic completion quality
- Refactoring support (rename, extract, code actions)

## Current Scaffold Status

This initial scaffold includes:

- TypeScript extension project with esbuild bundling
- Extension manifest (`package.json`) with commands and configuration
- Activation entrypoint in `src/extension.ts`
- Initial parity-status command and language-service restart helper command
- Core feature probe module (`src/parityProbe.ts`) for C#/VB.NET parity checks

## Current Status (Feb 2026)

Implemented today:

- Project scaffold with TypeScript + esbuild
- Parity probe for C# vs VB.NET language feature availability
- Guided remediation command for VB.NET parity gaps
- Startup tooling checks and install prompts for required .NET extensions
- Live status bar indicator (`VB parity: OK` / `VB parity: N gaps`)
- Configurable status refresh debounce (`vsextensionforvb.statusRefreshDelayMs`)
- Optional language-client bridge scaffold for .NET LSP integration

## Current Support Boundary

This extension currently adds VB.NET parity measurement and diagnostics (feature probing and gap reporting).

It does not replace the underlying .NET language server. Actual editor intelligence (Go to Definition, IntelliSense, refactorings) is still provided by installed .NET tooling (for example C# Dev Kit / C# extension stack).

In short:

- C# language intelligence already exists in VS Code via Microsoft tooling.
- This project adds parity-focused support for VB.NET by detecting and surfacing where VB behavior does or does not match C#.

## Commands

- `VSExtensionForVB: Show .NET Language Parity Status`
- `VSExtensionForVB: Remediate VB.NET Parity Gaps`
- `VSExtensionForVB: Restart .NET Language Services`
- `VSExtensionForVB: Restart Language Client Bridge`
- `VSExtensionForVB: Apply Roslyn Bridge Preset`

When you run the parity status command, it now probes provider availability for:

- Definition (`F12` behavior)
- Completion (IntelliSense)
- References
- Rename refactoring
- Code actions

Detailed results are emitted to the output channel `VSExtensionForVB`.

The remediation command runs the same probe, identifies VB.NET gaps relative to C#, and provides guided actions:

- Open recommended .NET tooling extension
- Restart language services
- Re-run probe
- Open extension settings

The extension also shows a live status bar indicator:

- `VB parity: OK` when no VB gaps are detected
- `VB parity: N gaps` when VB is behind C#
- Click the indicator to open the remediation flow

## Configuration

- `vsextensionforvb.enableForCSharp`
- `vsextensionforvb.enableForVisualBasic`
- `vsextensionforvb.preferredLanguageServer`
- `vsextensionforvb.autoCheckToolingOnStartup`
- `vsextensionforvb.promptToInstallMissingTooling`
- `vsextensionforvb.statusRefreshDelayMs`
- `vsextensionforvb.enableLanguageClientBridge`
- `vsextensionforvb.languageClientServerCommand`
- `vsextensionforvb.languageClientServerArgs`
- `vsextensionforvb.languageClientTraceLevel`

## Language-Client Bridge Scaffold

The bridge is intentionally opt-in and disabled by default.

To enable it:

1. Set `vsextensionforvb.enableLanguageClientBridge` to `true`.
2. Provide a server command in `vsextensionforvb.languageClientServerCommand`.
3. Optionally add arguments in `vsextensionforvb.languageClientServerArgs`.

When configured, the extension starts a `vscode-languageclient` instance targeting both `csharp` and `vb` documents.

`VSExtensionForVB: Apply Roslyn Bridge Preset` provides a guided setup flow that:

- Prompts for a local Roslyn server executable path
- Seeds common bridge args (`--stdio` by default)
- Enables the bridge and sets trace level to `messages`
- Restarts the bridge from updated settings

## Development

Install dependencies (already run during scaffolding):

- `npm install`

Build:

- `npm run compile`

Watch mode:

- `npm run watch`

Run tests:

- `npm test`

## Troubleshooting

### Test host says “Code is currently being updated”

If `npm test` fails with that message on Windows, a VS Code updater process is usually still running.

Quick fix:

1. Close all VS Code windows.
2. Wait for update/setup processes to complete (or end `CodeSetup-stable-*` in Task Manager).
3. Run `npm test` again.

## Next Milestones

1. Add richer parity diagnostics that map each failing feature to concrete remediation guidance.
2. Introduce telemetry/metrics hooks (opt-in) to track parity trend over time across workspaces.
3. Expand test coverage with integration-style probes for C# and VB.NET documents.
