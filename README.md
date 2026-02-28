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
- Companion VB language server scaffold (`server/VSExtensionForVB.LanguageServer`)

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
- `VSExtensionForVB: Check Language Client Bridge Compatibility`

When you run the parity status command, it probes provider availability in parallel (with a 5 s timeout per probe) for:

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
- `vsextensionforvb.enableBridgeForCSharp`
- `vsextensionforvb.enableBridgeForVisualBasic`
- `vsextensionforvb.languageClientServerCommand`
- `vsextensionforvb.languageClientServerArgs`
- `vsextensionforvb.languageClientTraceLevel`
- `vsextensionforvb.autoBootstrapRoslynBridge`

## Language-Client Bridge Scaffold

Automatic bootstrap is enabled by default. After install, the extension attempts to configure a working bridge automatically on first activation **only when no compatible server command is already set**. Existing user configuration is never overwritten.

The extension prefers its bundled companion VB server and configures:

- `languageClientServerCommand = dotnet`
- `languageClientServerArgs = [<bundled-server-dll>, --stdio]`
- `enableLanguageClientBridge = true`
- `enableBridgeForVisualBasic = true`
- `enableBridgeForCSharp = false`

This is designed to be install-and-go (no manual preset command required).

If no bundled server is available in a dev scenario, bootstrap falls back to the companion server project (`dotnet run --project ... -- --stdio`) when present.

If the configured server points to the bundled C# extension Roslyn executable, or if the path does not exist on disk, the extension auto-disables bridge startup to prevent LSP error loops. The Linux Roslyn executable (no `.exe` suffix) is also detected and blocked correctly.

Manual override (optional):

1. Set `vsextensionforvb.enableLanguageClientBridge` to `true`.
2. Provide a server command in `vsextensionforvb.languageClientServerCommand`.
3. Optionally add arguments in `vsextensionforvb.languageClientServerArgs`.

When configured, the extension starts a `vscode-languageclient` instance targeting both `csharp` and `vb` documents.

`VSExtensionForVB: Apply Roslyn Bridge Preset` provides a guided setup flow that:

- Prompts for a local Roslyn server executable path
- Seeds common bridge args (`--stdio` by default)
- Enables the bridge and sets trace level to `messages`
- Restarts the bridge from updated settings

`VSExtensionForVB: Check Language Client Bridge Compatibility` validates the configured bridge server path and offers quick actions when the setup is incompatible.

The companion server source lives in:

- `server/VSExtensionForVB.LanguageServer`

## Development

Install dependencies (already run during scaffolding):

- `npm install`

Build:

- `npm run compile`

Watch mode:

- `npm run watch`

Run tests:

- `npm test`

## Quick Manual Test Workspace

Use `test-workspace` for a fast manual check of parity commands.

Included files:

- `test-workspace/Sample.cs`
- `test-workspace/Sample.vb`

Manual flow:

1. Press `F5` to launch an Extension Development Host.
2. In that new window, open the `test-workspace` folder.
3. Run `VSExtensionForVB: Show .NET Language Parity Status`.
4. Run `VSExtensionForVB: Remediate VB.NET Parity Gaps`.
5. Inspect output channel `VSExtensionForVB` and status bar parity state.

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
