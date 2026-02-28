# VB.NET Companion

VB.NET Companion narrows the editor experience gap between C# and VB.NET in VS Code. It adds parity diagnostics, guided remediation, and a bundled Roslyn-powered language server that brings **cross-project reference counting (CodeLens)** and **Go to Definition** from VB.NET into C# projects and vice versa.

## Features

### Cross-Project Reference Counting (CodeLens)

The bundled Roslyn language server loads your solution or projects and shows accurate reference counts directly in the editor — including cross-project references between VB.NET and C# code.

- C# method defined in `CSharpLib` called from `VbConsumer.vb` → shows **1 reference**
- VB class referencing C# types → correctly counted
- Works for classes, methods, properties, and constructors

### Parity Probe

Probes language feature availability in parallel for both C# and VB.NET:

- Go to Definition (`F12`)
- IntelliSense / completions
- References
- Rename refactoring
- Code actions

Results surface in the output channel and status bar indicator (`VB parity: OK` / `VB parity: N gaps`).

### Guided Remediation

When parity gaps are detected, the remediation command provides targeted actions:

- Open recommended .NET tooling extension on the Marketplace
- Restart language services
- Re-run probe
- Open extension settings

### Status Bar Indicator

Live status bar item shows current parity health at a glance. Click to open the full remediation flow.

### Auto-Bootstrap Roslyn Bridge

On first activation, the extension automatically configures a working Roslyn bridge using the bundled companion server — no manual setup required.

## Current Support Boundary

This extension adds parity diagnostics and cross-project references via its bundled Roslyn server. It does not replace the underlying .NET language server for C# — that is still provided by C# Dev Kit / C# extension. VB.NET intelligence is delivered through this extension's bridge.

## Commands

- `VB.NET Companion: Show .NET Language Parity Status`
- `VB.NET Companion: Remediate VB.NET Parity Gaps`
- `VB.NET Companion: Restart .NET Language Services`
- `VB.NET Companion: Restart Language Client Bridge`
- `VB.NET Companion: Apply Roslyn Bridge Preset`
- `VB.NET Companion: Check Language Client Bridge Compatibility`

When you run the parity status command, it probes provider availability in parallel (with a 5 s timeout per probe) for:

- Definition (`F12` behavior)
- Completion (IntelliSense)
- References
- Rename refactoring
- Code actions

Detailed results are emitted to the output channel `VB.NET Companion`.

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

- `vbnetcompanion.enableForCSharp`
- `vbnetcompanion.enableForVisualBasic`
- `vbnetcompanion.preferredLanguageServer`
- `vbnetcompanion.autoCheckToolingOnStartup`
- `vbnetcompanion.promptToInstallMissingTooling`
- `vbnetcompanion.statusRefreshDelayMs`
- `vbnetcompanion.enableLanguageClientBridge`
- `vbnetcompanion.enableBridgeForCSharp`
- `vbnetcompanion.enableBridgeForVisualBasic`
- `vbnetcompanion.languageClientServerCommand`
- `vbnetcompanion.languageClientServerArgs`
- `vbnetcompanion.languageClientTraceLevel`
- `vbnetcompanion.autoBootstrapRoslynBridge`

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

1. Set `vbnetcompanion.enableLanguageClientBridge` to `true`.
2. Provide a server command in `vbnetcompanion.languageClientServerCommand`.
3. Optionally add arguments in `vbnetcompanion.languageClientServerArgs`.

When configured, the extension starts a `vscode-languageclient` instance targeting both `csharp` and `vb` documents.

`VB.NET Companion: Apply Roslyn Bridge Preset` provides a guided setup flow that:

- Prompts for a local Roslyn server executable path
- Seeds common bridge args (`--stdio` by default)
- Enables the bridge and sets trace level to `messages`
- Restarts the bridge from updated settings

`VB.NET Companion: Check Language Client Bridge Compatibility` validates the configured bridge server path and offers quick actions when the setup is incompatible.

The companion server source lives in:

- `server/VBNetCompanion.LanguageServer`

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
3. Run `VB.NET Companion: Show .NET Language Parity Status`.
4. Run `VB.NET Companion: Remediate VB.NET Parity Gaps`.
5. Inspect output channel `VB.NET Companion` and status bar parity state.

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
