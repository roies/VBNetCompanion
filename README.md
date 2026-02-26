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

## Development

Install dependencies (already run during scaffolding):

- `npm install`

Build:

- `npm run compile`

Watch mode:

- `npm run watch`

Run tests:

- `npm test`

## Next Milestones

1. Add a language-client bridge for Roslyn/LSP capabilities relevant to VB.NET.
2. Implement feature capability checks for F12, IntelliSense, and refactoring by language.
3. Surface actionable diagnostics when VB.NET capability is below C# parity.
