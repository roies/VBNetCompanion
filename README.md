# VB.NET Companion

**VB.NET Companion** is the missing bridge for mixed-language .NET solutions in VS Code. Unlike standalone VB.NET language servers that replace your existing tooling, this extension works *alongside* C# Dev Kit to add the features that VB.NET has always been missing: **cross-language navigation**, **cross-project reference counts**, and **live parity diagnostics** that tell you exactly where VB.NET falls behind C#.

## Why VB.NET Companion?

Most VB.NET extensions for VS Code require you to throw out your existing C# setup and start over. VB.NET Companion is different:

- **No conflicts.** Installs alongside C# Dev Kit without touching your existing configuration.
- **Cross-language aware.** The bundled Roslyn server understands your whole solution — VB.NET *and* C# projects together.
- **Parity-first.** The only extension that actively tracks where VB.NET lags behind C# and guides you to fix it.

## Features

### Cross-Language Go to Definition (F12)

Press `F12` on any VB.NET symbol that resolves to a C# type and jump straight to its definition — including C# classes with implicit constructors (e.g. `New GreeterService()`).

Works across:
- VB.NET → C# types, methods, properties, constructors
- C# → VB.NET types

### Cross-Project Reference Counting (CodeLens)

The bundled Roslyn server loads your full solution and shows accurate reference counts in the editor — counting references that cross language and project boundaries.

- C# method called from a VB.NET project → shows the correct count
- VB.NET class used from C# → correctly included
- Works for classes, methods, properties, and constructors

### Live Parity Diagnostics

VB.NET Companion is the **only VS Code extension** that measures the live feature gap between C# and VB.NET in your workspace. It probes both languages in parallel and surfaces the results in a status bar indicator:

- `VB parity: OK` — VB.NET and C# are on equal footing
- `VB parity: N gaps` — N features available in C# are missing for VB.NET

Click the indicator to open the full remediation flow.

Probed features:
- Go to Definition (`F12`)
- IntelliSense / completions
- References
- Rename refactoring
- Code actions

### Guided Remediation

When parity gaps are detected, a targeted remediation command walks you through fixing them:

- Open the recommended .NET tooling extension on the Marketplace
- Restart language services
- Re-run the probe
- Open extension settings

### Zero-Config Setup

On first activation, the extension automatically bootstraps the Roslyn bridge using the bundled companion server. No manual configuration required — install and start coding.

## Getting Started

1. Install the extension.
2. Open a folder containing a `.sln`, `.slnx`, or `.vbproj` file.
3. The Roslyn bridge starts automatically.
4. Run **VB.NET Companion: Show .NET Language Parity Status** to see a full parity report.

> **Platform:** Windows (win-x64). The bundled companion server is a .NET 8 win-x64 binary. macOS and Linux support is planned.

## Commands

| Command | Description |
|---|---|
| `VB.NET Companion: Show .NET Language Parity Status` | Probe and display the C# vs VB.NET feature parity report |
| `VB.NET Companion: Remediate VB.NET Parity Gaps` | Run the guided fix flow for detected gaps |
| `VB.NET Companion: Restart .NET Language Services` | Restart all .NET language services |
| `VB.NET Companion: Restart Language Client Bridge` | Restart the companion Roslyn bridge |
| `VB.NET Companion: Apply Roslyn Bridge Preset` | Guided setup for a custom Roslyn server path |
| `VB.NET Companion: Check Language Client Bridge Compatibility` | Validate the configured bridge setup |

## Configuration

| Setting | Default | Description |
|---|---|---|
| `vbnetcompanion.enableForCSharp` | `true` | Enable parity checks for C# files |
| `vbnetcompanion.enableForVisualBasic` | `true` | Enable parity checks for VB.NET files |
| `vbnetcompanion.preferredLanguageServer` | `auto` | Preferred .NET language service (`auto`, `csharp-dev-kit`, `omnisharp`) |
| `vbnetcompanion.autoCheckToolingOnStartup` | `true` | Check for required tooling on startup |
| `vbnetcompanion.promptToInstallMissingTooling` | `true` | Show install prompts when tooling is missing |
| `vbnetcompanion.enableLanguageClientBridge` | `false` | Enable the Roslyn language client bridge |
| `vbnetcompanion.enableBridgeForCSharp` | `false` | Route C# documents through the bridge |
| `vbnetcompanion.enableBridgeForVisualBasic` | `true` | Route VB.NET documents through the bridge |
| `vbnetcompanion.autoBootstrapRoslynBridge` | `true` | Auto-configure the bridge on first activation |
| `vbnetcompanion.languageClientTraceLevel` | `off` | LSP trace verbosity (`off`, `messages`, `verbose`) |

## Requirements

- VS Code 1.109.0 or later
- [C# Dev Kit](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.csdevkit) or [C# extension](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.csharp) (recommended for C# side of parity checks)
- .NET SDK installed and on `PATH`
