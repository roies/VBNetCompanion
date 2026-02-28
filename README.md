# VB.NET Companion

> **Full IDE-grade language support for VB.NET in VS Code** — powered by a bundled Roslyn server that understands your whole solution, VB.NET *and* C# together.

VB.NET has been a second-class citizen in VS Code for too long. VB.NET Companion fixes that by shipping a complete language server with **18 LSP features** including IntelliSense, hover docs, rename, call hierarchy, inlay hints, and more — all working cross-language across your entire solution.

---

## ✨ Features at a Glance

| Feature | Shortcut / Trigger |
|---|---|
| IntelliSense completions | Typing `.` or space |
| Hover documentation | Hover over any symbol |
| Go to Definition | `F12` |
| Find All References | `Shift+F12` |
| Rename symbol | `F2` |
| Signature help | `(` or `,` |
| Code actions & quick fixes | `Ctrl+.` / lightbulb |
| Inlay hints | Inline parameter names |
| Go to Implementation | `Ctrl+F12` |
| Call hierarchy | Right-click → Peek Call Hierarchy |
| Workspace symbol search | `Ctrl+T` |
| Cross-project CodeLens | Reference counts above every symbol |
| Document symbols (Outline) | Explorer → Outline panel |
| Document highlights | Click any symbol |
| Folding ranges | Gutter fold markers |
| Semantic token colors | Full syntax highlighting |
| Live diagnostics | On open and on change |
| VB.NET parity report | Status bar indicator |

---

## 🔍 IntelliSense & Completions

Full member-access completions powered by Roslyn — including cross-language members from C# projects referenced by your VB.NET code. Completions trigger on `.` and space, with accurate type-aware filtering.

- VB.NET → C# members and vice versa
- Extension methods, generic types, overloads
- Works with unsaved in-memory edits

---

## 📖 Hover Documentation

Hover over any symbol to see its full signature and XML doc summary in a markdown popup. Works for:

- VB.NET and C# symbols in the same solution
- Framework types (navigates to the metadata stub)
- Properties, fields, parameters, enum members

---

## 🚀 Go to Definition (`F12`)

Press `F12` on any symbol and navigate to its source — including cross-language and cross-project jumps:

- VB.NET → C# types, methods, properties, constructors
- C# → VB.NET types
- Framework symbols → generated metadata stub file

---

## 🔗 Find All References (`Shift+F12`)

`SymbolFinder.FindReferencesAsync` across every project in the solution. Reference results include declaration site, all usages, and cross-language hits.

---

## ✏️ Rename (`F2`)

Rename any symbol — type, method, property, field, local variable, or parameter — and every reference across the entire solution is updated atomically. Cross-language renames included.

---

## 🖊️ Signature Help

Displays the full overload list when you type `(` or `,` inside a call. Shows:
- Active parameter highlighted
- All overloads listed
- XML `<summary>` and `<param>` documentation per overload

---

## 💡 Code Actions & Quick Fixes (`Ctrl+.`)

Click the lightbulb or press `Ctrl+.` to see context-sensitive actions:

- **Add XML doc comment** — inserts a fully-formed `''' <summary>` / `<param>` / `<returns>` block (or `///` for C#) with correct indentation and parameter names, on any undocumented declaration
- **Suppress warning** — inserts `#Disable Warning BCXXXXX` (VB) or `#pragma warning disable` (C#) above the offending line for any diagnostic at the cursor
- **Remove unused imports** — detects `Imports` / `using` statements flagged by the compiler and removes them all at once

---

## 🏷️ Inlay Hints

Parameter name labels appear inline at call sites so you always know what each argument means without having to hover:

```vb
ProcessText(input: rawData, maxLength: 100, trim: True)
'           ^^^^^^           ^^^^^^^^^       ^^^^   ← inlay hints
```

Hints are automatically suppressed for:
- Single-character parameter names
- `params` arrays
- Already-named arguments (VB `:=` / C# `name:`)
- Arguments whose variable name matches the parameter name

---

## 📍 Go to Implementation (`Ctrl+F12`)

Navigate from an interface member or abstract method to its concrete implementation(s). Also works on concrete methods — navigates to the method's own declaration when no override exists. Cross-project implementations are resolved via Roslyn's source definition lookup.

---

## 🌳 Call Hierarchy

Right-click any method or property → **Peek Call Hierarchy** to explore:

- **Incoming calls** — every method in the solution that calls this one, grouped by caller with the exact call-site ranges highlighted
- **Outgoing calls** — every method that *this* body calls, including cross-language and cross-project callees

---

## 🔎 Workspace Symbol Search (`Ctrl+T`)

Type any symbol name to search across the entire solution instantly. Finds types, methods, properties, fields, events, and namespaces across all VB.NET and C# projects. Results are deduplicated and capped at 200 for performance.

---

## 🔢 Cross-Project CodeLens

Accurate reference counts displayed above every class, method, and property — counting references that cross language and project boundaries:

- C# method called from a VB.NET project → correct count
- VB.NET class used from C# → correctly included

---

## 📐 Document Outline & Symbols

The Explorer **Outline** panel and `Ctrl+Shift+O` symbol picker are fully populated with all declarations in the current file — classes, modules, methods, properties, fields, events, enums, and interfaces.

---

## 🎨 Semantic Token Colors

Full semantic highlighting powered by Roslyn's classifier: namespaces, classes, interfaces, structs, enums, type parameters, methods, properties, fields, variables, parameters, and enum members — each styled distinctly by your theme.

---

## 📊 VB.NET Parity Status

VB.NET Companion is the **only VS Code extension** that actively measures the feature gap between C# and VB.NET in your workspace:

- `VB parity: OK` — both languages on equal footing
- `VB parity: N gaps` — N features available in C# are missing for VB.NET

Click the status bar item to run the full parity probe and launch the guided remediation flow.

---

## ⚡ Zero-Config Setup

On first activation the extension automatically detects your solution file (`.sln`, `.slnx`) and configures the bundled Roslyn bridge. No manual setup required — install and start coding.

---

## Getting Started

1. Install the extension.
2. Open a folder containing a `.sln`, `.slnx`, or `.vbproj` file.
3. The Roslyn bridge starts automatically and loads all projects.
4. Open any `.vb` file — all features activate immediately.

> **Platform:** Windows (win-x64). The bundled companion server is a .NET 8 win-x64 binary. macOS and Linux support is planned.

---

## Commands

| Command | Description |
|---|---|
| `VB.NET Companion: Show .NET Language Parity Status` | Probe and display the C# vs VB.NET feature parity report |
| `VB.NET Companion: Remediate VB.NET Parity Gaps` | Run the guided fix flow for detected gaps |
| `VB.NET Companion: Restart .NET Language Services` | Restart all .NET language services |
| `VB.NET Companion: Restart Language Client Bridge` | Restart the companion Roslyn bridge |
| `VB.NET Companion: Apply Roslyn Bridge Preset` | Guided setup for a custom Roslyn server path |
| `VB.NET Companion: Check Language Client Bridge Compatibility` | Validate the configured bridge setup |

---

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

---

## Requirements

- VS Code 1.109.0 or later
- .NET SDK installed and on `PATH`
- [C# Dev Kit](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.csdevkit) (recommended — used for C# side of parity checks)

---

## Release Notes

See [CHANGELOG](CHANGELOG.md) for the full history.
