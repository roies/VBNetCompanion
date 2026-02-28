# Change Log

All notable changes to the "vbnet-companion" extension will be documented in this file.

Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.

## [Unreleased]

- No pending unreleased changes.

## [0.1.11] - 2026-02-28

### Fixed

- **Hover and Signature Help were not activating**: the `hoverProvider`, `documentSymbolProvider`, `documentHighlightProvider`, and `signatureHelpProvider` capabilities were accidentally placed as siblings of `capabilities` in the `initialize` response rather than inside it. VS Code never received them, so it never sent the corresponding requests.
- **Hover position mismatch with live edits**: hover now recomputes the character offset against the live in-memory text buffer instead of the on-disk snapshot, so it correctly resolves symbols while you are editing.

## [0.1.10] - 2026-02-28

### Fixed

- **Server crash on concurrent F12 to BCL types** (e.g. `Console.WriteLine`): two simultaneous Go-to-Definition requests racing to write the same metadata stub file (`System.Console.vb`) caused an `IOException` that crashed the server process. The stub is now written only when the file does not yet exist, with the `IOException` swallowed for the concurrent-write case.
- **Corrupt LSP stream / server disconnect**: `PushDiagnosticsAsync` ran as a background `Task.Run` task and wrote to `stdout` concurrently with the main request-handling loop, interleaving `Content-Length` headers with JSON payload bytes. All stdout writes are now serialized through a `SemaphoreSlim outputGate` gate, eliminating the race entirely.

## [0.1.9] - 2026-02-28

### Added

- **Hover documentation** (`textDocument/hover`): hovering over any symbol shows its type signature in a code fence, plus XML doc `<summary>` text and `<param>` descriptions rendered as Markdown.
- **Push diagnostics** (`textDocument/publishDiagnostics`): compiler errors and warnings from Roslyn are pushed to the editor whenever a file is opened or changed.
- **Document symbols** (`textDocument/documentSymbol`): the Outline panel and breadcrumbs now list all classes, methods, properties, fields, enums, and other declared symbols in the current VB.NET or C# file.
- **Document highlights** (`textDocument/documentHighlight`): pressing or clicking a symbol highlights all other occurrences of that symbol in the same file via Roslyn `SymbolFinder.FindReferencesAsync`.
- **Signature help** (`textDocument/signatureHelp`): typing `(` or `,` shows the parameter list for the active method or constructor overload, including the active parameter and XML doc summary.

## [0.1.8] - 2026-02-28

### Added

- **Semantic token colorization for VB.NET**: types, methods, properties, enums, interfaces, namespaces, and other symbols in `.vb` files are now colored using the same VS Code theme colors as C#. Uses Roslyn's `Classifier.GetClassifiedSpansAsync` to produce accurate per-token type information including static and readonly modifiers.

## [0.1.7] - 2026-02-28

### Fixed

- **F12 on .NET BCL types (e.g. `ConsoleColor`, `Console`, `String`) navigates to nowhere**: when a symbol has only metadata locations (no source file in the solution), the server now generates a readable VB.NET stub in a temp directory and navigates there. The stub shows the type's public members (fields/enum values, properties, methods, events) with their signatures and XML doc summaries, formatted as valid VB.NET syntax.

## [0.1.6] - 2026-02-28

### Fixed

- **IntelliSense missing for .NET types and member access**: completion now uses Roslyn's `Recommender.GetRecommendedSymbolsAtPositionAsync` to produce context-aware suggestions. Typing `ConsoleColor.` now shows `Cyan`, `Red`, `Green`, etc. with their types; typing in any expression context shows all in-scope symbols (locals, parameters, imported types, .NET framework types). The previous implementation only returned 20 hardcoded VB keywords plus symbols declared in the current file.

## [0.1.5] - 2026-02-28

### Fixed

- **Extension updates never took effect**: the auto-bootstrap only ran once per workspace. After the initial setup it stored a `roies.vbnet-companion-0.1.0` server path in workspace settings and never updated it, so all fixes in v0.1.1–v0.1.4 were silently bypassed. The bootstrap now detects when the configured server path points to an older companion extension version and automatically redirects to the current extension's server binary.

## [0.1.4] - 2026-02-28

### Fixed

- **Go to Definition fails for C#-defined types from VB.NET files**: pressing F12 on a C# type (e.g. `DataAnalyzer`, `StringHelper`) from a consuming VB.NET file now correctly navigates to the C# source. In cross-language P2P workspaces, Roslyn may surface the target type as a metadata symbol without an in-source location. A new `FindDeclarationsAsync` fallback scans the full solution by symbol name and kind so the source file is always found.
- **Silent failure when VB file is not tracked in Roslyn solution**: `textDocument/definition` requests on files not yet indexed by the Roslyn workspace now emit a diagnostic log entry listing the loaded projects, making the root cause immediately visible in the output channel.

## [0.1.3] - 2026-02-28

### Fixed

- **Duplicate project load errors on `.slnx` workspaces**: `OpenProjectAsync` automatically loads transitive project references. Subsequent attempts to open those same projects explicitly threw `'X' is already part of the workspace`. Projects now check `workspace.CurrentSolution` before opening and skip any already loaded as transitive dependencies.
- **`roslynSolution` missing transitively-loaded projects**: solution is now always synced from `workspace.CurrentSolution` (which includes all transitive loads) rather than from the individual `project.Solution` returned per open call.

## [0.1.2] - 2026-02-28

### Fixed

- **Go to Definition on type names navigates to comment line**: pressing F12 on a type name (e.g. `DataAnalyzer`) no longer lands on a comment containing that word. Comment lines (`'`, `//`, `/*`) are now skipped before matching.
- **Go to Definition on static method navigates to wrong file**: pressing F12 on `DataAnalyzer.CalculateStatistics` now navigates to `DataAnalyzer.cs` instead of `Program.cs`. The receiver token (`DataAnalyzer`) is now used as a type-filter fallback even when the receiver is a static class rather than a local variable.
- **`IsMethodDeclarationLine` matching call sites instead of declarations**: added a negative lookbehind `(?<!\.)` to the C# method pattern so `DataAnalyzer.CalculateStatistics(` call sites are never mistaken for declarations. Added class/struct/interface/enum/module type declaration patterns so F12 on a type name finds the class declaration.

## [0.1.1] - 2026-02-28

### Fixed

- **Go to Definition (F12) on implicit constructors**: pressing F12 on `New GreeterService()` in VB.NET where `GreeterService` has no explicit constructor now correctly navigates to the class declaration in C#. Previously the implicit Roslyn constructor symbol had no source location and the request returned nothing silently.
- Added diagnostic logging to the definition handler so failures are now visible in the output channel.

### Changed

- README rewritten with marketplace-facing messaging: highlights the parity-first and cross-language differentiators, adds Getting Started flow, Commands and Configuration tables, and removes internal development notes.

## [0.1.0] - 2026-02-28

### Added

- **Cross-project CodeLens reference counting** via bundled Roslyn language server (`VBNetCompanion.LanguageServer`).
  - Loads `.sln` and `.slnx` solution files to resolve references across VB.NET and C# projects.
  - Uses `SymbolFinder.FindReferencesAsync` against the full Roslyn solution for accurate counts.
  - Falls back to single-file text search (scoped to current file only) when no project context is available.
- Roslyn workspace loads `.slnx` (XML format) by parsing it directly with `XDocument`, since `OpenSolutionAsync` only supports `.sln` text format.
- MSBuild auto-registration via `dotnet --list-sdks` when `MSBuildLocator.RegisterDefaults()` is unavailable (required for deployed executables).

### Fixed

- `Assembly.Location` returning empty string inside single-file bundles caused `BuildHostProcessManager.GetNetCoreBuildHostPath()` to crash. Fixed by publishing as a standard directory deployment (DLLs as loose files) instead of a single-file bundle.
- Text-fallback CodeLens was scanning all open documents, producing false cross-file matches for same-named symbols (e.g. `GetMessage` appearing in unrelated C# and VB files). Scoped to current file only.
- C# methods were not found by the VB-only keyword regex in `FindDeclarations`. Fixed by using the Roslyn semantic model (`GetDeclaredSymbol`) for symbol enumeration when project context is available.

### Changed

- Extension renamed from `VSExtensionForVB` to `VB.NET Companion`.
- All command IDs updated from `vsextensionforvb.*` to `vbnetcompanion.*`.
- All configuration keys updated from `vsextensionforvb.*` to `vbnetcompanion.*`.
- Version promoted from `0.1.0-beta` to `0.1.0` (stable release).

## [0.1.0-beta] - 2026-02-26

This beta delivers the first end-to-end parity-focused foundation for VB.NET alongside C# in VS Code.

### Added

- Project scaffold with TypeScript + esbuild build pipeline.
- C# vs VB.NET parity probing for:
	- Go to Definition (`F12`)
	- IntelliSense/completions
	- References
	- Rename
	- Code actions
- Guided remediation command for detected VB.NET parity gaps.
- Startup checks and prompts for missing .NET tooling.
- Live status bar parity indicator with configurable debounce.
- Optional language-client bridge scaffold with Roslyn preset support.
- Command-line argument parsing with quoted/escaped argument support for bridge setup.
- Unit tests for parser behavior.

### Changed

- Initial beta release (`0.1.0-beta`).

### Notes

- This beta adds parity diagnostics and bridge scaffolding; it does not replace the underlying .NET language server.
- C# support remains provided by Microsoft .NET tooling while this extension focuses on narrowing VB.NET parity gaps.