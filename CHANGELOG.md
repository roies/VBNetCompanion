# Change Log

All notable changes to the "vbnet-companion" extension will be documented in this file.

Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.

## [Unreleased]

- No pending unreleased changes.

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