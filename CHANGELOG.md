# Change Log

All notable changes to the "vbnet-companion" extension will be documented in this file.

Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.

## [Unreleased]

- No pending unreleased changes.

## [0.1.30]

- **fix:** Probe documents no longer appear as visible tabs or trigger "Save?" confirm dialogs. Untitled docs are reverted before closing; real workspace files opened during probing are automatically cleaned up.
- **fix:** Explicitly clear `RuntimeIdentifier` and set `UseCurrentRuntimeIdentifier=false` in MSBuildWorkspace properties to prevent .NET 10+ SDK from injecting the host RID (`win`) into projects that were not restored with it.

## [0.1.29] - 2026-03-01

### Fixed

- **Massive project load failures with .NET 10 SDK (`RuntimeIdentifier` errors)**: `MSBuildWorkspace.Create()` was called without global properties, causing MSBuild to apply the host platform's default runtime evaluation. On .NET 10 SDK, this injected `RuntimeIdentifier` checks into old .NET Framework projects that don't list `win` in their `RuntimeIdentifiers`, producing hundreds of cascading failures ("Your project file doesn't list 'win' as a RuntimeIdentifier"). Now configures MSBuildWorkspace with design-time build properties (`DesignTimeBuild=true`, `BuildingInsideVisualStudio=true`, `SkipCompilerExecution=true`) matching what Visual Studio and OmniSharp use for IDE evaluation, producing a more forgiving and accurate project load.

## [0.1.28] - 2026-03-01

### Fixed

- **Synthetic `ProbeClass` document left open as a visible tab**: when no real workspace `.cs` or `.vb` file is found (or `findFiles` returns empty before workspace indexing completes), the parity probe creates an untitled in-memory document containing `ProbeClass`. This document was never closed, leaving a confusing untitled tab in the customer's editor. Now automatically closes synthetic (untitled) probe documents after probing completes.

## [0.1.27] - 2026-03-01

### Fixed

- **Bridge starts old server version after extension update (race condition)**: during bootstrap, each sequential `config.update()` fired `onDidChangeConfiguration` before subsequent settings (like the server command path) were written. The config handler's `restartFromConfiguration()` read the **old** `languageClientServerCommand` (e.g. v0.1.22 path) while the bootstrap was still writing the new path (v0.1.26). Added a `bridgeBootstrapPending` guard that suppresses config-change-triggered bridge restarts until bootstrap completes and all settings are written, then starts the bridge exactly once with the correct values.

## [0.1.26] - 2026-03-01

### Fixed

- **`enableBridgeForCSharp` always reset to `false`**: `autoBootstrapRoslynBridge()` forced `enableBridgeForCSharp = false` on every version-update re-bootstrap. Now preserves the existing value when the bridge was already configured, and defaults to `true` on first install.
- **VB parity probe false negatives on designer/generated files**: the probe searched for a single `.vb` file and could pick auto-generated files like `*.designer.vb` where no meaningful symbols exist. Now searches up to 20 candidates and filters out designer/generated files (`*.designer.*`, `*.generated.*`, `*.g.cs`, `*.g.vb`).
- **Parity probe position selection on real workspace files**: when the synthetic `localValue` identifier is absent, the probe now scans for VB/C# declaration keywords (`Sub`, `Function`, `Property`, `Class`, `Dim`, etc.) and positions on the identifier name, instead of falling back to the first non-whitespace character (which could be a comment or attribute).

## [0.1.25] - 2026-03-01

### Fixed

- **Document Symbol crash ("selectionRange must be contained in fullRange")**: the `textDocument/documentSymbol` handler computed `selectionRange` from `sym.Name.Length`, which could exceed the source span (e.g. for constructors where `sym.Name` is `.ctor` but source is `New`). Now uses the syntax-node span for `range` and the identifier source span for `selectionRange`, guaranteeing containment. Same fix applied to call hierarchy handlers.
- **MSBuild SDK resolver `MissingMethodException` with .NET 10**: the Roslyn/MSBuild packages (v4.13.0) were from the .NET 9 era and called `System.Text.Json.Utf8JsonReader` APIs that don't exist in .NET 10. Upgraded all Roslyn packages to **v5.0.0** and `Microsoft.Build.Locator` to **v1.11.2** for full .NET 10 SDK compatibility.
- **Deprecated `Workspace.WorkspaceFailed` event**: replaced with `RegisterWorkspaceFailedHandler` (Roslyn 5.0 API).

## [0.1.24] - 2026-03-01

### Fixed

- **Server crash on machines with only .NET 10+ runtime (exit code 1)**: the language server was published as framework-dependent targeting `net8.0` with no roll-forward policy. On machines without .NET 8 runtime, the exe failed immediately. Added `<RollForward>LatestMajor</RollForward>` to the project so the server runs on any .NET 8+ runtime.
- **SDK version detection now matches actual runtime**: `TryFindMSBuildPath()` used a hardcoded major version `8`. Now uses `Environment.Version.Major` to match the SDK to whichever runtime is actually executing, preventing MSBuild/runtime mismatches after roll-forward.

## [0.1.23] - 2026-03-01

### Fixed

- **MSBuild SDK version mismatch causing project load failures**: `TryFindMSBuildPath()` was selecting the newest SDK (e.g., .NET 10), but the language server targets net8.0. The .NET 10 SDK's `WorkloadManifestReader` calls APIs absent from .NET 8 (`MissingMethodException`), causing cascading project failures. Now prefers SDKs matching the server's major version (8.x), falling back to the latest only when necessary.
- **No user notification for workspace load failures**: hundreds of `WorkspaceDiag` warnings were logged but never surfaced. Added `SummarizeWorkspaceLoadIssuesAsync()` that categorises failures (SDK resolver errors, missing references, other) and sends a `window/showMessage` warning with actionable guidance.
- **VB.NET code actions returning nothing when projects fail to load**: when Roslyn can't resolve a document (project load failed), `HandleCodeActionAsync` now provides text-based fallback code actions ("Comment out line", "Wrap in region") instead of an empty response.

## [0.1.22] - 2026-02-28

### Fixed

- **Server crash on startup (exit code 1)**: workspace loading was blocking the `initialize` response. If loading took long enough for the client to time out, the client closed the connection; any subsequent `LogAsync` / `SendAsync` call then threw `IOException`, which escaped `EnsureRoslynWorkspaceLoadedAsync`'s catch block and crashed the process. Fixed by:
  1. Sending the `initialize` response immediately and loading the Roslyn workspace in a background `Task.Run`.
  2. Wrapping the `LogAsync` call inside `EnsureRoslynWorkspaceLoadedAsync`'s catch block with its own try-catch so a broken pipe cannot propagate.
  3. Wrapping the entire dispatch loop body in a top-level try-catch so any future unhandled exception logs to stderr and replies with a JSON-RPC error instead of crashing the server process.

## [0.1.21] - 2026-02-28

### Added

- **Document formatting** (`textDocument/formatting`, `textDocument/rangeFormatting`): Format Document and Format Selection now work for VB.NET and C# files via Roslyn's `Formatter.FormatAsync`. Changes are returned as minimal `TextEdit` diffs using `SourceText.GetTextChanges`.
- **Selection ranges** (`textDocument/selectionRange`): `Shift+Alt+→` (Expand Selection) now walks from the innermost syntax node to the file root through Roslyn's `SyntaxNode` ancestor chain, enabling incremental structural selection.
- **Document links** (`textDocument/documentLink`): URLs (`https://` and `http://`) in comments and string literals are surfaced as clickable links in the editor. Scanned via Roslyn trivia and string token walks; falls back to full-text regex scan when Roslyn is unavailable.
- **Type hierarchy** (`textDocument/prepareTypeHierarchy`, `typeHierarchy/supertypes`, `typeHierarchy/subtypes`): right-click a type → Show Type Hierarchy. Supertypes shows the base class (excluding `System.Object`) and implemented interfaces; subtypes uses `SymbolFinder.FindDerivedClassesAsync` / `FindImplementationsAsync` to discover subclasses and implementations across the whole solution (capped at 200).
- **VB.NET debug configuration** (`vbnet` debug type): contributes `Launch (VB.NET)` and `Attach (VB.NET)` configuration snippets. The `vbnet` type is rewritten to `coreclr` at resolve time so the C# extension's DAP adapter handles execution — no extra binary required.
- **Cross-platform support**: removed the `"os": ["win32"]` restriction. The extension now installs on Linux and macOS. The bundled Roslyn server `.dll` is launched via `dotnet <dll>` on non-Windows. Platform-specific publish paths (`publish/linux-x64`, `publish/osx-arm64`) are resolved first when present. Added `build:server:linux`, `build:server:osx`, and `build:server:all` npm scripts for producing platform-targeted builds.

## [0.1.20] - 2026-02-28

### Fixed

- Suppressed built-in language client error popup notifications ("couldn't create connection to server", "Server initialization failed", "Pending response rejected since connection got disposed") by adding a custom `errorHandler` and setting `revealOutputChannelOn: Never` on the bridge client. Errors are still logged to the output channel.

## [0.1.19] - 2026-03-01

### Changed

- Reduced log verbosity: removed all per-request success logs (SemanticTokens, Hover, Completion, CodeLens per-symbol, DocumentHighlight, SignatureHelp, FoldingRange, InlayHint, WorkspaceSymbol, CallHierarchy, etc.). Only errors, warnings, workspace loading, diagnostics summary, and intentional user actions (rename) are logged.
- Removed background popup notifications for auto-bootstrap and bridge restart actions; these are now silent (logged to output channel only).

## [0.1.18] - 2026-02-28

### Changed

- Expanded Marketplace keyword tags: added `vbnet`, `vb`, `basic`, `dotnet`, `c#`, `lsp`, `completions`, `hover`, `go to definition`, `find references`, `rename`, `inlay hints`, `call hierarchy`, `code actions`, `codelens`, `mixed language`, `cross-language`, `solution`.

## [0.1.17] - 2026-02-28

### Changed

- Rewrote README for VS Marketplace visibility: feature-first ordering by developer desirability, full capability table, per-feature sections with code examples, and updated Getting Started and Requirements.

## [0.1.16] - 2026-02-28

### Fixed

- **Code actions / lightbulb not appearing**: VS Code does not echo back diagnostics sourced from `textDocument/publishDiagnostics` in the `codeAction` request's `context.diagnostics`, so the suppress action was always empty. The suppress logic now reads directly from `semanticModel.GetDiagnostics()` filtered to the cursor line — no dependency on the client echoing diagnostics back.
- **Add XML doc comment**: new always-available `refactor.rewrite` action. When the cursor is on any class, method, sub, function, or property declaration that has no existing doc comment, the lightbulb offers “Add XML doc comment”. Inserts `''' <summary>` / `''' <param>` / `''' <returns>` (VB) or `/// ...` (C#) with correct indentation and parameter names from the symbol.

## [0.1.15] - 2026-02-28

### Added

- **Inlay hints** (`textDocument/inlayHint`): parameter-name hints appear inline at each call-site argument. Hints are suppressed for single-character parameter names, params arrays, already-named arguments (VB `:=` / C# `name:`), and arguments whose text trivially matches the parameter name.
- **Workspace symbol search** (`workspace/symbol`): Ctrl+T / ⌘+T now searches all symbols (types, methods, properties, fields, events, namespaces) across the full solution via `SymbolFinder.FindDeclarationsAsync`. Results are capped at 200 and deduplicated across projects.
- **Call hierarchy** (`textDocument/prepareCallHierarchy`, `callHierarchy/incomingCalls`, `callHierarchy/outgoingCalls`): right-click → Peek Call Hierarchy shows all callers of a method/property (incoming) and all methods called from its body (outgoing). Cross-project symbols are resolved via `FindSourceDefinitionAsync`.
- **Code actions / quick fixes** (`textDocument/codeAction`): two actions provided — “Remove N unused import(s)” (detects BC50001 / CS8019 diagnostics, `source.organizeImports` kind) and “Suppress [code]” which inserts a `#Disable Warning` (VB) or `#pragma warning disable` (C#) comment above the offending line.

## [0.1.14] - 2026-02-28

### Fixed

- **Ctrl+F12 still silent for cross-project concrete methods**: The v0.1.13 fallback in `HandleImplementationAsync` checked `symbol.Locations.Where(l => l.IsInSource)`, but cross-project symbols (e.g. a VB file calling a C# library method) have only metadata locations at that point. The fallback now mirrors `HandleDefinitionAsync`: if no in-source locations exist, it first calls `SymbolFinder.FindSourceDefinitionAsync` to get the real source symbol, then maps its locations. Ctrl+F12 on any concrete method — local or cross-project — now navigates to its definition.

## [0.1.13] - 2026-03-01

### Fixed

- **Duplicate hover / IntelliSense responses**: `startFromConfiguration` now stops any already-running `LanguageClient` before creating a new one. Previously, writing bootstrap settings triggered `onDidChangeConfiguration → restartFromConfiguration` (server #1) while the `.finally()` callback independently called `startFromConfiguration` (server #2), causing every provider to return results twice.
- **Ctrl+F12 (Go to Implementation) silent on concrete methods**: `HandleImplementationAsync` now falls back to the symbol's own source declaration(s) when both `FindImplementationsAsync` and `FindOverridesAsync` return empty. Concrete methods that are neither overrides nor interface implementations now navigate to their own definition instead of silently doing nothing.

## [0.1.12] - 2026-02-28

### Added

- **Rename** (`textDocument/rename`): rename any symbol (type, method, property, field, local, parameter) across all files in the solution using Roslyn `SymbolFinder.FindReferencesAsync`. Duplicate edits at the same location are deduplicated before returning the `WorkspaceEdit`.
- **Folding ranges** (`textDocument/foldingRange`): classes, modules, methods, loops, conditionals, try/catch, and other block constructs are foldable in the editor gutter. Works for both VB.NET and C# files by inspecting Roslyn syntax node class names.
- **Go to Implementation** (`textDocument/implementation`): navigates to concrete implementations of an interface member or virtual method via `SymbolFinder.FindImplementationsAsync`. Falls back to `FindOverridesAsync` when called on a virtual/abstract member that has no interface implementations.

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