# VBNetCompanion Language Server

This is the Roslyn-powered LSP server bundled with the VB.NET Companion VS Code extension. It provides full IDE-grade language support for VB.NET (and C#) across an entire solution, using [Microsoft.CodeAnalysis (Roslyn)](https://github.com/dotnet/roslyn) via `MSBuildWorkspace`.

---

## Architecture

- **Transport:** stdin/stdout (`--stdio` flag), JSON-RPC framing per the LSP spec.
- **Workspace:** `MSBuildWorkspace.Create()` loads `.sln` or `.slnx` solution files. `.slnx` (modern XML format) is parsed via `XDocument` because Roslyn's `OpenSolutionAsync` only supports the classic `.sln` format.
- **Live edits:** Document changes are tracked in an in-memory dictionary and applied to Roslyn documents via `.WithText()` so all features reflect unsaved edits.
- **Fallback:** Every feature has a text/regex-based fallback that activates when the Roslyn workspace fails to load, keeping the server functional (at reduced accuracy) on unsupported project types.
- **Thread safety:** An `outputGate` semaphore serializes all stdout writes to prevent LSP stream corruption from concurrent request handlers.
- **Cross-language:** All features work on both `.vb` and `.cs` files within the same solution.

---

## LSP Capabilities

### Navigation
| Capability | Method | Notes |
|---|---|---|
| Go to Definition | `textDocument/definition` | Roslyn semantic lookup; generates readable metadata stub files for binary-only types |
| Find All References | `textDocument/references` | `SymbolFinder.FindReferencesAsync` across entire solution; text-based single-file fallback |
| Go to Implementation | `textDocument/implementation` | Resolves virtual/abstract members to concrete implementations |
| Document Highlights | `textDocument/documentHighlight` | All occurrences of symbol in current file |
| Workspace Symbol Search | `workspace/symbol` | `SymbolFinder.FindDeclarationsAsync`; capped at 200 results |

### IntelliSense
| Capability | Method | Notes |
|---|---|---|
| Completions | `textDocument/completion` | `Recommender.GetRecommendedSymbolsAtPositionAsync`; keyword/symbol text fallback |
| Signature Help | `textDocument/signatureHelp` | Full overload list with active parameter highlighting; XML `<param>` docs per overload |
| Hover Documentation | `textDocument/hover` | XML doc summaries + parameter types from `symbol.GetDocumentationCommentXml()` |
| Inlay Hints | `textDocument/inlayHint` | Parameter names at call sites; suppressed for single-char params, `params` arrays, already-named args |

### Refactoring & Actions
| Capability | Method | Notes |
|---|---|---|
| Rename | `textDocument/rename` | `SymbolFinder.FindReferencesAsync` across entire solution; atomic cross-language renames |
| Code Actions | `textDocument/codeAction` | "Add XML doc comment", "Suppress warning" (`#Disable Warning`/`#pragma warning disable`), "Remove unused imports" |

### Analysis
| Capability | Method | Notes |
|---|---|---|
| Diagnostics (push) | `textDocument/publishDiagnostics` | `semanticModel.GetDiagnostics()` pushed on open and on change |
| CodeLens | `textDocument/codeLens` | Per-symbol reference counts; text-based fallback scoped to single file only |
| Semantic Tokens (full + delta) | `textDocument/semanticTokens/full`, `.../delta` | Roslyn `Classification.GetClassifiedSpansAsync`; delta-encoded |

### Structure
| Capability | Method | Notes |
|---|---|---|
| Document Symbols | `textDocument/documentSymbol` | All declarations enumerated via `semanticModel.GetDeclaredSymbol()` |
| Call Hierarchy | `callHierarchy/incomingCalls`, `callHierarchy/outgoingCalls` | Incoming via `SymbolFinder`; outgoing via method body AST walk |
| Type Hierarchy | `typeHierarchy/supertypes`, `typeHierarchy/subtypes` | Base class + interfaces (supertypes); derived classes/implementations (subtypes) |
| Folding Ranges | `textDocument/foldingRange` | Detects `Block`, `Declaration`, and `Statement` syntax node kinds |
| Selection Ranges | `textDocument/selectionRange` | Walks the syntax node ancestor chain |

### Formatting & Links
| Capability | Method | Notes |
|---|---|---|
| Document Formatting | `textDocument/formatting` | `Roslyn.Formatter.FormatAsync` |
| Range Formatting | `textDocument/rangeFormatting` | Same formatter scoped to selection |
| Document Links | `textDocument/documentLink` | Scans comment and string trivia for `http://` / `https://` URLs via regex |

---

## Metadata Stubs

For binary-only types (e.g. `System.String`, third-party NuGet types), the server generates a temporary `.vb` or `.cs` stub file in `%TEMP%` containing:
- Public member signatures
- XML doc summaries extracted from metadata
- Formatted as valid syntax so Go to Definition navigates to a readable file

These stubs are **never written to the user's workspace** — they are extension-internal scratch files only.

---

## Run Manually

```bash
dotnet run --project server/VBNetCompanion.LanguageServer/VBNetCompanion.LanguageServer.csproj -- --stdio
```

Or, after publishing:

```bash
dotnet server/publish/VBNetCompanion.LanguageServer.dll --stdio
```
