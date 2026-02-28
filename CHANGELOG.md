# Change Log

All notable changes to the "vbnet-companion" extension will be documented in this file.

Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.

## [Unreleased]

- No pending unreleased changes.

## [0.1.0] - 2026-02-28

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