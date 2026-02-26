# Change Log

All notable changes to the "vsextensionforvb" extension will be documented in this file.

Check [Keep a Changelog](http://keepachangelog.com/) for recommendations on how to structure this file.

## [Unreleased]

- No pending unreleased changes.

## [0.1.0-beta] - 2026-02-26

- Scaffolded TypeScript VS Code extension with esbuild build pipeline.
- Added .NET parity probe for C# and VB.NET feature checks (definition, completion, references, rename, code actions).
- Added guided remediation flow for VB.NET parity gaps.
- Added startup tooling checks and install prompts for missing .NET extensions.
- Added live status bar parity indicator with configurable refresh debounce.
- Added optional language-client bridge scaffold with restart and Roslyn preset commands.
- Added command-line argument parser with quote/escape support for bridge preset arguments.
- Added parser-focused unit tests and validated test suite pass.