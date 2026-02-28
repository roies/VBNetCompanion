# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest  | ✅        |

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub Issues.**

To report a security vulnerability, open a
[GitHub Security Advisory](https://github.com/roies/VBNetCompanion/security/advisories/new)
for this repository. This keeps the report private until a fix is available.

Include as much of the following as possible:

- A description of the vulnerability and its potential impact
- Steps to reproduce or a proof-of-concept
- The version of VB.NET Companion affected
- Any relevant logs or screenshots

You can expect an acknowledgement within **48 hours** and a resolution or
status update within **7 days**.

## Scope

VB.NET Companion is a VS Code extension that bundles a .NET 8 language server
binary (`VBNetCompanion.LanguageServer`). Relevant security areas include:

- The bundled Roslyn language server process
- Extension activation and file system access
- Any external process spawning (`languageClientServerCommand`)
- Dependency vulnerabilities in npm or NuGet packages
