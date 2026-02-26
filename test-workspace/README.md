# Test Workspace

This folder contains both lightweight single-file samples and a multi-project C#/VB interop solution.

Files:
- `Sample.cs`
- `Sample.vb`
- `InteropSample.slnx`
- `CSharpLib/`
- `VbConsumer/`

Use this folder in an Extension Development Host window to run:
- `VSExtensionForVB: Show .NET Language Parity Status`
- `VSExtensionForVB: Remediate VB.NET Parity Gaps`

## Cross-project C# ↔ VB check

1. Open `InteropSample.slnx`.
2. Open `VbConsumer/Class1.vb`.
3. Run `F12` on `FormatMessage` and `Add` calls.
4. Verify navigation goes to `CSharpLib/Class1.cs`.
5. Run Find References on `GreeterService.FormatMessage` to verify references from VB are included.
