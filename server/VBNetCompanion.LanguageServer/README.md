# VSExtensionForVB Companion Language Server

This is a lightweight companion LSP server used by the VS Code extension bridge.

Current scope:
- `textDocument/definition` for simple VB symbol lookup in open documents
- `textDocument/completion` with VB keywords and in-document symbols

Run manually:
- `dotnet run --project server/VSExtensionForVB.LanguageServer/VSExtensionForVB.LanguageServer.csproj -- --stdio`
