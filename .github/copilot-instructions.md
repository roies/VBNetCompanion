<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->
- [x] Verify that the copilot-instructions.md file in the .github directory is created.

- [x] Clarify Project Requirements

- [x] Scaffold the Project

- [x] Customize the Project

- [x] Install Required Extensions

- [x] Compile the Project

- [x] Create and Run Task

- [x] Launch the Project

- [x] Ensure Documentation is Complete

- Work through each checklist item systematically.
- Keep communication concise and focused.
- Follow development best practices.

## Critical Rule — Never Modify the User's Project

This extension must **NEVER** modify, write, or create any files in the user's workspace or project directory. It must not alter the user's build system, project files (.csproj, .vbproj, .sln), or any on-disk configuration.

Specifically, you must **never**:
- Write override `.props` or `.targets` files and inject them via `CustomBeforeMicrosoftCommonProps` / `CustomAfterMicrosoftCommonTargets` environment variables.
- Set `Environment.SetEnvironmentVariable()` with MSBuild property names (e.g., `RuntimeIdentifier`, `EnableNETAnalyzers`, `BuildProjectReferences`) that alter how the user's projects are evaluated.
- Create stub directories or files (e.g., VSToolsPath stubs) that intercept the user's MSBuild imports.
- Modify, create, or delete any file under the user's workspace root.

**Allowed approaches** for configuring MSBuild design-time evaluation:
- Pass properties through the `MSBuildWorkspace.Create(properties)` dictionary — these are Roslyn-internal and process-scoped.
- Write temporary files only to `%TEMP%` for extension-internal purposes (e.g., metadata decompilation stubs for Go-To-Definition) that do NOT affect the user's build.
- Set `DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER=1` (process-scoped, does not affect the user's build system).
