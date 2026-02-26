import * as fs from 'node:fs';
import * as path from 'node:path';
import * as vscode from 'vscode';
import { buildParityReport, getParityGaps, runDotnetParityProbe } from './parityProbe';
import { assessBridgeServerCompatibility, DotnetLanguageClientBridge } from './languageClientBridge';

let languageClientBridge: DotnetLanguageClientBridge | undefined;
const ROSLYN_BRIDGE_BOOTSTRAP_KEY = 'vsextensionforvb.roslynBridgeBootstrapCompleted';

export function activate(context: vscode.ExtensionContext) {
	const outputChannel = vscode.window.createOutputChannel('VSExtensionForVB');
	languageClientBridge = new DotnetLanguageClientBridge(outputChannel);
	const parityStatusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
	parityStatusBarItem.name = 'VSExtensionForVB Parity';
	parityStatusBarItem.command = 'vsextensionforvb.remediateParityGaps';
	parityStatusBarItem.tooltip = 'Click to run VB.NET parity remediation';
	parityStatusBarItem.text = '$(sync~spin) VB parity: checking...';
	parityStatusBarItem.show();
	context.subscriptions.push(outputChannel);
	context.subscriptions.push(parityStatusBarItem);
	let refreshTimer: NodeJS.Timeout | undefined;
	let refreshInFlight = false;
	let refreshQueued = false;
	void promptForMissingDotnetTooling();
	void autoBootstrapRoslynBridge(context, outputChannel)
		.finally(() => languageClientBridge?.startFromConfiguration());
	scheduleParityStatusRefresh(0);

	const showParityStatusCommand = vscode.commands.registerCommand('vsextensionforvb.showParityStatus', async () => {
		const csharpExt = vscode.extensions.getExtension('ms-dotnettools.csdevkit') ?? vscode.extensions.getExtension('ms-dotnettools.csharp');
		const csharpState = csharpExt ? 'Installed' : 'Missing';

		const vbFiles = await vscode.workspace.findFiles('**/*.{vb,vbproj}', '**/{bin,obj,node_modules}/**', 1);
		const vbContext = vbFiles.length > 0 ? 'Detected' : 'Not detected in workspace';
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const enableForCSharp = config.get<boolean>('enableForCSharp', true);
		const enableForVisualBasic = config.get<boolean>('enableForVisualBasic', true);

		const probeSummaries = await runDotnetParityProbe(enableForCSharp, enableForVisualBasic);
		const report = buildParityReport(probeSummaries);

		const message = [
			'.NET Language Parity Snapshot',
			`C# language tooling: ${csharpState}`,
			`VB.NET workspace context: ${vbContext}`,
			`Probe targets: ${enableForCSharp ? 'C# ' : ''}${enableForVisualBasic ? 'VB.NET' : ''}`.trim(),
			'Detailed parity report written to output channel: VSExtensionForVB.'
		].join('\n');

		outputChannel.clear();
		outputChannel.appendLine(report);
		outputChannel.show(true);
		await executeParityStatusRefresh();
		void vscode.window.showInformationMessage(message);
	});

	const remediateParityGapsCommand = vscode.commands.registerCommand('vsextensionforvb.remediateParityGaps', async () => {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const enableForCSharp = config.get<boolean>('enableForCSharp', true);
		const enableForVisualBasic = config.get<boolean>('enableForVisualBasic', true);

		const probeSummaries = await runDotnetParityProbe(enableForCSharp, enableForVisualBasic);
		const report = buildParityReport(probeSummaries);
		const parityGaps = getParityGaps(probeSummaries).filter((gap) => gap.direction === 'vb_missing_vs_csharp');

		outputChannel.clear();
		outputChannel.appendLine(report);
		outputChannel.show(true);

		if (parityGaps.length === 0) {
			await executeParityStatusRefresh();
			void vscode.window.showInformationMessage('No VB.NET parity gaps were detected against C# in the latest probe run.');
			return;
		}

		const gapList = parityGaps.map((gap) => gap.feature).join(', ');
		const action = await vscode.window.showWarningMessage(
			`VB.NET parity gaps detected for: ${gapList}`,
			'Open C# Dev Kit',
			'Restart .NET Services',
			'Re-run Probe',
			'Open Settings'
		);

		switch (action) {
			case 'Open C# Dev Kit':
				await openExtensionOrSearch('ms-dotnettools.csdevkit');
				break;
			case 'Restart .NET Services':
				await vscode.commands.executeCommand('vsextensionforvb.restartDotnetLanguageServices');
				break;
			case 'Re-run Probe':
				await vscode.commands.executeCommand('vsextensionforvb.showParityStatus');
				break;
			case 'Open Settings':
				await vscode.commands.executeCommand('workbench.action.openSettings', 'vsextensionforvb');
				break;
			default:
				break;
		}

		await executeParityStatusRefresh();
	});

	const restartDotnetLanguageServicesCommand = vscode.commands.registerCommand('vsextensionforvb.restartDotnetLanguageServices', async () => {
		const candidates = [
			'csharp.restartServer',
			'roslyn.restartServer',
			'dotnet.restartServer'
		] as const;

		for (const command of candidates) {
			try {
				await vscode.commands.executeCommand(command);
				void vscode.window.showInformationMessage(`Executed language-service restart via command: ${command}`);
				return;
			} catch {
				continue;
			}
		}

		void vscode.window.showWarningMessage('No known .NET language-service restart command is currently available.');
	});

	const restartLanguageClientBridgeCommand = vscode.commands.registerCommand('vsextensionforvb.restartLanguageClientBridge', async () => {
		if (!languageClientBridge) {
			void vscode.window.showWarningMessage('Language client bridge is not initialized.');
			return;
		}

		await languageClientBridge.restartFromConfiguration();
		void vscode.window.showInformationMessage('Language client bridge restarted from current settings.');
	});

	const applyRoslynBridgePresetCommand = vscode.commands.registerCommand('vsextensionforvb.applyRoslynBridgePreset', async () => {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const currentCommand = (config.get<string>('languageClientServerCommand', '') ?? '').trim();
		const defaultRoslynPath = process.platform === 'win32'
			? 'C:\\tools\\roslyn\\Microsoft.CodeAnalysis.LanguageServer.exe'
			: '/usr/local/bin/roslyn-language-server';

		const serverCommand = await vscode.window.showInputBox({
			prompt: 'Path to local Roslyn language server executable',
			value: currentCommand || defaultRoslynPath,
			ignoreFocusOut: true,
			validateInput: (value) => value.trim().length > 0 ? undefined : 'Server command path is required.'
		});

		if (!serverCommand) {
			return;
		}

		const argsText = await vscode.window.showInputBox({
			prompt: 'Optional Roslyn server arguments (supports quotes)',
			value: '--stdio',
			ignoreFocusOut: true
		});

		if (argsText === undefined) {
			return;
		}

		const parsedArgs = parseCommandLineArgs(argsText);

		await config.update('enableLanguageClientBridge', true, vscode.ConfigurationTarget.Workspace);
		await config.update('enableBridgeForCSharp', false, vscode.ConfigurationTarget.Workspace);
		await config.update('enableBridgeForVisualBasic', false, vscode.ConfigurationTarget.Workspace);
		await config.update('languageClientServerCommand', serverCommand.trim(), vscode.ConfigurationTarget.Workspace);
		await config.update('languageClientServerArgs', parsedArgs, vscode.ConfigurationTarget.Workspace);
		await config.update('languageClientTraceLevel', 'messages', vscode.ConfigurationTarget.Workspace);
		await context.globalState.update(ROSLYN_BRIDGE_BOOTSTRAP_KEY, true);

		await languageClientBridge?.restartFromConfiguration();

		const action = await vscode.window.showInformationMessage(
			'Applied Roslyn bridge preset to workspace settings and restarted the bridge.',
			'Open Settings'
		);

		if (action === 'Open Settings') {
			await vscode.commands.executeCommand('workbench.action.openSettings', 'vsextensionforvb.languageClient');
		}
	});

	const checkLanguageClientBridgeCompatibilityCommand = vscode.commands.registerCommand('vsextensionforvb.checkLanguageClientBridgeCompatibility', async () => {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const serverCommand = (config.get<string>('languageClientServerCommand', '') ?? '').trim();
		const compatibility = assessBridgeServerCompatibility(serverCommand);

		if (compatibility.isCompatible) {
			void vscode.window.showInformationMessage(`Bridge compatibility check passed. ${compatibility.message}`);
			return;
		}

		const action = await vscode.window.showWarningMessage(
			`Bridge compatibility check failed: ${compatibility.message}`,
			'Open Settings',
			'Disable Bridge'
		);

		if (action === 'Open Settings') {
			await vscode.commands.executeCommand('workbench.action.openSettings', 'vsextensionforvb.languageClientServerCommand');
			return;
		}

		if (action === 'Disable Bridge') {
			const target = vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0
				? vscode.ConfigurationTarget.Workspace
				: vscode.ConfigurationTarget.Global;
			await config.update('enableLanguageClientBridge', false, target);
			void vscode.window.showInformationMessage('Disabled language client bridge due to incompatible configuration.');
		}
	});

	const showReferencesFromBridgeCommand = vscode.commands.registerCommand('vsextensionforvb.showReferencesFromBridge', async (rawUri: unknown, rawPosition: unknown, rawLocations: unknown) => {
		const uri = parseProtocolUri(rawUri);
		const position = parseProtocolPosition(rawPosition);
		const locations = parseProtocolLocations(rawLocations);

		if (!uri || !position) {
			void vscode.window.showWarningMessage('Unable to open references: invalid reference payload from language server.');
			return;
		}

		await vscode.commands.executeCommand('editor.action.showReferences', uri, position, locations);
	});

	context.subscriptions.push(showParityStatusCommand, remediateParityGapsCommand, restartDotnetLanguageServicesCommand, restartLanguageClientBridgeCommand, applyRoslynBridgePresetCommand, checkLanguageClientBridgeCompatibilityCommand, showReferencesFromBridgeCommand);
 	context.subscriptions.push({
		dispose: () => {
			if (refreshTimer) {
				clearTimeout(refreshTimer);
			}
		}
	});

	const triggerStatusRefresh = () => {
		scheduleParityStatusRefresh();
	};

	context.subscriptions.push(vscode.window.onDidChangeActiveTextEditor(triggerStatusRefresh));
	context.subscriptions.push(vscode.workspace.onDidOpenTextDocument(triggerStatusRefresh));
	context.subscriptions.push(vscode.workspace.onDidChangeConfiguration((event) => {
		if (event.affectsConfiguration('vsextensionforvb')) {
			triggerStatusRefresh();
			if (
				event.affectsConfiguration('vsextensionforvb.enableLanguageClientBridge') ||
				event.affectsConfiguration('vsextensionforvb.enableBridgeForCSharp') ||
				event.affectsConfiguration('vsextensionforvb.enableBridgeForVisualBasic') ||
				event.affectsConfiguration('vsextensionforvb.languageClientServerCommand') ||
				event.affectsConfiguration('vsextensionforvb.languageClientServerArgs') ||
				event.affectsConfiguration('vsextensionforvb.languageClientTraceLevel')
			) {
				void languageClientBridge?.restartFromConfiguration();
			}
		}
	}));

	function scheduleParityStatusRefresh(delayMs = 300): void {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const configuredDelay = config.get<number>('statusRefreshDelayMs', 300);
		const effectiveDelay = Math.min(5000, Math.max(50, Math.trunc(configuredDelay ?? 300)));
		const resolvedDelay = delayMs === 300 ? effectiveDelay : delayMs;

		if (refreshTimer) {
			clearTimeout(refreshTimer);
		}

		refreshTimer = setTimeout(() => {
			void executeParityStatusRefresh();
		}, resolvedDelay);
	}

	async function executeParityStatusRefresh(): Promise<void> {
		if (refreshInFlight) {
			refreshQueued = true;
			return;
		}

		refreshInFlight = true;
		try {
			await refreshParityStatusBar();
		} finally {
			refreshInFlight = false;
			if (refreshQueued) {
				refreshQueued = false;
				scheduleParityStatusRefresh(150);
			}
		}
	}

	async function refreshParityStatusBar(): Promise<void> {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const enableForCSharp = config.get<boolean>('enableForCSharp', true);
		const enableForVisualBasic = config.get<boolean>('enableForVisualBasic', true);
		const configuredDelay = config.get<number>('statusRefreshDelayMs', 300);
		const effectiveDelay = Math.min(5000, Math.max(50, Math.trunc(configuredDelay ?? 300)));
		const preferredLanguageServer = config.get<string>('preferredLanguageServer', 'auto');

		if (!enableForVisualBasic) {
			parityStatusBarItem.text = '$(circle-slash) VB parity: disabled';
			parityStatusBarItem.tooltip = `VB.NET probing is disabled in settings. Language server: ${preferredLanguageServer}. Refresh debounce: ${effectiveDelay}ms.`;
			return;
		}

		parityStatusBarItem.text = '$(sync~spin) VB parity: checking...';
		parityStatusBarItem.tooltip = `Running parity probe... Language server: ${preferredLanguageServer}. Refresh debounce: ${effectiveDelay}ms.`;

		try {
			const summaries = await runDotnetParityProbe(enableForCSharp, enableForVisualBasic);
			const vbGaps = getParityGaps(summaries).filter((gap) => gap.direction === 'vb_missing_vs_csharp');
			if (vbGaps.length === 0) {
				parityStatusBarItem.text = '$(check) VB parity: OK';
				parityStatusBarItem.tooltip = `No VB.NET parity gaps detected against C# in latest probe. Language server: ${preferredLanguageServer}. Refresh debounce: ${effectiveDelay}ms.`;
				return;
			}

			parityStatusBarItem.text = `$(warning) VB parity: ${vbGaps.length} gap${vbGaps.length === 1 ? '' : 's'}`;
			parityStatusBarItem.tooltip = `Detected VB.NET parity gaps: ${vbGaps.map((gap) => gap.feature).join(', ')}. Click to remediate. Language server: ${preferredLanguageServer}. Refresh debounce: ${effectiveDelay}ms.`;
		} catch (error) {
			const detail = error instanceof Error ? error.message : String(error);
			parityStatusBarItem.text = '$(error) VB parity: error';
			parityStatusBarItem.tooltip = `Parity probe failed: ${detail}. Language server: ${preferredLanguageServer}. Refresh debounce: ${effectiveDelay}ms.`;
		}
	}

	async function promptForMissingDotnetTooling(): Promise<void> {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		if (!config.get<boolean>('autoCheckToolingOnStartup', true)) {
			return;
		}
		if (!config.get<boolean>('promptToInstallMissingTooling', true)) {
			return;
		}

		const hasCsDevKit = !!vscode.extensions.getExtension('ms-dotnettools.csdevkit');
		const hasCSharp = !!vscode.extensions.getExtension('ms-dotnettools.csharp');
		if (hasCsDevKit || hasCSharp) {
			return;
		}

		const action = await vscode.window.showWarningMessage(
			'No Microsoft .NET language extension is installed. Install one to enable navigation, IntelliSense, and refactoring parity checks.',
			'Open C# Dev Kit',
			'Open C#',
			'Disable Prompt'
		);

		if (action === 'Open C# Dev Kit') {
			await openExtensionOrSearch('ms-dotnettools.csdevkit');
			return;
		}

		if (action === 'Open C#') {
			await openExtensionOrSearch('ms-dotnettools.csharp');
			return;
		}

		if (action === 'Disable Prompt') {
			await config.update('promptToInstallMissingTooling', false, vscode.ConfigurationTarget.Global);
		}
	}
}

export async function deactivate(): Promise<void> {
	await languageClientBridge?.stop();
	languageClientBridge = undefined;
}

async function openExtensionOrSearch(extensionId: string): Promise<void> {
	try {
		await vscode.commands.executeCommand('extension.open', extensionId);
	} catch {
		await vscode.commands.executeCommand('workbench.extensions.search', `@id:${extensionId}`);
	}
}

export function parseCommandLineArgs(input: string): string[] {
	const args: string[] = [];
	let current = '';
	let quote: '"' | "'" | undefined;
	let escaping = false;

	for (let index = 0; index < input.length; index++) {
		const character = input[index];
		const nextCharacter = index + 1 < input.length ? input[index + 1] : undefined;

		if (escaping) {
			current += character;
			escaping = false;
			continue;
		}

		if (character === '\\') {
			const canEscapeInQuote = !!quote && (nextCharacter === quote || nextCharacter === '\\');
			const canEscapeOutsideQuote = !quote && !!nextCharacter && (/\s/.test(nextCharacter) || nextCharacter === '"' || nextCharacter === "'" || nextCharacter === '\\');

			if (canEscapeInQuote || canEscapeOutsideQuote) {
				escaping = true;
			} else {
				current += character;
			}
			continue;
		}

		if (quote) {
			if (character === quote) {
				quote = undefined;
			} else {
				current += character;
			}
			continue;
		}

		if (character === '"' || character === "'") {
			quote = character;
			continue;
		}

		if (/\s/.test(character)) {
			if (current.length > 0) {
				args.push(current);
				current = '';
			}
			continue;
		}

		current += character;
	}

	if (escaping) {
		current += '\\';
	}

	if (current.length > 0) {
		args.push(current);
	}

	return args;
}

async function autoBootstrapRoslynBridge(context: vscode.ExtensionContext, outputChannel: vscode.OutputChannel): Promise<void> {
	const config = vscode.workspace.getConfiguration('vsextensionforvb');
	if (!config.get<boolean>('autoBootstrapRoslynBridge', true)) {
		return;
	}

	const companionServerLaunch = detectCompanionServerLaunch(context);
	const existingServerCommand = (config.get<string>('languageClientServerCommand', '') ?? '').trim();
	const existingServerArgs = config.get<string[]>('languageClientServerArgs', []) ?? [];
	const existingCompatibility = assessBridgeServerCompatibility(existingServerCommand);
	const bridgeEnabled = config.get<boolean>('enableLanguageClientBridge', false);
	const bridgeForVisualBasic = config.get<boolean>('enableBridgeForVisualBasic', false);
	const bridgeForCSharp = config.get<boolean>('enableBridgeForCSharp', false);

	if (companionServerLaunch && shouldApplyCompanionBootstrap(
		existingServerCommand,
		existingServerArgs,
		existingCompatibility.isCompatible,
		bridgeEnabled,
		bridgeForVisualBasic,
		bridgeForCSharp,
		companionServerLaunch
	)) {
		const target = vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0
			? vscode.ConfigurationTarget.Workspace
			: vscode.ConfigurationTarget.Global;

		await config.update('enableLanguageClientBridge', true, target);
		await config.update('enableBridgeForCSharp', false, target);
		await config.update('enableBridgeForVisualBasic', true, target);
		await config.update('languageClientServerCommand', companionServerLaunch.command, target);
		await config.update('languageClientServerArgs', companionServerLaunch.args, target);
		await context.globalState.update(ROSLYN_BRIDGE_BOOTSTRAP_KEY, true);

		outputChannel.appendLine(`[Bridge] Applied companion VB bridge profile: ${companionServerLaunch.description}`);
		void vscode.window.showInformationMessage('VSExtensionForVB automatically enabled VB bridge support.');
		return;
	}

	const alreadyBootstrapped = context.globalState.get<boolean>(ROSLYN_BRIDGE_BOOTSTRAP_KEY, false);
	if (alreadyBootstrapped) {
		return;
	}

	if (existingServerCommand.length > 0) {
		await context.globalState.update(ROSLYN_BRIDGE_BOOTSTRAP_KEY, true);
		return;
	}

	const target = vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0
		? vscode.ConfigurationTarget.Workspace
		: vscode.ConfigurationTarget.Global;

	const detectedServerCommand = detectRoslynServerCommandPath();
	if (!detectedServerCommand) {
		outputChannel.appendLine('[Bridge] Automatic bootstrap skipped: no compatible bridge server detected.');
		return;
	}

	await config.update('enableLanguageClientBridge', true, target);
	await config.update('enableBridgeForCSharp', false, target);
	await config.update('enableBridgeForVisualBasic', false, target);
	await config.update('languageClientServerCommand', detectedServerCommand, target);
	await config.update('languageClientServerArgs', ['--stdio'], target);
	await context.globalState.update(ROSLYN_BRIDGE_BOOTSTRAP_KEY, true);

	outputChannel.appendLine(`[Bridge] Automatically configured Roslyn bridge with: ${detectedServerCommand}`);
	void vscode.window.showInformationMessage('VSExtensionForVB auto-configured bridge settings.');
}

function detectRoslynServerCommandPath(): string | undefined {
	const csharpExtension = vscode.extensions.getExtension('ms-dotnettools.csharp');
	if (!csharpExtension) {
		return undefined;
	}

	const candidates = [
		path.join(csharpExtension.extensionPath, '.roslyn', 'Microsoft.CodeAnalysis.LanguageServer.exe'),
		path.join(csharpExtension.extensionPath, '.roslyn', 'Microsoft.CodeAnalysis.LanguageServer')
	];

	for (const candidate of candidates) {
		if (fs.existsSync(candidate)) {
			return candidate;
		}
	}

	return undefined;
}

function shouldApplyCompanionBootstrap(
	existingServerCommand: string,
	existingServerArgs: string[],
	existingServerCompatible: boolean,
	bridgeEnabled: boolean,
	bridgeForVisualBasic: boolean,
	bridgeForCSharp: boolean,
	companionServerLaunch: CompanionServerLaunch
): boolean {
	if (!existingServerCommand || !existingServerCompatible) {
		return true;
	}

	if (existingServerCommand !== companionServerLaunch.command) {
		return true;
	}

	if (!areArgsEqual(existingServerArgs, companionServerLaunch.args)) {
		return true;
	}

	if (!bridgeEnabled || !bridgeForVisualBasic || bridgeForCSharp) {
		return true;
	}

	return false;
}

function areArgsEqual(left: string[], right: string[]): boolean {
	if (left.length !== right.length) {
		return false;
	}

	for (let index = 0; index < left.length; index++) {
		if (left[index] !== right[index]) {
			return false;
		}
	}

	return true;
}

function parseProtocolUri(value: unknown): vscode.Uri | undefined {
	if (value instanceof vscode.Uri) {
		return value;
	}

	if (typeof value === 'string' && value.trim().length > 0) {
		try {
			return vscode.Uri.parse(value);
		} catch {
			return undefined;
		}
	}

	return undefined;
}

function parseProtocolPosition(value: unknown): vscode.Position | undefined {
	if (!value || typeof value !== 'object') {
		return undefined;
	}

	const lineValue = Reflect.get(value, 'line');
	const characterValue = Reflect.get(value, 'character');
	if (typeof lineValue !== 'number' || typeof characterValue !== 'number') {
		return undefined;
	}

	return new vscode.Position(lineValue, characterValue);
}

function parseProtocolRange(value: unknown): vscode.Range | undefined {
	if (!value || typeof value !== 'object') {
		return undefined;
	}

	const start = parseProtocolPosition(Reflect.get(value, 'start'));
	const end = parseProtocolPosition(Reflect.get(value, 'end'));
	if (!start || !end) {
		return undefined;
	}

	return new vscode.Range(start, end);
}

function parseProtocolLocations(value: unknown): vscode.Location[] {
	if (!Array.isArray(value)) {
		return [];
	}

	const locations: vscode.Location[] = [];
	for (const item of value) {
		if (!item || typeof item !== 'object') {
			continue;
		}

		const uri = parseProtocolUri(Reflect.get(item, 'uri'));
		const range = parseProtocolRange(Reflect.get(item, 'range'));
		if (!uri || !range) {
			continue;
		}

		locations.push(new vscode.Location(uri, range));
	}

	return locations;
}

function detectCompanionServerProjectPath(): string | undefined {
	const workspaceFolderPath = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
	if (!workspaceFolderPath) {
		return undefined;
	}

	const projectPath = path.join(workspaceFolderPath, 'server', 'VSExtensionForVB.LanguageServer', 'VSExtensionForVB.LanguageServer.csproj');
	if (fs.existsSync(projectPath)) {
		return projectPath;
	}

	return undefined;
}

type CompanionServerLaunch = {
	command: string;
	args: string[];
	description: string;
};

function detectCompanionServerLaunch(context: vscode.ExtensionContext): CompanionServerLaunch | undefined {
	const serverRelativePath = path.join('server', 'VSExtensionForVB.LanguageServer');
	const serverDllName = 'VSExtensionForVB.LanguageServer.dll';
	const serverExeName = 'VSExtensionForVB.LanguageServer.exe';
	const candidateExePaths = [
		path.join(context.extensionPath, serverRelativePath, 'publish', serverExeName),
		path.join(context.extensionPath, serverRelativePath, 'bin', 'Release', 'net8.0', serverExeName),
		path.join(context.extensionPath, serverRelativePath, 'bin', 'Debug', 'net8.0', serverExeName)
	];

	if (process.platform === 'win32') {
		for (const candidate of candidateExePaths) {
			if (fs.existsSync(candidate)) {
				const siblingDllPath = candidate.replace(/\.exe$/i, '.dll');
				if (fs.existsSync(siblingDllPath)) {
					try {
						fs.rmSync(siblingDllPath);
					} catch {
						// Keep launch resilient even if cleanup fails.
					}
				}

				return {
					command: candidate,
					args: ['--stdio'],
					description: candidate
				};
			}
		}
	}

	const candidateDllPaths = [
		path.join(context.extensionPath, serverRelativePath, 'publish', serverDllName),
		path.join(context.extensionPath, serverRelativePath, 'bin', 'Release', 'net8.0', serverDllName),
		path.join(context.extensionPath, serverRelativePath, 'bin', 'Debug', 'net8.0', serverDllName)
	];

	for (const candidate of candidateDllPaths) {
		if (fs.existsSync(candidate)) {
			return {
				command: 'dotnet',
				args: [candidate, '--stdio'],
				description: candidate
			};
		}
	}

	const projectPath = detectCompanionServerProjectPath();
	if (projectPath) {
		return {
			command: 'dotnet',
			args: ['run', '--project', projectPath, '--', '--stdio'],
			description: projectPath
		};
	}

	return undefined;
}
