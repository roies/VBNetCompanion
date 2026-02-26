import * as vscode from 'vscode';
import { buildParityReport, getParityGaps, runDotnetParityProbe } from './parityProbe';
import { DotnetLanguageClientBridge } from './languageClientBridge';

let languageClientBridge: DotnetLanguageClientBridge | undefined;

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
	void languageClientBridge.startFromConfiguration();
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
			prompt: 'Optional Roslyn server arguments (space-separated)',
			value: '--stdio',
			ignoreFocusOut: true
		});

		if (argsText === undefined) {
			return;
		}

		const parsedArgs = argsText
			.split(' ')
			.map((arg) => arg.trim())
			.filter((arg) => arg.length > 0);

		await config.update('enableLanguageClientBridge', true, vscode.ConfigurationTarget.Workspace);
		await config.update('languageClientServerCommand', serverCommand.trim(), vscode.ConfigurationTarget.Workspace);
		await config.update('languageClientServerArgs', parsedArgs, vscode.ConfigurationTarget.Workspace);
		await config.update('languageClientTraceLevel', 'messages', vscode.ConfigurationTarget.Workspace);

		await languageClientBridge?.restartFromConfiguration();

		const action = await vscode.window.showInformationMessage(
			'Applied Roslyn bridge preset to workspace settings and restarted the bridge.',
			'Open Settings'
		);

		if (action === 'Open Settings') {
			await vscode.commands.executeCommand('workbench.action.openSettings', 'vsextensionforvb.languageClient');
		}
	});

	context.subscriptions.push(showParityStatusCommand, remediateParityGapsCommand, restartDotnetLanguageServicesCommand, restartLanguageClientBridgeCommand, applyRoslynBridgePresetCommand);
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
