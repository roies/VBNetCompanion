import * as vscode from 'vscode';
import type { LanguageClient } from 'vscode-languageclient/node.js';

type LanguageClientModule = typeof import('vscode-languageclient/node.js');

export type BridgeCompatibility = {
	isCompatible: boolean;
	message: string;
	hint?: string;
};

export class DotnetLanguageClientBridge {
	private client: LanguageClient | undefined;

	constructor(private readonly outputChannel: vscode.OutputChannel) {}

	public async startFromConfiguration(): Promise<void> {
		const config = vscode.workspace.getConfiguration('vsextensionforvb');
		const enabled = config.get<boolean>('enableLanguageClientBridge', false);
		if (!enabled) {
			this.outputChannel.appendLine('[Bridge] Language client bridge is disabled.');
			return;
		}

		const command = (config.get<string>('languageClientServerCommand', '') ?? '').trim();
		if (!command) {
			this.outputChannel.appendLine('[Bridge] Enabled, but no server command is configured. Running in scaffold mode.');
			return;
		}

		const compatibility = assessBridgeServerCompatibility(command);
		if (!compatibility.isCompatible) {
			const target = vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0
				? vscode.ConfigurationTarget.Workspace
				: vscode.ConfigurationTarget.Global;
			await config.update('enableLanguageClientBridge', false, target);
			this.outputChannel.appendLine(`[Bridge] Disabled bridge startup: ${compatibility.message}`);
			void vscode.window.showWarningMessage(`VSExtensionForVB disabled the bridge: ${compatibility.message}`);
			return;
		}

		const args = config.get<string[]>('languageClientServerArgs', []) ?? [];
		const enableBridgeForCSharp = config.get<boolean>('enableBridgeForCSharp', false);
		const enableBridgeForVisualBasic = config.get<boolean>('enableBridgeForVisualBasic', false);
		const traceLevel = config.get<'off' | 'messages' | 'verbose'>('languageClientTraceLevel', 'off');
		const documentSelector: string[] = [];
		if (enableBridgeForCSharp) {
			documentSelector.push('csharp');
		}
		if (enableBridgeForVisualBasic) {
			documentSelector.push('vb');
		}
		if (documentSelector.length === 0) {
			this.outputChannel.appendLine('[Bridge] Bridge is enabled, but no bridge document languages are selected.');
			return;
		}

		const languageClientModule = await import('vscode-languageclient/node.js') as LanguageClientModule;
		const { LanguageClient, Trace } = languageClientModule;

		const client = new LanguageClient(
			'vsextensionforvb.dotnetBridge',
			'VSExtensionForVB .NET Bridge',
			{
				command,
				args,
				transport: languageClientModule.TransportKind.stdio
			},
			{
				documentSelector,
				outputChannel: this.outputChannel,
				traceOutputChannel: this.outputChannel,
				synchronize: {
					configurationSection: ['vsextensionforvb']
				}
			}
		);

		client.setTrace(this.toTrace(traceLevel, Trace));
		this.client = client;
		await client.start();
		this.outputChannel.appendLine(`[Bridge] Started language client bridge using command: ${command}`);
	}

	public async restartFromConfiguration(): Promise<void> {
		await this.stop();
		await this.startFromConfiguration();
	}

	public async stop(): Promise<void> {
		if (!this.client) {
			return;
		}

		await this.client.stop();
		this.client = undefined;
		this.outputChannel.appendLine('[Bridge] Stopped language client bridge.');
	}

	private toTrace(level: 'off' | 'messages' | 'verbose', traceEnum: { Off: number; Messages: number; Verbose: number }): number {
		switch (level) {
			case 'messages':
				return traceEnum.Messages;
			case 'verbose':
				return traceEnum.Verbose;
			case 'off':
			default:
				return traceEnum.Off;
		}
	}
}

export function assessBridgeServerCompatibility(commandPath: string): BridgeCompatibility {
	if (!commandPath.trim()) {
		return {
			isCompatible: false,
			message: 'No bridge server command is configured.',
			hint: 'Set vsextensionforvb.languageClientServerCommand to a compatible standalone LSP server.'
		};
	}

	if (isBundledCSharpRoslynPath(commandPath)) {
		return {
			isCompatible: false,
			message: 'Bundled C# Roslyn server path is not supported as a standalone bridge endpoint on this setup.',
			hint: 'Use a dedicated standalone server command and keep VB routing off unless server supports Visual Basic.'
		};
	}

	return {
		isCompatible: true,
		message: 'Configured server path passed basic compatibility checks.'
	};
}

function isBundledCSharpRoslynPath(commandPath: string): boolean {
	const normalized = commandPath.replace(/\\/g, '/').toLowerCase();
	return normalized.includes('/.vscode/extensions/ms-dotnettools.csharp-') && normalized.endsWith('/microsoft.codeanalysis.languageserver.exe');
}
