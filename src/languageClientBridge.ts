import * as fs from 'node:fs';
import * as path from 'node:path';
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
	private isRestarting = false;

	constructor(private readonly outputChannel: vscode.OutputChannel) {}

	public async startFromConfiguration(): Promise<void> {
		const config = vscode.workspace.getConfiguration('vbnetcompanion');
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
			void vscode.window.showWarningMessage(`VB.NET Companion disabled the bridge: ${compatibility.message}`);
			return;
		}

		const configuredArgs = config.get<string[]>('languageClientServerArgs', []) ?? [];
		const resolvedLaunch = this.resolveBridgeLaunch(command, configuredArgs);
		const args = resolvedLaunch.args;
		const resolvedCommand = resolvedLaunch.command;
		if (resolvedLaunch.reason) {
			this.outputChannel.appendLine(`[Bridge] ${resolvedLaunch.reason}`);
		}
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
			'vbnetcompanion.dotnetBridge',
			'VB.NET Companion Bridge',
			{
				command: resolvedCommand,
				args,
				transport: languageClientModule.TransportKind.stdio
			},
			{
				documentSelector,
				outputChannel: this.outputChannel,
				traceOutputChannel: this.outputChannel,
				synchronize: {
					configurationSection: ['vbnetcompanion']
				}
			}
		);

		client.setTrace(this.toTrace(traceLevel, Trace));
		this.client = client;
		try {
			await client.start();
			this.outputChannel.appendLine(`[Bridge] Started language client bridge using command: ${resolvedCommand}`);
			return;
		} catch (error) {
			await this.stop();
			const detail = error instanceof Error ? error.message : String(error);
			this.outputChannel.appendLine(`[Bridge] Failed to start bridge: ${detail}`);

			const fallback = this.buildCompanionProjectFallback(resolvedCommand, args);
			if (!fallback) {
				throw error;
			}

			this.outputChannel.appendLine('[Bridge] Retrying bridge with companion project fallback command.');
			const fallbackClient = new LanguageClient(
				'vbnetcompanion.dotnetBridge',
				'VB.NET Companion Bridge',
				{
					command: fallback.command,
					args: fallback.args,
					transport: languageClientModule.TransportKind.stdio
				},
				{
					documentSelector,
					outputChannel: this.outputChannel,
					traceOutputChannel: this.outputChannel,
					synchronize: {
						configurationSection: ['vbnetcompanion']
					}
				}
			);

			fallbackClient.setTrace(this.toTrace(traceLevel, Trace));
			this.client = fallbackClient;
			try {
				await fallbackClient.start();
				this.outputChannel.appendLine(`[Bridge] Started language client bridge with fallback command: ${fallback.command}`);
			} catch (fallbackError) {
				const primaryMsg = error instanceof Error ? error.message : String(error);
				const fallbackMsg = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
				this.outputChannel.appendLine(`[Bridge] Fallback also failed: ${fallbackMsg}`);
				throw new Error(`Bridge startup failed. Primary: ${primaryMsg}. Fallback: ${fallbackMsg}`);
			}
		}
	}

	public async restartFromConfiguration(): Promise<void> {
		if (this.isRestarting) {
			return;
		}
		this.isRestarting = true;
		try {
			await this.stop();
			await this.startFromConfiguration();
		} finally {
			this.isRestarting = false;
		}
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

	private buildCompanionProjectFallback(command: string, args: string[]): { command: string; args: string[] } | undefined {
		const combined = `${command} ${args.join(' ')}`.toLowerCase();
		if (!combined.includes('vbnetcompanion.languageserver')) {
			return undefined;
		}

		const projectPath = this.findCompanionProjectPath();
		if (!projectPath) {
			return undefined;
		}
		return {
			command: 'dotnet',
			args: ['run', '--project', projectPath, '--', '--stdio']
		};
	}

	private resolveBridgeLaunch(command: string, args: string[]): { command: string; args: string[]; reason?: string } {
		// If the configured command exists on disk, use it directly.
		if (command.toLowerCase() !== 'dotnet' && fs.existsSync(command)) {
			return { command, args };
		}

		// Binary doesn't exist — fall back to `dotnet run` against the companion project if available.
		const looksLikeCompanionBinary = `${command} ${args.join(' ')}`.toLowerCase().includes('vbnetcompanion.languageserver');
		if (!looksLikeCompanionBinary) {
			return { command, args };
		}

		const projectPath = this.findCompanionProjectPath();
		if (!projectPath) {
			return { command, args };
		}

		return {
			command: 'dotnet',
			args: ['run', '--project', projectPath, '--', '--stdio'],
			reason: 'Configured binary not found; falling back to companion project launch via dotnet run.'
		};
	}

	private findCompanionProjectPath(): string | undefined {
		const workspaceFolders = vscode.workspace.workspaceFolders ?? [];
		for (const folder of workspaceFolders) {
			const folderPath = folder.uri.fsPath;
			const candidates = [
				path.join(folderPath, 'server', 'VBNetCompanion.LanguageServer', 'VBNetCompanion.LanguageServer.csproj'),
				path.join(folderPath, '..', 'server', 'VBNetCompanion.LanguageServer', 'VBNetCompanion.LanguageServer.csproj'),
				path.join(folderPath, '..', '..', 'server', 'VBNetCompanion.LanguageServer', 'VBNetCompanion.LanguageServer.csproj')
			];

			for (const candidate of candidates) {
				const normalized = path.normalize(candidate);
				if (fs.existsSync(normalized)) {
					return normalized;
				}
			}
		}

		return undefined;
	}
}

export function assessBridgeServerCompatibility(commandPath: string): BridgeCompatibility {
	if (!commandPath.trim()) {
		return {
			isCompatible: false,
			message: 'No bridge server command is configured.',
			hint: 'Set vbnetcompanion.languageClientServerCommand to a compatible standalone LSP server.'
		};
	}

	if (isBundledCSharpRoslynPath(commandPath)) {
		return {
			isCompatible: false,
			message: 'Bundled C# Roslyn server path is not supported as a standalone bridge endpoint on this setup.',
			hint: 'Use a dedicated standalone server command and keep VB routing off unless server supports Visual Basic.'
		};
	}

	const resolvedCommand = commandPath.trim();
	if (resolvedCommand.toLowerCase() !== 'dotnet' && !fs.existsSync(resolvedCommand)) {
		return {
			isCompatible: false,
			message: `Configured server command does not exist at the specified path: ${resolvedCommand}`,
			hint: 'Verify the path in vbnetcompanion.languageClientServerCommand or use the Apply Roslyn Bridge Preset command.'
		};
	}

	return {
		isCompatible: true,
		message: 'Configured server path passed basic compatibility checks.'
	};
}

function isBundledCSharpRoslynPath(commandPath: string): boolean {
	const normalized = commandPath.replace(/\\/g, '/').toLowerCase();
	if (!normalized.includes('/.vscode/extensions/ms-dotnettools.csharp-')) {
		return false;
	}
	return normalized.endsWith('/microsoft.codeanalysis.languageserver.exe') ||
		normalized.endsWith('/microsoft.codeanalysis.languageserver');
}
