import * as vscode from 'vscode';
import type { LanguageClient } from 'vscode-languageclient/node.js';

type LanguageClientModule = typeof import('vscode-languageclient/node.js');

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

		const args = config.get<string[]>('languageClientServerArgs', []) ?? [];
		const traceLevel = config.get<'off' | 'messages' | 'verbose'>('languageClientTraceLevel', 'off');

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
				documentSelector: [
					{ language: 'csharp' },
					{ language: 'vb' }
				],
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
