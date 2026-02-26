import * as vscode from 'vscode';

export type DotnetLanguage = 'csharp' | 'vb';

export type FeatureName = 'definition' | 'completion' | 'references' | 'rename' | 'codeActions';

export type FeatureProbeResult = {
	feature: FeatureName;
	available: boolean;
	detail: string;
};

export type LanguageProbeSummary = {
	language: DotnetLanguage;
	documentUri: vscode.Uri;
	results: FeatureProbeResult[];
};

const FEATURE_ORDER: FeatureName[] = ['definition', 'completion', 'references', 'rename', 'codeActions'];

export type ParityGap = {
	feature: FeatureName;
	direction: 'vb_missing_vs_csharp' | 'csharp_missing_vs_vb';
	csharpDetail: string;
	vbDetail: string;
};

export async function runDotnetParityProbe(enableCSharp: boolean, enableVb: boolean): Promise<LanguageProbeSummary[]> {
	const summaries: LanguageProbeSummary[] = [];

	if (enableCSharp) {
		summaries.push(await probeLanguage('csharp'));
	}

	if (enableVb) {
		summaries.push(await probeLanguage('vb'));
	}

	return summaries;
}

export function buildParityReport(summaries: LanguageProbeSummary[]): string {
	if (summaries.length === 0) {
		return 'No languages are enabled for probing. Enable at least one language in extension settings.';
	}

	const byLanguage = new Map<DotnetLanguage, LanguageProbeSummary>();
	for (const summary of summaries) {
		byLanguage.set(summary.language, summary);
	}

	const lines: string[] = [];
	lines.push('=== .NET Feature Parity Probe ===');

	for (const summary of summaries) {
		lines.push('');
		lines.push(`[${labelForLanguage(summary.language)}] document: ${summary.documentUri.toString()}`);
		for (const result of summary.results) {
			const icon = result.available ? '✅' : '⚠️';
			lines.push(`  ${icon} ${result.feature}: ${result.detail}`);
		}
	}

	const gaps = getParityGaps(summaries);
	if (gaps.length > 0) {
		lines.push('');
		lines.push('--- Parity Delta (VB.NET vs C#) ---');
		for (const gap of gaps) {
			if (gap.direction === 'vb_missing_vs_csharp') {
				lines.push(`  ❌ ${gap.feature}: available in C# but not currently available in VB.NET`);
			} else {
				lines.push(`  ℹ️ ${gap.feature}: available in VB.NET but currently unavailable in C#`);
			}
		}
	} else if (byLanguage.has('csharp') && byLanguage.has('vb')) {
		lines.push('');
		lines.push('--- Parity Delta (VB.NET vs C#) ---');
		for (const feature of FEATURE_ORDER) {
			lines.push(`  ✅ ${feature}: parity state matches`);
		}
	}

	return lines.join('\n');
}

export function getParityGaps(summaries: LanguageProbeSummary[]): ParityGap[] {
	const byLanguage = new Map<DotnetLanguage, LanguageProbeSummary>();
	for (const summary of summaries) {
		byLanguage.set(summary.language, summary);
	}

	const csharp = byLanguage.get('csharp');
	const vb = byLanguage.get('vb');
	if (!csharp || !vb) {
		return [];
	}

	const gaps: ParityGap[] = [];
	for (const feature of FEATURE_ORDER) {
		const csharpFeature = csharp.results.find((item) => item.feature === feature);
		const vbFeature = vb.results.find((item) => item.feature === feature);
		if (!csharpFeature || !vbFeature) {
			continue;
		}

		if (csharpFeature.available && !vbFeature.available) {
			gaps.push({
				feature,
				direction: 'vb_missing_vs_csharp',
				csharpDetail: csharpFeature.detail,
				vbDetail: vbFeature.detail
			});
		} else if (!csharpFeature.available && vbFeature.available) {
			gaps.push({
				feature,
				direction: 'csharp_missing_vs_vb',
				csharpDetail: csharpFeature.detail,
				vbDetail: vbFeature.detail
			});
		}
	}

	return gaps;
}

async function probeLanguage(language: DotnetLanguage): Promise<LanguageProbeSummary> {
	const document = await getOrCreateProbeDocument(language);
	const position = selectProbePosition(document);
	const range = new vscode.Range(position, position);

	const results: FeatureProbeResult[] = [];
	results.push(await probeDefinition(document.uri, position));
	results.push(await probeCompletion(document.uri, position));
	results.push(await probeReferences(document.uri, position));
	results.push(await probeRename(document.uri, position));
	results.push(await probeCodeActions(document.uri, range));

	return {
		language,
		documentUri: document.uri,
		results
	};
}

async function probeDefinition(uri: vscode.Uri, position: vscode.Position): Promise<FeatureProbeResult> {
	try {
		const result = await vscode.commands.executeCommand<(vscode.Location | vscode.LocationLink)[]>(
			'vscode.executeDefinitionProvider',
			uri,
			position
		);
		const count = result?.length ?? 0;
		return {
			feature: 'definition',
			available: count > 0,
			detail: count > 0 ? `${count} definition result(s)` : 'No definition result'
		};
	} catch (error) {
		return failedResult('definition', error);
	}
}

async function probeCompletion(uri: vscode.Uri, position: vscode.Position): Promise<FeatureProbeResult> {
	try {
		const result = await vscode.commands.executeCommand<vscode.CompletionList>(
			'vscode.executeCompletionItemProvider',
			uri,
			position,
			'.',
			20
		);
		const count = result?.items.length ?? 0;
		return {
			feature: 'completion',
			available: count > 0,
			detail: count > 0 ? `${count} completion item(s)` : 'No completion results'
		};
	} catch (error) {
		return failedResult('completion', error);
	}
}

async function probeReferences(uri: vscode.Uri, position: vscode.Position): Promise<FeatureProbeResult> {
	try {
		const result = await vscode.commands.executeCommand<vscode.Location[]>(
			'vscode.executeReferenceProvider',
			uri,
			position
		);
		const count = result?.length ?? 0;
		return {
			feature: 'references',
			available: count > 0,
			detail: count > 0 ? `${count} reference result(s)` : 'No reference results'
		};
	} catch (error) {
		return failedResult('references', error);
	}
}

async function probeRename(uri: vscode.Uri, position: vscode.Position): Promise<FeatureProbeResult> {
	try {
		const result = await vscode.commands.executeCommand<vscode.WorkspaceEdit>(
			'vscode.executeDocumentRenameProvider',
			uri,
			position,
			'parityProbeRenamedSymbol'
		);
		const hasChanges = !!result && workspaceEditHasChanges(result);
		return {
			feature: 'rename',
			available: hasChanges,
			detail: hasChanges ? 'Rename provider returned edits' : 'Rename provider returned no edits'
		};
	} catch (error) {
		return failedResult('rename', error);
	}
}

async function probeCodeActions(uri: vscode.Uri, range: vscode.Range): Promise<FeatureProbeResult> {
	try {
		const result = await vscode.commands.executeCommand<(vscode.Command | vscode.CodeAction)[]>(
			'vscode.executeCodeActionProvider',
			uri,
			range
		);
		const count = result?.length ?? 0;
		return {
			feature: 'codeActions',
			available: count > 0,
			detail: count > 0 ? `${count} code action(s)` : 'No code actions returned'
		};
	} catch (error) {
		return failedResult('codeActions', error);
	}
}

async function getOrCreateProbeDocument(language: DotnetLanguage): Promise<vscode.TextDocument> {
	const glob = language === 'csharp' ? '**/*.cs' : '**/*.vb';
	const existing = await vscode.workspace.findFiles(glob, '**/{bin,obj,node_modules,.git}/**', 1);
	if (existing.length > 0) {
		return vscode.workspace.openTextDocument(existing[0]);
	}

	const content = language === 'csharp'
		? [
			'class ProbeClass',
			'{',
			'    void TestMethod()',
			'    {',
			'        int localValue = 1;',
			'        localValue.ToString();',
			'    }',
			'}'
		].join('\n')
		: [
			'Class ProbeClass',
			'    Sub TestMethod()',
			'        Dim localValue As Integer = 1',
			'        localValue.ToString()',
			'    End Sub',
			'End Class'
		].join('\n');

	return vscode.workspace.openTextDocument({
		language,
		content
	});
}

function selectProbePosition(document: vscode.TextDocument): vscode.Position {
	for (let lineIndex = 0; lineIndex < document.lineCount; lineIndex++) {
		const line = document.lineAt(lineIndex).text;
		const firstNonWhitespace = line.search(/\S/);
		if (firstNonWhitespace >= 0) {
			return new vscode.Position(lineIndex, firstNonWhitespace);
		}
	}

	return new vscode.Position(0, 0);
}

function workspaceEditHasChanges(edit: vscode.WorkspaceEdit): boolean {
	for (const [uri, edits] of edit.entries()) {
		if (edits.length > 0) {
			return true;
		}
		if (edit.get(uri).length > 0) {
			return true;
		}
	}

	return false;
}

function failedResult(feature: FeatureName, error: unknown): FeatureProbeResult {
	const detail = error instanceof Error ? error.message : String(error);
	return {
		feature,
		available: false,
		detail: `Provider error: ${detail}`
	};
}

function labelForLanguage(language: DotnetLanguage): string {
	return language === 'csharp' ? 'C#' : 'VB.NET';
}