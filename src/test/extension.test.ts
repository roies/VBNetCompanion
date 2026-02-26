import * as assert from 'assert';

// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
import * as vscode from 'vscode';
import { parseCommandLineArgs } from '../extension';
// import * as myExtension from '../../extension';

suite('Extension Test Suite', () => {
	vscode.window.showInformationMessage('Start all tests.');

	test('Sample test', () => {
		assert.strictEqual(-1, [1, 2, 3].indexOf(5));
		assert.strictEqual(-1, [1, 2, 3].indexOf(0));
	});

	test('parseCommandLineArgs handles basic args', () => {
		assert.deepStrictEqual(parseCommandLineArgs('--stdio --logLevel info'), ['--stdio', '--logLevel', 'info']);
	});

	test('parseCommandLineArgs handles quoted values', () => {
		assert.deepStrictEqual(
			parseCommandLineArgs('--path "C:\\Program Files\\Roslyn" --name "vb bridge"'),
			['--path', 'C:\\Program Files\\Roslyn', '--name', 'vb bridge']
		);
	});

	test('parseCommandLineArgs handles escaped characters', () => {
		assert.deepStrictEqual(
			parseCommandLineArgs('--flag value\\ with\\ spaces --literal \\"quoted\\"'),
			['--flag', 'value with spaces', '--literal', '"quoted"']
		);
	});
});
