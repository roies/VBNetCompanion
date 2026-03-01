using System.Collections.Concurrent;
using System.Collections.Immutable;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Build.Locator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Classification;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Formatting;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.Recommendations;
using Microsoft.CodeAnalysis.Text;

const string JsonRpcVersion = "2.0";

var documents = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var standardInput = Console.OpenStandardInput();
var standardOutput = Console.OpenStandardOutput();
var workspaceRootPath = string.Empty;
var workspaceLoadGate = new SemaphoreSlim(1, 1);
// All writes to stdout must go through this gate to prevent concurrent-write
// stream corruption (e.g. background PushDiagnosticsAsync vs. main response).
var outputGate = new SemaphoreSlim(1, 1);
MSBuildWorkspace? roslynWorkspace = null;
Solution? roslynSolution = null;

async Task SendAsync(JsonObject message)
{
	await outputGate.WaitAsync();
	try
	{
		await WriteMessageAsync(standardOutput, message);
	}
	finally
	{
		outputGate.Release();
	}
}

async Task LogAsync(string message, int type = 4 /* 1=error,2=warn,3=info,4=log */)
{
	var notification = new JsonObject
	{
		["jsonrpc"] = JsonRpcVersion,
		["method"] = "window/logMessage",
		["params"] = new JsonObject
		{
			["type"] = type,
			["message"] = $"[VBNetCompanion.LanguageServer] {message}"
		}
	};
	await SendAsync(notification);
}

while (true)
{
	var incoming = await ReadMessageAsync(standardInput);
	if (incoming is null)
	{
		break;
	}

	var root = incoming.RootElement;
	if (!root.TryGetProperty("method", out var methodElement))
	{
		continue;
	}

	var method = methodElement.GetString();
	if (string.IsNullOrWhiteSpace(method))
	{
		continue;
	}

	root.TryGetProperty("id", out var idElement);
	JsonNode? idNode = idElement.ValueKind is JsonValueKind.Undefined
		? null
		: JsonNode.Parse(idElement.GetRawText());

	try
	{
	if (method == "initialize")
	{
		if (TryReadInitializeWorkspacePath(root, out var resolvedWorkspacePath))
		{
			workspaceRootPath = resolvedWorkspacePath;
			// Kick off workspace loading in the background so the initialize response
			// is returned immediately and the client doesn't time out waiting for it.
			_ = Task.Run(() => EnsureRoslynWorkspaceLoadedAsync(forceReload: true));
		}

		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = new JsonObject
			{
				["capabilities"] = new JsonObject
				{
					["textDocumentSync"] = 1,
					["definitionProvider"] = true,
					["referencesProvider"] = true,
					["codeLensProvider"] = new JsonObject
					{
						["resolveProvider"] = false
					},
					["completionProvider"] = new JsonObject
					{
						["resolveProvider"] = false,
						["triggerCharacters"] = new JsonArray(".", " ")
					},
					["semanticTokensProvider"] = new JsonObject
					{
						["full"] = true,
						["range"] = false,
						["legend"] = new JsonObject
						{
							["tokenTypes"] = new JsonArray(
								"namespace", "class", "enum", "interface", "struct",
								"typeParameter", "type", "parameter", "variable",
								"property", "enumMember", "event", "method"
							),
							["tokenModifiers"] = new JsonArray("static", "readonly", "declaration")
						}
					},
					["hoverProvider"] = true,
					["documentSymbolProvider"] = true,
					["documentHighlightProvider"] = true,
					["signatureHelpProvider"] = new JsonObject
					{
						["triggerCharacters"] = new JsonArray("(", ",")
					},
					["renameProvider"] = true,
					["foldingRangeProvider"] = true,
					["implementationProvider"] = true,
					["inlayHintProvider"] = true,
					["workspaceSymbolProvider"] = true,
					["callHierarchyProvider"] = true,
					["codeActionProvider"] = new JsonObject { ["resolveProvider"] = false },
					["documentFormattingProvider"] = true,
					["documentRangeFormattingProvider"] = true,
					["selectionRangeProvider"] = true,
					["documentLinkProvider"] = new JsonObject { ["resolveProvider"] = false },
					["typeHierarchyProvider"] = true
				},
				["serverInfo"] = new JsonObject
				{
					["name"] = "VBNetCompanion.LanguageServer",
					["version"] = "0.1.0"
				}
			}
		};

		await SendAsync(response);
		continue;
	}

	if (method == "initialized")
	{
		continue;
	}

	if (method == "shutdown")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = null
		};
		await SendAsync(response);
		continue;
	}

	if (method == "exit")
	{
		break;
	}

	if (method == "textDocument/didOpen")
	{
		if (TryReadDidOpen(root, out var uri, out var text))
		{
			documents[uri] = text;
			await UpdateRoslynDocumentTextAsync(uri, text);
			_ = Task.Run(() => PushDiagnosticsAsync(uri));
		}
		continue;
	}

	if (method == "textDocument/didChange")
	{
		if (TryReadDidChange(root, out var uri, out var text))
		{
			documents[uri] = text;
			await UpdateRoslynDocumentTextAsync(uri, text);
			_ = Task.Run(() => PushDiagnosticsAsync(uri));
		}
		continue;
	}

	if (method == "textDocument/didClose")
	{
		if (TryReadUri(root, out var uri))
		{
			documents.TryRemove(uri, out _);
		}
		continue;
	}

	if (method == "textDocument/definition")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleDefinitionAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/completion")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleCompletionAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/references")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleReferencesAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/codeLens")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleCodeLensAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/semanticTokens/full")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleSemanticTokensAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/hover")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleHoverAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/documentSymbol")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleDocumentSymbolAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/documentHighlight")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleDocumentHighlightAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/signatureHelp")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleSignatureHelpAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/rename")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleRenameAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/foldingRange")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleFoldingRangeAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/implementation")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleImplementationAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/inlayHint")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleInlayHintAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "workspace/symbol")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleWorkspaceSymbolAsync(root)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/prepareCallHierarchy")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandlePrepareCallHierarchyAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "callHierarchy/incomingCalls")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleIncomingCallsAsync(root)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "callHierarchy/outgoingCalls")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleOutgoingCallsAsync(root)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/codeAction")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleCodeActionAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/formatting")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleFormattingAsync(root, documents, range: null)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/rangeFormatting")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleFormattingAsync(root, documents, range: TryReadFormattingRange(root))
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/selectionRange")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleSelectionRangeAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/documentLink")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleDocumentLinkAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "textDocument/prepareTypeHierarchy")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandlePrepareTypeHierarchyAsync(root, documents)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "typeHierarchy/supertypes")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleTypeHierarchySupertypesAsync(root)
		};
		await SendAsync(response);
		continue;
	}

	if (method == "typeHierarchy/subtypes")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = await HandleTypeHierarchySubtypesAsync(root)
		};
		await SendAsync(response);
		continue;
	}

	if (idNode is not null)
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["error"] = new JsonObject
			{
				["code"] = -32601,
				["message"] = $"Method not implemented: {method}"
			}
		};
		await SendAsync(response);
	}
	} // end try
	catch (Exception loopEx)
	{
		// Log to stderr so the crash is visible in the VS Code output channel even if stdout is broken.
		await Console.Error.WriteLineAsync($"[VBNetCompanion.LanguageServer] Unhandled exception in dispatch loop for '{method}': {loopEx}");
		try
		{
			await LogAsync($"Unhandled exception dispatching '{method}': {loopEx.Message}", 1);
			// If the message had an id, reply with a JSON-RPC error so the client doesn't hang.
			if (idNode is not null)
			{
				var errResponse = new JsonObject
				{
					["jsonrpc"] = JsonRpcVersion,
					["id"] = idNode,
					["error"] = new JsonObject
					{
						["code"] = -32603,
						["message"] = $"Internal server error: {loopEx.Message}"
					}
				};
				await SendAsync(errResponse);
			}
		}
		catch
		{
			// If we can't even log, keep going; don't crash the process.
		}
	}
}

return;

async Task<JsonNode?> HandleDefinitionAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var semanticResult = await TryHandleDefinitionWithRoslynAsync(requestRoot);
	if (semanticResult is not null)
	{
		return semanticResult;
	}

	if (!TryReadPosition(requestRoot, out var uri, out var line, out var character))
	{
		return null;
	}

	if (!documents.TryGetValue(uri, out var text))
	{
		return null;
	}

	var lines = SplitLines(text);
	if (line < 0 || line >= lines.Count)
	{
		return null;
	}

	var symbol = GetWordAt(lines[line], character);
	if (string.IsNullOrWhiteSpace(symbol))
	{
		return null;
	}

	var definition = FindDefinition(lines, symbol);
	if (definition is null)
	{
		var crossFileDefinition = await TryHandleDefinitionHeuristicAsync(uri, lines, line, character, symbol);
		if (crossFileDefinition is not null)
		{
			return crossFileDefinition;
		}

		return null;
	}

	return new JsonObject
	{
		["uri"] = uri,
		["range"] = new JsonObject
		{
			["start"] = new JsonObject
			{
				["line"] = definition.Value.Line,
				["character"] = definition.Value.StartCharacter
			},
			["end"] = new JsonObject
			{
				["line"] = definition.Value.Line,
				["character"] = definition.Value.EndCharacter
			}
		}
	};
}

async Task<JsonNode?> TryHandleDefinitionHeuristicAsync(string sourceUri, IReadOnlyList<string> sourceLines, int line, int character, string symbol)
{
	if (string.IsNullOrWhiteSpace(workspaceRootPath) || line < 0 || line >= sourceLines.Count)
	{
		return null;
	}

	var lineText = sourceLines[line];
	if (!TryGetWordRangeAt(lineText, character, out var symbolStart, out _))
	{
		return null;
	}

	var receiver = TryGetReceiverForMemberAccess(lineText, symbolStart);
	var receiverType = !string.IsNullOrWhiteSpace(receiver)
		? FindLocalTypeForReceiver(sourceLines, receiver)
		: null;

	// For static method calls like DataAnalyzer.CalculateStatistics(), receiver is the class name
	// but FindLocalTypeForReceiver returns null (it's not a local variable).
	// Fall back to using the receiver token itself as the type filter.
	var typeFilter = receiverType ?? receiver;

	var candidates = Directory
		.EnumerateFiles(workspaceRootPath, "*.*", SearchOption.AllDirectories)
		.Where(path => path.EndsWith(".cs", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".vb", StringComparison.OrdinalIgnoreCase))
		.Where(path => !path.Contains("\\bin\\", StringComparison.OrdinalIgnoreCase) && !path.Contains("\\obj\\", StringComparison.OrdinalIgnoreCase));

	foreach (var candidatePath in candidates)
	{
		var candidateLines = await File.ReadAllLinesAsync(candidatePath);
		if (candidateLines.Length == 0)
		{
			continue;
		}

		if (!string.IsNullOrWhiteSpace(typeFilter))
		{
			var hasContainingType = candidateLines.Any(candidateLine => Regex.IsMatch(candidateLine, $@"\b(class|struct|module)\s+{Regex.Escape(typeFilter)}\b", RegexOptions.IgnoreCase));
			if (!hasContainingType)
			{
				continue;
			}
		}

		for (var lineIndex = 0; lineIndex < candidateLines.Length; lineIndex++)
		{
			// Skip comment lines — they should never be navigation targets.
			var trimmedCandidate = candidateLines[lineIndex].TrimStart();
			if (trimmedCandidate.StartsWith("//") || trimmedCandidate.StartsWith("'") || trimmedCandidate.StartsWith("/*"))
			{
				continue;
			}

			if (!IsMethodDeclarationLine(candidateLines[lineIndex], symbol))
			{
				continue;
			}

			var column = candidateLines[lineIndex].IndexOf(symbol, StringComparison.OrdinalIgnoreCase);
			if (column < 0)
			{
				continue;
			}

			return new JsonObject
			{
				["uri"] = new Uri(candidatePath).AbsoluteUri,
				["range"] = CreateRange(lineIndex, column, lineIndex, column + symbol.Length)
			};
		}
	}

	return null;
}

static bool TryGetWordRangeAt(string line, int character, out int start, out int end)
{
	start = 0;
	end = 0;
	if (string.IsNullOrEmpty(line))
	{
		return false;
	}

	var safeCharacter = Math.Clamp(character, 0, line.Length);
	start = safeCharacter;
	while (start > 0 && IsWordChar(line[start - 1]))
	{
		start--;
	}

	end = safeCharacter;
	while (end < line.Length && IsWordChar(line[end]))
	{
		end++;
	}

	return end > start;
}

static string? TryGetReceiverForMemberAccess(string line, int memberStart)
{
	var index = memberStart - 1;
	while (index >= 0 && char.IsWhiteSpace(line[index]))
	{
		index--;
	}

	if (index < 0 || line[index] != '.')
	{
		return null;
	}

	index--;
	while (index >= 0 && char.IsWhiteSpace(line[index]))
	{
		index--;
	}

	if (index < 0 || !IsWordChar(line[index]))
	{
		return null;
	}

	var end = index + 1;
	while (index >= 0 && IsWordChar(line[index]))
	{
		index--;
	}

	var start = index + 1;
	return line[start..end];
}

static string? FindLocalTypeForReceiver(IReadOnlyList<string> lines, string receiver)
{
	foreach (var line in lines)
	{
		var explicitNewPattern = new Regex($@"\b{Regex.Escape(receiver)}\b\s+As\s+New\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.IgnoreCase);
		var explicitTypePattern = new Regex($@"\b{Regex.Escape(receiver)}\b\s+As\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.IgnoreCase);

		var explicitNewMatch = explicitNewPattern.Match(line);
		if (explicitNewMatch.Success)
		{
			return explicitNewMatch.Groups[1].Value;
		}

		var explicitTypeMatch = explicitTypePattern.Match(line);
		if (explicitTypeMatch.Success)
		{
			return explicitTypeMatch.Groups[1].Value;
		}
	}

	return null;
}

static bool IsMethodDeclarationLine(string line, string methodName)
{
	if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(methodName))
	{
		return false;
	}

	// Skip comment lines.
	var trimmed = line.TrimStart();
	if (trimmed.StartsWith("//") || trimmed.StartsWith("'") || trimmed.StartsWith("/*") || trimmed.StartsWith("*"))
	{
		return false;
	}

	// C# method/constructor declaration: name must NOT be immediately preceded by a dot,
	// which excludes call-sites like DataAnalyzer.CalculateStatistics(...).
	var csharpMethodPattern = $@"(?<!\.){Regex.Escape(methodName)}\s*\(";

	// Type declaration pattern for C# and VB (class, struct, interface, enum, Module).
	var typeDeclarationPattern = $@"\b(class|struct|interface|enum|module)\s+{Regex.Escape(methodName)}\b";

	// VB sub/function declaration.
	var vbPattern = $@"\b(Function|Sub)\s+{Regex.Escape(methodName)}\b";

	return Regex.IsMatch(line, csharpMethodPattern, RegexOptions.IgnoreCase)
		|| Regex.IsMatch(line, typeDeclarationPattern, RegexOptions.IgnoreCase)
		|| Regex.IsMatch(line, vbPattern, RegexOptions.IgnoreCase);
}

// Legend indices must match the tokenTypes array declared in the initialize capabilities.
static int? ClassificationToTokenType(string classification) => classification switch
{
	ClassificationTypeNames.NamespaceName       => 0,
	ClassificationTypeNames.ClassName           => 1,
	ClassificationTypeNames.RecordClassName     => 1,
	ClassificationTypeNames.ModuleName          => 1,  // VB Module ≈ static class
	ClassificationTypeNames.EnumName            => 2,
	ClassificationTypeNames.InterfaceName       => 3,
	ClassificationTypeNames.StructName          => 4,
	ClassificationTypeNames.RecordStructName    => 4,
	ClassificationTypeNames.TypeParameterName   => 5,
	ClassificationTypeNames.DelegateName        => 6,
	ClassificationTypeNames.ParameterName       => 7,
	ClassificationTypeNames.LocalName           => 8,
	ClassificationTypeNames.FieldName           => 8,
	ClassificationTypeNames.ConstantName        => 8,
	ClassificationTypeNames.PropertyName        => 9,
	ClassificationTypeNames.EnumMemberName      => 10,
	ClassificationTypeNames.EventName           => 11,
	ClassificationTypeNames.MethodName          => 12,
	ClassificationTypeNames.ExtensionMethodName => 12,
	_ => (int?)null
};

async Task<JsonNode> HandleSemanticTokensAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonObject { ["data"] = new JsonArray() };

	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return empty;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return empty;

	// Apply live edits from the in-memory buffer.
	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var sourceText = await roslynDoc.GetTextAsync();
		var spans = await Classifier.GetClassifiedSpansAsync(roslynDoc, new Microsoft.CodeAnalysis.Text.TextSpan(0, sourceText.Length));

		// Group by text span — Roslyn can emit multiple classifications for the same span
		// (e.g. MethodName + StaticSymbol). Merge into a single entry with modifiers.
		var bySpan = new Dictionary<Microsoft.CodeAnalysis.Text.TextSpan, (int tokenType, int modifiers)>();

		foreach (var span in spans.OrderBy(s => s.TextSpan.Start))
		{
			if (span.ClassificationType == ClassificationTypeNames.StaticSymbol)
			{
				if (bySpan.TryGetValue(span.TextSpan, out var existing))
					bySpan[span.TextSpan] = (existing.tokenType, existing.modifiers | 1); // set static modifier
				continue;
			}

			var tokenType = ClassificationToTokenType(span.ClassificationType);
			if (tokenType is null) continue;

			var modifiers = 0;
			if (span.ClassificationType == ClassificationTypeNames.ConstantName)
				modifiers |= 2; // readonly modifier

			if (bySpan.TryGetValue(span.TextSpan, out var prev))
				bySpan[span.TextSpan] = (tokenType.Value, prev.modifiers | modifiers);
			else
				bySpan[span.TextSpan] = (tokenType.Value, modifiers);
		}

		// Encode as LSP delta-encoded token data.
		var data = new JsonArray();
		var prevLine = 0;
		var prevChar = 0;

		foreach (var (textSpan, (tokenType, modifiers)) in bySpan.OrderBy(kv => kv.Key.Start))
		{
			var linePos = sourceText.Lines.GetLinePosition(textSpan.Start);
			var line = linePos.Line;
			var startChar = linePos.Character;
			var length = textSpan.Length;
			if (length == 0) continue;

			var deltaLine = line - prevLine;
			var deltaStart = deltaLine == 0 ? startChar - prevChar : startChar;

			data.Add(deltaLine);
			data.Add(deltaStart);
			data.Add(length);
			data.Add(tokenType);
			data.Add(modifiers);

			prevLine = line;
			prevChar = startChar;
		}

		return new JsonObject { ["data"] = data };
	}
	catch (Exception ex)
	{
		await LogAsync($"SemanticTokens failed: {ex.Message}", 2);
		return empty;
	}
}

// ─── Hover ───────────────────────────────────────────────────────────────────
async Task<JsonNode?> HandleHoverAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null) return null;

	var doc = context.Value.Document;
	var position = context.Value.Position;

	// Apply live text and recompute position against it so hover works even while editing.
	if (TryReadPosition(requestRoot, out var hoverUri, out var hoverLine, out var hoverChar)
		&& documents.TryGetValue(hoverUri, out var liveText))
	{
		var liveSourceText = SourceText.From(liveText, Encoding.UTF8);
		doc = doc.WithText(liveSourceText);
		if (hoverLine >= 0 && hoverLine < liveSourceText.Lines.Count)
		{
			var safeChar = Math.Clamp(hoverChar, 0, liveSourceText.Lines[hoverLine].Span.Length);
			position = liveSourceText.Lines.GetPosition(new LinePosition(hoverLine, safeChar));
		}
	}

	var symbol = await ResolveSymbolAtPositionAsync(doc, position);
	if (symbol is null) return null;
	var sb = new StringBuilder();
	var display = symbol.ToDisplayString(SymbolDisplayFormat.CSharpErrorMessageFormat);
	sb.AppendLine("```vb");
	sb.AppendLine(display);
	sb.AppendLine("```");

	try
	{
		var xml = symbol.GetDocumentationCommentXml();
		if (!string.IsNullOrWhiteSpace(xml))
		{
			var xdoc = XDocument.Parse(xml);
			var summary = xdoc.Descendants("summary").FirstOrDefault()?.Value?.Trim();
			if (!string.IsNullOrEmpty(summary))
			{
				sb.AppendLine();
				sb.AppendLine(summary);
			}
			foreach (var param in xdoc.Descendants("param"))
			{
				var name = param.Attribute("name")?.Value;
				var desc = param.Value.Trim();
				if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(desc))
					sb.AppendLine($"\n*{name}*: {desc}");
			}
		}
	}
	catch { /* XML parse failure is non-fatal */ }

	return new JsonObject
	{
		["contents"] = new JsonObject
		{
			["kind"] = "markdown",
			["value"] = sb.ToString().Trim()
		}
	};
}

// ─── Diagnostics (push notification) ─────────────────────────────────────────
async Task PushDiagnosticsAsync(string uri)
{
	try
	{
		await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
		if (roslynSolution is null || !TryGetFilePathFromUri(uri, out var filePath)) return;

		var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
		var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
		if (roslynDoc is null) return;

		if (documents.TryGetValue(uri, out var liveText))
			roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

		var semanticModel = await roslynDoc.GetSemanticModelAsync();
		if (semanticModel is null) return;

		var diagArray = new JsonArray();
		foreach (var diag in semanticModel.GetDiagnostics())
		{
			if (diag.Severity < DiagnosticSeverity.Warning) continue;
			var loc = diag.Location;
			if (loc.Kind != LocationKind.SourceFile) continue;

			var lineSpan = loc.GetLineSpan();
			var s = lineSpan.StartLinePosition;
			var e = lineSpan.EndLinePosition;
			var lspSeverity = diag.Severity switch
			{
				DiagnosticSeverity.Error   => 1,
				DiagnosticSeverity.Warning => 2,
				DiagnosticSeverity.Info    => 3,
				_                          => 4
			};
			diagArray.Add(new JsonObject
			{
				["range"]    = CreateRange(s.Line, s.Character, e.Line, e.Character),
				["severity"] = lspSeverity,
				["code"]     = diag.Id,
				["source"]   = "VBNetCompanion",
				["message"]  = diag.GetMessage()
			});
		}

		var notification = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["method"]  = "textDocument/publishDiagnostics",
			["params"]  = new JsonObject
			{
				["uri"]         = uri,
				["diagnostics"] = diagArray
			}
		};
		await SendAsync(notification);
		await LogAsync($"Diagnostics: {diagArray.Count} issues for {Path.GetFileName(filePath)}");
	}
	catch (Exception ex)
	{
		await LogAsync($"Diagnostics push failed: {ex.Message}", 2);
	}
}

// ─── Document Symbols ─────────────────────────────────────────────────────────
async Task<JsonNode> HandleDocumentSymbolAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return empty;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return empty;

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var syntaxRoot    = await roslynDoc.GetSyntaxRootAsync();
		var semanticModel = await roslynDoc.GetSemanticModelAsync();
		var docText       = await roslynDoc.GetTextAsync();
		var symbols       = new JsonArray();
		var seen          = new HashSet<ISymbol>(SymbolEqualityComparer.Default);

		foreach (var node in syntaxRoot!.DescendantNodes())
		{
			var sym = semanticModel!.GetDeclaredSymbol(node);
			if (sym is null || sym.IsImplicitlyDeclared || !seen.Add(sym)) continue;

			var loc = sym.Locations.FirstOrDefault(l =>
				l.IsInSource && string.Equals(l.SourceTree?.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
			if (loc is null) continue;

			// Use the full syntax-node span for "range" (the whole declaration)
			// and the identifier span (loc.SourceSpan) for "selectionRange".
			// This guarantees selectionRange ⊆ range, which VS Code requires.
			var nodeStart = docText.Lines.GetLinePosition(node.Span.Start);
			var nodeEnd   = docText.Lines.GetLinePosition(node.Span.End);
			var selStart  = docText.Lines.GetLinePosition(loc.SourceSpan.Start);
			var selEnd    = docText.Lines.GetLinePosition(loc.SourceSpan.End);

			var kind = sym switch
			{
				INamedTypeSymbol { TypeKind: TypeKind.Class }     => 5,
				INamedTypeSymbol { TypeKind: TypeKind.Enum }      => 10,
				INamedTypeSymbol { TypeKind: TypeKind.Interface } => 11,
				INamedTypeSymbol { TypeKind: TypeKind.Struct }    => 23,
				INamedTypeSymbol                                   => 2,   // Module → Module
				IMethodSymbol { MethodKind: MethodKind.Constructor } => 9,
				IMethodSymbol                                      => 6,
				IPropertySymbol                                    => 7,
				IFieldSymbol { IsConst: true }                     => 14,
				IFieldSymbol                                       => 8,
				IEventSymbol                                       => 24,
				INamespaceSymbol                                   => 3,
				_                                                  => 13
			};

			symbols.Add(new JsonObject
			{
				["name"]           = sym.Name,
				["kind"]           = kind,
				["range"]          = CreateRange(nodeStart.Line, nodeStart.Character, nodeEnd.Line, nodeEnd.Character),
				["selectionRange"] = CreateRange(selStart.Line, selStart.Character, selEnd.Line, selEnd.Character)
			});
		}

		return symbols;
	}
	catch (Exception ex)
	{
		await LogAsync($"DocumentSymbol failed: {ex.Message}", 2);
		return empty;
	}
}

// ─── Document Highlights ──────────────────────────────────────────────────────
async Task<JsonNode?> HandleDocumentHighlightAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null) return null;

	var doc = context.Value.Document;
	if (TryReadUri(requestRoot, out var hlUri) && documents.TryGetValue(hlUri, out var liveText))
		doc = doc.WithText(SourceText.From(liveText, Encoding.UTF8));

	var symbol = await ResolveSymbolAtPositionAsync(doc, context.Value.Position);
	if (symbol is null) return null;

	var docText    = await doc.GetTextAsync();
	var highlights = new JsonArray();
	var seen       = new HashSet<int>();

	var refs = await SymbolFinder.FindReferencesAsync(symbol, context.Value.Solution);
	foreach (var referencedSymbol in refs)
	{
		// Declaration locations (write kind = 2)
		foreach (var declLoc in referencedSymbol.Definition.Locations.Where(l => l.IsInSource))
		{
			if (!string.Equals(declLoc.SourceTree?.FilePath, doc.FilePath, StringComparison.OrdinalIgnoreCase)) continue;
			if (!seen.Add(declLoc.SourceSpan.Start)) continue;
			var s = docText.Lines.GetLinePosition(declLoc.SourceSpan.Start);
			var e = docText.Lines.GetLinePosition(declLoc.SourceSpan.End);
			highlights.Add(new JsonObject { ["range"] = CreateRange(s.Line, s.Character, e.Line, e.Character), ["kind"] = 2 });
		}
		// Reference locations (text kind = 1)
		foreach (var refLoc in referencedSymbol.Locations)
		{
			if (!string.Equals(refLoc.Location.SourceTree?.FilePath, doc.FilePath, StringComparison.OrdinalIgnoreCase)) continue;
			if (!seen.Add(refLoc.Location.SourceSpan.Start)) continue;
			var s = docText.Lines.GetLinePosition(refLoc.Location.SourceSpan.Start);
			var e = docText.Lines.GetLinePosition(refLoc.Location.SourceSpan.End);
			highlights.Add(new JsonObject { ["range"] = CreateRange(s.Line, s.Character, e.Line, e.Character), ["kind"] = 1 });
		}
	}

	return highlights;
}

// ─── Signature Help ───────────────────────────────────────────────────────────
async Task<JsonNode?> HandleSignatureHelpAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null) return null;

	if (!TryReadPosition(requestRoot, out var uri, out var line, out var character)) return null;
	if (!TryGetFilePathFromUri(uri, out var filePath)) return null;

	var docId    = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return null;

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var sourceText = await roslynDoc.GetTextAsync();
		if (line < 0 || line >= sourceText.Lines.Count) return null;

		var safeChar = Math.Clamp(character, 0, sourceText.Lines[line].Span.Length);
		var position = sourceText.Lines.GetPosition(new LinePosition(line, safeChar));
		var rawText  = sourceText.ToString();

		// Walk backward to find the opening '(' for the current call and count commas.
		var depth       = 0;
		var activeParam = 0;
		var callPos     = -1;
		for (var i = position - 1; i >= 0; i--)
		{
			var ch = rawText[i];
			if (ch == ')' || ch == ']') { depth++; continue; }
			if (ch == '(' || ch == '[')
			{
				if (depth > 0) { depth--; continue; }
				callPos = i;
				break;
			}
			if (ch == ',' && depth == 0) activeParam++;
		}
		if (callPos < 0) return null;

		// Resolve the symbol at the position just before '('.
		var symPosition = Math.Max(0, callPos - 1);
		var symbol = await ResolveSymbolAtPositionAsync(roslynDoc, symPosition);

		IEnumerable<IMethodSymbol> methods = symbol switch
		{
			IMethodSymbol ms                                  => new[] { ms },
			INamedTypeSymbol ts when ts.TypeKind != TypeKind.Enum => ts.InstanceConstructors.Cast<IMethodSymbol>(),
			_                                                  => []
		};

		var signatures     = new JsonArray();
		var activeSignature = 0;
		var sigIdx          = 0;

		foreach (var m in methods)
		{
			var paramLabels = new JsonArray();
			foreach (var p in m.Parameters)
				paramLabels.Add(new JsonObject { ["label"] = $"{p.Name} As {p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}" });

			var isVoid  = m.ReturnsVoid || m.MethodKind == MethodKind.Constructor;
			var paramStr = string.Join(", ", m.Parameters.Select(p =>
				$"{p.Name} As {p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}"));
			var label   = isVoid
				? $"Sub {m.Name}({paramStr})"
				: $"Function {m.Name}({paramStr}) As {m.ReturnType.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}";

			var sig = new JsonObject { ["label"] = label, ["parameters"] = paramLabels };
			try
			{
				var xml = m.GetDocumentationCommentXml();
				if (!string.IsNullOrWhiteSpace(xml))
				{
					var xdoc    = XDocument.Parse(xml);
					var summary = xdoc.Descendants("summary").FirstOrDefault()?.Value?.Trim();
					if (!string.IsNullOrEmpty(summary)) sig["documentation"] = summary;
				}
			}
			catch { }

			signatures.Add(sig);
			if (m.Parameters.Length > activeParam) activeSignature = sigIdx;
			sigIdx++;
		}

		if (signatures.Count == 0) return null;

		return new JsonObject
		{
			["signatures"]      = signatures,
			["activeSignature"] = activeSignature,
			["activeParameter"] = Math.Min(activeParam, ((JsonArray)signatures[activeSignature]!["parameters"]!).Count - 1)
		};
	}
	catch (Exception ex)
	{
		await LogAsync($"SignatureHelp failed: {ex.Message}", 2);
		return null;
	}
}

// ─── Rename ───────────────────────────────────────────────────────────────────
async Task<JsonNode?> HandleRenameAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null) return null;

	// Read the new name from params.newName
	if (!requestRoot.TryGetProperty("params", out var p) || !p.TryGetProperty("newName", out var newNameEl))
		return null;
	var newName = newNameEl.GetString();
	if (string.IsNullOrWhiteSpace(newName)) return null;

	var doc = context.Value.Document;
	if (TryReadPosition(requestRoot, out var renUri, out var renLine, out var renChar)
		&& documents.TryGetValue(renUri, out var liveText))
	{
		var lst = SourceText.From(liveText, Encoding.UTF8);
		doc = doc.WithText(lst);
	}

	var symbol = await ResolveSymbolAtPositionAsync(doc, context.Value.Position);
	if (symbol is null) return null;

	await LogAsync($"Rename: renaming '{symbol.Name}' → '{newName}'");

	var refs = await SymbolFinder.FindReferencesAsync(symbol, context.Value.Solution);
	var changesByUri = new Dictionary<string, List<JsonObject>>(StringComparer.OrdinalIgnoreCase);

	foreach (var referencedSymbol in refs)
	{
		// Declaration sites
		foreach (var loc in referencedSymbol.Definition.Locations.Where(l => l.IsInSource))
		{
			var fileUri = new Uri(loc.SourceTree!.FilePath).AbsoluteUri;
			if (!changesByUri.TryGetValue(fileUri, out var list))
				changesByUri[fileUri] = list = new List<JsonObject>();
			var ls = loc.GetLineSpan();
			list.Add(new JsonObject
			{
				["range"]   = CreateRange(ls.StartLinePosition.Line, ls.StartLinePosition.Character, ls.EndLinePosition.Line, ls.EndLinePosition.Character),
				["newText"] = newName
			});
		}
		// Reference sites
		foreach (var refLoc in referencedSymbol.Locations)
		{
			var loc = refLoc.Location;
			if (!loc.IsInSource) continue;
			var fileUri = new Uri(loc.SourceTree!.FilePath).AbsoluteUri;
			if (!changesByUri.TryGetValue(fileUri, out var list))
				changesByUri[fileUri] = list = new List<JsonObject>();
			var ls = loc.GetLineSpan();
			list.Add(new JsonObject
			{
				["range"]   = CreateRange(ls.StartLinePosition.Line, ls.StartLinePosition.Character, ls.EndLinePosition.Line, ls.EndLinePosition.Character),
				["newText"] = newName
			});
		}
	}

	if (changesByUri.Count == 0) return null;

	var documentChanges = new JsonObject();
	foreach (var (fileUri, edits) in changesByUri)
	{
		// Deduplicate edits at identical ranges (declaration can appear in both sets)
		var seen    = new HashSet<string>();
		var unique  = new JsonArray();
		foreach (var edit in edits)
		{
			var key = edit.ToJsonString();
			if (seen.Add(key)) unique.Add(edit);
		}
		documentChanges[fileUri] = unique;
	}

	await LogAsync($"Rename: {changesByUri.Values.Sum(v => v.Count)} edits across {changesByUri.Count} file(s)");
	return new JsonObject { ["changes"] = documentChanges };
}

// ─── Folding Ranges ───────────────────────────────────────────────────────────
async Task<JsonNode> HandleFoldingRangeAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return empty;

	var docId    = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return empty;

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var syntaxRoot = await roslynDoc.GetSyntaxRootAsync();
		var srcText    = await roslynDoc.GetTextAsync();
		if (syntaxRoot is null) return empty;

		var ranges = new JsonArray();
		var seen   = new HashSet<int>(); // track start lines to avoid duplicates

		foreach (var node in syntaxRoot.DescendantNodesAndSelf())
		{
			// Use the runtime class name (e.g. "ClassBlockSyntax", "MethodDeclarationSyntax")
			// rather than Kind(), which requires a language-specific SyntaxNode subtype.
			var kind      = node.GetType().Name;
			var foldKind  = (string?)null;
			var isFoldable = false;

			// VB syntax class names end in "Syntax" and contain "Block"
			// C# class names end in "Syntax" and contain "Declaration" or "Statement"
			if (kind.Contains("Block") || kind.Contains("Declaration") || kind.Contains("Statement"))
			{
				isFoldable = true;
				if (kind.Contains("Class") || kind.Contains("Module") || kind.Contains("Structure")
					|| kind.Contains("Namespace") || kind.Contains("Interface") || kind.Contains("Enum"))
					foldKind = "region";
				else if (kind.Contains("Method") || kind.Contains("Sub") || kind.Contains("Function")
						|| kind.Contains("Constructor") || kind.Contains("Property") || kind.Contains("Accessor"))
					foldKind = "region";
				else
					foldKind = null; // VSCode default (collapses to "...")
			}

			if (!isFoldable) continue;

			var startLine = srcText.Lines.GetLinePosition(node.SpanStart).Line;
			var endLine   = srcText.Lines.GetLinePosition(node.Span.End).Line;

			// Only emit multi-line ranges that haven't been started already
			if (endLine <= startLine || !seen.Add(startLine)) continue;

			var entry = new JsonObject
			{
				["startLine"] = startLine,
				["endLine"]   = endLine - 1   // LSP convention: endLine is the last line shown when folded
			};
			if (foldKind is not null) entry["kind"] = foldKind;
			ranges.Add(entry);
		}

		return ranges;
	}
	catch (Exception ex)
	{
		await LogAsync($"FoldingRange failed: {ex.Message}", 2);
		return empty;
	}
}

// ─── Go to Implementation ─────────────────────────────────────────────────────
async Task<JsonNode> HandleImplementationAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();

	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null) return empty;

	var doc = context.Value.Document;
	if (TryReadPosition(requestRoot, out var implUri, out var implLine, out var implChar)
		&& documents.TryGetValue(implUri, out var liveText))
	{
		var lst = SourceText.From(liveText, Encoding.UTF8);
		doc = doc.WithText(lst);
	}

	var symbol = await ResolveSymbolAtPositionAsync(doc, context.Value.Position);
	if (symbol is null) return empty;

	var locations = new JsonArray();
	var seen      = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

	try
	{
		var impls = await SymbolFinder.FindImplementationsAsync(symbol, context.Value.Solution);
		foreach (var impl in impls)
		{
			foreach (var loc in impl.Locations.Where(l => l.IsInSource))
			{
				var node = CreateLocationNode(loc);
				var key  = node.ToJsonString();
				if (seen.Add(key)) locations.Add(node);
			}
		}

		// If no implementations were found (e.g. called on a concrete type rather than interface),
		// fall back to overrides (for virtual/abstract members).
		if (locations.Count == 0)
		{
			var overrides = await SymbolFinder.FindOverridesAsync(symbol, context.Value.Solution);
			foreach (var ov in overrides)
			{
				foreach (var loc in ov.Locations.Where(l => l.IsInSource))
				{
					var node = CreateLocationNode(loc);
					var key  = node.ToJsonString();
					if (seen.Add(key)) locations.Add(node);
				}
			}
		}
	}
	catch (Exception ex)
	{
		await LogAsync($"Implementation failed: {ex.Message}", 2);
	}

	// Final fallback: return the symbol's own declaration so Ctrl+F12 is never silent.
	// For cross-project / cross-language symbols the resolved symbol may have only metadata
	// locations, so we first attempt FindSourceDefinitionAsync (same as HandleDefinitionAsync).
	if (locations.Count == 0)
	{
		var fallbackSymbol = symbol;
		if (!symbol.Locations.Any(l => l.IsInSource))
		{
			var sourceDef = await SymbolFinder.FindSourceDefinitionAsync(symbol, context.Value.Solution);
			if (sourceDef is not null) fallbackSymbol = sourceDef;
		}
		foreach (var loc in fallbackSymbol.Locations.Where(l => l.IsInSource))
		{
			var node = CreateLocationNode(loc);
			if (seen.Add(node.ToJsonString())) locations.Add(node);
		}
	}

	return locations;
}

async Task<JsonNode> HandleCompletionAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);

	if (roslynSolution is not null
		&& TryReadPosition(requestRoot, out var uri, out var line, out var character)
		&& TryGetFilePathFromUri(uri, out var filePath))
	{
		var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
		var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;

		if (roslynDoc is not null)
		{
			// Apply any in-memory unsaved edits so completion reflects current typing.
			if (documents.TryGetValue(uri, out var liveText))
			{
				roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));
			}

			try
			{
				var sourceText = await roslynDoc.GetTextAsync();
				if (line >= 0 && line < sourceText.Lines.Count)
				{
					var safeChar = Math.Clamp(character, 0, sourceText.Lines[line].Span.Length);
					var position = sourceText.Lines.GetPosition(new LinePosition(line, safeChar));
					var lineText = sourceText.Lines[line].ToString();
					var isMemberAccess = safeChar > 0 && lineText[safeChar - 1] == '.';

					// For member access (dot trigger), use position-1 so Recommender
					// has the full receiver expression in context.
					var completionPosition = isMemberAccess ? position - 1 : position;

					var recommended = await Recommender.GetRecommendedSymbolsAtPositionAsync(
						roslynDoc, completionPosition);

					var items = new JsonArray();
					var seen = new HashSet<string>(StringComparer.Ordinal);

					foreach (var sym in recommended)
					{
						if (!sym.CanBeReferencedByName || !seen.Add(sym.Name))
							continue;

						var kind = sym switch
						{
							INamedTypeSymbol { TypeKind: TypeKind.Enum } => 13,     // Enum
							INamedTypeSymbol { TypeKind: TypeKind.Interface } => 8, // Interface
							INamedTypeSymbol { TypeKind: TypeKind.Struct } => 22,   // Struct
							INamedTypeSymbol => 7,                                  // Class
							IMethodSymbol => 3,                                      // Method
							IPropertySymbol => 10,                                   // Property
							IFieldSymbol { IsConst: true } => 21,                   // Constant
							IFieldSymbol => 5,                                       // Field
							IEventSymbol => 23,                                      // Event
							ILocalSymbol => 6,                                       // Variable
							IParameterSymbol => 6,                                   // Variable
							INamespaceSymbol => 9,                                   // Module/Namespace
							_ => 6
						};

						var detail = sym switch
						{
							IMethodSymbol m => m.ReturnType.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat),
							IPropertySymbol p => p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat),
							IFieldSymbol f => f.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat),
							ILocalSymbol l => l.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat),
							IParameterSymbol p => p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat),
							INamedTypeSymbol t => t.ContainingNamespace?.ToDisplayString() ?? "",
							_ => ""
						};

						var entry = new JsonObject { ["label"] = sym.Name, ["kind"] = kind };
						if (!string.IsNullOrEmpty(detail))
							entry["detail"] = detail;
						items.Add(entry);
					}

					if (items.Count > 0)
					{

						return new JsonObject { ["isIncomplete"] = false, ["items"] = items };
					}
				}
			}
			catch (Exception ex)
			{
				await LogAsync($"Completion: Roslyn failed: {ex.Message}", 2);
			}
		}
	}

	return HandleCompletionFallback(requestRoot, documents);
}

static JsonNode HandleCompletionFallback(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var keywords = new[]
	{
		"Class", "Module", "Namespace", "Sub", "Function", "Property", "Dim", "As", "If", "Then", "Else",
		"End", "Public", "Private", "Friend", "Protected", "Imports", "Return", "ByVal", "ByRef", "New"
	};

	var items = new JsonArray();
	foreach (var keyword in keywords)
		items.Add(new JsonObject { ["label"] = keyword, ["kind"] = 14 });

	if (TryReadUri(requestRoot, out var uri) && documents.TryGetValue(uri, out var text))
	{
		foreach (var symbol in ExtractDeclaredSymbols(SplitLines(text)).Distinct(StringComparer.OrdinalIgnoreCase))
			items.Add(new JsonObject { ["label"] = symbol, ["kind"] = 6 });
	}

	return new JsonObject { ["isIncomplete"] = false, ["items"] = items };
}

async Task<JsonNode> HandleReferencesAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var semanticResult = await TryHandleReferencesWithRoslynAsync(requestRoot);
	if (semanticResult is not null)
	{
		return semanticResult;
	}

	if (!TryReadPosition(requestRoot, out var uri, out var line, out var character))
	{
		return new JsonArray();
	}

	if (!documents.TryGetValue(uri, out var sourceText))
	{
		return new JsonArray();
	}

	var sourceLines = SplitLines(sourceText);
	if (line < 0 || line >= sourceLines.Count)
	{
		return new JsonArray();
	}

	var symbol = GetWordAt(sourceLines[line], character);
	if (string.IsNullOrWhiteSpace(symbol))
	{
		return new JsonArray();
	}

	var declarations = FindDeclarations(sourceLines);
	var context = ResolveReferenceContext(uri, line, character, symbol, declarations);
	var includeDeclaration = TryReadIncludeDeclaration(requestRoot, out var include) ? include : true;
	var locations = new JsonArray();

	foreach (var document in documents)
	{
		var lines = SplitLines(document.Value);
		for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
		{
			if (context.IsLocalScope) {
				var sameDocument = string.Equals(document.Key, context.Uri, StringComparison.OrdinalIgnoreCase);
				if (!sameDocument || lineIndex < context.ScopeStartLine || lineIndex > context.ScopeEndLine)
				{
					continue;
				}
			}

			foreach (var occurrence in FindSymbolOccurrences(lines[lineIndex], symbol))
			{
				if (context.IsLocalScope && IsMemberAccessOccurrence(lines[lineIndex], occurrence.StartCharacter))
				{
					continue;
				}

				var isDeclaration = includeDeclaration == false
					&& context.Declaration is not null
					&& string.Equals(document.Key, context.Uri, StringComparison.OrdinalIgnoreCase)
					&& context.Declaration.Value.Line == lineIndex
					&& context.Declaration.Value.StartCharacter == occurrence.StartCharacter
					&& context.Declaration.Value.EndCharacter == occurrence.EndCharacter;

				if (isDeclaration)
				{
					continue;
				}

				locations.Add(new JsonObject
				{
					["uri"] = document.Key,
					["range"] = new JsonObject
					{
						["start"] = new JsonObject
						{
							["line"] = lineIndex,
							["character"] = occurrence.StartCharacter
						},
						["end"] = new JsonObject
						{
							["line"] = lineIndex,
							["character"] = occurrence.EndCharacter
						}
					}
				});
			}
		}
	}

	return locations;
}

async Task<JsonNode> HandleCodeLensAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	if (!TryReadUri(requestRoot, out var uri))
	{
		return new JsonArray();
	}

	// Try to resolve the Roslyn document for cross-project reference counting.
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	Document? roslynDoc = null;
	if (roslynSolution is not null && TryGetFilePathFromUri(uri, out var codeLensFilePath))
	{
		var docId = roslynSolution.GetDocumentIdsWithFilePath(codeLensFilePath).FirstOrDefault();
		roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	}

	if (!documents.TryGetValue(uri, out var sourceText))
	{
		if (roslynDoc is null) return new JsonArray();
		var roslynText = await roslynDoc.GetTextAsync();
		sourceText = roslynText.ToString();
	}

	// Roslyn-first path: enumerate declared symbols directly via the semantic model.
	// This handles both C# and VB without language-specific regex.
	if (roslynDoc is not null && roslynSolution is not null)
	{
		try
		{
			var syntaxRoot = await roslynDoc.GetSyntaxRootAsync();
			var semanticModel = await roslynDoc.GetSemanticModelAsync();
			var docText = await roslynDoc.GetTextAsync();
			var roslynLenses = new JsonArray();
			var seen = new HashSet<ISymbol>(SymbolEqualityComparer.Default);

			foreach (var node in syntaxRoot!.DescendantNodes())
			{
				var symbol = semanticModel!.GetDeclaredSymbol(node);
				if (symbol is null || symbol.IsImplicitlyDeclared) continue;
				if (symbol is not INamedTypeSymbol
					and not IPropertySymbol
					and not IMethodSymbol { MethodKind: MethodKind.Ordinary or MethodKind.Constructor })
					continue;
				if (!seen.Add(symbol)) continue;

				var symbolLoc = symbol.Locations.FirstOrDefault(l => l.IsInSource);
				if (symbolLoc is null) continue;

				var spanStart = docText.Lines.GetLinePosition(symbolLoc.SourceSpan.Start);
				var spanEnd = docText.Lines.GetLinePosition(symbolLoc.SourceSpan.End);

				var refs = await SymbolFinder.FindReferencesAsync(symbol, roslynSolution);
				var refCount = 0;
				var refLocations = new JsonArray();
				foreach (var refSym in refs)
				{
					foreach (var refLoc in refSym.Locations)
					{
						refCount++;
						refLocations.Add(CreateLocationNode(refLoc.Location));
					}
				}

				var title = refCount == 1 ? "1 reference" : $"{refCount} references";
				roslynLenses.Add(new JsonObject
				{
					["range"] = CreateRange(spanStart.Line, spanStart.Character, spanEnd.Line, spanEnd.Character),
					["command"] = new JsonObject
					{
						["title"] = title,
						["command"] = "vbnetcompanion.showReferencesFromBridge",
						["arguments"] = new JsonArray
						{
							uri,
							new JsonObject
							{
								["line"] = spanStart.Line,
								["character"] = spanStart.Character
							},
							refLocations
						}
					}
				});
			}
			return roslynLenses;
		}
		catch (Exception ex)
		{
			await LogAsync($"CodeLens Roslyn enum failed: {ex.Message}", 2);
		}
	}

	// Fallback: text-based VB keyword regex (used when Roslyn workspace is unavailable).
	var lines = SplitLines(sourceText);
	var declarations = FindDeclarations(lines);
	var lenses = new JsonArray();

	foreach (var declaration in declarations)
	{
		if (string.Equals(declaration.Kind, "Dim", StringComparison.OrdinalIgnoreCase))
		{
			continue;
		}

		var referenceLocations = new JsonArray();
		var referenceCount = 0;

		// Without Roslyn we can't disambiguate same-named symbols across files,
		// so only search the current file to avoid false cross-file matches.
		var singleFileDocuments = documents
			.Where(d => string.Equals(d.Key, uri, StringComparison.OrdinalIgnoreCase))
			.ToList();

		foreach (var document in singleFileDocuments)
		{
			var documentLines = SplitLines(document.Value);
			for (var lineIndex = 0; lineIndex < documentLines.Count; lineIndex++)
			{
				foreach (var occurrence in FindSymbolOccurrences(documentLines[lineIndex], declaration.Symbol))
				{
					if (IsMemberAccessOccurrence(documentLines[lineIndex], occurrence.StartCharacter))
					{
						continue;
					}

					var isDeclaration =
						string.Equals(document.Key, uri, StringComparison.OrdinalIgnoreCase)
						&& lineIndex == declaration.Line
						&& occurrence.StartCharacter == declaration.StartCharacter
						&& occurrence.EndCharacter == declaration.EndCharacter;

					if (isDeclaration) continue;

					referenceCount++;
					referenceLocations.Add(CreateLocation(document.Key, lineIndex, occurrence.StartCharacter, occurrence.EndCharacter));
				}
			}
		}

		var title = referenceCount == 1 ? "1 reference" : $"{referenceCount} references";
		lenses.Add(new JsonObject
		{
			["range"] = CreateRange(declaration.Line, declaration.StartCharacter, declaration.Line, declaration.EndCharacter),
			["command"] = new JsonObject
			{
				["title"] = title,
				["command"] = "vbnetcompanion.showReferencesFromBridge",
				["arguments"] = new JsonArray
				{
					uri,
					new JsonObject
					{
						["line"] = declaration.Line,
						["character"] = declaration.StartCharacter
					},
					referenceLocations
				}
			}
		});
	}

	return lenses;
}

static (int Line, int StartCharacter, int EndCharacter)? FindDefinition(IReadOnlyList<string> lines, string symbol)
{
	var patterns = new[]
	{
		$@"\bClass\s+{Regex.Escape(symbol)}\b",
		$@"\bModule\s+{Regex.Escape(symbol)}\b",
		$@"\bSub\s+{Regex.Escape(symbol)}\b",
		$@"\bFunction\s+{Regex.Escape(symbol)}\b",
		$@"\bProperty\s+{Regex.Escape(symbol)}\b",
		$@"\bDim\s+{Regex.Escape(symbol)}\b"
	};

	for (var i = 0; i < lines.Count; i++)
	{
		foreach (var pattern in patterns)
		{
			var match = Regex.Match(lines[i], pattern, RegexOptions.IgnoreCase);
			if (match.Success)
			{
				var symbolIndex = lines[i].IndexOf(symbol, match.Index, StringComparison.OrdinalIgnoreCase);
				if (symbolIndex >= 0)
				{
					return (i, symbolIndex, symbolIndex + symbol.Length);
				}
			}
		}
	}

	return null;
}

static IEnumerable<string> ExtractDeclaredSymbols(IReadOnlyList<string> lines)
{
	var pattern = new Regex("\\b(?:Class|Module|Sub|Function|Property|Dim)\\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.IgnoreCase);
	foreach (var line in lines)
	{
		foreach (Match match in pattern.Matches(line))
		{
			if (match.Groups.Count > 1)
			{
				yield return match.Groups[1].Value;
			}
		}
	}
}

static string GetWordAt(string line, int character)
{
	if (string.IsNullOrEmpty(line))
	{
		return string.Empty;
	}

	var safeCharacter = Math.Clamp(character, 0, line.Length);
	var left = safeCharacter;
	while (left > 0 && IsWordChar(line[left - 1]))
	{
		left--;
	}

	var right = safeCharacter;
	while (right < line.Length && IsWordChar(line[right]))
	{
		right++;
	}

	if (right <= left)
	{
		return string.Empty;
	}

	return line[left..right];
}

static bool IsWordChar(char character)
{
	return char.IsLetterOrDigit(character) || character == '_';
}

static IEnumerable<(int StartCharacter, int EndCharacter)> FindSymbolOccurrences(string line, string symbol)
{
	if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(symbol))
	{
		yield break;
	}

	var startIndex = 0;
	while (startIndex < line.Length)
	{
		var matchIndex = line.IndexOf(symbol, startIndex, StringComparison.OrdinalIgnoreCase);
		if (matchIndex < 0)
		{
			yield break;
		}

		var endIndex = matchIndex + symbol.Length;
		var leftBoundary = matchIndex == 0 || !IsWordChar(line[matchIndex - 1]);
		var rightBoundary = endIndex >= line.Length || !IsWordChar(line[endIndex]);

		if (leftBoundary && rightBoundary)
		{
			yield return (matchIndex, endIndex);
		}

		startIndex = endIndex;
	}
}

static bool IsMemberAccessOccurrence(string line, int symbolStartCharacter)
{
	for (var index = symbolStartCharacter - 1; index >= 0; index--)
	{
		var character = line[index];
		if (char.IsWhiteSpace(character))
		{
			continue;
		}

		return character == '.';
	}

	return false;
}

static IReadOnlyList<(string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)> FindDeclarations(IReadOnlyList<string> lines)
{
	var declarations = new List<(string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)>();
	var pattern = new Regex("\\b(Class|Module|Sub|Function|Property|Dim)\\s+([A-Za-z_][A-Za-z0-9_]*)", RegexOptions.IgnoreCase);
	var procedureScopes = FindProcedureScopes(lines);

	for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
	{
		foreach (Match match in pattern.Matches(lines[lineIndex]))
		{
			if (match.Groups.Count <= 2 || !match.Groups[2].Success)
			{
				continue;
			}

			var kind = match.Groups[1].Value;

			var symbolGroup = match.Groups[2];
			var scopeStartLine = 0;
			var scopeEndLine = lines.Count - 1;
			if (string.Equals(kind, "Dim", StringComparison.OrdinalIgnoreCase))
			{
				var containingScope = procedureScopes.FirstOrDefault(scope => lineIndex >= scope.StartLine && lineIndex <= scope.EndLine);
				var foundContainingScope = procedureScopes.Any(scope => lineIndex >= scope.StartLine && lineIndex <= scope.EndLine);
				if (foundContainingScope)
				{
					scopeStartLine = containingScope.StartLine;
					scopeEndLine = containingScope.EndLine;
				}
			}

			declarations.Add((
				kind,
				symbolGroup.Value,
				lineIndex,
				symbolGroup.Index,
				symbolGroup.Index + symbolGroup.Length,
				scopeStartLine,
				scopeEndLine));
		}
	}

	return declarations;
}

static IReadOnlyList<(int StartLine, int EndLine)> FindProcedureScopes(IReadOnlyList<string> lines)
{
	var startPattern = new Regex("^\\s*(?:(?:Public|Private|Friend|Protected|Shared|Partial|Overloads|Overrides|Overridable|MustOverride|NotOverridable|Async|Iterator|Static)\\s+)*(Sub|Function)\\b", RegexOptions.IgnoreCase);
	var endPattern = new Regex("^\\s*End\\s+(Sub|Function)\\b", RegexOptions.IgnoreCase);
	var scopes = new List<(int StartLine, int EndLine)>();
	var stack = new Stack<(string Kind, int StartLine)>();

	for (var lineIndex = 0; lineIndex < lines.Count; lineIndex++)
	{
		var line = lines[lineIndex];
		var startMatch = startPattern.Match(line);
		if (startMatch.Success)
		{
			stack.Push((startMatch.Groups[1].Value, lineIndex));
		}

		var endMatch = endPattern.Match(line);
		if (endMatch.Success && stack.Count > 0)
		{
			var endKind = endMatch.Groups[1].Value;
			var start = stack.Pop();
			if (string.Equals(start.Kind, endKind, StringComparison.OrdinalIgnoreCase))
			{
				scopes.Add((start.StartLine, lineIndex));
			}
		}
	}

	return scopes;
}

static (string Uri, string Symbol, bool IsLocalScope, int ScopeStartLine, int ScopeEndLine, (string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)? Declaration) ResolveReferenceContext(
	string uri,
	int line,
	int character,
	string symbol,
	IReadOnlyList<(string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)> declarations)
{
	var localAtCursor = declarations
		.Where(declaration => string.Equals(declaration.Kind, "Dim", StringComparison.OrdinalIgnoreCase))
		.Where(declaration => declaration.Line == line)
		.Where(declaration => character >= declaration.StartCharacter && character <= declaration.EndCharacter)
		.Cast<(string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)?>()
		.FirstOrDefault();

	if (localAtCursor is not null)
	{
		var localDeclarationAtCursor = localAtCursor.Value;
		return (uri, symbol, true, localDeclarationAtCursor.ScopeStartLine, localDeclarationAtCursor.ScopeEndLine, localDeclarationAtCursor);
	}

	var localInScope = declarations
		.Where(declaration => string.Equals(declaration.Kind, "Dim", StringComparison.OrdinalIgnoreCase))
		.Where(declaration => string.Equals(declaration.Symbol, symbol, StringComparison.OrdinalIgnoreCase))
		.Where(declaration => line >= declaration.ScopeStartLine && line <= declaration.ScopeEndLine)
		.OrderByDescending(declaration => declaration.Line)
		.Cast<(string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)?>()
		.FirstOrDefault();

	if (localInScope is not null)
	{
		var localDeclarationInScope = localInScope.Value;
		return (uri, symbol, true, localDeclarationInScope.ScopeStartLine, localDeclarationInScope.ScopeEndLine, localDeclarationInScope);
	}

	var matchingDeclaration = declarations
		.Where(item => string.Equals(item.Symbol, symbol, StringComparison.OrdinalIgnoreCase))
		.OrderBy(item => item.Line)
		.Cast<(string Kind, string Symbol, int Line, int StartCharacter, int EndCharacter, int ScopeStartLine, int ScopeEndLine)?>()
		.FirstOrDefault();

	if (matchingDeclaration is not null)
	{
		return (uri, symbol, false, 0, int.MaxValue, matchingDeclaration.Value);
	}

	return (uri, symbol, false, 0, int.MaxValue, null);
}

async Task<JsonNode?> TryHandleDefinitionWithRoslynAsync(JsonElement requestRoot)
{
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null)
	{
		return null;
	}

	var symbol = await ResolveSymbolAtPositionAsync(context.Value.Document, context.Value.Position);
	if (symbol is null)
	{
		await LogAsync("Definition: no symbol at position");
		return null;
	}

	await LogAsync($"Definition: resolved symbol '{symbol.Name}' kind={symbol.Kind} implicitlyDeclared={symbol.IsImplicitlyDeclared}");

	// When the resolved symbol is an implicitly-declared constructor (e.g. `New GreeterService()` where
	// GreeterService has no explicit constructor), navigate to the containing type instead.
	var navigationSymbol = symbol;
	if (navigationSymbol is Microsoft.CodeAnalysis.IMethodSymbol { MethodKind: Microsoft.CodeAnalysis.MethodKind.Constructor, IsImplicitlyDeclared: true } implicitCtor
		&& implicitCtor.ContainingType is not null)
	{
		navigationSymbol = implicitCtor.ContainingType;
		await LogAsync($"Definition: implicit ctor → navigating to containing type '{navigationSymbol.Name}'");
	}

	var sourceLocation = navigationSymbol
		.Locations
		.Where(location => location.IsInSource && location.SourceTree is not null)
		.FirstOrDefault();

	// If still no source location, try to resolve through the source definition (handles metadata stubs,
	// including cross-language P2P references where VB sees C# types as metadata).
	if (sourceLocation is null)
	{
		var locationKinds = string.Join(", ", navigationSymbol.Locations.Select(l => l.IsInSource ? "src" : "meta"));
		await LogAsync($"Definition: no direct source location for '{navigationSymbol.Name}' — locations: [{locationKinds}], trying FindSourceDefinitionAsync");

		var sourceDef = await SymbolFinder.FindSourceDefinitionAsync(navigationSymbol, context.Value.Solution);
		if (sourceDef is not null)
		{
			sourceLocation = sourceDef.Locations
				.Where(location => location.IsInSource && location.SourceTree is not null)
				.FirstOrDefault();
			if (sourceLocation is not null)
			{
				await LogAsync($"Definition: found source definition via FindSourceDefinitionAsync for '{sourceDef.Name}'");
			}
			else
			{
				await LogAsync($"Definition: FindSourceDefinitionAsync returned '{sourceDef.Name}' but still no in-source location");
			}
		}
		else
		{
			await LogAsync($"Definition: FindSourceDefinitionAsync returned null for '{navigationSymbol.Name}'");
		}
	}

	// Last-resort: search by name across all solution projects.
	// This is needed when cross-language P2P references produce metadata symbols
	// that FindSourceDefinitionAsync cannot re-link (e.g., VB consuming a C# project).
	if (sourceLocation is null)
	{
		await LogAsync($"Definition: trying FindDeclarationsAsync fallback for '{navigationSymbol.Name}'");
		foreach (var project in context.Value.Solution.Projects)
		{
			var decls = await SymbolFinder.FindDeclarationsAsync(
				project, navigationSymbol.Name, ignoreCase: false,
				filter: SymbolFilter.Type | SymbolFilter.Member);
			var matchingDecl = decls.FirstOrDefault(d =>
				d.Kind == navigationSymbol.Kind &&
				string.Equals(d.Name, navigationSymbol.Name, StringComparison.Ordinal));
			if (matchingDecl is not null)
			{
				sourceLocation = matchingDecl.Locations
					.Where(l => l.IsInSource && l.SourceTree is not null)
					.FirstOrDefault();
				if (sourceLocation is not null)
				{
					await LogAsync($"Definition: found via FindDeclarationsAsync in project '{project.Name}'");
					break;
				}
			}
		}
	}

	// Final fallback: generate a metadata-as-source stub so F12 on .NET BCL types
	// (e.g. ConsoleColor, Console, String) navigates to a readable declaration file.
	if (sourceLocation is null)
	{
		var metaNode = await TryGenerateMetadataStubLocationAsync(navigationSymbol);
		if (metaNode is not null)
		{
			await LogAsync($"Definition: navigating to metadata stub for '{navigationSymbol.Name}'");
			return metaNode;
		}
		await LogAsync($"Definition: no source location found for '{navigationSymbol.Name}' after all fallbacks");
		return null;
	}

	return CreateLocationNode(sourceLocation);
}

async Task<JsonNode?> TryGenerateMetadataStubLocationAsync(ISymbol symbol)
{
	// Only handle purely metadata symbols (no in-source location at all).
	if (symbol.Locations.Any(l => l.IsInSource))
		return null;

	// For members, navigate to the containing type's stub; for types, use the type itself.
	var typeSymbol = symbol as INamedTypeSymbol ?? symbol.ContainingType;
	if (typeSymbol is null)
		return null;

	var stubContent = BuildMetadataStub(typeSymbol);

	var stubDir = Path.Combine(Path.GetTempPath(), "VBNetCompanion", "metadata");
	Directory.CreateDirectory(stubDir);

	var qualifiedName = typeSymbol.ContainingNamespace is { IsGlobalNamespace: false } ns
		? $"{ns.ToDisplayString()}.{typeSymbol.Name}"
		: typeSymbol.Name;
	var safeName = Regex.Replace(qualifiedName, @"[^\w.]", "_");
	var stubPath = Path.Combine(stubDir, safeName + ".vb");

	// Write the stub only if it doesn't already exist; swallow IOException from
	// concurrent requests racing to write the same file.
	if (!File.Exists(stubPath))
	{
		try { await File.WriteAllTextAsync(stubPath, stubContent, Encoding.UTF8); }
		catch (IOException) { /* another concurrent F12 already wrote it — that's fine */ }
	}

	// Position the cursor on the member line (for fields/methods) or the type declaration line.
	var targetName = symbol.Name;
	var lines = stubContent.Split('\n');
	var targetLine = 0;
	var typePattern = new Regex($@"\b(Class|Enum|Interface|Structure|Module)\s+{Regex.Escape(typeSymbol.Name)}\b", RegexOptions.IgnoreCase);

	for (var i = 0; i < lines.Length; i++)
	{
		var trimmed = lines[i].TrimStart();
		if (trimmed.StartsWith("'")) continue;

		// For a member symbol, land on that specific member's line.
		if (!ReferenceEquals(symbol, typeSymbol) && lines[i].Contains(targetName, StringComparison.OrdinalIgnoreCase))
		{
			targetLine = i;
			break;
		}
		// For the type itself, land on the type declaration keyword.
		if (ReferenceEquals(symbol, typeSymbol) && typePattern.IsMatch(lines[i]))
		{
			targetLine = i;
			break;
		}
	}

	var col = targetLine < lines.Length
		? Math.Max(0, lines[targetLine].IndexOf(targetName, StringComparison.OrdinalIgnoreCase))
		: 0;

	return new JsonObject
	{
		["uri"] = new Uri(stubPath).AbsoluteUri,
		["range"] = CreateRange(targetLine, col, targetLine, col + targetName.Length)
	};
}

static string BuildMetadataStub(INamedTypeSymbol typeSymbol)
{
	var sb = new StringBuilder();
	var assemblyName = typeSymbol.ContainingAssembly?.Name ?? "unknown";
	var namespaceName = typeSymbol.ContainingNamespace is { IsGlobalNamespace: false } ns
		? ns.ToDisplayString() : null;

	sb.AppendLine($"' [metadata] {typeSymbol.ToDisplayString()} — {assemblyName}");
	sb.AppendLine("' Auto-generated by VB.NET Companion for navigation. Read-only.");
	sb.AppendLine();

	if (namespaceName is not null)
	{
		sb.AppendLine($"Namespace {namespaceName}");
		sb.AppendLine();
	}

	var indent = namespaceName is not null ? "    " : "";

	AppendXmlDocComment(sb, typeSymbol, indent);

	var keyword = typeSymbol.TypeKind switch
	{
		TypeKind.Enum => "Enum",
		TypeKind.Interface => "Interface",
		TypeKind.Struct => "Structure",
		_ => typeSymbol.IsStatic ? "Module" : "Class"
	};

	sb.AppendLine($"{indent}Public {keyword} {typeSymbol.Name}");

	foreach (var member in typeSymbol.GetMembers().OrderBy(m => m.Name))
	{
		if (member.IsImplicitlyDeclared || !member.CanBeReferencedByName) continue;
		if (member.DeclaredAccessibility is not Accessibility.Public and not Accessibility.Protected) continue;

		AppendXmlDocComment(sb, member, indent + "    ");

		switch (member)
		{
			case IFieldSymbol field when typeSymbol.TypeKind == TypeKind.Enum:
				var constVal = field.ConstantValue is not null ? $" = {field.ConstantValue}" : "";
				sb.AppendLine($"{indent}    {field.Name}{constVal}");
				break;
			case IFieldSymbol field:
				var fieldMod = field.IsConst ? "Const" : "Shared ReadOnly";
				sb.AppendLine($"{indent}    Public {fieldMod} {field.Name} As {field.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}");
				break;
			case IPropertySymbol prop:
				var propMod = prop.IsReadOnly ? "ReadOnly " : prop.IsWriteOnly ? "WriteOnly " : "";
				var propShared = prop.IsStatic ? "Shared " : "";
				sb.AppendLine($"{indent}    Public {propShared}{propMod}Property {prop.Name} As {prop.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}");
				break;
			case IMethodSymbol method when method.MethodKind is MethodKind.Ordinary or MethodKind.Constructor:
				var isVoid = method.ReturnsVoid || method.MethodKind == MethodKind.Constructor;
				var sfKeyword = isVoid ? "Sub" : "Function";
				var methodName = method.MethodKind == MethodKind.Constructor ? "New" : method.Name;
				var parms = string.Join(", ", method.Parameters.Select(p =>
					$"{p.Name} As {p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}"));
				var retSuffix = isVoid ? "" : $" As {method.ReturnType.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}";
				var methodShared = method.IsStatic ? "Shared " : "";
				sb.AppendLine($"{indent}    Public {methodShared}{sfKeyword} {methodName}({parms}){retSuffix}");
				break;
			case IEventSymbol evt:
				sb.AppendLine($"{indent}    Public Event {evt.Name} As {evt.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}");
				break;
		}
	}

	sb.AppendLine($"{indent}End {keyword}");
	if (namespaceName is not null)
	{
		sb.AppendLine();
		sb.AppendLine("End Namespace");
	}

	return sb.ToString();
}

static void AppendXmlDocComment(StringBuilder sb, ISymbol symbol, string indent)
{
	try
	{
		var xml = symbol.GetDocumentationCommentXml();
		if (string.IsNullOrWhiteSpace(xml)) return;
		var doc = XDocument.Parse(xml);
		var summary = doc.Descendants("summary").FirstOrDefault()?.Value?.Trim();
		if (string.IsNullOrWhiteSpace(summary)) return;
		sb.AppendLine($"{indent}''' <summary>{summary}</summary>");
	}
	catch { /* XML parse can fail for some metadata symbols */ }
}

async Task<JsonNode?> TryHandleReferencesWithRoslynAsync(JsonElement requestRoot)
{
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null)
	{
		return null;
	}

	var symbol = await ResolveSymbolAtPositionAsync(context.Value.Document, context.Value.Position);
	if (symbol is null)
	{
		return null;
	}

	var includeDeclaration = TryReadIncludeDeclaration(requestRoot, out var include) ? include : true;
	var locations = new JsonArray();
	var references = await SymbolFinder.FindReferencesAsync(symbol, context.Value.Solution);

	foreach (var referencedSymbol in references)
	{
		foreach (var referenceLocation in referencedSymbol.Locations)
		{
			if (!includeDeclaration && referenceLocation.IsImplicit)
			{
				continue;
			}

			locations.Add(CreateLocationNode(referenceLocation.Location));
		}

		if (includeDeclaration)
		{
			foreach (var declarationLocation in referencedSymbol.Definition.Locations.Where(location => location.IsInSource))
			{
				locations.Add(CreateLocationNode(declarationLocation));
			}
		}
	}

	return locations;
}

async Task<ISymbol?> ResolveSymbolAtPositionAsync(Document document, int position)
{
	var syntaxRoot = await document.GetSyntaxRootAsync();
	if (syntaxRoot is null)
	{
		return null;
	}

	var safePosition = Math.Clamp(position, 0, Math.Max(0, syntaxRoot.FullSpan.End - 1));
	var token = syntaxRoot.FindToken(safePosition, findInsideTrivia: true);
	if (!token.Span.Contains(safePosition) && safePosition > 0)
	{
		token = syntaxRoot.FindToken(safePosition - 1, findInsideTrivia: true);
	}

	var semanticModel = await document.GetSemanticModelAsync();
	if (semanticModel is null)
	{
		return null;
	}

	var workspace = roslynWorkspace ?? document.Project.Solution.Workspace;
	var symbol = await SymbolFinder.FindSymbolAtPositionAsync(semanticModel, safePosition, workspace);
	if (symbol is not null)
	{
		return symbol;
	}

	var currentNode = token.Parent;
	while (currentNode is not null)
	{
		symbol = semanticModel.GetSymbolInfo(currentNode).Symbol ?? semanticModel.GetDeclaredSymbol(currentNode);
		if (symbol is not null)
		{
			return symbol;
		}

		var candidates = semanticModel.GetSymbolInfo(currentNode).CandidateSymbols;
		symbol = candidates.FirstOrDefault();
		if (symbol is not null)
		{
			return symbol;
		}

		currentNode = currentNode.Parent;
	}

	return null;
}

async Task<(Document Document, int Position, Solution Solution)?> TryResolveRoslynContextAsync(JsonElement requestRoot)
{
	if (!TryReadPosition(requestRoot, out var uri, out var line, out var character))
	{
		return null;
	}

	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null)
	{
		return null;
	}

	if (!TryGetFilePathFromUri(uri, out var filePath))
	{
		return null;
	}

	var documentId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	if (documentId is null)
	{
		var projectNames = string.Join(", ", roslynSolution.Projects.Select(p => p.Name));
		await LogAsync($"Definition: file not found in Roslyn solution (projects: {projectNames}) — {filePath}", 2);
		return null;
	}

	var document = roslynSolution.GetDocument(documentId);
	if (document is null)
	{
		return null;
	}

	var text = await document.GetTextAsync();
	if (line < 0 || line >= text.Lines.Count)
	{
		return null;
	}

	var safeCharacter = Math.Clamp(character, 0, text.Lines[line].Span.Length);
	var position = text.Lines.GetPosition(new LinePosition(line, safeCharacter));
	return (document, position, roslynSolution);
}

async Task EnsureRoslynWorkspaceLoadedAsync(bool forceReload)
{
	if (string.IsNullOrWhiteSpace(workspaceRootPath))
	{
		return;
	}

	if (!forceReload && roslynSolution is not null)
	{
		return;
	}

	await workspaceLoadGate.WaitAsync();
	try
	{
		if (!forceReload && roslynSolution is not null)
		{
			return;
		}

		if (!MSBuildLocator.IsRegistered)
		{
			var msbuildPath = TryFindMSBuildPath();
			if (msbuildPath is not null)
			{
				await LogAsync($"Registering MSBuild from: {msbuildPath}");
				// Warn if the selected SDK doesn't match the server's runtime major version.
				var sdkDirName = Path.GetFileName(msbuildPath.TrimEnd(Path.DirectorySeparatorChar));
				var runtimeMajor = Environment.Version.Major;
				if (Version.TryParse(sdkDirName.Split('-')[0], out var selectedSdkVersion) && selectedSdkVersion.Major != runtimeMajor)
				{
					await LogAsync($"Warning: Selected SDK {sdkDirName} (major {selectedSdkVersion.Major}) differs from server runtime (major {runtimeMajor}). " +
						"This may cause MissingMethodException errors.", 2);
				}
				MSBuildLocator.RegisterMSBuildPath(msbuildPath);
			}
			else
			{
				await LogAsync("Falling back to MSBuildLocator.RegisterDefaults()");
				MSBuildLocator.RegisterDefaults();
			}
		}

		// Configure MSBuildWorkspace with design-time build properties.
		// This mirrors what Visual Studio / OmniSharp use for IDE evaluation and prevents
		// false failures such as "Your project file doesn't list 'win' as a RuntimeIdentifier"
		// that occur when MSBuild applies the host platform's default RID to old .NET Framework projects.
		var workspaceProperties = new Dictionary<string, string>
		{
			{ "DesignTimeBuild", "true" },
			{ "BuildingInsideVisualStudio", "true" },
			{ "SkipCompilerExecution", "true" },
			{ "ProvideCommandLineArgs", "true" },
			{ "ShouldUnsetParentConfigurationAndPlatform", "false" },
		};
		var workspace = MSBuildWorkspace.Create(workspaceProperties);
		var workspaceFailures = new ConcurrentBag<string>();
		workspace.RegisterWorkspaceFailedHandler(e => workspaceFailures.Add($"[{e.Diagnostic.Kind}] {e.Diagnostic.Message}"));

		var solutionPath = Directory
			.GetFiles(workspaceRootPath, "*.sln", SearchOption.TopDirectoryOnly)
			.FirstOrDefault();

		if (!string.IsNullOrWhiteSpace(solutionPath))
		{
			try
			{
				roslynSolution = await workspace.OpenSolutionAsync(solutionPath);
				roslynWorkspace = workspace;
				foreach (var f in workspaceFailures) await LogAsync($"WorkspaceDiag: {f}", 2);
				await LogAsync($"Roslyn solution loaded: {solutionPath} ({roslynSolution.Projects.Count()} projects)");
				await SummarizeWorkspaceLoadIssuesAsync(workspaceFailures);
				return;
			}
			catch (Exception ex)
			{
				await LogAsync($"Solution load failed ({solutionPath}): {ex.Message} — falling back to project load", 2);
			}
		}

		// Try to read project paths from a .slnx file (XML format not supported by OpenSolutionAsync).
		var slnxPath = Directory
			.GetFiles(workspaceRootPath, "*.slnx", SearchOption.TopDirectoryOnly)
			.FirstOrDefault();

		var slnxProjectPaths = new List<string>();
		if (!string.IsNullOrWhiteSpace(slnxPath))
		{
			try
			{
				var slnxDoc = XDocument.Load(slnxPath);
				foreach (var projectElement in slnxDoc.Descendants("Project"))
				{
					var relPath = projectElement.Attribute("Path")?.Value;
					if (string.IsNullOrWhiteSpace(relPath)) continue;
					var absPath = Path.GetFullPath(Path.Combine(workspaceRootPath, relPath));
					if (File.Exists(absPath)) slnxProjectPaths.Add(absPath);
				}
				await LogAsync($"Parsed .slnx: {slnxProjectPaths.Count} project(s) from {slnxPath}");
			}
			catch (Exception ex)
			{
				await LogAsync($".slnx parse failed: {ex.Message}", 2);
			}
		}

		// Fall back to scanning the workspace for projects, excluding bin/obj.
		var excludedSegments = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "bin", "obj", ".git", "node_modules" };
		var scannedProjectPaths = Directory
			.GetFiles(workspaceRootPath, "*.csproj", SearchOption.AllDirectories)
			.Concat(Directory.GetFiles(workspaceRootPath, "*.vbproj", SearchOption.AllDirectories))
			.Where(p => !Path.GetFullPath(p).Split(Path.DirectorySeparatorChar).Any(seg => excludedSegments.Contains(seg)))
			.ToList();

		var projectPaths = slnxProjectPaths.Count > 0 ? slnxProjectPaths : scannedProjectPaths;

		if (projectPaths.Count == 0)
		{
			await LogAsync("No projects found to load", 2);
			return;
		}

		await LogAsync($"Loading {projectPaths.Count} project(s): {string.Join(", ", projectPaths.Select(Path.GetFileName))}");
		foreach (var projectPath in projectPaths)
		{
			// Skip projects already loaded as transitive dependencies of a previously opened project.
			var normalizedPath = Path.GetFullPath(projectPath);
			var alreadyLoaded = workspace.CurrentSolution.Projects
				.Any(p => string.Equals(p.FilePath, normalizedPath, StringComparison.OrdinalIgnoreCase));
			if (alreadyLoaded)
			{
				await LogAsync($"Skipping (already loaded as transitive dep): {Path.GetFileName(projectPath)}");
				roslynSolution = workspace.CurrentSolution;
				continue;
			}

			try
			{
				await LogAsync($"Opening project: {projectPath}");
				await workspace.OpenProjectAsync(projectPath);
				// Always take the full workspace solution — it includes all transitively loaded projects.
				roslynSolution = workspace.CurrentSolution;
				var loadedProject = roslynSolution.Projects
					.FirstOrDefault(p => string.Equals(p.FilePath, normalizedPath, StringComparison.OrdinalIgnoreCase));
				await LogAsync($"Opened project: {Path.GetFileName(projectPath)} ({loadedProject?.Documents.Count() ?? 0} docs)");
			}
			catch (Exception ex)
			{
				await LogAsync($"OpenProjectAsync failed for {Path.GetFileName(projectPath)}: {ex.Message}", 1);
			}
		}

		roslynWorkspace = workspace;
		// Final sync — ensures roslynSolution reflects all projects including transitive ones.
		roslynSolution = workspace.CurrentSolution;
		foreach (var f in workspaceFailures) await LogAsync($"WorkspaceDiag: {f}", 2);
		await LogAsync($"Roslyn projects loaded: {roslynSolution?.Projects.Count() ?? 0} project(s)");
		await SummarizeWorkspaceLoadIssuesAsync(workspaceFailures);
	}
	catch (Exception ex)
	{
		roslynWorkspace = null;
		roslynSolution = null;
		try { await LogAsync($"Roslyn workspace load failed: {ex}", 1); } catch { /* pipe may be broken; don't crash */ }
	}
	finally
	{
		workspaceLoadGate.Release();
	}
}

async Task SummarizeWorkspaceLoadIssuesAsync(ConcurrentBag<string> failures)
{
	if (failures.IsEmpty) return;

	var failureList = failures.ToList();
	var sdkResolverFailures = failureList.Count(f => f.Contains("SdkResolverFailure", StringComparison.OrdinalIgnoreCase)
		|| f.Contains("MissingMethodException", StringComparison.OrdinalIgnoreCase)
		|| f.Contains("SDK Resolver Failure", StringComparison.OrdinalIgnoreCase));
	var missingRefWarnings = failureList.Count(f => f.Contains("without a matching metadata reference", StringComparison.OrdinalIgnoreCase));
	var otherFailures = failureList.Count - sdkResolverFailures - missingRefWarnings;

	var parts = new List<string>();

	if (sdkResolverFailures > 0)
	{
		parts.Add($"{sdkResolverFailures} MSBuild SDK resolver failure(s) — this usually means the installed .NET SDK version is incompatible with the language server runtime. " +
			"Install a .NET 8 SDK or ensure the server targets a framework matching your SDK.");
	}

	if (missingRefWarnings > 0)
	{
		parts.Add($"{missingRefWarnings} project reference(s) could not be resolved. Cross-project navigation and IntelliSense may be degraded.");
	}

	if (otherFailures > 0)
	{
		parts.Add($"{otherFailures} other workspace diagnostic(s).");
	}

	if (parts.Count == 0) return;

	var summary = $"Workspace loaded with issues: {string.Join(" ", parts)}";
	await LogAsync(summary, sdkResolverFailures > 0 ? 1 : 2);

	// Send a user-visible notification for critical issues.
	if (sdkResolverFailures > 0)
	{
		var notification = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["method"] = "window/showMessage",
			["params"] = new JsonObject
			{
				["type"] = 2, // Warning
				["message"] = $"VB.NET Companion: {sdkResolverFailures} project(s) failed to load due to .NET SDK version mismatch. " +
					"Install a .NET 8.x SDK for full compatibility, or check the output channel for details."
			}
		};
		await SendAsync(notification);
	}
}

async Task UpdateRoslynDocumentTextAsync(string uri, string text)
{
	if (roslynSolution is null)
	{
		return;
	}

	if (!TryGetFilePathFromUri(uri, out var filePath))
	{
		return;
	}

	await workspaceLoadGate.WaitAsync();
	try
	{
		if (roslynSolution is null)
		{
			return;
		}

		var documentIds = roslynSolution.GetDocumentIdsWithFilePath(filePath);
		if (documentIds.Length == 0)
		{
			return;
		}

		var nextSolution = roslynSolution;
		var updatedText = SourceText.From(text, Encoding.UTF8);
		foreach (var documentId in documentIds)
		{
			nextSolution = nextSolution.WithDocumentText(documentId, updatedText, PreservationMode.PreserveIdentity);
		}

		roslynSolution = nextSolution;
		roslynWorkspace?.TryApplyChanges(nextSolution);
	}
	finally
	{
		workspaceLoadGate.Release();
	}
}

bool TryReadInitializeWorkspacePath(JsonElement requestRoot, out string path)
{
	path = string.Empty;
	if (!requestRoot.TryGetProperty("params", out var paramsElement))
	{
		return false;
	}

	if (paramsElement.TryGetProperty("rootUri", out var rootUriElement))
	{
		var rootUri = rootUriElement.GetString();
		if (!string.IsNullOrWhiteSpace(rootUri) && TryGetFilePathFromUri(rootUri, out var rootPathFromUri))
		{
			path = rootPathFromUri;
			return true;
		}
	}

	if (paramsElement.TryGetProperty("workspaceFolders", out var foldersElement) && foldersElement.ValueKind == JsonValueKind.Array)
	{
		foreach (var folder in foldersElement.EnumerateArray())
		{
			if (!folder.TryGetProperty("uri", out var uriElement))
			{
				continue;
			}

			var folderUri = uriElement.GetString();
			if (!string.IsNullOrWhiteSpace(folderUri) && TryGetFilePathFromUri(folderUri, out var rootPathFromFolder))
			{
				path = rootPathFromFolder;
				return true;
			}
		}
	}

	return false;
}

bool TryGetFilePathFromUri(string uri, out string filePath)
{
	filePath = string.Empty;
	if (string.IsNullOrWhiteSpace(uri))
	{
		return false;
	}

	var candidate = uri.Trim();

	if (candidate.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
	{
		if (!Uri.TryCreate(candidate, UriKind.Absolute, out var parsedUri) || !parsedUri.IsFile)
		{
			return false;
		}

		candidate = Uri.UnescapeDataString(parsedUri.LocalPath);
	}

	var hasLeadingSlashDrivePattern =
		candidate.Length >= 4 &&
		candidate[0] == '/' &&
		char.IsLetter(candidate[1]) &&
		candidate[2] == ':' &&
		(candidate[3] == '/' || candidate[3] == '\\');

	if (hasLeadingSlashDrivePattern)
	{
		candidate = candidate[1..];
	}

	if (!Path.IsPathRooted(candidate))
	{
		return false;
	}

	try
	{
		filePath = Path.GetFullPath(candidate);
		return !string.IsNullOrWhiteSpace(filePath);
	}
	catch
	{
		filePath = string.Empty;
		return false;
	}
}

JsonObject CreateLocationNode(Location location)
{
	if (location.SourceTree is null || string.IsNullOrWhiteSpace(location.SourceTree.FilePath))
	{
		return new JsonObject();
	}

	var lineSpan = location.GetLineSpan();
	var start = lineSpan.StartLinePosition;
	var end = lineSpan.EndLinePosition;
	var uri = new Uri(location.SourceTree.FilePath).AbsoluteUri;

	return new JsonObject
	{
		["uri"] = uri,
		["range"] = CreateRange(start.Line, start.Character, end.Line, end.Character)
	};
}

static JsonObject CreateLocation(string uri, int startLine, int startCharacter, int endCharacter)
{
	return new JsonObject
	{
		["uri"] = uri,
		["range"] = CreateRange(startLine, startCharacter, startLine, endCharacter)
	};
}

static JsonObject CreateRange(int startLine, int startCharacter, int endLine, int endCharacter)
{
	return new JsonObject
	{
		["start"] = new JsonObject
		{
			["line"] = startLine,
			["character"] = startCharacter
		},
		["end"] = new JsonObject
		{
			["line"] = endLine,
			["character"] = endCharacter
		}
	};
}

/// <summary>
/// Build a selectionRange from start + nameLength, clamped so it never exceeds
/// the enclosing fullRange (start→end).  VS Code rejects symbols whose
/// selectionRange is not contained in the fullRange.
/// </summary>
static JsonObject ClampedSelectionRange(LinePosition start, LinePosition end, int nameLength)
{
	int selEndLine = start.Line;
	int selEndChar = start.Character + nameLength;

	// If selectionRange end would exceed the fullRange end, clamp it.
	if (selEndLine > end.Line || (selEndLine == end.Line && selEndChar > end.Character))
	{
		selEndLine = end.Line;
		selEndChar = end.Character;
	}
	return CreateRange(start.Line, start.Character, selEndLine, selEndChar);
}

static IReadOnlyList<string> SplitLines(string text)
{
	return text.Replace("\r\n", "\n").Split('\n');
}

static bool TryReadDidOpen(JsonElement root, out string uri, out string text)
{
	uri = string.Empty;
	text = string.Empty;
	if (!root.TryGetProperty("params", out var paramsElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("textDocument", out var documentElement))
	{
		return false;
	}

	if (!documentElement.TryGetProperty("uri", out var uriElement) || !documentElement.TryGetProperty("text", out var textElement))
	{
		return false;
	}

	uri = uriElement.GetString() ?? string.Empty;
	text = textElement.GetString() ?? string.Empty;
	return !string.IsNullOrWhiteSpace(uri);
}

static bool TryReadDidChange(JsonElement root, out string uri, out string text)
{
	uri = string.Empty;
	text = string.Empty;
	if (!root.TryGetProperty("params", out var paramsElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("textDocument", out var documentElement))
	{
		return false;
	}

	if (!documentElement.TryGetProperty("uri", out var uriElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("contentChanges", out var changesElement) || changesElement.ValueKind != JsonValueKind.Array)
	{
		return false;
	}

	var firstChange = changesElement.EnumerateArray().FirstOrDefault();
	if (firstChange.ValueKind == JsonValueKind.Undefined || !firstChange.TryGetProperty("text", out var textElement))
	{
		return false;
	}

	uri = uriElement.GetString() ?? string.Empty;
	text = textElement.GetString() ?? string.Empty;
	return !string.IsNullOrWhiteSpace(uri);
}

static bool TryReadPosition(JsonElement root, out string uri, out int line, out int character)
{
	uri = string.Empty;
	line = 0;
	character = 0;
	if (!root.TryGetProperty("params", out var paramsElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("textDocument", out var documentElement))
	{
		return false;
	}

	if (!documentElement.TryGetProperty("uri", out var uriElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("position", out var positionElement))
	{
		return false;
	}

	if (!positionElement.TryGetProperty("line", out var lineElement) || !positionElement.TryGetProperty("character", out var characterElement))
	{
		return false;
	}

	uri = uriElement.GetString() ?? string.Empty;
	line = lineElement.GetInt32();
	character = characterElement.GetInt32();
	return !string.IsNullOrWhiteSpace(uri);
}

static bool TryReadUri(JsonElement root, out string uri)
{
	uri = string.Empty;
	if (!root.TryGetProperty("params", out var paramsElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("textDocument", out var documentElement))
	{
		return false;
	}

	if (!documentElement.TryGetProperty("uri", out var uriElement))
	{
		return false;
	}

	uri = uriElement.GetString() ?? string.Empty;
	return !string.IsNullOrWhiteSpace(uri);
}

static bool TryReadIncludeDeclaration(JsonElement root, out bool includeDeclaration)
{
	includeDeclaration = true;
	if (!root.TryGetProperty("params", out var paramsElement))
	{
		return false;
	}

	if (!paramsElement.TryGetProperty("context", out var contextElement))
	{
		return false;
	}

	if (!contextElement.TryGetProperty("includeDeclaration", out var includeDeclarationElement))
	{
		return false;
	}

	if (includeDeclarationElement.ValueKind != JsonValueKind.True && includeDeclarationElement.ValueKind != JsonValueKind.False)
	{
		return false;
	}

	includeDeclaration = includeDeclarationElement.GetBoolean();
	return true;
}

static async Task<JsonDocument?> ReadMessageAsync(Stream input)
{
	var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
	while (true)
	{
		var line = await ReadHeaderLineAsync(input);
		if (line is null)
		{
			return null;
		}

		if (line.Length == 0)
		{
			break;
		}

		var separator = line.IndexOf(':');
		if (separator <= 0)
		{
			continue;
		}

		var name = line[..separator].Trim();
		var value = line[(separator + 1)..].Trim();
		headers[name] = value;
	}

	if (!headers.TryGetValue("Content-Length", out var contentLengthRaw) || !int.TryParse(contentLengthRaw, out var contentLength))
	{
		return null;
	}

	var payload = new byte[contentLength];
	var read = 0;
	while (read < contentLength)
	{
		var bytesRead = await input.ReadAsync(payload.AsMemory(read, contentLength - read));
		if (bytesRead == 0)
		{
			return null;
		}
		read += bytesRead;
	}

	return JsonDocument.Parse(payload);
}

static async Task<string?> ReadHeaderLineAsync(Stream input)
{
	var bytes = new List<byte>();
	while (true)
	{
		var buffer = new byte[1];
		var read = await input.ReadAsync(buffer.AsMemory(0, 1));
		if (read == 0)
		{
			return null;
		}

		if (buffer[0] == (byte)'\n')
		{
			break;
		}

		if (buffer[0] != (byte)'\r')
		{
			bytes.Add(buffer[0]);
		}
	}

	return Encoding.UTF8.GetString(bytes.ToArray());
}

static string? TryFindMSBuildPath()
{
	// Try to find the best .NET SDK by running `dotnet --list-sdks`.
	// We strongly prefer an SDK whose major version matches the server's actual
	// running runtime to avoid System.Text.Json / MSBuild version mismatches
	// that cause MissingMethodException at runtime.  If no matching major-version
	// SDK is available we fall back to the latest installed one.
	// Format: "10.0.102 [C:\Program Files\dotnet\sdk]"
	var runtimeMajorVersion = Environment.Version.Major;

	try
	{
		var process = new System.Diagnostics.Process
		{
			StartInfo = new System.Diagnostics.ProcessStartInfo
			{
				FileName = "dotnet",
				Arguments = "--list-sdks",
				RedirectStandardOutput = true,
				UseShellExecute = false,
				CreateNoWindow = true
			}
		};
		process.Start();
		var output = process.StandardOutput.ReadToEnd();
		process.WaitForExit(5000);

		string? bestCompatiblePath = null;
		Version? bestCompatibleVersion = null;
		string? bestFallbackPath = null;
		Version? bestFallbackVersion = null;

		foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
		{
			var trimmed = line.Trim();
			// Format: "10.0.102 [C:\Program Files\dotnet\sdk]"
			var bracketStart = trimmed.IndexOf('[');
			var bracketEnd = trimmed.IndexOf(']');
			if (bracketStart < 0 || bracketEnd <= bracketStart) continue;

			var versionStr = trimmed[..bracketStart].Trim();
			var sdkBase = trimmed[(bracketStart + 1)..bracketEnd].Trim();
			var sdkPath = Path.Combine(sdkBase, versionStr);

			if (!Directory.Exists(sdkPath)) continue;
			if (!Version.TryParse(versionStr.Split('-')[0], out var version)) continue;

			if (version.Major == runtimeMajorVersion)
			{
				if (bestCompatibleVersion is null || version > bestCompatibleVersion)
				{
					bestCompatibleVersion = version;
					bestCompatiblePath = sdkPath;
				}
			}
			else
			{
				if (bestFallbackVersion is null || version > bestFallbackVersion)
				{
					bestFallbackVersion = version;
					bestFallbackPath = sdkPath;
				}
			}
		}

		// Prefer the compatible SDK; fall back to the latest available.
		return bestCompatiblePath ?? bestFallbackPath;
	}
	catch
	{
		return null;
	}
}

// ─── Inlay Hints ─────────────────────────────────────────────────────────────
async Task<JsonNode> HandleInlayHintAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var hints = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return hints;

	if (!TryReadHintRange(requestRoot, out var rStartLine, out var rStartChar, out var rEndLine, out var rEndChar))
		return hints;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return hints;

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var syntaxRoot    = await roslynDoc.GetSyntaxRootAsync();
		var semanticModel = await roslynDoc.GetSemanticModelAsync();
		var docText       = await roslynDoc.GetTextAsync();
		if (syntaxRoot is null || semanticModel is null) return hints;

		var totalLines  = docText.Lines.Count;
		var safeEndLine = Math.Min(rEndLine, totalLines - 1);
		var safeEndChar = Math.Min(rEndChar, docText.Lines[safeEndLine].Span.Length);
		var spanStart   = docText.Lines.GetPosition(new LinePosition(Math.Min(rStartLine, totalLines - 1), rStartChar));
		var spanEnd     = docText.Lines.GetPosition(new LinePosition(safeEndLine, safeEndChar));
		var rangeSpan   = TextSpan.FromBounds(Math.Min(spanStart, spanEnd), Math.Max(spanStart, spanEnd));

		foreach (var node in syntaxRoot.DescendantNodes(n => n.Span.IntersectsWith(rangeSpan)))
		{
			var typeName = node.GetType().Name;
			if (typeName != "InvocationExpressionSyntax"     &&
			    typeName != "ObjectCreationExpressionSyntax" &&
			    typeName != "ImplicitObjectCreationExpressionSyntax") continue;
			if (!node.Span.IntersectsWith(rangeSpan)) continue;

			if (semanticModel.GetSymbolInfo(node).Symbol is not IMethodSymbol methodSym) continue;
			if (methodSym.Parameters.Length == 0) continue;

			var argListNode = node.ChildNodes().FirstOrDefault(n => n.GetType().Name.Contains("ArgumentList"));
			if (argListNode is null) continue;

			var argNodes = argListNode.ChildNodes()
				.Where(n => !n.GetType().Name.StartsWith("Omitted"))
				.ToList();

			for (int i = 0; i < argNodes.Count && i < methodSym.Parameters.Length; i++)
			{
				var argNode = argNodes[i];
				var param   = methodSym.Parameters[i];
				if (param.IsParams || param.Name.Length <= 1) continue;

				// Skip already-named arguments (VB NamedArgumentSyntax; C# ArgumentSyntax.NameColon)
				bool isNamed = argNode.GetType().Name.StartsWith("Named") ||
				              argNode.GetType().GetProperty("NameColon")?.GetValue(argNode) is not null;
				if (isNamed) continue;

				var exprNode = argNode.ChildNodes().FirstOrDefault() ?? argNode;
				var exprText = exprNode.ToString().Trim().TrimStart('_');
				if (string.Equals(exprText, param.Name, StringComparison.OrdinalIgnoreCase)) continue;

				var argStartPos = docText.Lines.GetLinePosition(argNode.SpanStart);
				hints.Add(new JsonObject
				{
					["position"]     = new JsonObject { ["line"] = argStartPos.Line, ["character"] = argStartPos.Character },
					["label"]        = param.Name + ":",
					["kind"]         = 2,
					["paddingRight"] = JsonValue.Create(true)
				});
			}
		}

		}
	catch (Exception ex)
	{
		await LogAsync($"InlayHint failed: {ex.Message}", 2);
	}

	return hints;
}

// ─── Workspace Symbols ────────────────────────────────────────────────────────
async Task<JsonNode> HandleWorkspaceSymbolAsync(JsonElement requestRoot)
{
	var results = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null) return results;

	if (!requestRoot.TryGetProperty("params", out var p) ||
	    !p.TryGetProperty("query", out var qEl))
		return results;
	var query = qEl.GetString() ?? string.Empty;
	if (string.IsNullOrWhiteSpace(query)) return results;

	var seen   = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
	var capped = false;
	try
	{
		foreach (var project in roslynSolution.Projects)
		{
			if (capped) break;
			var decls = await SymbolFinder.FindDeclarationsAsync(project, query, ignoreCase: true);
			foreach (var sym in decls)
			{
				if (sym.IsImplicitlyDeclared) continue;
				var loc = sym.Locations.FirstOrDefault(l => l.IsInSource);
				if (loc is null) continue;
				var key = $"{sym.ContainingNamespace?.ToDisplayString()}::{sym.Name}::{loc.GetLineSpan().StartLinePosition.Line}";
				if (!seen.Add(key)) continue;

				results.Add(new JsonObject
				{
					["name"]     = sym.Name,
					["kind"]     = SymbolToLspKind(sym),
					["location"] = CreateLocationNode(loc)
				});
				if (results.Count >= 200) { capped = true; break; }
			}
		}
	}
	catch (Exception ex)
	{
		await LogAsync($"WorkspaceSymbol failed: {ex.Message}", 2);
	}

	return results;
}

// ─── Call Hierarchy ───────────────────────────────────────────────────────────
async Task<JsonNode> HandlePrepareCallHierarchyAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null) return empty;

	if (!TryReadPosition(requestRoot, out var uri, out _, out _)) return empty;
	var document = context.Value.Document;
	if (documents.TryGetValue(uri, out var liveText))
		document = document.WithText(SourceText.From(liveText, Encoding.UTF8));

	var symbol = await ResolveSymbolAtPositionAsync(document, context.Value.Position);
	if (symbol is null) return empty;

	var loc = symbol.Locations.FirstOrDefault(l => l.IsInSource);
	if (loc is null)
	{
		var srcDef = await SymbolFinder.FindSourceDefinitionAsync(symbol, context.Value.Solution);
		if (srcDef is not null) loc = srcDef.Locations.FirstOrDefault(l => l.IsInSource);
		if (loc is null) return empty;
	}

	var lineSpan = loc.GetLineSpan();
	var start    = lineSpan.StartLinePosition;
	var end      = lineSpan.EndLinePosition;
	var symUri   = new Uri(loc.SourceTree!.FilePath).AbsoluteUri;

	var item = new JsonObject
	{
		["name"]           = symbol.Name,
		["kind"]           = SymbolToLspKind(symbol),
		["uri"]            = symUri,
		["range"]          = CreateRange(start.Line, start.Character, end.Line, end.Character),
		["selectionRange"] = ClampedSelectionRange(start, end, symbol.Name.Length),
		["data"]           = new JsonObject { ["uri"] = symUri, ["line"] = start.Line, ["character"] = start.Character }
	};

	return new JsonArray { item };
}

async Task<JsonNode> HandleIncomingCallsAsync(JsonElement requestRoot)
{
	var results = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null) return results;

	if (!TryReadCallHierarchyItemData(requestRoot, out var fileUri, out var fileLine, out var fileChar)) return results;
	if (!TryGetFilePathFromUri(fileUri, out var filePath)) return results;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return results;

	var docText  = await roslynDoc.GetTextAsync();
	var safeChar = Math.Clamp(fileChar, 0, docText.Lines[fileLine].Span.Length);
	var position = docText.Lines.GetPosition(new LinePosition(fileLine, safeChar));

	var symbol = await ResolveSymbolAtPositionAsync(roslynDoc, position);
	if (symbol is null) return results;


	try
	{
		var refs        = await SymbolFinder.FindReferencesAsync(symbol, roslynSolution);
		var byContainer = new Dictionary<ISymbol, (JsonObject Item, JsonArray Ranges)>(SymbolEqualityComparer.Default);

		foreach (var refGroup in refs)
		{
			foreach (var refLoc in refGroup.Locations)
			{
				if (!refLoc.Location.IsInSource) continue;
				var callerDoc = roslynSolution.GetDocument(refLoc.Document.Id);
				if (callerDoc is null) continue;

				var callerModel = await callerDoc.GetSemanticModelAsync();
				var callerRoot  = await callerDoc.GetSyntaxRootAsync();
				if (callerModel is null || callerRoot is null) continue;

				var tokenParent  = callerRoot.FindToken(refLoc.Location.SourceSpan.Start).Parent;
				var containerSym = FindContainingMethodSymbol(callerModel, tokenParent);
				if (containerSym is null) continue;

				var cLoc = containerSym.Locations.FirstOrDefault(l => l.IsInSource);
				if (cLoc is null) continue;

				var refLineSpan = refLoc.Location.GetLineSpan();
				var refRange    = CreateRange(
					refLineSpan.StartLinePosition.Line, refLineSpan.StartLinePosition.Character,
					refLineSpan.EndLinePosition.Line,   refLineSpan.EndLinePosition.Character);

				if (!byContainer.TryGetValue(containerSym, out var entry))
				{
					var cLineSpan = cLoc.GetLineSpan();
					var cStart    = cLineSpan.StartLinePosition;
					var cEnd      = cLineSpan.EndLinePosition;
					var cUri      = new Uri(cLoc.SourceTree!.FilePath).AbsoluteUri;
					var callerItem = new JsonObject
					{
						["name"]           = containerSym.Name,
						["kind"]           = SymbolToLspKind(containerSym),
						["uri"]            = cUri,
						["range"]          = CreateRange(cStart.Line, cStart.Character, cEnd.Line, cEnd.Character),
						["selectionRange"] = ClampedSelectionRange(cStart, cEnd, containerSym.Name.Length),
						["data"]           = new JsonObject { ["uri"] = cUri, ["line"] = cStart.Line, ["character"] = cStart.Character }
					};
					entry = (callerItem, new JsonArray());
					byContainer[containerSym] = entry;
				}
				entry.Ranges.Add(refRange);
			}
		}

		foreach (var (callerItem, ranges) in byContainer.Values)
			results.Add(new JsonObject { ["from"] = callerItem, ["fromRanges"] = ranges });

	}
	catch (Exception ex)
	{
		await LogAsync($"IncomingCalls failed: {ex.Message}", 2);
	}

	return results;
}

async Task<JsonNode> HandleOutgoingCallsAsync(JsonElement requestRoot)
{
	var results = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null) return results;

	if (!TryReadCallHierarchyItemData(requestRoot, out var fileUri, out var fileLine, out var fileChar)) return results;
	if (!TryGetFilePathFromUri(fileUri, out var filePath)) return results;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return results;

	var docText  = await roslynDoc.GetTextAsync();
	var safeChar = Math.Clamp(fileChar, 0, docText.Lines[fileLine].Span.Length);
	var position = docText.Lines.GetPosition(new LinePosition(fileLine, safeChar));

	var symbol = await ResolveSymbolAtPositionAsync(roslynDoc, position);
	if (symbol is null) return results;


	try
	{
		var syntaxRoot    = await roslynDoc.GetSyntaxRootAsync();
		var semanticModel = await roslynDoc.GetSemanticModelAsync();
		if (syntaxRoot is null || semanticModel is null) return results;

		// Find the method block that contains the resolved position
		var token    = syntaxRoot.FindToken(position);
		var declNode = token.Parent;
		while (declNode is not null)
		{
			var tn = declNode.GetType().Name;
			if (tn.Contains("MethodBlock")   || tn.Contains("MethodDeclaration")  ||
			    tn.Contains("SubBlock")       || tn.Contains("FunctionBlock")      ||
			    tn.Contains("ConstructorBlock") || tn.Contains("ConstructorDeclaration") ||
			    tn.Contains("AccessorDeclaration"))
				break;
			declNode = declNode.Parent;
		}
		if (declNode is null) return results;

		var calleeToRanges = new Dictionary<ISymbol, (JsonObject Item, JsonArray Ranges)>(SymbolEqualityComparer.Default);

		foreach (var invNode in declNode.DescendantNodes())
		{
			var tn = invNode.GetType().Name;
			if (tn != "InvocationExpressionSyntax"     &&
			    tn != "ObjectCreationExpressionSyntax" &&
			    tn != "ImplicitObjectCreationExpressionSyntax") continue;

			var calledSym = semanticModel.GetSymbolInfo(invNode).Symbol;
			if (calledSym is null) continue;

			var calleeLoc = calledSym.Locations.FirstOrDefault(l => l.IsInSource);
			if (calleeLoc is null)
			{
				var srcDef = await SymbolFinder.FindSourceDefinitionAsync(calledSym, roslynSolution);
				if (srcDef is null) continue;
				calledSym = srcDef;
				calleeLoc = srcDef.Locations.FirstOrDefault(l => l.IsInSource);
				if (calleeLoc is null) continue;
			}

			var nodeLineSpan = invNode.GetLocation().GetLineSpan();
			var nodeRange    = CreateRange(
				nodeLineSpan.StartLinePosition.Line, nodeLineSpan.StartLinePosition.Character,
				nodeLineSpan.EndLinePosition.Line,   nodeLineSpan.EndLinePosition.Character);

			if (!calleeToRanges.TryGetValue(calledSym, out var entry))
			{
				var cLineSpan = calleeLoc.GetLineSpan();
				var cStart    = cLineSpan.StartLinePosition;
				var cEnd      = cLineSpan.EndLinePosition;
				var cUri      = new Uri(calleeLoc.SourceTree!.FilePath).AbsoluteUri;
				entry = (new JsonObject
				{
					["name"]           = calledSym.Name,
					["kind"]           = SymbolToLspKind(calledSym),
					["uri"]            = cUri,
					["range"]          = CreateRange(cStart.Line, cStart.Character, cEnd.Line, cEnd.Character),
					["selectionRange"] = ClampedSelectionRange(cStart, cEnd, calledSym.Name.Length),
					["data"]           = new JsonObject { ["uri"] = cUri, ["line"] = cStart.Line, ["character"] = cStart.Character }
				}, new JsonArray());
				calleeToRanges[calledSym] = entry;
			}
			entry.Ranges.Add(nodeRange);
		}

		foreach (var (calleeItem, ranges) in calleeToRanges.Values)
			results.Add(new JsonObject { ["to"] = calleeItem, ["fromRanges"] = ranges });

	}
	catch (Exception ex)
	{
		await LogAsync($"OutgoingCalls failed: {ex.Message}", 2);
	}

	return results;
}

// ─── Code Actions ─────────────────────────────────────────────────────────────
async Task<JsonNode> HandleCodeActionAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var actions = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return actions;

	if (!TryReadHintRange(requestRoot, out var startLine, out _, out _, out _))
		return actions;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null)
	{
		// Roslyn couldn't resolve this document (project may have failed to load).
		// Offer basic text-based code actions so the user isn't left with nothing.
		if (documents.TryGetValue(uri, out var rawText) && !string.IsNullOrEmpty(rawText))
		{
			var isVbFallback = filePath.EndsWith(".vb", StringComparison.OrdinalIgnoreCase);
			AppendFallbackCodeActions(actions, uri, rawText, startLine, isVbFallback);
		}
		return actions;
	}

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	var isVb = filePath.EndsWith(".vb", StringComparison.OrdinalIgnoreCase);

	try
	{
		var docText       = await roslynDoc.GetTextAsync();
		var semanticModel = await roslynDoc.GetSemanticModelAsync();
		var syntaxRoot    = await roslynDoc.GetSyntaxRootAsync();
		var allDiags      = semanticModel?.GetDiagnostics() ?? ImmutableArray<Diagnostic>.Empty;

		// ── 1) Remove all unused imports ────────────────────────────────────
		var unusedImportDiags = allDiags.Where(d => d.Id == "BC50001" || d.Id == "CS8019").ToList();
		if (unusedImportDiags.Count > 0)
		{
			var edits = new JsonArray();
			foreach (var diag in unusedImportDiags)
			{
				var diagLine = diag.Location.GetLineSpan().StartLinePosition.Line;
				edits.Add(new JsonObject
				{
					["range"]   = CreateRange(diagLine, 0, diagLine + 1, 0),
					["newText"] = string.Empty
				});
			}
			actions.Add(new JsonObject
			{
				["title"] = $"Remove {unusedImportDiags.Count} unused import(s)",
				["kind"]  = "source.organizeImports",
				["edit"]  = new JsonObject { ["changes"] = new JsonObject { [uri] = edits } }
			});
		}

		// ── 2) Suppress diagnostics at the cursor line (from semantic model) ─
		var cursorDiags = allDiags
			.Where(d => d.Severity >= DiagnosticSeverity.Warning &&
			            d.Location.GetLineSpan().StartLinePosition.Line == startLine)
			.GroupBy(d => d.Id)
			.Select(g => g.First())
			.ToList();

		var lineText = startLine < docText.Lines.Count ? docText.Lines[startLine].ToString() : string.Empty;
		var indent   = new string(' ', lineText.Length - lineText.TrimStart().Length);

		foreach (var diag in cursorDiags)
		{
			var suppress = isVb ? $"#Disable Warning {diag.Id}" : $"#pragma warning disable {diag.Id}";
			actions.Add(new JsonObject
			{
				["title"] = $"Suppress {diag.Id}: {diag.GetMessage().Split('.')[0].Trim()}",
				["kind"]  = "quickfix",
				["diagnostics"] = new JsonArray { BuildLspDiagnosticNode(diag) },
				["edit"]  = new JsonObject
				{
					["changes"] = new JsonObject
					{
						[uri] = new JsonArray
						{
							new JsonObject
							{
								["range"]   = CreateRange(startLine, 0, startLine, 0),
								["newText"] = $"{indent}{suppress}{Environment.NewLine}"
							}
						}
					}
				}
			});
		}

		// ── 3) Add XML doc comment (always offered for declarations) ─────────
		if (syntaxRoot is not null && semanticModel is not null &&
		    startLine < docText.Lines.Count)
		{
			var cursorPos = docText.Lines.GetPosition(new LinePosition(startLine, 0));
			var declNode  = syntaxRoot
				.DescendantNodes()
				.Where(n => n.SpanStart >= cursorPos || n.Span.Contains(cursorPos))
				.FirstOrDefault(n =>
				{
					var tn = n.GetType().Name;
					return tn.Contains("MethodBlock")      || tn.Contains("SubBlock")          ||
					       tn.Contains("FunctionBlock")    || tn.Contains("SubStatement")       ||
					       tn.Contains("FunctionStatement")|| tn.Contains("MethodDeclaration")  ||
					       tn.Contains("ClassBlock")       || tn.Contains("ClassStatement")     ||
					       tn.Contains("ClassDeclaration") || tn.Contains("InterfaceBlock")     ||
					       tn.Contains("PropertyBlock")    || tn.Contains("PropertyStatement")  ||
					       tn.Contains("PropertyDeclaration");
				});

			if (declNode is not null)
			{
				var sym = semanticModel.GetDeclaredSymbol(declNode);
				if (sym is not null)
				{
					// Only offer if there is no existing doc comment
					var trivia = declNode.GetLeadingTrivia();
					var hasDoc  = trivia.Any(t =>
						t.GetType().Name.Contains("DocumentationCommentTrivia") ||
						t.ToString().TrimStart().StartsWith("'") && t.ToString().Contains("<summary>") ||
						t.ToString().TrimStart().StartsWith("/") && t.ToString().Contains("<summary>"));

					if (!hasDoc)
					{
						var declLine  = docText.Lines.GetLinePosition(declNode.SpanStart).Line;
						var declText  = declLine < docText.Lines.Count ? docText.Lines[declLine].ToString() : string.Empty;
						var declIndent = new string(' ', declText.Length - declText.TrimStart().Length);

						var commentLines = new List<string>();
						var prefix = isVb ? "'''" : "///";

						commentLines.Add($"{declIndent}{prefix} <summary>");
						commentLines.Add($"{declIndent}{prefix} {sym.Name}");
						commentLines.Add($"{declIndent}{prefix} </summary>");

						if (sym is IMethodSymbol methodSym2)
						{
							foreach (var p in methodSym2.Parameters)
								commentLines.Add($"{declIndent}{prefix} <param name=\"{p.Name}\"></param>");
							if (methodSym2.ReturnType.SpecialType != SpecialType.System_Void &&
							    methodSym2.MethodKind != MethodKind.Constructor)
								commentLines.Add($"{declIndent}{prefix} <returns></returns>");
						}

						var commentText = string.Join(Environment.NewLine, commentLines) + Environment.NewLine;

						actions.Add(new JsonObject
						{
							["title"] = $"Add XML doc comment",
							["kind"]  = "refactor.rewrite",
							["edit"]  = new JsonObject
							{
								["changes"] = new JsonObject
								{
									[uri] = new JsonArray
									{
										new JsonObject
										{
											["range"]   = CreateRange(declLine, 0, declLine, 0),
											["newText"] = commentText
										}
									}
								}
							}
						});
					}
				}
			}
		}

	
	}
	catch (Exception ex)
	{
		await LogAsync($"CodeAction failed: {ex.Message}", 2);
	}

	return actions;
}

// ─── Tier 3 Helpers ───────────────────────────────────────────────────────────
static bool TryReadCallHierarchyItemData(JsonElement requestRoot, out string uri, out int line, out int character)
{
	uri = string.Empty; line = character = 0;
	if (!requestRoot.TryGetProperty("params", out var p)) return false;
	if (!p.TryGetProperty("item", out var item)) return false;
	if (!item.TryGetProperty("data", out var data)) return false;
	if (!data.TryGetProperty("uri", out var uriEl)) return false;
	uri       = uriEl.GetString() ?? string.Empty;
	line      = data.TryGetProperty("line", out var lineEl) ? lineEl.GetInt32() : 0;
	character = data.TryGetProperty("character", out var charEl) ? charEl.GetInt32() : 0;
	return !string.IsNullOrWhiteSpace(uri);
}

static ISymbol? FindContainingMethodSymbol(SemanticModel model, SyntaxNode? node)
{
	while (node is not null)
	{
		var sym = model.GetDeclaredSymbol(node);
		if (sym is IMethodSymbol || sym is IPropertySymbol)
			return sym;
		node = node.Parent;
	}
	return null;
}

static bool TryReadHintRange(JsonElement requestRoot, out int startLine, out int startChar, out int endLine, out int endChar)
{
	startLine = startChar = endLine = endChar = 0;
	if (!requestRoot.TryGetProperty("params", out var p)) return false;
	if (!p.TryGetProperty("range", out var r)) return false;
	if (!r.TryGetProperty("start", out var s) || !r.TryGetProperty("end", out var e)) return false;
	startLine = s.TryGetProperty("line", out var sl) ? sl.GetInt32() : 0;
	startChar = s.TryGetProperty("character", out var sc) ? sc.GetInt32() : 0;
	endLine   = e.TryGetProperty("line", out var el) ? el.GetInt32() : 0;
	endChar   = e.TryGetProperty("character", out var ec) ? ec.GetInt32() : 0;
	return true;
}

static JsonObject BuildLspDiagnosticNode(Diagnostic diag)
{
	var span  = diag.Location.GetLineSpan();
	var start = span.StartLinePosition;
	var end   = span.EndLinePosition;
	return new JsonObject
	{
		["range"]    = CreateRange(start.Line, start.Character, end.Line, end.Character),
		["severity"] = diag.Severity switch { DiagnosticSeverity.Error => 1, DiagnosticSeverity.Warning => 2, DiagnosticSeverity.Info => 3, _ => 4 },
		["code"]     = diag.Id,
		["message"]  = diag.GetMessage()
	};
}

static int SymbolToLspKind(ISymbol sym) => sym switch{
	INamedTypeSymbol { TypeKind: TypeKind.Class }        => 5,
	INamedTypeSymbol { TypeKind: TypeKind.Enum }         => 10,
	INamedTypeSymbol { TypeKind: TypeKind.Interface }    => 11,
	INamedTypeSymbol { TypeKind: TypeKind.Struct }       => 23,
	INamedTypeSymbol                                      => 2,
	IMethodSymbol { MethodKind: MethodKind.Constructor } => 9,
	IMethodSymbol                                         => 6,
	IPropertySymbol                                       => 7,
	IFieldSymbol { IsConst: true }                        => 14,
	IFieldSymbol                                          => 8,
	IEventSymbol                                          => 24,
	INamespaceSymbol                                      => 3,
	_                                                     => 13
};

// ─── Fallback Code Actions (text-based, no Roslyn project needed) ─────────────
void AppendFallbackCodeActions(JsonArray actions, string uri, string text, int cursorLine, bool isVb)
{
	var lines = text.Split('\n');
	if (cursorLine >= lines.Length) return;

	var lineText = lines[cursorLine];
	var indent = new string(' ', lineText.Length - lineText.TrimStart().Length);

	// Offer to comment out the current line
	var commentPrefix = isVb ? "' " : "// ";
	var trimmedLine = lineText.TrimStart();
	if (!trimmedLine.StartsWith(commentPrefix.TrimEnd()) && trimmedLine.Length > 0)
	{
		actions.Add(new JsonObject
		{
			["title"] = "Comment out line",
			["kind"] = "quickfix",
			["edit"] = new JsonObject
			{
				["changes"] = new JsonObject
				{
					[uri] = new JsonArray
					{
						new JsonObject
						{
							["range"] = CreateRange(cursorLine, 0, cursorLine, lineText.TrimEnd('\r').Length),
							["newText"] = $"{indent}{commentPrefix}{trimmedLine.TrimEnd('\r')}"
						}
					}
				}
			}
		});
	}

	// Offer to add a Region around the selection
	var regionStart = isVb ? $"{indent}#Region \"Region\"" : $"{indent}#region Region";
	var regionEnd = isVb ? $"{indent}#End Region" : $"{indent}#endregion";
	actions.Add(new JsonObject
	{
		["title"] = "Wrap in region",
		["kind"] = "refactor",
		["edit"] = new JsonObject
		{
			["changes"] = new JsonObject
			{
				[uri] = new JsonArray
				{
					new JsonObject
					{
						["range"] = CreateRange(cursorLine, 0, cursorLine, 0),
						["newText"] = $"{regionStart}{Environment.NewLine}"
					},
					new JsonObject
					{
						["range"] = CreateRange(cursorLine + 1, 0, cursorLine + 1, 0),
						["newText"] = $"{regionEnd}{Environment.NewLine}"
					}
				}
			}
		}
	});
}

// ─── Formatting ───────────────────────────────────────────────────────────────
async Task<JsonNode> HandleFormattingAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents, (int startLine, int startChar, int endLine, int endChar)? range)
{
	var empty = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return empty;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return empty;

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var oldText = await roslynDoc.GetTextAsync();
		Document formattedDoc;
		if (range is not null)
		{
			var r = range.Value;
			var safeEndLine = Math.Clamp(r.endLine, 0, oldText.Lines.Count - 1);
			var safeEndChar = Math.Clamp(r.endChar, 0, oldText.Lines[safeEndLine].Span.Length);
			var spanStart = oldText.Lines.GetPosition(new LinePosition(Math.Clamp(r.startLine, 0, oldText.Lines.Count - 1), r.startChar));
			var spanEnd   = oldText.Lines.GetPosition(new LinePosition(safeEndLine, safeEndChar));
			formattedDoc = await Formatter.FormatAsync(roslynDoc, TextSpan.FromBounds(spanStart, spanEnd));
		}
		else
		{
			formattedDoc = await Formatter.FormatAsync(roslynDoc);
		}

		var newText = await formattedDoc.GetTextAsync();
		var changes = newText.GetTextChanges(oldText);

		var edits = new JsonArray();
		foreach (var change in changes)
		{
			var startPos = oldText.Lines.GetLinePosition(change.Span.Start);
			var endPos   = oldText.Lines.GetLinePosition(change.Span.End);
			edits.Add(new JsonObject
			{
				["range"]   = CreateRange(startPos.Line, startPos.Character, endPos.Line, endPos.Character),
				["newText"] = change.NewText ?? string.Empty
			});
		}
		return edits;
	}
	catch (Exception ex)
	{
		await LogAsync($"Formatting failed: {ex.Message}", 2);
		return empty;
	}
}

static (int startLine, int startChar, int endLine, int endChar)? TryReadFormattingRange(JsonElement requestRoot)
{
	if (!requestRoot.TryGetProperty("params", out var p)) return null;
	if (!p.TryGetProperty("range", out var r)) return null;
	if (!r.TryGetProperty("start", out var s) || !r.TryGetProperty("end", out var e)) return null;
	return (
		s.TryGetProperty("line", out var sl) ? sl.GetInt32() : 0,
		s.TryGetProperty("character", out var sc) ? sc.GetInt32() : 0,
		e.TryGetProperty("line", out var el) ? el.GetInt32() : 0,
		e.TryGetProperty("character", out var ec) ? ec.GetInt32() : 0
	);
}

// ─── Selection Ranges ─────────────────────────────────────────────────────────
async Task<JsonNode> HandleSelectionRangeAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null || !TryReadUri(requestRoot, out var uri) || !TryGetFilePathFromUri(uri, out var filePath))
		return empty;

	if (!requestRoot.TryGetProperty("params", out var pEl) || !pEl.TryGetProperty("positions", out var positionsEl))
		return empty;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return empty;

	if (documents.TryGetValue(uri, out var liveText))
		roslynDoc = roslynDoc.WithText(SourceText.From(liveText, Encoding.UTF8));

	try
	{
		var syntaxRoot = await roslynDoc.GetSyntaxRootAsync();
		var srcText    = await roslynDoc.GetTextAsync();
		if (syntaxRoot is null) return empty;

		var results = new JsonArray();
		foreach (var pos in positionsEl.EnumerateArray())
		{
			var ln  = pos.TryGetProperty("line", out var lEl) ? lEl.GetInt32() : 0;
			var ch  = pos.TryGetProperty("character", out var cEl) ? cEl.GetInt32() : 0;
			var safeChar = Math.Clamp(ch, 0, srcText.Lines[Math.Clamp(ln, 0, srcText.Lines.Count - 1)].Span.Length);
			var offset   = srcText.Lines.GetPosition(new LinePosition(Math.Clamp(ln, 0, srcText.Lines.Count - 1), safeChar));

			var token = syntaxRoot.FindToken(Math.Clamp(offset, 0, Math.Max(0, syntaxRoot.FullSpan.End - 1)));
			var node  = token.Parent;

			// Collect ancestors innermost-first, then build SelectionRange chain outward.
			// LSP requires each entry's `parent` to be WIDER than the entry itself.
			var ancestors = new List<SyntaxNode>();
			var cur = node;
			while (cur is not null)
			{
				ancestors.Add(cur);
				cur = cur.Parent;
			}

			// Build from outermost → innermost so each entry's parent is already the outer entry.
			JsonObject? chain = null;
			for (var i = ancestors.Count - 1; i >= 0; i--)
			{
				var span    = ancestors[i].Span;
				var startLp = srcText.Lines.GetLinePosition(span.Start);
				var endLp   = srcText.Lines.GetLinePosition(span.End);
				var entry   = new JsonObject
				{
					["range"] = CreateRange(startLp.Line, startLp.Character, endLp.Line, endLp.Character)
				};
				if (chain is not null)
					entry["parent"] = chain;  // chain is the broader (outer) parent
				chain = entry;
			}

			results.Add(chain ?? new JsonObject { ["range"] = CreateRange(ln, ch, ln, ch) });
		}
		return results;
	}
	catch (Exception ex)
	{
		await LogAsync($"SelectionRange failed: {ex.Message}", 2);
		return empty;
	}
}

// ─── Document Links ───────────────────────────────────────────────────────────
async Task<JsonNode> HandleDocumentLinkAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();
	if (!TryReadUri(requestRoot, out var uri)) return empty;

	string? sourceText = null;
	if (documents.TryGetValue(uri, out var liveText))
	{
		sourceText = liveText;
	}
	else
	{
		await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
		if (roslynSolution is not null && TryGetFilePathFromUri(uri, out var fp))
		{
			var docId = roslynSolution.GetDocumentIdsWithFilePath(fp).FirstOrDefault();
			var doc   = docId is not null ? roslynSolution.GetDocument(docId) : null;
			if (doc is not null)
				sourceText = (await doc.GetTextAsync()).ToString();
		}
	}

	if (sourceText is null) return empty;

	try
	{
		// Scan only comment trivia and string literals to avoid false positives.
		Document? roslynDocForLinks = null;
		await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
		if (roslynSolution is not null && TryGetFilePathFromUri(uri, out var filePath2))
		{
			var docId2 = roslynSolution.GetDocumentIdsWithFilePath(filePath2).FirstOrDefault();
			roslynDocForLinks = docId2 is not null ? roslynSolution.GetDocument(docId2) : null;
			if (roslynDocForLinks is not null && documents.TryGetValue(uri, out var lt))
				roslynDocForLinks = roslynDocForLinks.WithText(SourceText.From(lt, Encoding.UTF8));
		}

		var links   = new JsonArray();
		var urlRegex = new Regex(@"https?://[^\s""'<>\)\]]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);

		SourceText? src = null;
		if (roslynDocForLinks is not null)
		{
			var root = await roslynDocForLinks.GetSyntaxRootAsync();
			src = await roslynDocForLinks.GetTextAsync();
			if (root is not null && src is not null)
			{
				var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
				// Walk trivia (comments) and string literals
				foreach (var trivia in root.DescendantTrivia())
				{
					var trivText = trivia.ToString();
					foreach (Match m in urlRegex.Matches(trivText))
					{
						var absoluteStart = trivia.FullSpan.Start + trivText.IndexOf(m.Value, StringComparison.Ordinal);
						var startLp = src.Lines.GetLinePosition(absoluteStart);
						var endLp   = src.Lines.GetLinePosition(absoluteStart + m.Length);
						var key     = m.Value;
						if (seen.Add(key))
							links.Add(new JsonObject { ["range"] = CreateRange(startLp.Line, startLp.Character, endLp.Line, endLp.Character), ["target"] = m.Value });
					}
				}
				foreach (var token in root.DescendantTokens())
				{
					var rawToken = token.ToString();
					// Only inspect string-like tokens (starts with a quote char)
					if (!rawToken.StartsWith("\"") && !rawToken.StartsWith("@\"") && !rawToken.StartsWith("$\"")) continue;
					var tokInner = rawToken.Trim('"', '@', '$');
					foreach (Match m in urlRegex.Matches(tokInner))
					{
						var innerOffset = rawToken.IndexOf(m.Value, StringComparison.Ordinal);
						if (innerOffset < 0) continue;
						var absoluteStart = token.SpanStart + innerOffset;
						var startLp = src.Lines.GetLinePosition(absoluteStart);
						var endLp   = src.Lines.GetLinePosition(absoluteStart + m.Length);
						var key     = $"{startLp.Line}:{startLp.Character}";
						if (seen.Add(key))
							links.Add(new JsonObject { ["range"] = CreateRange(startLp.Line, startLp.Character, endLp.Line, endLp.Character), ["target"] = m.Value });
					}
				}
				return links;
			}
		}

		// Fallback: full text scan when Roslyn unavailable
		var lines = SplitLines(sourceText);
		for (var li = 0; li < lines.Count; li++)
		{
			foreach (Match m in urlRegex.Matches(lines[li]))
				links.Add(new JsonObject { ["range"] = CreateRange(li, m.Index, li, m.Index + m.Length), ["target"] = m.Value });
		}
		return links;
	}
	catch (Exception ex)
	{
		await LogAsync($"DocumentLink failed: {ex.Message}", 2);
		return empty;
	}
}

// ─── Type Hierarchy ───────────────────────────────────────────────────────────
async Task<JsonNode> HandlePrepareTypeHierarchyAsync(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var empty = new JsonArray();
	var context = await TryResolveRoslynContextAsync(requestRoot);
	if (context is null) return empty;

	if (!TryReadPosition(requestRoot, out var uri, out _, out _)) return empty;
	var document = context.Value.Document;
	if (documents.TryGetValue(uri, out var liveText))
		document = document.WithText(SourceText.From(liveText, Encoding.UTF8));

	var symbol = await ResolveSymbolAtPositionAsync(document, context.Value.Position);
	// Type hierarchy only applies to named types
	if (symbol is not INamedTypeSymbol typeSymbol)
	{
		// If the symbol is a member, try the containing type
		if (symbol?.ContainingType is INamedTypeSymbol containingType)
			typeSymbol = containingType;
		else
			return empty;
	}

	var loc = typeSymbol.Locations.FirstOrDefault(l => l.IsInSource);
	if (loc is null)
	{
		var srcDef = await SymbolFinder.FindSourceDefinitionAsync(typeSymbol, context.Value.Solution);
		if (srcDef is not null) loc = srcDef.Locations.FirstOrDefault(l => l.IsInSource);
		if (loc is null) return empty;
	}

	var lineSpan = loc.GetLineSpan();
	var start    = lineSpan.StartLinePosition;
	var end      = lineSpan.EndLinePosition;
	var symUri   = new Uri(loc.SourceTree!.FilePath).AbsoluteUri;

	var item = new JsonObject
	{
		["name"]           = typeSymbol.Name,
		["kind"]           = SymbolToLspKind(typeSymbol),
		["uri"]            = symUri,
		["range"]          = CreateRange(start.Line, start.Character, end.Line, end.Character),
		["selectionRange"] = CreateRange(start.Line, start.Character, start.Line, start.Character + typeSymbol.Name.Length),
		["data"]           = new JsonObject { ["uri"] = symUri, ["line"] = start.Line, ["character"] = start.Character }
	};

	return new JsonArray { item };
}

async Task<JsonNode> HandleTypeHierarchySupertypesAsync(JsonElement requestRoot)
{
	var results = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null) return results;

	// Reuse same item.data structure as call hierarchy
	if (!TryReadCallHierarchyItemData(requestRoot, out var fileUri, out var fileLine, out var fileChar)) return results;
	if (!TryGetFilePathFromUri(fileUri, out var filePath)) return results;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return results;

	var docText  = await roslynDoc.GetTextAsync();
	var safeChar = Math.Clamp(fileChar, 0, docText.Lines[fileLine].Span.Length);
	var position = docText.Lines.GetPosition(new LinePosition(fileLine, safeChar));

	var symbol = await ResolveSymbolAtPositionAsync(roslynDoc, position);
	if (symbol is not INamedTypeSymbol typeSymbol) return results;

	try
	{
		var supertypes = new List<INamedTypeSymbol>();

		// Base class (excluding System.Object / System.ValueType / System.Enum etc.)
		if (typeSymbol.BaseType is INamedTypeSymbol baseType &&
		    baseType.SpecialType != SpecialType.System_Object &&
		    baseType.SpecialType != SpecialType.System_ValueType &&
		    baseType.SpecialType != SpecialType.System_Enum)
		{
			supertypes.Add(baseType);
		}

		// Implemented interfaces
		supertypes.AddRange(typeSymbol.Interfaces);

		foreach (var supertype in supertypes)
		{
			var loc = supertype.Locations.FirstOrDefault(l => l.IsInSource);
			if (loc is null)
			{
				var srcDef = await SymbolFinder.FindSourceDefinitionAsync(supertype, roslynSolution);
				if (srcDef is not null) loc = srcDef.Locations.FirstOrDefault(l => l.IsInSource);
			}
			if (loc is null) continue;

			var lineSpan = loc.GetLineSpan();
			var s        = lineSpan.StartLinePosition;
			var e        = lineSpan.EndLinePosition;
			var sUri     = new Uri(loc.SourceTree!.FilePath).AbsoluteUri;

			results.Add(new JsonObject
			{
				["name"]           = supertype.Name,
				["kind"]           = SymbolToLspKind(supertype),
				["uri"]            = sUri,
				["range"]          = CreateRange(s.Line, s.Character, e.Line, e.Character),
				["selectionRange"] = CreateRange(s.Line, s.Character, s.Line, s.Character + supertype.Name.Length),
				["data"]           = new JsonObject { ["uri"] = sUri, ["line"] = s.Line, ["character"] = s.Character }
			});
		}
	}
	catch (Exception ex)
	{
		await LogAsync($"TypeHierarchy supertypes failed: {ex.Message}", 2);
	}

	return results;
}

async Task<JsonNode> HandleTypeHierarchySubtypesAsync(JsonElement requestRoot)
{
	var results = new JsonArray();
	await EnsureRoslynWorkspaceLoadedAsync(forceReload: false);
	if (roslynSolution is null) return results;

	if (!TryReadCallHierarchyItemData(requestRoot, out var fileUri, out var fileLine, out var fileChar)) return results;
	if (!TryGetFilePathFromUri(fileUri, out var filePath)) return results;

	var docId = roslynSolution.GetDocumentIdsWithFilePath(filePath).FirstOrDefault();
	var roslynDoc = docId is not null ? roslynSolution.GetDocument(docId) : null;
	if (roslynDoc is null) return results;

	var docText  = await roslynDoc.GetTextAsync();
	var safeChar = Math.Clamp(fileChar, 0, docText.Lines[fileLine].Span.Length);
	var position = docText.Lines.GetPosition(new LinePosition(fileLine, safeChar));

	var symbol = await ResolveSymbolAtPositionAsync(roslynDoc, position);
	if (symbol is not INamedTypeSymbol typeSymbol) return results;

	try
	{
		var subtypes = new List<INamedTypeSymbol>();

		if (typeSymbol.TypeKind == TypeKind.Interface)
		{
			// Implementations of the interface
			var impls = await SymbolFinder.FindImplementationsAsync(typeSymbol, roslynSolution);
			subtypes.AddRange(impls.OfType<INamedTypeSymbol>());
		}
		else
		{
			// Direct subclasses (not transitive)
			var derived = await SymbolFinder.FindDerivedClassesAsync(typeSymbol, roslynSolution, transitive: false);
			subtypes.AddRange(derived);
		}

		var capped = false;
		foreach (var subtype in subtypes)
		{
			if (capped) break;
			var loc = subtype.Locations.FirstOrDefault(l => l.IsInSource);
			if (loc is null) continue;

			var lineSpan = loc.GetLineSpan();
			var s        = lineSpan.StartLinePosition;
			var e        = lineSpan.EndLinePosition;
			var sUri     = new Uri(loc.SourceTree!.FilePath).AbsoluteUri;

			results.Add(new JsonObject
			{
				["name"]           = subtype.Name,
				["kind"]           = SymbolToLspKind(subtype),
				["uri"]            = sUri,
				["range"]          = CreateRange(s.Line, s.Character, e.Line, e.Character),
				["selectionRange"] = CreateRange(s.Line, s.Character, s.Line, s.Character + subtype.Name.Length),
				["data"]           = new JsonObject { ["uri"] = sUri, ["line"] = s.Line, ["character"] = s.Character }
			});
			if (results.Count >= 200) capped = true;
		}
	}
	catch (Exception ex)
	{
		await LogAsync($"TypeHierarchy subtypes failed: {ex.Message}", 2);
	}

	return results;
}

static async Task WriteMessageAsync(Stream output, JsonObject message)
{
	var payload = Encoding.UTF8.GetBytes(message.ToJsonString());
	var header = Encoding.ASCII.GetBytes($"Content-Length: {payload.Length}\r\n\r\n");
	await output.WriteAsync(header);
	await output.WriteAsync(payload);
	await output.FlushAsync();
}
