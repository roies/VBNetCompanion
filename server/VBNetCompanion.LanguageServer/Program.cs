using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Microsoft.Build.Locator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.Text;

const string JsonRpcVersion = "2.0";

var documents = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var standardInput = Console.OpenStandardInput();
var standardOutput = Console.OpenStandardOutput();
var workspaceRootPath = string.Empty;
var workspaceLoadGate = new SemaphoreSlim(1, 1);
MSBuildWorkspace? roslynWorkspace = null;
Solution? roslynSolution = null;

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
	await WriteMessageAsync(standardOutput, notification);
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

	if (method == "initialize")
	{
		if (TryReadInitializeWorkspacePath(root, out var resolvedWorkspacePath))
		{
			workspaceRootPath = resolvedWorkspacePath;
			await EnsureRoslynWorkspaceLoadedAsync(forceReload: true);
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
						["triggerCharacters"] = new JsonArray(".")
					}
				},
				["serverInfo"] = new JsonObject
				{
					["name"] = "VBNetCompanion.LanguageServer",
					["version"] = "0.1.0"
				}
			}
		};

		await WriteMessageAsync(standardOutput, response);
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
		await WriteMessageAsync(standardOutput, response);
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
		}
		continue;
	}

	if (method == "textDocument/didChange")
	{
		if (TryReadDidChange(root, out var uri, out var text))
		{
			documents[uri] = text;
			await UpdateRoslynDocumentTextAsync(uri, text);
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
		await WriteMessageAsync(standardOutput, response);
		continue;
	}

	if (method == "textDocument/completion")
	{
		var response = new JsonObject
		{
			["jsonrpc"] = JsonRpcVersion,
			["id"] = idNode,
			["result"] = HandleCompletion(root, documents)
		};
		await WriteMessageAsync(standardOutput, response);
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
		await WriteMessageAsync(standardOutput, response);
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
		await WriteMessageAsync(standardOutput, response);
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
		await WriteMessageAsync(standardOutput, response);
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

		if (!string.IsNullOrWhiteSpace(receiverType))
		{
			var hasContainingType = candidateLines.Any(candidateLine => Regex.IsMatch(candidateLine, $@"\b(class|module)\s+{Regex.Escape(receiverType)}\b", RegexOptions.IgnoreCase));
			if (!hasContainingType)
			{
				continue;
			}
		}

		for (var lineIndex = 0; lineIndex < candidateLines.Length; lineIndex++)
		{
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

	var csharpPattern = $@"\b{Regex.Escape(methodName)}\s*\(";
	var vbPattern = $@"\b(Function|Sub)\s+{Regex.Escape(methodName)}\b";

	return Regex.IsMatch(line, csharpPattern, RegexOptions.IgnoreCase) || Regex.IsMatch(line, vbPattern, RegexOptions.IgnoreCase);
}

static JsonNode HandleCompletion(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
	var keywords = new[]
	{
		"Class", "Module", "Namespace", "Sub", "Function", "Property", "Dim", "As", "If", "Then", "Else",
		"End", "Public", "Private", "Friend", "Protected", "Imports", "Return", "ByVal", "ByRef", "New"
	};

	var items = new JsonArray();
	foreach (var keyword in keywords)
	{
		items.Add(new JsonObject
		{
			["label"] = keyword,
			["kind"] = 14
		});
	}

	if (TryReadUri(requestRoot, out var uri) && documents.TryGetValue(uri, out var text))
	{
		foreach (var symbol in ExtractDeclaredSymbols(SplitLines(text)).Distinct(StringComparer.OrdinalIgnoreCase))
		{
			items.Add(new JsonObject
			{
				["label"] = symbol,
				["kind"] = 6
			});
		}
	}

	return new JsonObject
	{
		["isIncomplete"] = false,
		["items"] = items
	};
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
		await LogAsync($"CodeLens: solution={roslynSolution.Projects.Count()}p, file={codeLensFilePath}, roslynDoc={roslynDoc?.Name ?? "null"}");
	}
	else
	{
		await LogAsync($"CodeLens: roslynSolution={(roslynSolution is null ? "null" : "loaded")}, uri={uri}", 2);
	}

	if (!documents.TryGetValue(uri, out var sourceText))
	{
		if (roslynDoc is null) return new JsonArray();
		var roslynText = await roslynDoc.GetTextAsync();
		sourceText = roslynText.ToString();
	}

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
		var resolvedViaRoslyn = false;

		// Prefer Roslyn: resolves cross-project references (VB → C# and vice versa).
		if (roslynDoc is not null && roslynSolution is not null)
		{
			try
			{
				var docText = await roslynDoc.GetTextAsync();
				if (declaration.Line < docText.Lines.Count)
				{
					var safeChar = Math.Clamp(declaration.StartCharacter, 0, docText.Lines[declaration.Line].Span.Length);
					var position = docText.Lines.GetPosition(new LinePosition(declaration.Line, safeChar));
					var symbol = await ResolveSymbolAtPositionAsync(roslynDoc, position);
					if (symbol is not null)
					{
						var refs = await SymbolFinder.FindReferencesAsync(symbol, roslynSolution);
						foreach (var referencedSymbol in refs)
						{
							foreach (var refLocation in referencedSymbol.Locations)
							{
								referenceCount++;
								referenceLocations.Add(CreateLocationNode(refLocation.Location));
							}
						}
						await LogAsync($"CodeLens Roslyn: '{declaration.Symbol}' → {referenceCount} ref(s) via {symbol.Kind}");
						resolvedViaRoslyn = true;
					}
					else
					{
						await LogAsync($"CodeLens Roslyn: '{declaration.Symbol}' at L{declaration.Line}:{safeChar} → symbol not resolved", 2);
					}
				}
			}
			catch
			{
				// Fall through to text-search.
				referenceLocations = new JsonArray();
				referenceCount = 0;
			}
		}

		// Fallback: text-search across open documents.
		if (!resolvedViaRoslyn)
		{
			foreach (var document in documents)
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
		return null;
	}

	var sourceLocation = symbol
		.Locations
		.Where(location => location.IsInSource && location.SourceTree is not null)
		.FirstOrDefault();

	if (sourceLocation is null)
	{
		return null;
	}

	return CreateLocationNode(sourceLocation);
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
				MSBuildLocator.RegisterMSBuildPath(msbuildPath);
			}
			else
			{
				await LogAsync("Falling back to MSBuildLocator.RegisterDefaults()");
				MSBuildLocator.RegisterDefaults();
			}
		}

		var workspace = MSBuildWorkspace.Create();
		var workspaceFailures = new ConcurrentBag<string>();
		workspace.WorkspaceFailed += (_, e) => workspaceFailures.Add($"[{e.Diagnostic.Kind}] {e.Diagnostic.Message}");

		var solutionPath = Directory
			.GetFiles(workspaceRootPath, "*.sln", SearchOption.TopDirectoryOnly)
			.Concat(Directory.GetFiles(workspaceRootPath, "*.slnx", SearchOption.TopDirectoryOnly))
			.FirstOrDefault();
		if (!string.IsNullOrWhiteSpace(solutionPath))
		{
			try
			{
				roslynSolution = await workspace.OpenSolutionAsync(solutionPath);
				roslynWorkspace = workspace;
				foreach (var f in workspaceFailures) await LogAsync($"WorkspaceDiag: {f}", 2);
				await LogAsync($"Roslyn solution loaded: {solutionPath} ({roslynSolution.Projects.Count()} projects)");
				return;
			}
			catch (Exception ex)
			{
				await LogAsync($"Solution load failed ({solutionPath}): {ex.Message} — falling back to project load", 2);
				// Fall back to project loading for environments where .slnx is unsupported.
			}
		}

		var projectPaths = Directory
			.GetFiles(workspaceRootPath, "*.csproj", SearchOption.AllDirectories)
			.Concat(Directory.GetFiles(workspaceRootPath, "*.vbproj", SearchOption.AllDirectories))
			.ToList();

		if (projectPaths.Count == 0)
		{
			return;
		}

		foreach (var projectPath in projectPaths)
		{
			var project = await workspace.OpenProjectAsync(projectPath);
			roslynSolution = project.Solution;
		}

		roslynWorkspace = workspace;
		foreach (var f in workspaceFailures) await LogAsync($"WorkspaceDiag: {f}", 2);
		await LogAsync($"Roslyn projects loaded: {roslynSolution?.Projects.Count() ?? 0} project(s) via individual project scan");
	}
	catch (Exception ex)
	{
		roslynWorkspace = null;
		roslynSolution = null;
		await LogAsync($"Roslyn workspace load failed: {ex.Message}", 1);
	}
	finally
	{
		workspaceLoadGate.Release();
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
	// Try to find the latest .NET SDK by running `dotnet --list-sdks`.
	// Format: "10.0.102 [C:\Program Files\dotnet\sdk]"
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

		var bestPath = (string?)null;
		var bestVersion = (Version?)null;

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

			if (bestVersion is null || version > bestVersion)
			{
				bestVersion = version;
				bestPath = sdkPath;
			}
		}

		return bestPath;
	}
	catch
	{
		return null;
	}
}

static async Task WriteMessageAsync(Stream output, JsonObject message)
{
	var payload = Encoding.UTF8.GetBytes(message.ToJsonString());
	var header = Encoding.ASCII.GetBytes($"Content-Length: {payload.Length}\r\n\r\n");
	await output.WriteAsync(header);
	await output.WriteAsync(payload);
	await output.FlushAsync();
}
