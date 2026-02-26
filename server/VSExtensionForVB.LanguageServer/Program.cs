using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

const string JsonRpcVersion = "2.0";

var documents = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var standardInput = Console.OpenStandardInput();
var standardOutput = Console.OpenStandardOutput();

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
					["completionProvider"] = new JsonObject
					{
						["resolveProvider"] = false,
						["triggerCharacters"] = new JsonArray(".")
					}
				},
				["serverInfo"] = new JsonObject
				{
					["name"] = "VSExtensionForVB.LanguageServer",
					["version"] = "0.1.0-beta"
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
		}
		continue;
	}

	if (method == "textDocument/didChange")
	{
		if (TryReadDidChange(root, out var uri, out var text))
		{
			documents[uri] = text;
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
			["result"] = HandleDefinition(root, documents)
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

static JsonNode? HandleDefinition(JsonElement requestRoot, ConcurrentDictionary<string, string> documents)
{
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

static async Task WriteMessageAsync(Stream output, JsonObject message)
{
	var payload = Encoding.UTF8.GetBytes(message.ToJsonString());
	var header = Encoding.ASCII.GetBytes($"Content-Length: {payload.Length}\r\n\r\n");
	await output.WriteAsync(header);
	await output.WriteAsync(payload);
	await output.FlushAsync();
}
