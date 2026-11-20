using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;

namespace Mcp4Office
{
    internal sealed class ToolAnnotations
    {
        public bool ReadOnlyHint { get; set; }
        public bool DestructiveHint { get; set; }
        public bool IdempotentHint { get; set; }
        public bool OpenWorldHint { get; set; }
    }

    internal sealed class ToolDefinition
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public JsonObject InputSchema { get; set; }
        public ToolAnnotations Annotations { get; set; }
        public Func<JsonObject, string> Handler { get; set; }
    }

    internal sealed class McpServer
    {
        private const string ServerName = "mcp4office";
        private const string ServerVersion = "0.5.5";
        private const int MaxLineChars = 2 * 1024 * 1024; // 2 MB max per riga JSON-RPC
        private static readonly string[] SupportedProtocolVersions = { "2025-03-26", "2024-11-05" };

        private readonly bool _readOnlyOnly;
        public bool IsReadOnly => _readOnlyOnly;
        private bool _initialized;
        private readonly Dictionary<string, ToolDefinition> _tools =
            new Dictionary<string, ToolDefinition>(StringComparer.Ordinal);

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = false,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            TypeInfoResolver = new DefaultJsonTypeInfoResolver()
        };

        private static readonly JsonSerializerOptions IndentedJsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            TypeInfoResolver = new DefaultJsonTypeInfoResolver()
        };

        public McpServer(bool readOnlyOnly = false)
        {
            _readOnlyOnly = readOnlyOnly;
        }

        public void RegisterTool(
            string name,
            string description,
            JsonObject inputSchema,
            ToolAnnotations annotations,
            Func<JsonObject, string> handler)
        {
            var ann = annotations ?? new ToolAnnotations { ReadOnlyHint = false, DestructiveHint = true };
            if (_readOnlyOnly && !ann.ReadOnlyHint)
                return;

            _tools[name] = new ToolDefinition
            {
                Name = name,
                Description = description,
                InputSchema = inputSchema ?? EmptySchema(),
                Annotations = ann,
                Handler = handler
            };
        }

        public static ToolAnnotations ReadOnly() => new ToolAnnotations
        {
            ReadOnlyHint = true,
            DestructiveHint = false,
            IdempotentHint = true,
            OpenWorldHint = false
        };

        public static ToolAnnotations Mutating(bool destructive = true, bool idempotent = false) => new ToolAnnotations
        {
            ReadOnlyHint = false,
            DestructiveHint = destructive,
            IdempotentHint = idempotent,
            OpenWorldHint = false
        };

        public static JsonObject EmptySchema()
        {
            return new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject(),
                ["additionalProperties"] = false
            };
        }

        public static JsonObject Schema(Action<JsonObject, JsonArray> configure)
        {
            var props = new JsonObject();
            var required = new JsonArray();
            configure(props, required);
            var schema = new JsonObject
            {
                ["type"] = "object",
                ["properties"] = props,
                ["additionalProperties"] = false
            };
            if (required.Count > 0)
                schema["required"] = required;
            return schema;
        }

        public static JsonObject Prop(string type, string description)
        {
            return new JsonObject
            {
                ["type"] = type,
                ["description"] = description
            };
        }

        public static JsonObject IntProp(string description, int min, int max)
        {
            return new JsonObject
            {
                ["type"] = "integer",
                ["description"] = description,
                ["minimum"] = min,
                ["maximum"] = max
            };
        }

        public static JsonObject EnumProp(string description, params string[] values)
        {
            var arr = new JsonArray();
            foreach (var v in values) arr.Add(v);
            return new JsonObject
            {
                ["type"] = "string",
                ["description"] = description,
                ["enum"] = arr
            };
        }

        public static string GetString(JsonObject args, string key, string defaultValue = "")
        {
            if (args == null || !args.TryGetPropertyValue(key, out var node) || node == null)
                return defaultValue;
            if (node is JsonValue v && v.TryGetValue<string>(out var s))
                return s;
            throw new ArgumentException($"Il parametro '{key}' deve essere una stringa.");
        }

        public static int GetInt(JsonObject args, string key, int defaultValue)
        {
            if (args == null || !args.TryGetPropertyValue(key, out var node) || node == null)
                return defaultValue;
            if (node is JsonValue v && v.TryGetValue<int>(out var i))
                return i;
            throw new ArgumentException($"Il parametro '{key}' deve essere un intero.");
        }

        public static double GetDouble(JsonObject args, string key, double defaultValue)
        {
            if (args == null || !args.TryGetPropertyValue(key, out var node) || node == null)
                return defaultValue;
            if (node is JsonValue v && v.TryGetValue<double>(out var d))
                return d;
            throw new ArgumentException($"Il parametro '{key}' deve essere un numero.");
        }

        public static bool GetBool(JsonObject args, string key, bool defaultValue = false)
        {
            if (args == null || !args.TryGetPropertyValue(key, out var node) || node == null)
                return defaultValue;
            if (node is JsonValue v && v.TryGetValue<bool>(out var b))
                return b;
            throw new ArgumentException($"Il parametro '{key}' deve essere un booleano.");
        }

        public string DumpToolsJson()
        {
            var arr = BuildToolsArray();
            return arr.ToJsonString(IndentedJsonOptions);
        }

        public void RunStdio(TextReader input, TextWriter output, TextWriter error)
        {
            while (true)
            {
                string line = ReadBoundedLine(input, MaxLineChars, out bool exceeded);
                if (line == null && !exceeded)
                    break;

                if (exceeded)
                {
                    var errResp = MakeError(null, -32600, $"Messaggio JSON-RPC troppo grande (limite {MaxLineChars} caratteri).");
                    output.WriteLine(errResp.ToJsonString(JsonOptions));
                    output.Flush();
                    continue;
                }

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                JsonNode response = null;
                try
                {
                    response = HandleMessage(line);
                }
                catch (Exception ex)
                {
                    error.WriteLine("[mcp4office] Errore gestione messaggio: " + ex.Message);
                    response = MakeError(null, -32603, "Errore interno del server: " + ex.Message);
                }

                if (response != null)
                {
                    output.WriteLine(response.ToJsonString(JsonOptions));
                    output.Flush();
                }
            }
        }

        private static string ReadBoundedLine(TextReader reader, int maxChars, out bool exceeded)
        {
            exceeded = false;
            var sb = new StringBuilder();
            int ch;
            while ((ch = reader.Read()) != -1)
            {
                if (ch == '\n')
                    return exceeded ? string.Empty : sb.ToString();
                if (ch == '\r')
                    continue;
                if (!exceeded)
                {
                    if (sb.Length >= maxChars)
                        exceeded = true;
                    else
                        sb.Append((char)ch);
                }
            }
            if (exceeded) return string.Empty;
            return sb.Length > 0 ? sb.ToString() : null;
        }

        internal JsonNode HandleMessage(string rawJson)
        {
            JsonObject msg;
            try
            {
                msg = JsonNode.Parse(rawJson) as JsonObject;
            }
            catch (Exception ex)
            {
                return MakeError(null, -32700, "JSON non valido: " + ex.Message);
            }

            if (msg == null)
                return MakeError(null, -32600, "Richiesta JSON-RPC non valida.");

            // Validazione envelope JSON-RPC 2.0
            if (!msg.TryGetPropertyValue("jsonrpc", out var verNode) ||
                !(verNode is JsonValue verVal) ||
                !verVal.TryGetValue<string>(out var verStr) ||
                verStr != "2.0")
            {
                return MakeError(null, -32600, "Campo 'jsonrpc' obbligatorio e deve essere '2.0'.");
            }

            JsonNode id = null;
            bool hasId = msg.TryGetPropertyValue("id", out id);
            if (hasId && !IsValidJsonRpcId(id))
                return MakeError(null, -32600, "Campo 'id' JSON-RPC deve essere una stringa o un numero.");

            if (!msg.TryGetPropertyValue("method", out var mNode) ||
                !(mNode is JsonValue mVal) ||
                !mVal.TryGetValue<string>(out var method) ||
                string.IsNullOrWhiteSpace(method))
            {
                return hasId ? MakeError(id, -32600, "Campo 'method' mancante o non valido.") : null;
            }

            JsonObject @params = null;
            if (msg.TryGetPropertyValue("params", out var pNode) && pNode != null)
            {
                @params = pNode as JsonObject;
                if (@params == null)
                    return hasId ? MakeError(id, -32602, "Campo 'params' deve essere un oggetto JSON.") : null;
            }

            // Notifiche JSON-RPC (senza id): nessuna risposta.
            if (!hasId)
                return null;

            switch (method)
            {
                case "initialize":
                {
                    string requestedVersion = null;
                    if (@params != null && @params.TryGetPropertyValue("protocolVersion", out var pvNode) && pvNode != null)
                    {
                        if (!(pvNode is JsonValue pvVal) || !pvVal.TryGetValue<string>(out requestedVersion))
                            return MakeError(id, -32602, "'protocolVersion' deve essere una stringa.");
                    }

                    // Negozia solo versioni realmente supportate.
                    string negotiated = SupportedProtocolVersions.Contains(requestedVersion, StringComparer.Ordinal)
                        ? requestedVersion
                        : SupportedProtocolVersions[0];

                    _initialized = true;
                    var result = new JsonObject
                    {
                        ["protocolVersion"] = negotiated,
                        ["capabilities"] = new JsonObject
                        {
                            ["tools"] = new JsonObject { ["listChanged"] = false }
                        },
                        ["serverInfo"] = new JsonObject
                        {
                            ["name"] = ServerName,
                            ["version"] = ServerVersion
                        }
                    };
                    return MakeResult(id, result);
                }

                case "ping":
                    return MakeResult(id, new JsonObject());

                case "tools/list":
                    if (!_initialized)
                        return MakeError(id, -32002, "Server non inizializzato: chiamare 'initialize' prima di 'tools/list'.");
                    return MakeResult(id, new JsonObject { ["tools"] = BuildToolsArray() });

                case "tools/call":
                    if (!_initialized)
                        return MakeError(id, -32002, "Server non inizializzato: chiamare 'initialize' prima di 'tools/call'.");
                    return HandleToolCall(id, @params);

                default:
                    return MakeError(id, -32601, $"Metodo non supportato: {method}");
            }
        }

        private static bool IsValidJsonRpcId(JsonNode id)
        {
            if (!(id is JsonValue v)) return false;
            return v.TryGetValue<string>(out _) || v.TryGetValue<long>(out _) || v.TryGetValue<double>(out _);
        }

        private JsonNode HandleToolCall(JsonNode id, JsonObject @params)
        {
            if (@params == null ||
                !@params.TryGetPropertyValue("name", out var nameNode) ||
                !(nameNode is JsonValue nameVal) ||
                !nameVal.TryGetValue<string>(out var toolName) ||
                string.IsNullOrWhiteSpace(toolName))
            {
                return MakeError(id, -32602, "Parametro 'name' obbligatorio per tools/call.");
            }

            if (!_tools.TryGetValue(toolName, out var tool))
                return MakeError(id, -32602, $"Tool sconosciuto: {toolName}");

            JsonObject arguments = new JsonObject();
            if (@params.TryGetPropertyValue("arguments", out var aNode) && aNode != null)
            {
                arguments = aNode as JsonObject;
                if (arguments == null)
                    return MakeError(id, -32602, "Il campo 'arguments' deve essere un oggetto JSON.");
            }

            // Validazione rigorosa contro lo schema PRIMA di toccare COM o disco.
            string validationError = ValidateArguments(tool.InputSchema, arguments);
            if (validationError != null)
                return MakeError(id, -32602, validationError);

            try
            {
                var text = tool.Handler(arguments) ?? string.Empty;
                var result = new JsonObject
                {
                    ["content"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "text",
                            ["text"] = text
                        }
                    },
                    ["isError"] = false
                };
                return MakeResult(id, result);
            }
            catch (Exception ex)
            {
                var result = new JsonObject
                {
                    ["content"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "text",
                            ["text"] = "Errore: " + ex.Message
                        }
                    },
                    ["isError"] = true
                };
                return MakeResult(id, result);
            }
            finally
            {
                ComHelper.CollectComGarbage();
            }
        }

        internal static string ValidateArguments(JsonObject schema, JsonObject args)
        {
            if (schema == null) return null;

            var props = schema.TryGetPropertyValue("properties", out var pNode) ? pNode as JsonObject : null;
            var required = schema.TryGetPropertyValue("required", out var rNode) ? rNode as JsonArray : null;
            bool disallowAdditional = schema.TryGetPropertyValue("additionalProperties", out var apNode) &&
                                      apNode is JsonValue apVal && apVal.TryGetValue<bool>(out var apBool) && !apBool;

            if (required != null)
            {
                foreach (var reqItem in required)
                {
                    string reqName = reqItem?.ToString();
                    if (string.IsNullOrEmpty(reqName)) continue;
                    if (!args.TryGetPropertyValue(reqName, out var val) || val == null)
                        return $"Parametro obbligatorio mancante: '{reqName}'.";
                }
            }

            foreach (var kv in args)
            {
                string key = kv.Key;
                JsonNode val = kv.Value;

                if (props == null || !props.TryGetPropertyValue(key, out var propSchemaNode) || !(propSchemaNode is JsonObject propSchema))
                {
                    if (disallowAdditional)
                        return $"Parametro non riconosciuto: '{key}'.";
                    continue;
                }

                if (val == null)
                    return $"Il parametro '{key}' non può essere null.";

                string expectedType = propSchema.TryGetPropertyValue("type", out var tNode) ? tNode?.ToString() : null;
                switch (expectedType)
                {
                    case "string":
                    {
                        if (!(val is JsonValue jv) || !jv.TryGetValue<string>(out var str))
                            return $"Il parametro '{key}' deve essere una stringa.";
                        if (propSchema.TryGetPropertyValue("enum", out var enumNode) && enumNode is JsonArray enumArr)
                        {
                            bool ok = enumArr.Any(e => string.Equals(e?.ToString(), str, StringComparison.Ordinal));
                            if (!ok)
                                return $"Valore non valido per '{key}': '{str}'. Valori ammessi: {string.Join(", ", enumArr.Select(e => e?.ToString()))}.";
                        }
                        break;
                    }
                    case "integer":
                    {
                        if (!(val is JsonValue jv) || !jv.TryGetValue<int>(out var intVal))
                            return $"Il parametro '{key}' deve essere un intero.";
                        if (propSchema.TryGetPropertyValue("minimum", out var minN) && minN is JsonValue minV && minV.TryGetValue<int>(out var minVal) && intVal < minVal)
                            return $"Il parametro '{key}' deve essere >= {minVal}.";
                        if (propSchema.TryGetPropertyValue("maximum", out var maxN) && maxN is JsonValue maxV && maxV.TryGetValue<int>(out var maxVal) && intVal > maxVal)
                            return $"Il parametro '{key}' deve essere <= {maxVal}.";
                        break;
                    }
                    case "number":
                    {
                        if (!(val is JsonValue jv) || !jv.TryGetValue<double>(out var numVal) || double.IsNaN(numVal) || double.IsInfinity(numVal))
                            return $"Il parametro '{key}' deve essere un numero finito.";
                        if (propSchema.TryGetPropertyValue("minimum", out var minN) && minN is JsonValue minV && minV.TryGetValue<double>(out var minVal) && numVal < minVal)
                            return $"Il parametro '{key}' deve essere >= {minVal}.";
                        if (propSchema.TryGetPropertyValue("maximum", out var maxN) && maxN is JsonValue maxV && maxV.TryGetValue<double>(out var maxVal) && numVal > maxVal)
                            return $"Il parametro '{key}' deve essere <= {maxVal}.";
                        break;
                    }
                    case "boolean":
                    {
                        if (!(val is JsonValue jv) || !jv.TryGetValue<bool>(out _))
                            return $"Il parametro '{key}' deve essere un booleano (true/false).";
                        break;
                    }
                    case "array":
                    {
                        if (!(val is JsonArray arr))
                            return $"Il parametro '{key}' deve essere un array.";
                        if (propSchema.TryGetPropertyValue("items", out var itemSchemaNode) && itemSchemaNode is JsonObject itemSchema)
                        {
                            string itemType = itemSchema.TryGetPropertyValue("type", out var itN) ? itN?.ToString() : null;
                            if (itemType == "string")
                            {
                                foreach (var elem in arr)
                                {
                                    if (!(elem is JsonValue ev) || !ev.TryGetValue<string>(out _))
                                        return $"Tutti gli elementi di '{key}' devono essere stringhe.";
                                }
                            }
                            else if (itemType == "array")
                            {
                                foreach (var elem in arr)
                                {
                                    if (!(elem is JsonArray))
                                        return $"Tutti gli elementi di '{key}' devono essere array (righe).";
                                }
                            }
                            else if (itemType == "object")
                            {
                                foreach (var elem in arr)
                                {
                                    if (!(elem is JsonObject))
                                        return $"Tutti gli elementi di '{key}' devono essere oggetti JSON.";
                                }
                            }
                        }
                        break;
                    }
                }
            }

            return null;
        }

        private JsonArray BuildToolsArray()
        {
            var arr = new JsonArray();
            foreach (var tool in _tools.Values.OrderBy(t => t.Name, StringComparer.Ordinal))
            {
                arr.Add(new JsonObject
                {
                    ["name"] = tool.Name,
                    ["description"] = tool.Description,
                    ["inputSchema"] = JsonNode.Parse(tool.InputSchema.ToJsonString(JsonOptions)),
                    ["annotations"] = new JsonObject
                    {
                        ["readOnlyHint"] = tool.Annotations.ReadOnlyHint,
                        ["destructiveHint"] = tool.Annotations.DestructiveHint,
                        ["idempotentHint"] = tool.Annotations.IdempotentHint,
                        ["openWorldHint"] = tool.Annotations.OpenWorldHint
                    }
                });
            }
            return arr;
        }

        private static JsonObject MakeResult(JsonNode id, JsonNode result)
        {
            return new JsonObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id?.DeepClone(),
                ["result"] = result
            };
        }

        private static JsonObject MakeError(JsonNode id, int code, string message)
        {
            return new JsonObject
            {
                ["jsonrpc"] = "2.0",
                ["id"] = id?.DeepClone(),
                ["error"] = new JsonObject
                {
                    ["code"] = code,
                    ["message"] = message
                }
            };
        }
    }
}
