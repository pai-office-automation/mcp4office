using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;
using System.Text.RegularExpressions;
using Tomlyn;
using Tomlyn.Model;

namespace Mcp4Office
{
    /// <summary>
    /// Configuratore sicuro per Antigravity, Claude Desktop (classico + Store/MSIX) e Codex.
    /// - JSON via System.Text.Json (profondita' illimitata).
    /// - TOML validato semanticamente con Tomlyn (TOML 1.0 completo) prima e dopo ogni modifica.
    /// - Rispetta sempre EntryOwnedByUs=false, fallisce su manifest corrotto, preserva voci se manifest assente,
    ///   supporta upgrade (.upgrading) preservando args utente, report UTF-8 con BOM leggibile da Inno Setup.
    /// </summary>
    internal static class ClientConfigurator
    {
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);
        private static readonly Encoding Utf8WithBom = new UTF8Encoding(true);

        private static readonly JsonSerializerOptions PrettyJson = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            TypeInfoResolver = new DefaultJsonTypeInfoResolver()
        };

        private sealed class ClientTarget
        {
            public string Name { get; set; }
            public string Kind { get; set; }
            public string ConfigPath { get; set; }
        }

        private sealed class ManifestEntry
        {
            public string ConfigPath { get; set; }
            public bool FileCreatedByUs { get; set; }
            public bool EntryOwnedByUs { get; set; }
            public string InstalledCommand { get; set; }
        }

        public static int Run(bool install, string exePath, string reportPath, string testRoot, TextWriter output)
        {
            string manifestPath = ResolveManifestPath(exePath, testRoot);
            string appDir = Path.GetDirectoryName(manifestPath) ?? AppDomain.CurrentDomain.BaseDirectory;
            string upgradeMarker = Path.Combine(appDir, ".upgrading");

            // Durante un upgrade/reinstallazione recente (< 5 minuti), l'uninstaller preserva config e manifest.
            // Se il marcatore .upgrading e' vecchio (> 5 minuti, residuo di un aggiornamento interrotto), viene ignorato e rimosso!
            if (!install && File.Exists(upgradeMarker))
            {
                bool isFresh = false;
                try
                {
                    var age = DateTime.UtcNow - File.GetLastWriteTimeUtc(upgradeMarker);
                    isFresh = age >= TimeSpan.FromMinutes(-1) && age <= TimeSpan.FromMinutes(5);
                }
                catch { isFresh = false; }

                if (isFresh)
                {
                    string msg = "[Upgrade] Rilevato aggiornamento in corso (.upgrading): configurazioni MCP e opzioni utente preservate.";
                    output?.WriteLine(msg);
                    WriteReport(reportPath, msg);
                    return 0;
                }
                else
                {
                    try { File.Delete(upgradeMarker); } catch { }
                }
            }

            try
            {
                Dictionary<string, ManifestEntry> manifest;
                bool manifestExisted;
                try
                {
                    manifest = LoadManifestStrict(manifestPath, out manifestExisted);
                }
                catch (Exception ex)
                {
                    string err = $"[Manifest] ERRORE: impossibile leggere {manifestPath} ({ex.Message}). Nessuna configurazione modificata.";
                    output?.WriteLine(err);
                    WriteReport(reportPath, err);
                    return 1;
                }

                var targets = DiscoverTargets(testRoot);
                var lines = new List<string>();
                int errors = 0;

                foreach (var t in targets)
                {
                    try
                    {
                        string status = t.Kind == "JSON"
                            ? (install
                                ? InstallJson(t.ConfigPath, exePath, manifest)
                                : UninstallJson(t.ConfigPath, manifest, manifestExisted))
                            : (install
                                ? InstallToml(t.ConfigPath, exePath, manifest)
                                : UninstallToml(t.ConfigPath, manifest, manifestExisted));

                        lines.Add($"[{t.Name}] {status}");
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        lines.Add($"[{t.Name}] ERRORE ({t.ConfigPath}): {ex.Message}");
                    }
                }

                if (install)
                {
                    try { SaveManifest(manifestPath, manifest); }
                    catch (Exception ex)
                    {
                        errors++;
                        lines.Add($"[Manifest] ERRORE salvataggio: {ex.Message}");
                    }
                }
                else if (errors == 0 && File.Exists(manifestPath))
                {
                    try { File.Delete(manifestPath); } catch { }
                }

                string summary = string.Join(Environment.NewLine, lines);
                output?.WriteLine(summary);
                WriteReport(reportPath, summary);
                return errors == 0 ? 0 : 1;
            }
            finally
            {
                // In fase di installazione rimuove SEMPRE .upgrading anche in caso di errore, cosi' non resta mai orfano.
                if (install && File.Exists(upgradeMarker))
                {
                    try { File.Delete(upgradeMarker); } catch { }
                }
            }
        }

        private static void WriteReport(string reportPath, string text)
        {
            if (string.IsNullOrWhiteSpace(reportPath)) return;
            try { File.WriteAllText(reportPath, text + Environment.NewLine, Utf8WithBom); }
            catch { }
        }

        private static string ResolveManifestPath(string exePath, string testRoot)
        {
            if (!string.IsNullOrWhiteSpace(testRoot))
                return Path.Combine(testRoot, "install-manifest.json");
            if (!string.IsNullOrWhiteSpace(exePath))
            {
                string dir = Path.GetDirectoryName(Path.GetFullPath(exePath));
                if (!string.IsNullOrEmpty(dir))
                    return Path.Combine(dir, "install-manifest.json");
            }
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "install-manifest.json");
        }

        private static List<ClientTarget> DiscoverTargets(string testRoot)
        {
            string userHome = !string.IsNullOrWhiteSpace(testRoot)
                ? Path.Combine(testRoot, "UserProfile")
                : Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string appData = !string.IsNullOrWhiteSpace(testRoot)
                ? Path.Combine(testRoot, "AppDataRoaming")
                : Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string localAppData = !string.IsNullOrWhiteSpace(testRoot)
                ? Path.Combine(testRoot, "AppDataLocal")
                : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            var list = new List<ClientTarget>
            {
                new ClientTarget { Name = "Antigravity", Kind = "JSON", ConfigPath = Path.Combine(userHome, @".gemini\config\mcp_config.json") },
                new ClientTarget { Name = "Antigravity IDE", Kind = "JSON", ConfigPath = Path.Combine(userHome, @".gemini\antigravity-ide\mcp_config.json") },
                new ClientTarget { Name = "Claude Desktop", Kind = "JSON", ConfigPath = Path.Combine(appData, @"Claude\claude_desktop_config.json") },
                new ClientTarget { Name = "Codex", Kind = "TOML", ConfigPath = Path.Combine(userHome, @".codex\config.toml") }
            };

            // Claude Code CLI / scheda Code di Claude Desktop (~/.claude.json, solo se gia' presente)
            string claudeCodeJson = Path.Combine(userHome, ".claude.json");
            if (File.Exists(claudeCodeJson))
            {
                list.Add(new ClientTarget
                {
                    Name = "Claude Code",
                    Kind = "JSON",
                    ConfigPath = claudeCodeJson
                });
            }

            // Claude Desktop Store/MSIX v6+ (%LOCALAPPDATA%\Claude-Data\claude_*\claude_desktop_config.json)
            string claudeDataDir = Path.Combine(localAppData, "Claude-Data");
            if (Directory.Exists(claudeDataDir))
            {
                try
                {
                    foreach (var sub in Directory.GetDirectories(claudeDataDir, "claude_*"))
                    {
                        list.Add(new ClientTarget
                        {
                            Name = "Claude Desktop (Claude-Data)",
                            Kind = "JSON",
                            ConfigPath = Path.Combine(sub, "claude_desktop_config.json")
                        });
                    }
                }
                catch { }
            }

            // Claude Desktop Store/MSIX legacy (%LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude)
            // Salta se la cartella contiene .user-data-left.* (migrata verso Claude-Data)
            string packagesDir = Path.Combine(localAppData, "Packages");
            if (Directory.Exists(packagesDir))
            {
                try
                {
                    foreach (var pkg in Directory.GetDirectories(packagesDir, "Claude_*"))
                    {
                        string msixClaudeDir = Path.Combine(pkg, @"LocalCache\Roaming\Claude");
                        if (Directory.Exists(msixClaudeDir) &&
                            Directory.GetFiles(msixClaudeDir, ".user-data-left.*").Length == 0)
                        {
                            list.Add(new ClientTarget
                            {
                                Name = "Claude Desktop (Store/MSIX)",
                                Kind = "JSON",
                                ConfigPath = Path.Combine(msixClaudeDir, "claude_desktop_config.json")
                            });
                        }
                    }
                }
                catch { }
            }

            return list;
        }

        private static Dictionary<string, ManifestEntry> LoadManifestStrict(string path, out bool existed)
        {
            var dict = new Dictionary<string, ManifestEntry>(StringComparer.OrdinalIgnoreCase);
            existed = File.Exists(path);
            if (!existed) return dict;

            string raw = File.ReadAllText(path, Utf8NoBom);
            var arr = JsonNode.Parse(raw) as JsonArray
                ?? throw new InvalidDataException("Il file install-manifest.json è corrotto (atteso array JSON).");

            foreach (var n in arr)
            {
                var o = n as JsonObject ?? throw new InvalidDataException("Voce non valida in install-manifest.json.");
                string cp = o["configPath"]?.ToString();
                if (string.IsNullOrEmpty(cp)) throw new InvalidDataException("Voce senza configPath in install-manifest.json.");
                dict[cp] = new ManifestEntry
                {
                    ConfigPath = cp,
                    FileCreatedByUs = o["fileCreatedByUs"]?.GetValue<bool>() ?? false,
                    EntryOwnedByUs = o["entryOwnedByUs"]?.GetValue<bool>() ?? false,
                    InstalledCommand = o["installedCommand"]?.ToString() ?? string.Empty
                };
            }
            return dict;
        }

        private static void SaveManifest(string path, Dictionary<string, ManifestEntry> manifest)
        {
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            var arr = new JsonArray();
            foreach (var e in manifest.Values)
            {
                arr.Add(new JsonObject
                {
                    ["configPath"] = e.ConfigPath,
                    ["fileCreatedByUs"] = e.FileCreatedByUs,
                    ["entryOwnedByUs"] = e.EntryOwnedByUs,
                    ["installedCommand"] = e.InstalledCommand
                });
            }
            AtomicWriteFile(path, arr.ToJsonString(PrettyJson) + Environment.NewLine, createBackup: false);
        }

        // ==================== JSON ====================

        private static string InstallJson(string configPath, string exePath, Dictionary<string, ManifestEntry> manifest)
        {
            string parentDir = Path.GetDirectoryName(configPath);
            if (string.IsNullOrEmpty(parentDir) || !Directory.Exists(parentDir))
                return "saltato (client non presente)";

            bool existed = File.Exists(configPath);
            JsonObject root;
            if (existed)
            {
                string raw = File.ReadAllText(configPath, Utf8NoBom);
                root = string.IsNullOrWhiteSpace(raw)
                    ? new JsonObject()
                    : (JsonNode.Parse(raw) as JsonObject ?? throw new InvalidDataException("Il file JSON radice non è un oggetto."));
            }
            else
            {
                root = new JsonObject();
            }

            if (!(root["mcpServers"] is JsonObject servers))
            {
                if (root.ContainsKey("mcpServers") && root["mcpServers"] != null)
                    throw new InvalidDataException("Il campo 'mcpServers' non è un oggetto JSON.");
                servers = new JsonObject();
                root["mcpServers"] = servers;
            }

            manifest.TryGetValue(configPath, out var existingMeta);
            bool alreadyHadEntry = servers.ContainsKey("mcp4office");

            if (alreadyHadEntry)
            {
                var existingObj = servers["mcp4office"] as JsonObject;
                string existingCmd = existingObj?["command"]?.ToString() ?? string.Empty;

                // Rispetta SEMPRE una voce non posseduta da noi o modificata manualmente dall'utente!
                bool isUserCustom =
                    (existingMeta != null && !existingMeta.EntryOwnedByUs) ||
                    (existingMeta != null && existingMeta.EntryOwnedByUs &&
                     !string.Equals(existingCmd, existingMeta.InstalledCommand, StringComparison.OrdinalIgnoreCase)) ||
                    (existingMeta == null && !string.IsNullOrEmpty(existingCmd) &&
                     !string.Equals(existingCmd, exePath, StringComparison.OrdinalIgnoreCase));

                if (isUserCustom)
                {
                    manifest[configPath] = new ManifestEntry
                    {
                        ConfigPath = configPath,
                        FileCreatedByUs = existingMeta?.FileCreatedByUs ?? false,
                        EntryOwnedByUs = false,
                        InstalledCommand = existingCmd
                    };
                    return $"MANTENUTA voce personalizzata (punta a: {existingCmd}, NON al Setup installato)";
                }

                if (existingObj == null)
                {
                    existingObj = new JsonObject { ["command"] = exePath, ["args"] = new JsonArray() };
                    servers["mcp4office"] = existingObj;
                }
                else
                {
                    existingObj["command"] = exePath;
                    if (!existingObj.ContainsKey("args"))
                        existingObj["args"] = new JsonArray();
                }

                manifest[configPath] = new ManifestEntry
                {
                    ConfigPath = configPath,
                    FileCreatedByUs = existingMeta?.FileCreatedByUs ?? false,
                    EntryOwnedByUs = true,
                    InstalledCommand = exePath
                };
            }
            else
            {
                servers["mcp4office"] = new JsonObject
                {
                    ["command"] = exePath,
                    ["args"] = new JsonArray()
                };
                manifest[configPath] = new ManifestEntry
                {
                    ConfigPath = configPath,
                    FileCreatedByUs = !existed,
                    EntryOwnedByUs = true,
                    InstalledCommand = exePath
                };
            }

            string outJson = root.ToJsonString(PrettyJson) + Environment.NewLine;
            if (!(JsonNode.Parse(outJson) is JsonObject))
                throw new InvalidDataException("Serializzazione JSON non valida.");

            // Non creare .bak per un file che abbiamo creato noi stessi
            bool shouldBackup = existed && !(manifest[configPath].FileCreatedByUs);
            AtomicWriteFile(configPath, outJson, createBackup: shouldBackup);
            return alreadyHadEntry ? "aggiornato (opzioni utente preservate)" : "configurato";
        }

        private static string UninstallJson(string configPath, Dictionary<string, ManifestEntry> manifest, bool manifestExisted)
        {
            if (!File.Exists(configPath))
                return "già assente";

            // Senza manifest o senza voce nel manifest: provenienza sconosciuta -> PRESERVA sempre!
            if (!manifestExisted || !manifest.TryGetValue(configPath, out var meta))
                return "provenienza sconosciuta (manifest assente): voce preservata per sicurezza";

            if (!meta.EntryOwnedByUs)
                return "voce personalizzata dall'utente preservata";

            string raw = File.ReadAllText(configPath, Utf8NoBom);
            if (string.IsNullOrWhiteSpace(raw)) return "vuoto";

            var root = JsonNode.Parse(raw) as JsonObject
                ?? throw new InvalidDataException("Il file JSON radice non è un oggetto.");

            if (!(root["mcpServers"] is JsonObject servers) || !servers.ContainsKey("mcp4office"))
                return "non presente";

            string currentCmd = (servers["mcp4office"] as JsonObject)?["command"]?.ToString() ?? string.Empty;
            if (!string.Equals(currentCmd, meta.InstalledCommand, StringComparison.OrdinalIgnoreCase))
                return "voce modificata dall'utente preservata";

            servers.Remove("mcp4office");

            if (meta.FileCreatedByUs && servers.Count == 0 && root.Count == 1)
            {
                File.Delete(configPath);
                return "file creato dall'installer rimosso";
            }

            string outJson = root.ToJsonString(PrettyJson) + Environment.NewLine;
            AtomicWriteFile(configPath, outJson, createBackup: false);
            return "rimosso";
        }

        // ==================== TOML (validato con Tomlyn 1.0) ====================

        internal static string EscapeTomlBasicString(string value)
        {
            var sb = new StringBuilder((value?.Length ?? 0) + 8);
            sb.Append('"');
            foreach (char ch in value ?? string.Empty)
            {
                switch (ch)
                {
                    case '\\': sb.Append(@"\\"); break;
                    case '"': sb.Append("\\\""); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\t': sb.Append("\\t"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\r': sb.Append("\\r"); break;
                    default:
                        if (ch < 0x20) sb.Append($"\\u{(int)ch:x4}");
                        else sb.Append(ch);
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }

        private static bool TryGetSemanticMcp4Office(TomlTable model, out string semanticCommand)
        {
            semanticCommand = null;
            if (model == null) return false;
            if (!model.TryGetValue("mcp_servers", out var serversObj) || !(serversObj is TomlTable serversTable))
                return false;
            if (!serversTable.TryGetValue("mcp4office", out var entryObj))
                return false;
            if (entryObj is TomlTable entryTable && entryTable.TryGetValue("command", out var cmdObj))
                semanticCommand = cmdObj?.ToString();
            return true;
        }

        private static TomlTable ParseTomlOrThrow(string raw, string context)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new TomlTable();
            if (!Toml.TryToModel(raw, out TomlTable model, out var diagnostics))
            {
                string firstErr = diagnostics != null && diagnostics.Count > 0 ? diagnostics[0].ToString() : "errore sintassi TOML";
                throw new InvalidDataException($"TOML non valido ({context}): {firstErr}");
            }
            return model;
        }

        internal static bool IsMcp4OfficeTomlHeader(string headerInner)
        {
            var segments = ParseTomlKeyPath(headerInner);
            return segments.Count >= 2 &&
                   string.Equals(segments[0], "mcp_servers", StringComparison.Ordinal) &&
                   string.Equals(segments[1], "mcp4office", StringComparison.Ordinal);
        }

        private static List<string> ParseTomlKeyPath(string inner)
        {
            var res = new List<string>();
            int i = 0, n = inner.Length;
            while (i < n)
            {
                while (i < n && char.IsWhiteSpace(inner[i])) i++;
                if (i >= n) break;
                if (inner[i] == '"' || inner[i] == '\'')
                {
                    char q = inner[i++];
                    var sb = new StringBuilder();
                    while (i < n && inner[i] != q)
                    {
                        if (q == '"' && inner[i] == '\\' && i + 1 < n)
                        {
                            i++;
                            if (inner[i] == 'u' && i + 4 < n &&
                                int.TryParse(inner.Substring(i + 1, 4), System.Globalization.NumberStyles.HexNumber, null, out int code))
                            {
                                sb.Append((char)code);
                                i += 5;
                            }
                            else sb.Append(inner[i++]);
                        }
                        else sb.Append(inner[i++]);
                    }
                    if (i < n && inner[i] == q) i++;
                    res.Add(sb.ToString());
                }
                else
                {
                    int start = i;
                    while (i < n && inner[i] != '.' && !char.IsWhiteSpace(inner[i])) i++;
                    res.Add(inner.Substring(start, i - start));
                }
                while (i < n && char.IsWhiteSpace(inner[i])) i++;
                if (i < n && inner[i] == '.') i++;
                else break;
            }
            return res;
        }

        internal static (string cleanedContent, bool foundSectionHeader) ScanAndStripMcp4OfficeToml(
            string raw, bool stripAllMcp4Office, out string lineEnding)
        {
            lineEnding = raw.Contains("\r\n") ? "\r\n" : "\n";
            var lines = raw.Length == 0 ? new string[0] : raw.Replace("\r\n", "\n").Split('\n');
            if (lines.Length > 0 && lines[lines.Length - 1].Length == 0)
            {
                var trimmed = new string[lines.Length - 1];
                Array.Copy(lines, trimmed, trimmed.Length);
                lines = trimmed;
            }

            var kept = new List<string>(lines.Length);
            bool inMultiDouble = false, inMultiSingle = false, inTargetSection = false, foundHeader = false;

            foreach (var line in lines)
            {
                if (!inMultiDouble && !inMultiSingle)
                {
                    string trim = line.TrimStart();
                    if (trim.StartsWith("[") && !trim.StartsWith("[["))
                    {
                        int closeIdx = FindClosingBracket(trim);
                        if (closeIdx > 1)
                        {
                            string headerInner = trim.Substring(1, closeIdx - 1);
                            if (IsMcp4OfficeTomlHeader(headerInner))
                            {
                                inTargetSection = true;
                                foundHeader = true;
                            }
                            else inTargetSection = false;
                        }
                    }
                    else if (trim.StartsWith("[[")) inTargetSection = false;
                }

                UpdateMultilineState(line, ref inMultiDouble, ref inMultiSingle);

                if (!inTargetSection || !stripAllMcp4Office)
                    kept.Add(line);
            }

            while (kept.Count > 0 && string.IsNullOrWhiteSpace(kept[kept.Count - 1]))
                kept.RemoveAt(kept.Count - 1);

            string result = kept.Count > 0 ? string.Join(lineEnding, kept) + lineEnding : string.Empty;
            return (result, foundHeader);
        }

        private static int FindClosingBracket(string s)
        {
            bool inD = false, inS = false;
            for (int i = 1; i < s.Length; i++)
            {
                char c = s[i];
                if (inD) { if (c == '\\') { i++; continue; } if (c == '"') inD = false; }
                else if (inS) { if (c == '\'') inS = false; }
                else { if (c == '"') inD = true; else if (c == '\'') inS = true; else if (c == ']') return i; else if (c == '#') return -1; }
            }
            return -1;
        }

        private static void UpdateMultilineState(string line, ref bool inMultiDouble, ref bool inMultiSingle)
        {
            bool inD = false, inS = false;
            int i = 0, n = line.Length;
            while (i < n)
            {
                if (inMultiDouble)
                {
                    if (line[i] == '\\') { i += 2; continue; }
                    if (i + 2 < n && line[i] == '"' && line[i + 1] == '"' && line[i + 2] == '"') { inMultiDouble = false; i += 3; continue; }
                    i++;
                }
                else if (inMultiSingle)
                {
                    if (i + 2 < n && line[i] == '\'' && line[i + 1] == '\'' && line[i + 2] == '\'') { inMultiSingle = false; i += 3; continue; }
                    i++;
                }
                else if (inD) { if (line[i] == '\\') { i += 2; continue; } if (line[i] == '"') inD = false; i++; }
                else if (inS) { if (line[i] == '\'') inS = false; i++; }
                else
                {
                    if (line[i] == '#') break;
                    if (i + 2 < n && line[i] == '"' && line[i + 1] == '"' && line[i + 2] == '"') { inMultiDouble = true; i += 3; }
                    else if (i + 2 < n && line[i] == '\'' && line[i + 1] == '\'' && line[i + 2] == '\'') { inMultiSingle = true; i += 3; }
                    else if (line[i] == '"') { inD = true; i++; }
                    else if (line[i] == '\'') { inS = true; i++; }
                    else i++;
                }
            }
        }

        private static string InstallToml(string tomlPath, string exePath, Dictionary<string, ManifestEntry> manifest)
        {
            string parentDir = Path.GetDirectoryName(tomlPath);
            if (string.IsNullOrEmpty(parentDir) || !Directory.Exists(parentDir))
                return "saltato (client non presente)";

            bool existed = File.Exists(tomlPath);
            string raw = existed ? File.ReadAllText(tomlPath, Utf8NoBom) : string.Empty;

            // Validazione semantica completa dell'input con Tomlyn
            TomlTable inputModel = ParseTomlOrThrow(raw, "input");
            bool semanticExists = TryGetSemanticMcp4Office(inputModel, out string existingCmd);

            var (stripped, foundStandardHeader) = ScanAndStripMcp4OfficeToml(raw, stripAllMcp4Office: true, out string eol);
            manifest.TryGetValue(tomlPath, out var existingMeta);

            if (semanticExists)
            {
                bool isUserCustom =
                    (existingMeta != null && !existingMeta.EntryOwnedByUs) ||
                    (existingMeta != null && existingMeta.EntryOwnedByUs &&
                     !string.Equals(existingCmd ?? "", existingMeta.InstalledCommand, StringComparison.OrdinalIgnoreCase)) ||
                    (existingMeta == null && !string.IsNullOrEmpty(existingCmd) &&
                     !string.Equals(existingCmd, exePath, StringComparison.OrdinalIgnoreCase));

                if (isUserCustom)
                {
                    manifest[tomlPath] = new ManifestEntry
                    {
                        ConfigPath = tomlPath,
                        FileCreatedByUs = existingMeta?.FileCreatedByUs ?? false,
                        EntryOwnedByUs = false,
                        InstalledCommand = existingCmd ?? string.Empty
                    };
                    return $"MANTENUTA voce personalizzata (punta a: {existingCmd}, NON al Setup installato)";
                }

                // Se esiste semanticamente ma NON in una sezione standard [mcp_servers.mcp4office] (es. inline table o chiave puntata),
                // non rischiare di duplicarla: fallisci pulitamente senza scrivere!
                if (!foundStandardHeader)
                    throw new InvalidDataException("La voce mcp_servers.mcp4office esiste già in forma TOML non standard (tabella inline o chiave puntata): aggiornarla manualmente.");
            }

            string outToml;
            if (semanticExists)
            {
                outToml = ReplaceCommandInTomlSection(raw, exePath, eol, out bool replaced);
                if (!replaced)
                    throw new InvalidDataException("Impossibile aggiornare la chiave 'command' nella sezione [mcp_servers.mcp4office] esistente.");
            }
            else
            {
                var sb = new StringBuilder(stripped);
                if (sb.Length > 0) sb.Append(eol);
                sb.Append("[mcp_servers.mcp4office]").Append(eol);
                sb.Append("command = ").Append(EscapeTomlBasicString(exePath)).Append(eol);
                sb.Append("args = []").Append(eol);
                sb.Append("startup_timeout_sec = 20").Append(eol);
                sb.Append("tool_timeout_sec = 60").Append(eol);
                outToml = sb.ToString();
            }

            // Validazione semantica indipendente post-generazione con Tomlyn!
            TomlTable outModel = ParseTomlOrThrow(outToml, "output");
            if (!TryGetSemanticMcp4Office(outModel, out string verifyCmd) ||
                !string.Equals(verifyCmd, exePath, StringComparison.Ordinal))
            {
                throw new InvalidDataException("Verifica Tomlyn post-generazione fallita: voce mcp4office non coerente.");
            }

            manifest[tomlPath] = new ManifestEntry
            {
                ConfigPath = tomlPath,
                FileCreatedByUs = existingMeta?.FileCreatedByUs ?? !existed,
                EntryOwnedByUs = true,
                InstalledCommand = exePath
            };

            bool shouldBackup = existed && !(manifest[tomlPath].FileCreatedByUs);
            AtomicWriteFile(tomlPath, outToml, createBackup: shouldBackup);
            return semanticExists ? "aggiornato (opzioni utente preservate)" : "configurato";
        }

        private static string ReplaceCommandInTomlSection(string raw, string newExePath, string eol, out bool replaced)
        {
            replaced = false;
            var lines = raw.Replace("\r\n", "\n").Split('\n');
            if (lines.Length > 0 && lines[lines.Length - 1].Length == 0)
            {
                var t = new string[lines.Length - 1];
                Array.Copy(lines, t, t.Length);
                lines = t;
            }
            bool inMD = false, inMS = false, inExactSection = false;
            var outLines = new List<string>(lines.Length);
            foreach (var line in lines)
            {
                if (!inMD && !inMS)
                {
                    string trim = line.TrimStart();
                    if (trim.StartsWith("[") && !trim.StartsWith("[["))
                    {
                        int closeIdx = FindClosingBracket(trim);
                        if (closeIdx > 1)
                        {
                            var segs = ParseTomlKeyPath(trim.Substring(1, closeIdx - 1));
                            inExactSection = segs.Count == 2 && segs[0] == "mcp_servers" && segs[1] == "mcp4office";
                        }
                    }
                    else if (trim.StartsWith("[[")) inExactSection = false;
                    else if (inExactSection && !replaced && Regex.IsMatch(trim, @"^command\s*="))
                    {
                        outLines.Add("command = " + EscapeTomlBasicString(newExePath));
                        replaced = true;
                        UpdateMultilineState(line, ref inMD, ref inMS);
                        continue;
                    }
                }
                UpdateMultilineState(line, ref inMD, ref inMS);
                outLines.Add(line);
            }
            return string.Join(eol, outLines) + eol;
        }

        private static string UninstallToml(string tomlPath, Dictionary<string, ManifestEntry> manifest, bool manifestExisted)
        {
            if (!File.Exists(tomlPath))
                return "già assente";

            if (!manifestExisted || !manifest.TryGetValue(tomlPath, out var meta))
                return "provenienza sconosciuta (manifest assente): voce preservata per sicurezza";

            if (!meta.EntryOwnedByUs)
                return "voce personalizzata dall'utente preservata";

            string raw = File.ReadAllText(tomlPath, Utf8NoBom);
            TomlTable model = ParseTomlOrThrow(raw, "uninstall-input");
            if (!TryGetSemanticMcp4Office(model, out string existingCmd))
                return "non presente";

            if (!string.Equals(existingCmd ?? "", meta.InstalledCommand, StringComparison.OrdinalIgnoreCase))
                return "voce modificata dall'utente preservata";

            var (cleaned, foundHeader) = ScanAndStripMcp4OfficeToml(raw, stripAllMcp4Office: true, out _);
            if (!foundHeader)
                throw new InvalidDataException("La voce mcp_servers.mcp4office è in forma TOML non standard: rimuoverla manualmente.");

            // Verifica che il TOML risultante sia valido e non contenga più mcp4office
            TomlTable afterModel = ParseTomlOrThrow(cleaned, "uninstall-output");
            if (TryGetSemanticMcp4Office(afterModel, out _))
                throw new InvalidDataException("Verifica post-rimozione fallita: mcp_servers.mcp4office ancora presente nel TOML.");

            if (meta.FileCreatedByUs && string.IsNullOrWhiteSpace(cleaned))
            {
                File.Delete(tomlPath);
                return "file creato dall'installer rimosso";
            }

            AtomicWriteFile(tomlPath, cleaned, createBackup: false);
            return "rimosso";
        }

        private static void AtomicWriteFile(string targetPath, string content, bool createBackup)
        {
            string dir = Path.GetDirectoryName(targetPath);
            string tmp = Path.Combine(dir, $".mcp4office.{Guid.NewGuid():N}.tmp");
            File.WriteAllText(tmp, content, Utf8NoBom);
            try
            {
                if (File.Exists(targetPath))
                {
                    string bak = targetPath + ".mcp4office.bak";
                    if (createBackup && !File.Exists(bak))
                        File.Copy(targetPath, bak, overwrite: false);

                    string replaceBak = Path.Combine(dir, $".mcp4office.{Guid.NewGuid():N}.old");
                    try
                    {
                        File.Replace(tmp, targetPath, replaceBak, ignoreMetadataErrors: true);
                    }
                    finally
                    {
                        if (File.Exists(replaceBak)) try { File.Delete(replaceBak); } catch { }
                    }
                }
                else
                {
                    File.Move(tmp, targetPath);
                }

                string verify = File.ReadAllText(targetPath, Utf8NoBom);
                if (!string.Equals(verify, content, StringComparison.Ordinal))
                    throw new IOException("Verifica post-scrittura fallita.");
            }
            finally
            {
                if (File.Exists(tmp)) try { File.Delete(tmp); } catch { }
            }
        }
    }
}
