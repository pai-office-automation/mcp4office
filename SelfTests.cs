using System;
using System.IO;
using System.Text;
using System.Text.Json.Nodes;
using Tomlyn;
using Tomlyn.Model;

namespace Mcp4Office
{
    internal static class SelfTests
    {
        private static int _passed;
        private static int _failed;

        public static int RunAll(TextWriter log)
        {
            _passed = 0;
            _failed = 0;

            Test(log, "Lifecycle: tools/list rifiutato pre-init (-32002), 32 tool dopo initialize", () =>
            {
                var srv = new McpServer();
                Program.RegisterAllTools(srv);
                var rPre = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}") as JsonObject;
                Assert(rPre["error"] != null && (int)rPre["error"]["code"] == -32002, "Atteso -32002 prima di initialize");
                var rInit = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"initialize\",\"params\":{\"protocolVersion\":\"2099-01-01\"}}") as JsonObject;
                Assert(rInit["result"]["protocolVersion"].ToString() == "2025-03-26", "Negoziazione versione fallita");
                var rPost = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":3,\"method\":\"tools/list\"}") as JsonObject;
                Assert(rPost["result"] != null && (rPost["result"]["tools"] as JsonArray).Count == 32, "Attesi 32 tool dopo initialize");
            });

            Test(log, "Schema validation: required/tipi/enum rifiutati senza chiamare handler", () =>
            {
                bool invoked = false;
                var srv = new McpServer();
                srv.RegisterTool("fake_write", "test",
                    McpServer.Schema((p, r) =>
                    {
                        p["text"] = McpServer.Prop("string", "t");
                        p["position"] = McpServer.EnumProp("pos", "cursor", "start", "end");
                        r.Add("text");
                    }),
                    McpServer.Mutating(),
                    _ => { invoked = true; return "OK"; });
                srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":0,\"method\":\"initialize\"}");

                var resp = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/call\",\"params\":{\"name\":\"fake_write\",\"arguments\":{}}}") as JsonObject;
                Assert(resp["error"] != null && (int)resp["error"]["code"] == -32602 && !invoked, "Campo mancante -> -32602");

                var respObj = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":2,\"method\":\"tools/call\",\"params\":{\"name\":\"fake_write\",\"arguments\":{\"text\":{\"a\":1}}}}") as JsonObject;
                Assert(respObj["error"] != null && (int)respObj["error"]["code"] == -32602 && !invoked, "Tipo oggetto -> -32602");

                var respEnum = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":3,\"method\":\"tools/call\",\"params\":{\"name\":\"fake_write\",\"arguments\":{\"text\":\"ciao\",\"position\":\"middle\"}}}") as JsonObject;
                Assert(respEnum["error"] != null && (int)respEnum["error"]["code"] == -32602 && !invoked, "Enum invalida -> -32602");

                var respEmpty = srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":4,\"method\":\"tools/call\",\"params\":{\"name\":\"fake_write\",\"arguments\":{\"text\":\"\"}}}") as JsonObject;
                Assert(respEmpty["result"] != null && !(bool)respEmpty["result"]["isError"] && invoked, "Stringa vuota esplicita deve passare");
            });

            Test(log, "Envelope JSON-RPC 2.0: rifiuta jsonrpc errato e id oggetto", () =>
            {
                var srv = new McpServer();
                Assert((int)(srv.HandleMessage("{\"id\":1,\"method\":\"ping\"}") as JsonObject)["error"]["code"] == -32600, "jsonrpc mancante");
                Assert((int)(srv.HandleMessage("{\"jsonrpc\":\"1.0\",\"id\":1,\"method\":\"ping\"}") as JsonObject)["error"]["code"] == -32600, "jsonrpc 1.0");
                Assert((int)(srv.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":{},\"method\":\"ping\"}") as JsonObject)["error"]["code"] == -32600, "id oggetto");
            });

            Test(log, "Flag --sola-lettura registra solo 16 tool ReadOnly", () =>
            {
                var ro = new McpServer(readOnlyOnly: true);
                Program.RegisterAllTools(ro);
                ro.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":0,\"method\":\"initialize\"}");
                var r = ro.HandleMessage("{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}") as JsonObject;
                var tools = r["result"]["tools"] as JsonArray;
                Assert(tools.Count == 16, $"Attesi 16 tool di sola lettura, trovati {tools.Count}");
                foreach (var t in tools)
                    Assert((bool)t["annotations"]["readOnlyHint"], $"Tool {t["name"]} non è readOnly!");
            });

            Test(log, "Excel ComputeRangeToken, formule assolute, date ISO, escape e diff fogli", () =>
            {
                var numVal = new object[,] { { 1.0 } };
                var strVal = new object[,] { { "1" } };
                var noF = new bool[,] { { false } };
                Assert(Tools.ExcelTools.ComputeRangeToken("C:\\a.xlsx", "F1", "A1", numVal, numVal, noF) !=
                       Tools.ExcelTools.ComputeRangeToken("C:\\a.xlsx", "F1", "A1", strVal, strVal, noF), "1.0 != '1'");

                var fText = new object[,] { { "=1+1" } };
                var yesF = new bool[,] { { true } };
                Assert(Tools.ExcelTools.ComputeRangeToken("C:\\a.xlsx", "F1", "A1", fText, fText, noF) !=
                       Tools.ExcelTools.ComputeRangeToken("C:\\a.xlsx", "F1", "A1", new object[,] { { 2.0 } }, fText, yesF), "Formula != testo");

                Assert(Tools.ExcelTools.IsPlausibleFormulaStart("=$A$1"), "=$A$1 ok");
                Assert(!Tools.ExcelTools.IsPlausibleFormulaStart("=== titolo"), "=== no");
                Assert(Tools.ExcelTools.TryParseIsoDate("2026-11-20", out var dt) && dt == new DateTime(2026, 11, 20), "ISO date ok");
                Assert(Tools.ExcelTools.FormatCellValue(unchecked((int)0x800A07D7)) == "#DIV/0!", "#DIV/0!");

                // Test ComputeSheetDiff + ComputeKeyedSheetDiff (riga inserita non fa slittare le altre!)
                var v1 = new object[,] { { "Codice", "Valore" }, { "A1", 100.0 }, { "A2", 200.0 } };
                var v2 = new object[,] { { "Codice", "Valore" }, { "NEW", 50.0 }, { "a1", 100.0 }, { "A2", 250.0 } };
                var v1Pad = new object[,] { { "Codice", "Valore" }, { "A1", 100.0 }, { "A2", 200.0 }, { null, null } };
                var m0 = new bool[4, 2];
                string keyed = Tools.ExcelTools.ComputeKeyedSheetDiff("F1", "F2", 1, 1, 4, 2, v1Pad, v1Pad, m0, v2, v2, m0, "Codice", true, true, 50);
                Assert(keyed.Contains("+ AGGIUNTA chiave 'NEW'") && keyed.Contains("~ MODIFICATA chiave 'A2'") && !keyed.Contains("MODIFICATA chiave 'A1'"), "Keyed diff Excel errato: " + keyed);

                // Test has_header=false: la riga 0 DEVE essere confrontata (Claude #3)
                var nh1 = new object[,] { { "R0", 10.0 }, { "R1", 20.0 } };
                var nh2 = new object[,] { { "R0", 99.0 }, { "R1", 20.0 } };
                var m2x2 = new bool[2, 2];
                string noHdr = Tools.ExcelTools.ComputeKeyedSheetDiff("F1", "F2", 1, 1, 2, 2, nh1, nh1, m2x2, nh2, nh2, m2x2, "A", false, true, 50);
                Assert(noHdr.Contains("~ MODIFICATA chiave 'R0'") && noHdr.Contains("has_header=false"), "has_header=false deve confrontare riga 0: " + noHdr);

                // Test intestazioni diverse: avviso esplicito anche se i dati sono uguali (Claude #4)
                var hd1 = new object[,] { { "Codice", "Prezzo" }, { "A1", 10.0 } };
                var hd2 = new object[,] { { "Codice", "Categoria" }, { "A1", 10.0 } };
                string hdrWarn = Tools.ExcelTools.ComputeKeyedSheetDiff("F1", "F2", 1, 1, 2, 2, hd1, hd1, m2x2, hd2, hd2, m2x2, "Codice", true, true, 50);
                Assert(hdrWarn.Contains("Avviso intestazioni diverse") && hdrWarn.Contains("'Prezzo' -> 'Categoria'"), "Avviso intestazioni diverse mancante: " + hdrWarn);

                // Test chiavi duplicate con seconda occorrenza modificata 20 -> 999 (Codex 0.5.6!)
                var dupA = new object[,] { { "Codice", "Valore" }, { "K1", 10.0 }, { "K1", 20.0 } };
                var dupB = new object[,] { { "Codice", "Valore" }, { "K1", 10.0 }, { "K1", 999.0 } };
                var m3x2 = new bool[3, 2];
                string dupDiff = Tools.ExcelTools.ComputeKeyedSheetDiff("F1", "F2", 1, 1, 3, 2, dupA, dupA, m3x2, dupB, dupB, m3x2, "Codice", true, true, 50);
                Assert(!dupDiff.Contains("NESSUNA DIFFERENZA") && dupDiff.Contains("20 -> 999") && dupDiff.Contains("Avviso chiavi duplicate"), "Chiave duplicata 20->999 non rilevata: " + dupDiff);

                // Test chiave vuota/spazi con dati non vuoti (20 -> 999) vs padding completamente vuoto (Codex 0.5.7!)
                var ekA = new object[,] { { "Codice", "Valore" }, { "K1", 10.0 }, { "   ", 20.0 }, { null, null } };
                var ekB = new object[,] { { "Codice", "Valore" }, { "K1", 10.0 }, { "", 999.0 }, { null, null } };
                string ekDiff = Tools.ExcelTools.ComputeKeyedSheetDiff("F1", "F2", 1, 1, 4, 2, ekA, ekA, m0, ekB, ekB, m0, "Codice", true, true, 50);
                Assert(!ekDiff.Contains("NESSUNA DIFFERENZA") && ekDiff.Contains("Avviso confronto incompleto") && ekDiff.Contains("righe 3"), "Chiave vuota con dati non segnalata: " + ekDiff);
            });

            Test(log, "Word token + diff a parole su paragrafi modificati + PowerPoint SlideID + Outlook InsertAfterBody", () =>
            {
                Assert(Tools.WordTools.ComputeSelectionToken("C:\\a.docx", 1, 0, 5, "T") !=
                       Tools.WordTools.ComputeSelectionToken("C:\\a.docx", 7, 0, 5, "T"), "StoryType");
                string tt1 = Tools.WordTools.ComputeTableToken("C:\\a.docx", 10, 3, 2, "A|B|");
                string tt2 = Tools.WordTools.ComputeTableToken("C:\\a.docx", 50, 3, 2, "A|B|");
                Assert(tt1 != tt2, "TableToken");
                string wDiff = Tools.WordTools.ComputeParagraphDiff(
                    "v1.docx", "Il contratto scade il giorno 10 marzo 2026.\nParagrafo invariato.",
                    "v2.docx", "Il contratto scade il giorno 15 aprile 2026.\nParagrafo invariato.\nNuova clausola.", 50);
                Assert(wDiff.Contains("~ [Doc1 §1 -> Doc2 §1]") && wDiff.Contains("[-10 marzo-]") && wDiff.Contains("[+15 aprile+]") && wDiff.Contains("+ [Doc2 §3] Nuova clausola."), "Word inline diff errato: " + wDiff);

                // Test ComputeSurgicalTextEdits: preserva prefisso/suffisso (es. grassetti intra-riga) e applica da destra a sinistra
                string sOld = "Responsabile IT: gestisce il team tecnico a Milano.";
                string sNew = "Responsabile IT: coordina la squadra tecnica a Milano.";
                var sEdits = Tools.WordTools.ComputeSurgicalTextEdits(sOld, sNew);
                Assert(sEdits.Count >= 1 && sEdits[0].start > 0, "Deve toccare solo le parole cambiate preservando 'Responsabile IT:'");
                string reconstructed = sOld;
                foreach (var ed in sEdits)
                    reconstructed = reconstructed.Substring(0, ed.start) + ed.replacement + reconstructed.Substring(ed.start + ed.length);
                Assert(reconstructed == sNew, "Ricostruzione surgical edits errata: " + reconstructed);
                Assert(Tools.PowerPointTools.ComputeNotesToken("C:\\p.pptx", 256, "n") !=
                       Tools.PowerPointTools.ComputeNotesToken("C:\\p.pptx", 257, "n"), "SlideID");
                string merged = Tools.OutlookTools.InsertAfterBody("<html><body lang=\"IT\"><p>Firma</p></body></html>", "<div>Ciao</div>");
                Assert(merged.Contains("<body lang=\"IT\"><div>Ciao</div><p>Firma</p>"), "InsertAfterBody");
            });

            Test(log, "PdfTools: ValidatePdfPath, radici esplicite invalide e reparse tag cloud vs symlink", () =>
            {
                AssertThrows<ArgumentException>(() => Tools.PdfTools.ConfigureAllowedRoots(new[] { @"\\server\share" }, isExplicit: true));
                Assert(Tools.PdfTools.IsDangerousReparseTag(0xA000000C) && Tools.PdfTools.IsDangerousReparseTag(0xA0000003), "Symlink/junction");
                Assert(!Tools.PdfTools.IsDangerousReparseTag(0x9000101A), "OneDrive cloud tag");
            });

            Test(log, "ClientConfigurator: JSON 18 livelli, MSIX, O'Neil, upgrade (.upgrading preserva args), 2x configure su voce custom", () =>
            {
                string root = Path.Combine(Path.GetTempPath(), "mcp4office_cc_1_" + Guid.NewGuid().ToString("N"));
                try
                {
                    string geminiDir = Path.Combine(root, @"UserProfile\.gemini\config");
                    string codexDir = Path.Combine(root, @"UserProfile\.codex");
                    string msixDir = Path.Combine(root, @"AppDataLocal\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude");
                    Directory.CreateDirectory(geminiDir);
                    Directory.CreateDirectory(codexDir);
                    Directory.CreateDirectory(msixDir);

                    string geminiFile = Path.Combine(geminiDir, "mcp_config.json");
                    string deepJson = "{\"mcpServers\":{\"deep\":{\"n\":12345678901234567890,\"d\":{\"a\":{\"b\":{\"c\":{\"d\":{\"e\":{\"f\":{\"g\":{\"h\":{\"i\":{\"j\":{\"k\":{\"l\":{\"m\":{\"n\":{\"o\":{\"p\":\"ok\"}}}}}}}}}}}}}}}}},\"mcp4office\":{\"command\":\"custom-original\",\"args\":[\"--sola-lettura\"]}}}";
                    File.WriteAllText(geminiFile, deepJson, new UTF8Encoding(false));

                    string tomlFile = Path.Combine(codexDir, "config.toml");
                    File.WriteAllText(tomlFile, "note = \"\"\"\n[mcp_servers.mcp4office]\nnot_a_table = true\n\"\"\"\n\n[mcp_servers.other]\ncommand = \"bar\"\n", new UTF8Encoding(false));

                    string exePath = @"C:\Users\O'Neil\AppData\Local\Programs\mcp4office\mcp4office.exe";
                    var sw = new StringWriter();
                    int rc1 = ClientConfigurator.Run(true, exePath, null, root, sw);
                    Assert(rc1 == 0, "1st install: " + sw.ToString());

                    // 2° configure consecutivo: la voce personalizzata in geminiFile DEVE restare 'custom-original' (Codex blocco #2!)
                    Assert(ClientConfigurator.Run(true, exePath, null, root, TextWriter.Null) == 0, "2nd install");
                    var g2 = JsonNode.Parse(File.ReadAllText(geminiFile)) as JsonObject;
                    Assert(g2["mcpServers"]["mcp4office"]["command"].ToString() == "custom-original", "Voce custom sovrascritta al 2° giro!");

                    // Verifica upgrade (.upgrading): utente aggiunge --sola-lettura su Claude MSIX -> unconfigure con .upgrading -> configure -> args PRESERVATI! (Claude #1, Codex #4!)
                    string msixFile = Path.Combine(msixDir, "claude_desktop_config.json");
                    var mObj = JsonNode.Parse(File.ReadAllText(msixFile)) as JsonObject;
                    mObj["mcpServers"]["mcp4office"]["args"] = new JsonArray("--sola-lettura");
                    File.WriteAllText(msixFile, mObj.ToJsonString(), new UTF8Encoding(false));

                    File.WriteAllText(Path.Combine(root, ".upgrading"), "upgrade");
                    Assert(ClientConfigurator.Run(false, exePath, null, root, TextWriter.Null) == 0, "upgrade unconfigure");
                    Assert(File.Exists(msixFile), "msixFile rimosso durante upgrade!");
                    Assert(ClientConfigurator.Run(true, exePath, null, root, TextWriter.Null) == 0, "upgrade reconfigure");
                    var mAfter = JsonNode.Parse(File.ReadAllText(msixFile)) as JsonObject;
                    Assert(mAfter["mcpServers"]["mcp4office"]["args"][0].ToString() == "--sola-lettura", "args --sola-lettura persi durante upgrade!");
                    Assert(!File.Exists(msixFile + ".mcp4office.bak"), "Non deve creare .bak su file creato dall'installer!");

                    // Test marcatore .upgrading vecchio (> 5 minuti) DEVE essere ignorato alla disinstallazione vera!
                    string staleMarker = Path.Combine(root, ".upgrading");
                    File.WriteAllText(staleMarker, "stale");
                    File.SetLastWriteTimeUtc(staleMarker, DateTime.UtcNow.AddMinutes(-10));

                    // Disinstallazione vera: ignora .upgrading scaduto, rimuove msixFile creato da noi, preserva custom-original in geminiFile, rimuove TOML + sottotabella
                    File.AppendAllText(tomlFile, "\n[mcp_servers.mcp4office.env]\nK = \"V\"\n", new UTF8Encoding(false));
                    Assert(ClientConfigurator.Run(false, exePath, null, root, TextWriter.Null) == 0, "real uninstall");
                    Assert(!File.Exists(staleMarker), "Marcatore .upgrading scaduto doveva essere eliminato!");
                    Assert(!File.Exists(msixFile), "msixFile doveva essere rimosso");
                    Assert((JsonNode.Parse(File.ReadAllText(geminiFile)) as JsonObject)["mcpServers"]["mcp4office"]["command"].ToString() == "custom-original", "Voce custom rimossa!");
                    string tFinal = File.ReadAllText(tomlFile);
                    Assert(Toml.TryToModel(tFinal, out TomlTable _, out _) && !tFinal.Contains("[mcp_servers.mcp4office.env]"), "TOML finale invalido o residui");
                }
                finally { try { Directory.Delete(root, true); } catch { } }
            });

            Test(log, "ClientConfigurator TOML edge cases (Codex #1 & #3): inline table, dotted root, unicode escape, manifest assente/corrotto", () =>
            {
                string[] weirdTomls =
                {
                    "[mcp_servers]\nmcp4office = { command = \"custom\", args = [] }\n",
                    "mcp_servers.mcp4office.command = \"custom\"\n",
                    "[mcp_servers.\"mcp4offic\\u0065\"]\ncommand = \"custom\"\n"
                };

                foreach (var src in weirdTomls)
                {
                    string root = Path.Combine(Path.GetTempPath(), "mcp4office_toml_" + Guid.NewGuid().ToString("N"));
                    try
                    {
                        string codexDir = Path.Combine(root, @"UserProfile\.codex");
                        Directory.CreateDirectory(codexDir);
                        string tomlFile = Path.Combine(codexDir, "config.toml");
                        File.WriteAllText(tomlFile, src, new UTF8Encoding(false));

                        // Installa: deve riconoscere semanticamente che esiste una voce custom, NON duplicarla e lasciare TOML 100% valido!
                        int rc = ClientConfigurator.Run(true, @"C:\App\mcp4office.exe", null, root, TextWriter.Null);
                        Assert(rc == 0, "Configure su TOML equivalente deve preservare (rc=0)");
                        string after = File.ReadAllText(tomlFile);
                        Assert(Toml.TryToModel(after, out TomlTable _, out _), "TOML corrotto dopo configure: " + after);
                        Assert(after == src, "Voce custom TOML non doveva essere modificata!");
                    }
                    finally { try { Directory.Delete(root, true); } catch { } }
                }

                // Manifest assente -> unconfigure NON deve cancellare voci esistenti!
                string rootNoMan = Path.Combine(Path.GetTempPath(), "mcp4office_noman_" + Guid.NewGuid().ToString("N"));
                try
                {
                    string codexDir = Path.Combine(rootNoMan, @"UserProfile\.codex");
                    Directory.CreateDirectory(codexDir);
                    string tomlFile = Path.Combine(codexDir, "config.toml");
                    string orig = "[mcp_servers.mcp4office]\ncommand = \"custom\"\n";
                    File.WriteAllText(tomlFile, orig, new UTF8Encoding(false));
                    Assert(ClientConfigurator.Run(false, @"C:\App\mcp4office.exe", null, rootNoMan, TextWriter.Null) == 0, "unconfigure senza manifest");
                    Assert(File.ReadAllText(tomlFile) == orig, "Senza manifest la voce NON deve essere eliminata!");

                    // Manifest corrotto -> exit non-zero e file intatto!
                    File.WriteAllText(Path.Combine(rootNoMan, "install-manifest.json"), "{ corrotto", new UTF8Encoding(false));
                    Assert(ClientConfigurator.Run(false, @"C:\App\mcp4office.exe", null, rootNoMan, TextWriter.Null) != 0, "Manifest corrotto deve dare exit non-zero");
                    Assert(File.ReadAllText(tomlFile) == orig, "Con manifest corrotto il file non deve cambiare!");
                }
                finally { try { Directory.Delete(rootNoMan, true); } catch { } }
            });

            Test(log, "ComHelper.Truncate non disabilita mai il limite con maxChars <= 0", () =>
            {
                string res = ComHelper.Truncate("abcdef", 0, out bool trunc);
                Assert(trunc && res.StartsWith("a"), "maxChars=0 clampato a 1");
            });

            log.WriteLine($"SELF-TEST COMPLETATI: {_passed} passati, {_failed} falliti.");
            return _failed == 0 ? 0 : 1;
        }

        private static void Test(TextWriter log, string name, Action body)
        {
            try { body(); _passed++; log.WriteLine($"[PASS] {name}"); }
            catch (Exception ex) { _failed++; log.WriteLine($"[FAIL] {name}: {ex.Message}"); }
        }

        private static void Assert(bool cond, string message) { if (!cond) throw new Exception(message); }
        private static void AssertThrows<TEx>(Action action) where TEx : Exception
        {
            try { action(); } catch (TEx) { return; }
            throw new Exception($"Attesa eccezione {typeof(TEx).Name}, nessuna sollevata.");
        }
    }
}
