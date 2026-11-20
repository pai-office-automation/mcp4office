using System;
using System.Text;
using System.Text.Json.Nodes;

namespace Mcp4Office.Tools
{
    internal static class WordTools
    {
        public static void Register(McpServer server)
        {
            server.RegisterTool(
                "word_status",
                "Mostra i documenti aperti in Microsoft Word e il doc_id (percorso) del documento attivo.",
                McpServer.EmptySchema(),
                McpServer.ReadOnly(),
                _ => ComHelper.RetryRead(GetStatus));

            server.RegisterTool(
                "word_read_selection",
                "Legge il testo attualmente selezionato in Word restituendo doc_id, story_type, start, end e state_token (necessari per word_replace_selection).",
                McpServer.Schema((props, _) =>
                {
                    props["max_chars"] = McpServer.IntProp("Lunghezza massima in caratteri (1-200000, default 50000).", 1, ComHelper.MaxAllowedChars);
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ReadSelection(McpServer.GetInt(args, "max_chars", 50000))));

            server.RegisterTool(
                "word_read_document",
                "Legge il testo del documento attivo di Word o di un documento aperto specifico (via doc_id/nome). Supporta offset per documenti lunghi.",
                McpServer.Schema((props, _) =>
                {
                    props["doc_id"] = McpServer.Prop("string", "Opzionale: doc_id (FullName) o nome di un documento aperto in Word (default: documento attivo).");
                    props["offset"] = McpServer.IntProp("Carattere iniziale da cui leggere (0-based, default 0).", 0, 50000000);
                    props["max_chars"] = McpServer.IntProp("Lunghezza massima in caratteri (1-200000, default 80000).", 1, ComHelper.MaxAllowedChars);
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ReadDocument(
                    McpServer.GetString(args, "doc_id"),
                    McpServer.GetInt(args, "offset", 0),
                    McpServer.GetInt(args, "max_chars", 80000))));

            server.RegisterTool(
                "word_compare_documents",
                "Confronta due documenti Word aperti paragrafo per paragrafo (evidenziando le parole cambiate [-vecchia-] [+nuova+] nei paragrafi modificati). Con open_word_compare=true apre un nuovo documento Word con Revisioni (disabilitato con --sola-lettura).",
                McpServer.Schema((props, req) =>
                {
                    props["doc1_id"] = McpServer.Prop("string", "doc_id (FullName) o nome del primo documento (originale) da word_status.");
                    props["doc2_id"] = McpServer.Prop("string", "doc_id (FullName) o nome del secondo documento (revisionato) da word_status.");
                    props["max_diffs"] = McpServer.IntProp("Numero massimo di differenze da mostrare (1-500, default 150).", 1, 500);
                    props["open_word_compare"] = McpServer.Prop("boolean", "Se true apre un nuovo documento Word con CompareDocuments (Revisioni / Track Changes). Non consentito in modalità --sola-lettura (default false).");
                    req.Add("doc1_id");
                    req.Add("doc2_id");
                }),
                McpServer.ReadOnly(),
                args => CompareDocuments(
                    McpServer.GetString(args, "doc1_id"),
                    McpServer.GetString(args, "doc2_id"),
                    McpServer.GetInt(args, "max_diffs", 150),
                    McpServer.GetBool(args, "open_word_compare", false),
                    server.IsReadOnly));

            server.RegisterTool(
                "word_replace_selection",
                "Sostituisce la selezione corrente in Word con testo semplice (niente Markdown) operando su Range.Duplicate. Richiede expected_doc ed expected_state_token da word_read_selection. Rifiuta selezioni che attraversano più celle di tabella.",
                McpServer.Schema((props, req) =>
                {
                    props["text"] = McpServer.Prop("string", "Nuovo testo semplice da inserire al posto della selezione (stringa vuota esplicita solo per cancellare).");
                    props["expected_doc"] = McpServer.Prop("string", "doc_id (FullName) del documento restituito da word_read_selection.");
                    props["expected_state_token"] = McpServer.Prop("string", "state_token restituito da word_read_selection; include documento, story, posizione e testo.");
                    req.Add("text");
                    req.Add("expected_doc");
                    req.Add("expected_state_token");
                }),
                McpServer.Mutating(destructive: true),
                args => ReplaceSelection(
                    McpServer.GetString(args, "text"),
                    McpServer.GetString(args, "expected_doc"),
                    McpServer.GetString(args, "expected_state_token")));

            server.RegisterTool(
                "word_insert_text",
                "Inserisce testo semplice (niente Markdown) nel documento attivo di Word verificando expected_doc. Con position='cursor' usa il cursore attuale (passa expected_state_token da word_read_selection se vuoi bloccare in caso di cursore spostato).",
                McpServer.Schema((props, req) =>
                {
                    props["text"] = McpServer.Prop("string", "Testo semplice da inserire.");
                    props["expected_doc"] = McpServer.Prop("string", "doc_id (FullName) del documento restituito da word_status/word_read_document.");
                    props["position"] = McpServer.EnumProp("Posizione di inserimento: 'cursor' (default), 'start' o 'end'.", "cursor", "start", "end");
                    props["expected_state_token"] = McpServer.Prop("string", "Opzionale per position='cursor': se fornito, verifica che cursore/selezione non siano cambiati rispetto a word_read_selection.");
                    req.Add("text");
                    req.Add("expected_doc");
                }),
                McpServer.Mutating(destructive: false),
                args => InsertText(
                    McpServer.GetString(args, "text"),
                    McpServer.GetString(args, "expected_doc"),
                    McpServer.GetString(args, "position", "cursor"),
                    McpServer.GetString(args, "expected_state_token")));

            server.RegisterTool(
                "word_list_tables",
                "Elenca le tabelle presenti nel documento attivo di Word (table_index, table_token stabile, righe x colonne e anteprima). Nota: il table_token dipende anche dalla posizione nel documento; se modifichi testo sopra una tabella, richiama word_list_tables per ottenere il nuovo token.",
                McpServer.EmptySchema(),
                McpServer.ReadOnly(),
                _ => ComHelper.RetryRead(ListTables));

            server.RegisterTool(
                "word_read_table",
                "Legge il contenuto di una tabella specifica in Word (table_index 1-based) restituendo doc_id e table_token. Nota: se modifichi testo sopra la tabella, il table_token cambia e va riletto con word_list_tables/word_read_table.",
                McpServer.Schema((props, req) =>
                {
                    props["table_index"] = McpServer.IntProp("Indice 1-based della tabella nel documento attivo.", 1, 500);
                    props["max_rows"] = McpServer.IntProp("Numero massimo di righe da leggere (1-200, default 100).", 1, 200);
                    req.Add("table_index");
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ReadTable(
                    McpServer.GetInt(args, "table_index", 1),
                    McpServer.GetInt(args, "max_rows", 100))));

            server.RegisterTool(
                "word_edit_table_cell",
                "Modifica il testo di una SINGOLA cella (table_index, row, col) preservando la struttura della tabella. Richiede expected_doc ed expected_table_token (da word_list_tables/word_read_table; se modifichi testo sopra la tabella, rileggi prima con word_list_tables). Se la cella è occupata e overwrite!=true blocca restituendo cell_state_token.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_doc"] = McpServer.Prop("string", "doc_id (FullName) restituito da word_list_tables/word_read_table.");
                    props["table_index"] = McpServer.IntProp("Indice 1-based della tabella.", 1, 500);
                    props["expected_table_token"] = McpServer.Prop("string", "table_token restituito da word_list_tables/word_read_table (impedisce scritture su un'altra tabella).");
                    props["row"] = McpServer.IntProp("Riga 1-based della cella.", 1, 5000);
                    props["col"] = McpServer.IntProp("Colonna 1-based della cella.", 1, 200);
                    props["text"] = McpServer.Prop("string", "Nuovo testo semplice da inserire nella cella.");
                    props["overwrite"] = McpServer.Prop("boolean", "Se true sovrascrive una cella occupata SOLO se expected_state_token coincide (default false).");
                    props["expected_state_token"] = McpServer.Prop("string", "Obbligatorio quando overwrite=true su cella occupata; se fornito viene sempre validato.");
                    req.Add("expected_doc");
                    req.Add("table_index");
                    req.Add("expected_table_token");
                    req.Add("row");
                    req.Add("col");
                    req.Add("text");
                }),
                McpServer.Mutating(destructive: true),
                args => EditTableCell(
                    McpServer.GetString(args, "expected_doc"),
                    McpServer.GetInt(args, "table_index", 0),
                    McpServer.GetString(args, "expected_table_token"),
                    McpServer.GetInt(args, "row", 0),
                    McpServer.GetInt(args, "col", 0),
                    McpServer.GetString(args, "text"),
                    McpServer.GetBool(args, "overwrite", false),
                    McpServer.GetString(args, "expected_state_token")));

            server.RegisterTool(
                "word_insert_table",
                "Crea una NUOVA tabella in Word da una matrice 2D rettangolare al cursore ('cursor') o in fondo al documento ('end'). Singolo Ctrl+Z.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_doc"] = McpServer.Prop("string", "doc_id (FullName) del documento attivo.");
                    props["rows"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Matrice rettangolare 2D di stringhe (max 50 righe x 15 colonne; prima riga = intestazione).",
                        ["items"] = new JsonObject { ["type"] = "array" }
                    };
                    props["position"] = McpServer.EnumProp("Posizione: 'cursor' (default) o 'end'.", "cursor", "end");
                    req.Add("expected_doc");
                    req.Add("rows");
                }),
                McpServer.Mutating(destructive: false),
                args => InsertTable(args));

            server.RegisterTool(
                "word_batch_edit_cells",
                "Modifica in UN SOLO COLPO e UN SOLO Ctrl+Z fino a 200 celle di una o più tabelle Word preservando integralmente tabelle, colonne e impaginazione (ideale per CV Europass, moduli e documenti multi-tabella). Esegue preflight completo su tutte le celle prima di scrivere; se celle occupate e overwrite!=true blocca restituendo batch_state_token. Nota: se modifichi il documento sopra le tabelle, rileggi i table_token con word_list_tables.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_doc"] = McpServer.Prop("string", "doc_id (FullName) da word_list_tables/word_read_table.");
                    props["edits"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Elenco di 1-200 modifiche: ogni elemento è {\"table_index\": 1, \"expected_table_token\": \"...\", \"row\": 1, \"col\": 2, \"text\": \"...\"}.",
                        ["items"] = new JsonObject { ["type"] = "object" }
                    };
                    props["overwrite"] = McpServer.Prop("boolean", "Se true sovrascrive celle occupate SOLO se expected_state_token coincide con batch_state_token (default false).");
                    props["expected_state_token"] = McpServer.Prop("string", "Obbligatorio quando overwrite=true su celle occupate; se fornito viene sempre validato.");
                    req.Add("expected_doc");
                    req.Add("edits");
                }),
                McpServer.Mutating(destructive: true),
                args => BatchEditTableCells(args));
        }

        internal static string ComputeTableToken(string docFullName, int rangeStart, int rows, int cols, string firstRowSignature)
        {
            string sig = firstRowSignature ?? string.Empty;
            return ComHelper.Sha256Short($"{docFullName?.Length ?? 0}:{docFullName}|st:{rangeStart}|{rows}x{cols}|{sig.Length}:{sig}");
        }

        internal static string ComputeCellToken(string docFullName, string tableToken, int row, int col, string cellText)
        {
            string t = cellText ?? string.Empty;
            return ComHelper.Sha256Short($"{docFullName?.Length ?? 0}:{docFullName}|tt:{tableToken}|r:{row}|c:{col}|{t.Length}:{t}");
        }

        internal static string ComputeSelectionToken(string docFullName, int storyType, int start, int end, string rawText)
        {
            string t = rawText ?? string.Empty;
            return ComHelper.Sha256Short($"{docFullName?.Length ?? 0}:{docFullName}|st:{storyType}|{start}:{end}|{t.Length}:{t}");
        }

        private static string GetStatus()
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic active = null;
            try
            {
                docs = app.Documents;
                int count = (int)docs.Count;
                if (count == 0)
                    return "Microsoft Word è aperto ma non ci sono documenti aperti.";

                var sb = new StringBuilder();
                active = app.ActiveDocument;
                sb.AppendLine($"Documento attivo: {active.Name}");
                sb.AppendLine($"doc_id: {active.FullName}");
                sb.AppendLine($"Documenti aperti ({count}):");
                for (int i = 1; i <= count; i++)
                {
                    dynamic d = docs.Item(i);
                    try { sb.AppendLine($"  - {d.Name} (doc_id: {d.FullName})"); }
                    finally { ComHelper.Release(d); }
                }
                return sb.ToString().TrimEnd();
            }
            finally
            {
                ComHelper.Release(active);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string ReadSelection(int maxChars)
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic sel = null;
            dynamic selRng = null;
            dynamic dupRng = null;
            dynamic doc = null;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                sel = app.Selection;
                selRng = sel.Range;
                dupRng = selRng.Duplicate;
                doc = dupRng.Document;

                string docFull = (string)doc.FullName;
                string docName = (string)doc.Name;
                int storyType = (int)dupRng.StoryType;
                int start = (int)dupRng.Start;
                int end = (int)dupRng.End;
                string raw = (string)(dupRng.Text ?? string.Empty);
                string token = ComputeSelectionToken(docFull, storyType, start, end, raw);

                string cleaned = NormalizeWordText(raw);
                var sb = new StringBuilder();
                sb.AppendLine($"doc_id: {docFull}");
                sb.AppendLine($"document: {docName}");
                sb.AppendLine($"story_type: {storyType} | start: {start} | end: {end}");
                sb.AppendLine($"state_token: {token}");
                sb.AppendLine();
                if (string.IsNullOrWhiteSpace(cleaned))
                    sb.Append("(Nessun testo selezionato: il cursore è fermo. Puoi usare questo state_token per inserire al cursore o usare word_read_document.)");
                else
                    sb.Append(ComHelper.Truncate(cleaned, maxChars, out _));

                return sb.ToString();
            }
            finally
            {
                ComHelper.Release(doc);
                ComHelper.Release(dupRng);
                ComHelper.Release(selRng);
                ComHelper.Release(sel);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static dynamic ResolveOpenDocument(dynamic docs, string docIdOrName)
        {
            int count = (int)docs.Count;
            if (count == 0)
                throw new InvalidOperationException("Nessun documento aperto in Word.");

            string query = (docIdOrName ?? string.Empty).Trim();
            for (int i = 1; i <= count; i++)
            {
                dynamic d = docs.Item(i);
                string fn = (string)d.FullName;
                string nm = (string)d.Name;
                if (string.Equals(fn, query, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(nm, query, StringComparison.OrdinalIgnoreCase))
                    return d;
                ComHelper.Release(d);
            }
            throw new InvalidOperationException($"Documento '{docIdOrName}' non trovato tra i documenti aperti in Word. Usa word_status per vedere i doc_id.");
        }

        private static string ReadDocument(string docId, int offset, int maxChars)
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic content = null;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                doc = string.IsNullOrWhiteSpace(docId) ? app.ActiveDocument : ResolveOpenDocument(docs, docId);
                string name = (string)doc.Name;
                string fullName = (string)doc.FullName;
                content = doc.Content;
                string raw = (string)(content.Text ?? string.Empty);

                string cleaned = NormalizeWordText(raw);
                int totalLen = cleaned.Length;
                int safeOffset = ComHelper.Clamp(offset, 0, totalLen);
                string slice = safeOffset < totalLen ? cleaned.Substring(safeOffset) : string.Empty;
                string body = ComHelper.Truncate(slice, maxChars, out bool truncated);

                var sb = new StringBuilder();
                sb.AppendLine($"doc_id: {fullName}");
                sb.AppendLine($"document: {name}");
                sb.AppendLine($"total_chars: {totalLen} | offset: {safeOffset} | has_more: {truncated}");
                sb.AppendLine();
                sb.Append(body);
                return sb.ToString();
            }
            finally
            {
                ComHelper.Release(content);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string CompareDocuments(string doc1Id, string doc2Id, int maxDiffs, bool openWordCompare, bool serverReadOnly)
        {
            if (openWordCompare && serverReadOnly)
                throw new InvalidOperationException("open_word_compare=true crea un nuovo documento Word e non è consentito in modalità --sola-lettura.");

            // Fase 1 (sola lettura, con RetryRead): legge i due documenti
            var (name1, txt1, name2, txt2) = ComHelper.RetryRead(() =>
            {
                dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
                dynamic docs = null, d1 = null, d2 = null, c1 = null, c2 = null;
                try
                {
                    docs = app.Documents;
                    d1 = ResolveOpenDocument(docs, doc1Id);
                    d2 = ResolveOpenDocument(docs, doc2Id);
                    c1 = d1.Content;
                    c2 = d2.Content;
                    return (
                        (string)d1.Name,
                        NormalizeWordText((string)(c1.Text ?? string.Empty)),
                        (string)d2.Name,
                        NormalizeWordText((string)(c2.Text ?? string.Empty))
                    );
                }
                finally
                {
                    ComHelper.Release(c2);
                    ComHelper.Release(c1);
                    ComHelper.Release(d2);
                    ComHelper.Release(d1);
                    ComHelper.Release(docs);
                    ComHelper.Release(app);
                }
            });

            string diffReport = ComputeParagraphDiff(name1, txt1, name2, txt2, maxDiffs);

            // Fase 2 (FUORI da RetryRead per non aprire mai due finestre di confronto in caso di errore COM)
            if (openWordCompare)
            {
                dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
                dynamic docs = null, d1 = null, d2 = null, cmpDoc = null;
                try
                {
                    docs = app.Documents;
                    d1 = ResolveOpenDocument(docs, doc1Id);
                    d2 = ResolveOpenDocument(docs, doc2Id);
                    cmpDoc = app.CompareDocuments(d1, d2);
                    diffReport += $"\n\n[Aperto nuovo documento di confronto Word con Revisioni: '{cmpDoc.Name}'.]";
                }
                catch (Exception ex)
                {
                    diffReport += $"\n\n[Avviso: CompareDocuments di Word non riuscito: {ex.Message}]";
                }
                finally
                {
                    ComHelper.Release(cmpDoc);
                    ComHelper.Release(d2);
                    ComHelper.Release(d1);
                    ComHelper.Release(docs);
                    ComHelper.Release(app);
                }
            }
            return diffReport;
        }

        internal static string ComputeParagraphDiff(string label1, string text1, string label2, string text2, int maxDiffs)
        {
            string[] p1 = (text1 ?? string.Empty).Split('\n');
            string[] p2 = (text2 ?? string.Empty).Split('\n');

            int prefix = 0;
            while (prefix < p1.Length && prefix < p2.Length && p1[prefix] == p2[prefix]) prefix++;
            int suffix = 0;
            while (suffix < (p1.Length - prefix) && suffix < (p2.Length - prefix) &&
                   p1[p1.Length - 1 - suffix] == p2[p2.Length - 1 - suffix]) suffix++;

            int m = p1.Length - prefix - suffix;
            int n = p2.Length - prefix - suffix;

            if (m == 0 && n == 0)
                return $"Confronto '{label1}' vs '{label2}': NESSUNA DIFFERENZA testuale ({p1.Length} paragrafi identici).";

            // Raccoglie operazioni grezze ('-', '+') per poter accoppiare i paragrafi modificati e fare il diff a parole
            var ops = new System.Collections.Generic.List<(char kind, int idx, string text)>();

            if ((long)m * n <= 2500000)
            {
                int[,] dp = new int[m + 1, n + 1];
                for (int i = m - 1; i >= 0; i--)
                {
                    for (int j = n - 1; j >= 0; j--)
                    {
                        dp[i, j] = p1[prefix + i] == p2[prefix + j]
                            ? dp[i + 1, j + 1] + 1
                            : Math.Max(dp[i + 1, j], dp[i, j + 1]);
                    }
                }

                int ii = 0, jj = 0;
                while (ii < m || jj < n)
                {
                    if (ii < m && jj < n && p1[prefix + ii] == p2[prefix + jj])
                    {
                        ii++; jj++;
                    }
                    else if (ii < m && (jj == n || dp[ii + 1, jj] >= dp[ii, jj + 1]))
                    {
                        ops.Add(('-', prefix + ii + 1, p1[prefix + ii]));
                        ii++;
                    }
                    else
                    {
                        ops.Add(('+', prefix + jj + 1, p2[prefix + jj]));
                        jj++;
                    }
                }
            }
            else
            {
                int maxLen = Math.Max(m, n);
                for (int k = 0; k < maxLen; k++)
                {
                    string s1 = k < m ? p1[prefix + k] : null;
                    string s2 = k < n ? p2[prefix + k] : null;
                    if (s1 == s2) continue;
                    if (s1 != null) ops.Add(('-', prefix + k + 1, s1));
                    if (s2 != null) ops.Add(('+', prefix + k + 1, s2));
                }
            }

            var diffs = new System.Collections.Generic.List<string>();
            int added = 0, removed = 0, modified = 0;

            for (int i = 0; i < ops.Count; i++)
            {
                if (ops[i].kind == '-' && i + 1 < ops.Count && ops[i + 1].kind == '+' &&
                    TryFormatWordLevelDiff(ops[i].text, ops[i + 1].text, out string inlineDiff))
                {
                    modified++;
                    if (diffs.Count < maxDiffs)
                        diffs.Add($"~ [Doc1 §{ops[i].idx} -> Doc2 §{ops[i + 1].idx}] {inlineDiff}");
                    i++; // consuma anche il '+' accoppiato
                }
                else if (ops[i].kind == '-')
                {
                    removed++;
                    if (diffs.Count < maxDiffs)
                        diffs.Add($"- [Doc1 §{ops[i].idx}] {ops[i].text}");
                }
                else
                {
                    added++;
                    if (diffs.Count < maxDiffs)
                        diffs.Add($"+ [Doc2 §{ops[i].idx}] {ops[i].text}");
                }
            }

            var sb = new StringBuilder();
            sb.AppendLine($"Confronto '{label1}' ({p1.Length} paragrafi) vs '{label2}' ({p2.Length} paragrafi):");
            sb.AppendLine($"Differenze: ~{modified} modificati, -{removed} rimossi, +{added} aggiunti.");
            int totalEntries = modified + removed + added;
            if (totalEntries > diffs.Count)
                sb.AppendLine($"(Mostrate le prime {diffs.Count} differenze su {totalEntries} totali)");
            sb.AppendLine();
            foreach (var d in diffs) sb.AppendLine(d);
            return ComHelper.Truncate(sb.ToString().TrimEnd(), 80000, out _);
        }

        private static bool TryFormatWordLevelDiff(string oldText, string newText, out string formatted)
        {
            formatted = null;
            if (string.IsNullOrWhiteSpace(oldText) || string.IsNullOrWhiteSpace(newText)) return false;

            string[] w1 = oldText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            string[] w2 = newText.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (w1.Length == 0 || w2.Length == 0 || (long)w1.Length * w2.Length > 250000) return false;

            int[,] dp = new int[w1.Length + 1, w2.Length + 1];
            for (int i = w1.Length - 1; i >= 0; i--)
                for (int j = w2.Length - 1; j >= 0; j--)
                    dp[i, j] = w1[i] == w2[j] ? dp[i + 1, j + 1] + 1 : Math.Max(dp[i + 1, j], dp[i, j + 1]);

            int common = dp[0, 0];
            double similarity = (2.0 * common) / (w1.Length + w2.Length);
            if (similarity < 0.4) return false; // Troppo diversi: meglio mostrarli come - e +

            var outParts = new System.Collections.Generic.List<string>();
            int ii = 0, jj = 0;
            while (ii < w1.Length || jj < w2.Length)
            {
                if (ii < w1.Length && jj < w2.Length && w1[ii] == w2[jj])
                {
                    outParts.Add(w1[ii]);
                    ii++; jj++;
                }
                else
                {
                    var rem = new System.Collections.Generic.List<string>();
                    var add = new System.Collections.Generic.List<string>();
                    while (ii < w1.Length && (jj == w2.Length || dp[ii + 1, jj] >= dp[ii, jj + 1]) && (jj == w2.Length || w1[ii] != w2[jj]))
                        rem.Add(w1[ii++]);
                    while (jj < w2.Length && (ii == w1.Length || dp[ii, jj + 1] > dp[ii + 1, jj]) && (ii == w1.Length || w1[ii] != w2[jj]))
                        add.Add(w2[jj++]);
                    if (rem.Count > 0) outParts.Add($"[-{string.Join(" ", rem)}-]");
                    if (add.Count > 0) outParts.Add($"[+{string.Join(" ", add)}+]");
                }
            }
            formatted = string.Join(" ", outParts);
            return true;
        }

        private static string ReplaceSelection(string text, string expectedDoc, string expectedStateToken)
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic sel = null;
            dynamic selRng = null;
            dynamic dupRng = null;
            dynamic doc = null;
            dynamic cells = null;
            dynamic undoRec = null;
            bool undoStarted = false;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                // Acquisisce un Range.Duplicate fisso e verifica il suo Document, StoryType e contenuto.
                sel = app.Selection;
                selRng = sel.Range;
                dupRng = selRng.Duplicate;
                doc = dupRng.Document;

                string actualDoc = (string)doc.FullName;
                if (!string.Equals(actualDoc, expectedDoc, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"Documento cambiato: atteso '{expectedDoc}', ma la selezione appartiene a '{actualDoc}'. Rileggi prima di scrivere.");
                }

                int storyType = (int)dupRng.StoryType;
                int start = (int)dupRng.Start;
                int end = (int)dupRng.End;
                string raw = (string)(dupRng.Text ?? string.Empty);
                string actualToken = ComputeSelectionToken(actualDoc, storyType, start, end, raw);
                if (!string.Equals(actualToken, expectedStateToken, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Selezione, story o contenuto cambiati in Word (atteso token '{expectedStateToken}', attuale '{actualToken}'). Rileggi con word_read_selection prima di scrivere.");
                }

                // Rifiuta sostituzioni che attraversano più celle di tabella (rovinerebbero la struttura).
                try
                {
                    if ((bool)dupRng.Information(12)) // 12 = wdWithInTable
                    {
                        cells = dupRng.Cells;
                        if ((int)cells.Count > 1)
                            throw new InvalidOperationException("La selezione attraversa più celle di tabella: sostituzione rifiutata per non corrompere la tabella.");
                    }
                }
                catch (InvalidOperationException) { throw; }
                catch { }

                try
                {
                    undoRec = app.UndoRecord;
                    undoRec.StartCustomRecord("mcp4office: sostituisci selezione");
                    undoStarted = true;
                }
                catch { undoStarted = false; }

                string normalized = (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r");

                // Protezione triplo clic: se la selezione termina con fine paragrafo \r e il nuovo testo no,
                // arretra la fine di 1 carattere per non fondere il paragrafo col successivo.
                if (raw.EndsWith("\r") && !normalized.EndsWith("\r") && end > start)
                {
                    dupRng.MoveEnd(1, -1); // 1 = wdCharacter
                }

                ReplaceRangePreservingFormatting(dupRng, text ?? string.Empty);

                string undoNote = undoStarted
                    ? "Annullabile con un singolo Ctrl+Z."
                    : "Nota: UndoRecord non disponibile in questa sessione; verificare l'annullamento in Word.";
                return $"Testo sostituito in '{doc.Name}' ({normalized.Length} caratteri, formattazione paragrafi/stili preservata). {undoNote}";
            }
            finally
            {
                if (undoStarted && undoRec != null)
                {
                    try { undoRec.EndCustomRecord(); } catch { }
                }
                ComHelper.Release(undoRec);
                ComHelper.Release(cells);
                ComHelper.Release(doc);
                ComHelper.Release(dupRng);
                ComHelper.Release(selRng);
                ComHelper.Release(sel);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string InsertText(string text, string expectedDoc, string position, string expectedStateToken)
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic sel = null;
            dynamic selRng = null;
            dynamic rng = null;
            dynamic rngDoc = null;
            dynamic undoRec = null;
            bool undoStarted = false;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                string normalized = (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r");
                string pos = (position ?? "cursor").Trim().ToLowerInvariant();

                if (pos == "cursor")
                {
                    sel = app.Selection;
                    selRng = sel.Range;
                    rng = selRng.Duplicate;
                    rngDoc = rng.Document;
                    string actualDoc = (string)rngDoc.FullName;
                    if (!string.Equals(actualDoc, expectedDoc, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException($"Documento cambiato: atteso '{expectedDoc}', attuale '{actualDoc}'.");

                    if (!string.IsNullOrWhiteSpace(expectedStateToken))
                    {
                        int storyType = (int)rng.StoryType;
                        int start = (int)rng.Start;
                        int end = (int)rng.End;
                        string raw = (string)(rng.Text ?? string.Empty);
                        string actualToken = ComputeSelectionToken(actualDoc, storyType, start, end, raw);
                        if (!string.Equals(actualToken, expectedStateToken.Trim(), StringComparison.Ordinal))
                            throw new InvalidOperationException($"Posizione del cursore cambiata (atteso token '{expectedStateToken}', attuale '{actualToken}').");
                    }

                    try
                    {
                        undoRec = app.UndoRecord;
                        undoRec.StartCustomRecord("mcp4office: inserisci testo");
                        undoStarted = true;
                    }
                    catch { undoStarted = false; }

                    rng.Collapse(0); // wdCollapseEnd
                    rng.InsertAfter(normalized);
                    string docName = (string)rngDoc.Name;
                    string undoNote = undoStarted ? "Annullabile con un singolo Ctrl+Z." : "Nota: UndoRecord non disponibile.";
                    return $"Testo inserito in '{docName}' (posizione: cursor, {normalized.Length} caratteri). {undoNote}";
                }
                else
                {
                    doc = app.ActiveDocument;
                    string actualDoc = (string)doc.FullName;
                    if (!string.Equals(actualDoc, expectedDoc, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException($"Documento cambiato: atteso '{expectedDoc}', attuale '{actualDoc}'.");

                    try
                    {
                        undoRec = app.UndoRecord;
                        undoRec.StartCustomRecord("mcp4office: inserisci testo");
                        undoStarted = true;
                    }
                    catch { undoStarted = false; }

                    if (pos == "start")
                    {
                        rng = doc.Range(0, 0);
                        rng.InsertBefore(normalized + "\r");
                    }
                    else
                    {
                        rng = doc.Content;
                        rng.Collapse(0);
                        rng.InsertAfter("\r" + normalized);
                    }

                    string undoNote = undoStarted ? "Annullabile con un singolo Ctrl+Z." : "Nota: UndoRecord non disponibile.";
                    return $"Testo inserito in '{doc.Name}' (posizione: {pos}, {normalized.Length} caratteri). {undoNote}";
                }
            }
            finally
            {
                if (undoStarted && undoRec != null)
                {
                    try { undoRec.EndCustomRecord(); } catch { }
                }
                ComHelper.Release(undoRec);
                ComHelper.Release(rngDoc);
                ComHelper.Release(rng);
                ComHelper.Release(selRng);
                ComHelper.Release(sel);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string GetTableTokenFromCom(string docFullName, dynamic tbl, int rCount, int cCount)
        {
            int rngStart = 0;
            try { rngStart = (int)tbl.Range.Start; } catch { }
            var firstRow = new StringBuilder();
            int cLim = Math.Min(cCount, 12);
            for (int c = 1; c <= cLim; c++)
            {
                try
                {
                    dynamic cell = tbl.Cell(1, c);
                    try { firstRow.Append(NormalizeCellText((string)(cell.Range.Text ?? string.Empty))).Append('|'); }
                    finally { ComHelper.Release(cell); }
                }
                catch { firstRow.Append("?|"); }
            }
            return ComputeTableToken(docFullName, rngStart, rCount, cCount, firstRow.ToString());
        }

        private static string ListTables()
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic tables = null;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                doc = app.ActiveDocument;
                string fullName = (string)doc.FullName;
                tables = doc.Tables;
                int count = (int)tables.Count;

                var sb = new StringBuilder();
                sb.AppendLine($"doc_id: {fullName}");
                sb.AppendLine($"Documento: {doc.Name} | Tabelle trovate: {count}");
                if (count == 0) return sb.ToString().TrimEnd();

                for (int i = 1; i <= count; i++)
                {
                    dynamic tbl = tables.Item(i);
                    dynamic rows = null;
                    dynamic cols = null;
                    try
                    {
                        rows = tbl.Rows;
                        cols = tbl.Columns;
                        int rCount = (int)rows.Count;
                        int cCount = (int)cols.Count;
                        string tTok = GetTableTokenFromCom(fullName, tbl, rCount, cCount);
                        string preview = string.Empty;
                        try
                        {
                            dynamic c1 = tbl.Cell(1, 1);
                            try { preview = NormalizeCellText((string)(c1.Range.Text ?? string.Empty)); }
                            finally { ComHelper.Release(c1); }
                            if (preview.Length > 60) preview = preview.Substring(0, 60) + "…";
                        }
                        catch { }
                        sb.AppendLine($"  Tabella {i} (table_token: {tTok}): {rCount} righe x {cCount} colonne" + (preview.Length > 0 ? $" — cella(1,1): \"{preview}\"" : ""));
                    }
                    finally
                    {
                        ComHelper.Release(cols);
                        ComHelper.Release(rows);
                        ComHelper.Release(tbl);
                    }
                }
                return sb.ToString().TrimEnd();
            }
            finally
            {
                ComHelper.Release(tables);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string ReadTable(int tableIndex, int maxRows)
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic tables = null;
            dynamic tbl = null;
            dynamic rows = null;
            dynamic cols = null;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                doc = app.ActiveDocument;
                string fullName = (string)doc.FullName;
                tables = doc.Tables;
                int totalTables = (int)tables.Count;
                if (tableIndex < 1 || tableIndex > totalTables)
                    throw new ArgumentOutOfRangeException(nameof(tableIndex), $"Tabella {tableIndex} inesistente (tabelle nel documento: {totalTables}).");

                tbl = tables.Item(tableIndex);
                rows = tbl.Rows;
                cols = tbl.Columns;
                int totalRows = (int)rows.Count;
                int totalCols = (int)cols.Count;
                int rLimit = Math.Min(totalRows, maxRows);
                int cLimit = Math.Min(totalCols, 30);
                string tTok = GetTableTokenFromCom(fullName, tbl, totalRows, totalCols);

                var sb = new StringBuilder();
                sb.AppendLine($"doc_id: {fullName}");
                sb.AppendLine($"table_index: {tableIndex} | table_token: {tTok}");
                sb.AppendLine($"Dimensioni: {totalRows} righe x {totalCols} colonne (lette {rLimit}x{cLimit}):");

                for (int r = 1; r <= rLimit; r++)
                {
                    var rowText = new System.Collections.Generic.List<string>(cLimit);
                    for (int c = 1; c <= cLimit; c++)
                    {
                        try
                        {
                            dynamic cell = tbl.Cell(r, c);
                            try
                            {
                                string txt = NormalizeCellText((string)(cell.Range.Text ?? string.Empty));
                                rowText.Add($"({r},{c}): {txt}");
                            }
                            finally { ComHelper.Release(cell); }
                        }
                        catch
                        {
                            rowText.Add($"({r},{c}): [cella unita/non indirizzabile]");
                        }
                    }
                    sb.AppendLine(string.Join(" | ", rowText));
                }

                return ComHelper.Truncate(sb.ToString().TrimEnd(), 80000, out _);
            }
            finally
            {
                ComHelper.Release(cols);
                ComHelper.Release(rows);
                ComHelper.Release(tbl);
                ComHelper.Release(tables);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string EditTableCell(string expectedDoc, int tableIndex, string expectedTableToken, int row, int col, string text, bool overwrite, string expectedToken)
        {
            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic tables = null;
            dynamic tbl = null;
            dynamic rows = null;
            dynamic cols = null;
            dynamic cell = null;
            dynamic cellRange = null;
            dynamic undoRec = null;
            bool undoStarted = false;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                doc = app.ActiveDocument;
                string actualDoc = (string)doc.FullName;
                if (!string.Equals(actualDoc, expectedDoc, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Documento cambiato: atteso '{expectedDoc}', attuale '{actualDoc}'.");

                tables = doc.Tables;
                int totalTables = (int)tables.Count;
                if (tableIndex < 1 || tableIndex > totalTables)
                    throw new ArgumentOutOfRangeException(nameof(tableIndex), $"Tabella {tableIndex} inesistente (tabelle: {totalTables}).");

                tbl = tables.Item(tableIndex);
                rows = tbl.Rows;
                cols = tbl.Columns;
                string actualTableToken = GetTableTokenFromCom(actualDoc, tbl, (int)rows.Count, (int)cols.Count);
                if (!string.Equals(actualTableToken, (expectedTableToken ?? string.Empty).Trim(), StringComparison.Ordinal))
                    throw new InvalidOperationException($"Identità della tabella {tableIndex} cambiata (atteso table_token '{expectedTableToken}', attuale '{actualTableToken}'). Rileggi con word_list_tables/word_read_table.");

                try
                {
                    cell = tbl.Cell(row, col);
                }
                catch
                {
                    throw new ArgumentOutOfRangeException(nameof(row), $"Cella ({row},{col}) non valida o unita nella tabella {tableIndex}.");
                }

                cellRange = cell.Range;
                string currentText = NormalizeCellText((string)(cellRange.Text ?? string.Empty));
                string currentToken = ComputeCellToken(actualDoc, actualTableToken, row, col, currentText);

                if (!string.IsNullOrWhiteSpace(expectedToken) && !string.Equals(expectedToken.Trim(), currentToken, StringComparison.Ordinal))
                    throw new InvalidOperationException($"Contenuto della cella ({row},{col}) cambiato (atteso '{expectedToken}', attuale '{currentToken}'). Rileggi la tabella.");

                if (!string.IsNullOrEmpty(currentText))
                {
                    if (!overwrite)
                    {
                        return $"SCRITTURA CELLA BLOCCATA: tabella {tableIndex} cella ({row},{col}) contiene già testo: \"{ComHelper.Truncate(currentText, 200, out _)}\"\n" +
                               $"cell_state_token: {currentToken}\n" +
                               "Chiedi conferma all'utente e richiama word_edit_table_cell con overwrite=true e questo expected_state_token.";
                    }
                    if (string.IsNullOrWhiteSpace(expectedToken))
                        throw new ArgumentException("Quando overwrite=true su cella occupata, 'expected_state_token' è obbligatorio.");
                }

                try
                {
                    undoRec = app.UndoRecord;
                    undoRec.StartCustomRecord($"mcp4office: modifica cella ({row},{col})");
                    undoStarted = true;
                }
                catch { undoStarted = false; }

                string normalized = (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r");
                ReplaceRangePreservingFormatting(cellRange, text ?? string.Empty);

                string undoNote = undoStarted ? "Annullabile con un singolo Ctrl+Z." : "Nota: UndoRecord non disponibile.";
                return $"Cella ({row},{col}) della tabella {tableIndex} in '{doc.Name}' aggiornata ({normalized.Length} caratteri, stili preservati). {undoNote}";
            }
            finally
            {
                if (undoStarted && undoRec != null)
                {
                    try { undoRec.EndCustomRecord(); } catch { }
                }
                ComHelper.Release(undoRec);
                ComHelper.Release(cellRange);
                ComHelper.Release(cell);
                ComHelper.Release(cols);
                ComHelper.Release(rows);
                ComHelper.Release(tbl);
                ComHelper.Release(tables);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        private static string InsertTable(JsonObject args)
        {
            string expectedDoc = McpServer.GetString(args, "expected_doc").Trim();
            string pos = McpServer.GetString(args, "position", "cursor").Trim().ToLowerInvariant();
            var rowsArr = args["rows"] as JsonArray;
            if (rowsArr == null || rowsArr.Count == 0 || rowsArr.Count > 50)
                throw new ArgumentException("Parametro 'rows' deve avere tra 1 e 50 righe.");

            var matrix = new System.Collections.Generic.List<System.Collections.Generic.List<string>>(rowsArr.Count);
            int cols = -1;
            foreach (var rNode in rowsArr)
            {
                var cArr = rNode as JsonArray;
                if (cArr == null || cArr.Count == 0 || cArr.Count > 15)
                    throw new ArgumentException("Ogni riga della tabella deve avere tra 1 e 15 colonne.");
                if (cols < 0) cols = cArr.Count;
                else if (cArr.Count != cols)
                    throw new ArgumentException($"Tabella non rettangolare: attese {cols} colonne, trovate {cArr.Count}.");
                var row = new System.Collections.Generic.List<string>(cols);
                foreach (var cNode in cArr) row.Add(cNode == null ? string.Empty : cNode.ToString());
                matrix.Add(row);
            }

            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic sel = null;
            dynamic targetRng = null;
            dynamic tables = null;
            dynamic tbl = null;
            dynamic undoRec = null;
            bool undoStarted = false;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                doc = app.ActiveDocument;
                string actualDoc = (string)doc.FullName;
                if (!string.Equals(actualDoc, expectedDoc, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Documento cambiato: atteso '{expectedDoc}', attuale '{actualDoc}'.");

                if (pos == "end")
                {
                    targetRng = doc.Content;
                    targetRng.Collapse(0);
                    targetRng.InsertAfter("\r");
                    targetRng.Collapse(0);
                }
                else
                {
                    sel = app.Selection;
                    targetRng = sel.Range.Duplicate;
                    targetRng.Collapse(0);
                }

                try
                {
                    undoRec = app.UndoRecord;
                    undoRec.StartCustomRecord("mcp4office: inserisci tabella");
                    undoStarted = true;
                }
                catch { undoStarted = false; }

                tables = doc.Tables;
                tbl = tables.Add(targetRng, matrix.Count, cols);
                try { tbl.Borders.Enable = 1; } catch { }

                for (int r = 0; r < matrix.Count; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        dynamic cell = tbl.Cell(r + 1, c + 1);
                        try
                        {
                            cell.Range.Text = matrix[r][c].Replace("\r\n", "\r").Replace("\n", "\r");
                            if (r == 0) cell.Range.Bold = 1;
                        }
                        finally { ComHelper.Release(cell); }
                    }
                }

                string undoNote = undoStarted ? "Annullabile con un singolo Ctrl+Z." : "Nota: UndoRecord non disponibile.";
                return $"Inserita tabella {matrix.Count}x{cols} in '{doc.Name}' (posizione: {pos}). {undoNote}";
            }
            finally
            {
                if (undoStarted && undoRec != null)
                {
                    try { undoRec.EndCustomRecord(); } catch { }
                }
                ComHelper.Release(undoRec);
                ComHelper.Release(tbl);
                ComHelper.Release(tables);
                ComHelper.Release(targetRng);
                ComHelper.Release(sel);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        internal static string ComputeBatchCellsToken(string docFullName, System.Collections.Generic.IEnumerable<(int tbl, string tTok, int r, int c, string txt)> cells)
        {
            var sb = new StringBuilder();
            sb.Append(docFullName?.Length ?? 0).Append(':').Append(docFullName).Append('|');
            foreach (var x in cells)
            {
                string t = x.txt ?? string.Empty;
                sb.Append(x.tbl).Append(':').Append(x.tTok).Append(':').Append(x.r).Append(',').Append(x.c).Append(':').Append(t.Length).Append(':').Append(t).Append('\x1e');
            }
            return ComHelper.Sha256Short(sb.ToString());
        }

        private static string BatchEditTableCells(JsonObject args)
        {
            string expectedDoc = McpServer.GetString(args, "expected_doc").Trim();
            bool overwrite = McpServer.GetBool(args, "overwrite", false);
            string expectedToken = McpServer.GetString(args, "expected_state_token").Trim();

            var editsArr = args["edits"] as JsonArray;
            if (editsArr == null || editsArr.Count < 1 || editsArr.Count > 200)
                throw new ArgumentException("Parametro 'edits' deve contenere tra 1 e 200 modifiche.");

            var parsed = new System.Collections.Generic.List<(int tableIndex, string tableToken, int row, int col, string text)>(editsArr.Count);
            var dupCheck = new System.Collections.Generic.HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < editsArr.Count; i++)
            {
                var obj = editsArr[i] as JsonObject
                    ?? throw new ArgumentException($"Elemento edits[{i}] deve essere un oggetto JSON.");
                int tIdx = McpServer.GetInt(obj, "table_index", 0);
                string tTok = McpServer.GetString(obj, "expected_table_token").Trim();
                int r = McpServer.GetInt(obj, "row", 0);
                int c = McpServer.GetInt(obj, "col", 0);
                if (!obj.ContainsKey("text") || obj["text"] == null)
                    throw new ArgumentException($"Elemento edits[{i}]: parametro 'text' obbligatorio.");
                string txt = McpServer.GetString(obj, "text");
                if (tIdx < 1 || r < 1 || c < 1 || string.IsNullOrEmpty(tTok))
                    throw new ArgumentException($"Elemento edits[{i}]: table_index, row, col (>=1) ed expected_table_token obbligatori.");
                string key = $"{tIdx}:{r}:{c}";
                if (!dupCheck.Add(key))
                    throw new ArgumentException($"Cella duplicata in 'edits': tabella {tIdx} ({r},{c}).");
                parsed.Add((tIdx, tTok, r, c, txt));
            }

            dynamic app = ComHelper.GetActiveApp("Word.Application", "Microsoft Word");
            dynamic docs = null;
            dynamic doc = null;
            dynamic tables = null;
            dynamic undoRec = null;
            bool undoStarted = false;
            try
            {
                docs = app.Documents;
                if ((int)docs.Count == 0)
                    throw new InvalidOperationException("Nessun documento aperto in Word.");

                doc = app.ActiveDocument;
                string actualDoc = (string)doc.FullName;
                if (!string.Equals(actualDoc, expectedDoc, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Documento cambiato: atteso '{expectedDoc}', attuale '{actualDoc}'.");

                tables = doc.Tables;
                int totalTables = (int)tables.Count;

                // Fase 1: Preflight completo su tutte le tabelle e celle prima di toccare nulla
                var tableTokenCache = new System.Collections.Generic.Dictionary<int, string>();
                var snapshot = new System.Collections.Generic.List<(int tbl, string tTok, int r, int c, string txt)>(parsed.Count);
                var occupied = new System.Collections.Generic.List<string>();

                foreach (var e in parsed)
                {
                    if (e.tableIndex > totalTables)
                        throw new ArgumentOutOfRangeException(nameof(args), $"Tabella {e.tableIndex} inesistente (tabelle nel documento: {totalTables}).");

                    dynamic tbl = tables.Item(e.tableIndex);
                    try
                    {
                        if (!tableTokenCache.TryGetValue(e.tableIndex, out string actualTableTok))
                        {
                            dynamic rows = tbl.Rows;
                            dynamic cols = tbl.Columns;
                            try { actualTableTok = GetTableTokenFromCom(actualDoc, tbl, (int)rows.Count, (int)cols.Count); }
                            finally { ComHelper.Release(cols); ComHelper.Release(rows); }
                            tableTokenCache[e.tableIndex] = actualTableTok;
                        }

                        if (!string.Equals(actualTableTok, e.tableToken, StringComparison.Ordinal))
                            throw new InvalidOperationException($"Identità della tabella {e.tableIndex} cambiata (atteso '{e.tableToken}', attuale '{actualTableTok}'). Rileggi con word_list_tables.");

                        dynamic cell;
                        try { cell = tbl.Cell(e.row, e.col); }
                        catch { throw new ArgumentOutOfRangeException(nameof(args), $"Cella ({e.row},{e.col}) non valida o unita nella tabella {e.tableIndex}."); }

                        dynamic rng = null;
                        try
                        {
                            rng = cell.Range;
                            string cur = NormalizeCellText((string)(rng.Text ?? string.Empty));
                            snapshot.Add((e.tableIndex, actualTableTok, e.row, e.col, cur));
                            if (!string.IsNullOrEmpty(cur))
                                occupied.Add($"T{e.tableIndex}({e.row},{e.col})");
                        }
                        finally
                        {
                            ComHelper.Release(rng);
                            ComHelper.Release(cell);
                        }
                    }
                    finally { ComHelper.Release(tbl); }
                }

                string batchToken = ComputeBatchCellsToken(actualDoc, snapshot);
                if (!string.IsNullOrEmpty(expectedToken) && !string.Equals(expectedToken, batchToken, StringComparison.Ordinal))
                    throw new InvalidOperationException($"Stato delle celle cambiato (atteso token '{expectedToken}', attuale '{batchToken}'). Rileggi le tabelle.");

                if (occupied.Count > 0)
                {
                    if (!overwrite)
                    {
                        string sample = string.Join(", ", occupied.GetRange(0, Math.Min(occupied.Count, 20)));
                        return $"SCRITTURA MULTIPLA BLOCCATA: {occupied.Count} celle su {parsed.Count} contengono già testo ({sample}).\n" +
                               $"batch_state_token: {batchToken}\n" +
                               "Chiedi conferma all'utente e, solo se autorizzato, richiama word_batch_edit_cells con overwrite=true e questo expected_state_token.";
                    }
                    if (string.IsNullOrEmpty(expectedToken))
                        throw new ArgumentException("Quando overwrite=true su celle occupate, 'expected_state_token' è obbligatorio.");
                }

                // Fase 2: Applicazione in un unico UndoRecord (un solo Ctrl+Z!)
                try
                {
                    undoRec = app.UndoRecord;
                    undoRec.StartCustomRecord($"mcp4office: aggiorna {parsed.Count} celle");
                    undoStarted = true;
                }
                catch { undoStarted = false; }

                var completedCells = new System.Collections.Generic.List<string>(parsed.Count);
                foreach (var e in parsed)
                {
                    dynamic tbl = tables.Item(e.tableIndex);
                    dynamic cell = null;
                    dynamic rng = null;
                    try
                    {
                        cell = tbl.Cell(e.row, e.col);
                        rng = cell.Range;
                        ReplaceRangePreservingFormatting(rng, e.text ?? string.Empty);
                        completedCells.Add($"T{e.tableIndex}({e.row},{e.col})");
                    }
                    catch (Exception writeEx)
                    {
                        string failedCell = $"T{e.tableIndex}({e.row},{e.col})";
                        string undoAction = undoStarted
                            ? "Premi un singolo Ctrl+Z in Word per annullare tutto il blocco prima di riprovare."
                            : "UndoRecord non attivo: verifica manualmente in Word le celle coinvolte.";
                        string partialInfo = completedCells.Count > 0
                            ? $"Celle già completate ({completedCells.Count}/{parsed.Count}): {string.Join(", ", completedCells)}; anche la cella fallita {failedCell} può risultare parzialmente modificata. {undoAction}"
                            : $"La cella fallita {failedCell} può risultare parzialmente modificata prima dell'errore. {undoAction}";
                        throw new InvalidOperationException(
                            $"Scrittura interrotta su {failedCell}: {writeEx.Message}. {partialInfo}",
                            writeEx);
                    }
                    finally
                    {
                        ComHelper.Release(rng);
                        ComHelper.Release(cell);
                        ComHelper.Release(tbl);
                    }
                }

                string undoNote = undoStarted ? "Tutte le modifiche sono annullabili con un singolo Ctrl+Z." : "Nota: UndoRecord non disponibile.";
                return $"Aggiornate {completedCells.Count} celle su {tableTokenCache.Count} tabelle in '{doc.Name}' preservando tabelle e impaginazione. {undoNote}";
            }
            finally
            {
                if (undoStarted && undoRec != null)
                {
                    try { undoRec.EndCustomRecord(); } catch { }
                }
                ComHelper.Release(undoRec);
                ComHelper.Release(tables);
                ComHelper.Release(doc);
                ComHelper.Release(docs);
                ComHelper.Release(app);
            }
        }

        /// <summary>
        /// Calcola le modifiche chirurgiche (da destra a sinistra) per trasformare oldText in newText
        /// toccando solo i token cambiati, preservando grassetti/corsivi/colori intra-paragrafo.
        /// </summary>
        internal static System.Collections.Generic.List<(int start, int length, string replacement)> ComputeSurgicalTextEdits(string oldText, string newText)
        {
            var edits = new System.Collections.Generic.List<(int start, int length, string replacement)>();
            string a = oldText ?? string.Empty;
            string b = newText ?? string.Empty;
            if (string.Equals(a, b, StringComparison.Ordinal)) return edits;
            if (a.Length == 0 || b.Length == 0)
            {
                edits.Add((0, a.Length, b));
                return edits;
            }

            var tokRegex = new System.Text.RegularExpressions.Regex(@"\S+|\s+");
            var m1 = tokRegex.Matches(a);
            var m2 = tokRegex.Matches(b);
            int n1 = m1.Count, n2 = m2.Count;

            if ((long)n1 * n2 > 250000)
            {
                edits.Add((0, a.Length, b));
                return edits;
            }

            int[,] dp = new int[n1 + 1, n2 + 1];
            for (int i = n1 - 1; i >= 0; i--)
                for (int j = n2 - 1; j >= 0; j--)
                    dp[i, j] = m1[i].Value == m2[j].Value
                        ? dp[i + 1, j + 1] + 1
                        : Math.Max(dp[i + 1, j], dp[i, j + 1]);

            double sim = (2.0 * dp[0, 0]) / Math.Max(1, n1 + n2);
            if (sim < 0.30)
            {
                edits.Add((0, a.Length, b));
                return edits;
            }

            int ii = 0, jj = 0;
            while (ii < n1 || jj < n2)
            {
                if (ii < n1 && jj < n2 && m1[ii].Value == m2[jj].Value)
                {
                    ii++; jj++;
                    continue;
                }

                int oldStart = ii < n1 ? m1[ii].Index : a.Length;
                int oldEnd = oldStart;
                var rep = new StringBuilder();

                while (ii < n1 && (jj == n2 || dp[ii + 1, jj] >= dp[ii, jj + 1]) && (jj == n2 || m1[ii].Value != m2[jj].Value))
                {
                    oldEnd = m1[ii].Index + m1[ii].Length;
                    ii++;
                }
                while (jj < n2 && (ii == n1 || dp[ii, jj + 1] > dp[ii + 1, jj]) && (ii == n1 || m1[ii].Value != m2[jj].Value))
                {
                    rep.Append(m2[jj].Value);
                    jj++;
                }

                edits.Add((oldStart, oldEnd - oldStart, rep.ToString()));
            }

            // Ordina da destra a sinistra così gli offset non slittano durante l'applicazione in Word
            edits.Sort((x, y) => y.start.CompareTo(x.start));
            return edits;
        }

        /// <summary>
        /// Crea un sotto-Range restando SEMPRE nella stessa Story di baseRange (MainText, Header, Footer, TextBox, Footnote, ecc.),
        /// evitando doc.Range(start, end) che punta sempre a wdMainTextStory.
        /// </summary>
        private static dynamic CreateSubRange(dynamic baseRange, int start, int end)
        {
            dynamic sub = baseRange.Duplicate;
            sub.SetRange(start, end);
            return sub;
        }

        /// <summary>
        /// Sostituisce il testo di un Range Word (selezione o cella tabella) preservando:
        /// 1) la Story originale (via Duplicate+SetRange, mai doc.Range);
        /// 2) i confini esatti di targetRange (non aggiunge/rimuove mai testo fuori dalla selezione);
        /// 3) stili, colori, dimensioni ed elenchi di ciascun paragrafo + grassetti/corsivi misti intra-paragrafo.
        /// </summary>
        private static void ReplaceRangePreservingFormatting(dynamic targetRange, string newText)
        {
            string normNew = (newText ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
            string[] newParas = normNew.Split('\n');

            dynamic paras = targetRange.Paragraphs;
            try
            {
                int oldCount = (int)paras.Count;
                if (oldCount <= 0)
                {
                    targetRange.Text = normNew.Replace("\n", "\r");
                    return;
                }

                int tStartOrig = (int)targetRange.Start;
                int tEndOrig = (int)targetRange.End;
                int common = Math.Min(oldCount, newParas.Length);

                // Se il nuovo testo ha PIU' paragrafi dell'originale, inserisce i paragrafi extra
                // restando RIGOROSAMENTE entro i confini di targetRange (mai oltre tEndOrig).
                if (newParas.Length > oldCount)
                {
                    dynamic lastPara = paras.Item(oldCount);
                    dynamic lastRng = null;
                    dynamic insRng = null;
                    try
                    {
                        lastRng = lastPara.Range;
                        string rawLast = (string)(lastRng.Text ?? string.Empty);
                        int stripLast = CountTrailingParagraphMarkers(rawLast);
                        int maxInsidePara = Math.Max((int)lastRng.Start, (int)lastRng.End - stripLast);
                        int pos = Math.Min(tEndOrig, maxInsidePara);
                        if (pos < tStartOrig) pos = tStartOrig;
                        insRng = CreateSubRange(targetRange, pos, pos);
                        var extra = new StringBuilder();
                        for (int k = oldCount; k < newParas.Length; k++)
                            extra.Append('\r').Append(newParas[k]);
                        insRng.Text = extra.ToString();
                    }
                    finally
                    {
                        ComHelper.Release(insRng);
                        ComHelper.Release(lastRng);
                        ComHelper.Release(lastPara);
                    }
                }
                else if (newParas.Length < oldCount)
                {
                    // Rimuove solo la porzione selezionata dei paragrafi in eccesso (dal basso verso l'alto),
                    // senza mai toccare testo fuori [tStart, tEnd] né il marcatore di fine cella (\r\a).
                    for (int p = oldCount; p > newParas.Length; p--)
                    {
                        dynamic pItem = paras.Item(p);
                        dynamic pRng = null;
                        dynamic delRng = null;
                        try
                        {
                            pRng = pItem.Range;
                            int curTStart = (int)targetRange.Start;
                            int curTEnd = (int)targetRange.End;
                            int prStart = (int)pRng.Start;
                            int prEnd = (int)pRng.End;
                            string rawP = (string)(pRng.Text ?? string.Empty);

                            int delStart = Math.Max(prStart, curTStart);
                            int maxEnd = rawP.EndsWith("\r\a") ? Math.Max(prStart, prEnd - 2) : prEnd;
                            int delEnd = Math.Min(maxEnd, curTEnd);

                            // Se il paragrafo è parzialmente selezionato in coda (curTEnd < prEnd) e il \r precedente
                            // cade dentro la selezione, includi quel \r per fondere il resto del paragrafo senza toccare oltre curTEnd.
                            if (delEnd < prEnd && delStart == prStart && p > 1 && prStart - 1 >= curTStart)
                                delStart = prStart - 1;

                            if (delEnd > delStart)
                            {
                                delRng = CreateSubRange(targetRange, delStart, delEnd);
                                delRng.Text = string.Empty;
                            }
                        }
                        finally
                        {
                            ComHelper.Release(delRng);
                            ComHelper.Release(pRng);
                            ComHelper.Release(pItem);
                        }
                    }
                }

                // Aggiorna i paragrafi comuni (1..common) da basso verso l'alto con diff chirurgico intra-paragrafo
                for (int p = common; p >= 1; p--)
                {
                    dynamic pItem = paras.Item(p);
                    dynamic pRng = null;
                    try
                    {
                        pRng = pItem.Range;
                        int tStart = (int)targetRange.Start;
                        int tEnd = (int)targetRange.End;
                        int pStart = Math.Max((int)pRng.Start, tStart);
                        int pEnd = Math.Min((int)pRng.End, tEnd);
                        if (pEnd < pStart) continue;

                        dynamic sliceRng = CreateSubRange(targetRange, pStart, pEnd);
                        try
                        {
                            string raw = (string)(sliceRng.Text ?? string.Empty);
                            int strip = CountTrailingParagraphMarkers(raw);
                            int innerEnd = Math.Max(pStart, pEnd - strip);
                            string oldInner = raw.Substring(0, Math.Max(0, raw.Length - strip));
                            string targetInner = newParas[p - 1];

                            if (string.Equals(oldInner, targetInner, StringComparison.Ordinal))
                                continue;

                            // Verifica coerenza offset Range vs stringa (es. assenza di field code/immagini inline)
                            if (innerEnd - pStart == oldInner.Length)
                            {
                                var edits = ComputeSurgicalTextEdits(oldInner, targetInner);
                                foreach (var ed in edits)
                                {
                                    dynamic edRng = CreateSubRange(targetRange, pStart + ed.start, pStart + ed.start + ed.length);
                                    try { edRng.Text = ed.replacement; }
                                    finally { ComHelper.Release(edRng); }
                                }
                            }
                            else
                            {
                                dynamic innerRng = CreateSubRange(targetRange, pStart, innerEnd);
                                try { innerRng.Text = targetInner; }
                                finally { ComHelper.Release(innerRng); }
                            }
                        }
                        finally { ComHelper.Release(sliceRng); }
                    }
                    finally
                    {
                        ComHelper.Release(pRng);
                        ComHelper.Release(pItem);
                    }
                }
            }
            finally
            {
                ComHelper.Release(paras);
            }
        }

        private static int CountTrailingParagraphMarkers(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return 0;
            if (raw.EndsWith("\r\a")) return 2;
            if (raw.EndsWith("\r") || raw.EndsWith("\a")) return 1;
            return 0;
        }

        private static string NormalizeCellText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            // Le celle Word terminano con \r\a
            return raw.Replace("\r\a", "").Replace("\a", "").Replace("\r", "\n").TrimEnd('\n');
        }

        private static string NormalizeWordText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            return raw.Replace("\r\a", "\t").Replace("\a", "").Replace("\r", "\n").TrimEnd('\n');
        }
    }
}
