using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace Mcp4Office.Tools
{
    internal static class ExcelTools
    {
        private const int MaxExcelRows = 1048576;
        private const int MaxExcelCols = 16384;
        private const int MaxWriteCells = 5000;

        private static readonly Regex IsoDateRegex =
            new Regex(@"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?$", RegexOptions.Compiled);

        private static readonly Dictionary<int, string> ExcelErrorCodes = new Dictionary<int, string>
        {
            { unchecked((int)0x800A07D0), "#NULL!" },
            { unchecked((int)0x800A07D7), "#DIV/0!" },
            { unchecked((int)0x800A07DF), "#VALUE!" },
            { unchecked((int)0x800A07E7), "#REF!" },
            { unchecked((int)0x800A07ED), "#NAME?" },
            { unchecked((int)0x800A07F4), "#NUM!" },
            { unchecked((int)0x800A07FA), "#N/A" }
        };

        public static void Register(McpServer server)
        {
            server.RegisterTool(
                "excel_status",
                "Mostra workbook_id (percorso), foglio attivo, fogli disponibili e indirizzo della selezione corrente in Excel.",
                McpServer.EmptySchema(),
                McpServer.ReadOnly(),
                _ => ComHelper.RetryRead(GetStatus));

            server.RegisterTool(
                "excel_read_range",
                "Legge valori (date ISO, errori #DIV/0! ecc.) e opzionalmente formule in inglese (=SUM(...)) da Excel. Supporta 'workbook' opzionale per leggere qualsiasi cartella aperta.",
                McpServer.Schema((props, _) =>
                {
                    props["workbook"] = McpServer.Prop("string", "Opzionale: workbook_id (FullName) o nome di una cartella aperta (default: cartella attiva).");
                    props["range"] = McpServer.Prop("string", "Indirizzo intervallo (es. 'A1:D20'). Obbligatorio se 'sheet' o 'workbook' è specificato (salvo use_used_range=true).");
                    props["sheet"] = McpServer.Prop("string", "Nome del foglio (opzionale; default foglio attivo).");
                    props["use_used_range"] = McpServer.Prop("boolean", "Se true e 'range' è omesso, legge UsedRange del foglio (default false).");
                    props["include_formulas"] = McpServer.Prop("boolean", "Se true mostra anche le formule standard inglesi (default false).");
                    props["max_rows"] = McpServer.IntProp("Numero massimo di righe da leggere (1-1000, default 200).", 1, 1000);
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ReadRange(
                    McpServer.GetString(args, "workbook"),
                    McpServer.GetString(args, "range"),
                    McpServer.GetString(args, "sheet"),
                    McpServer.GetBool(args, "use_used_range", false),
                    McpServer.GetBool(args, "include_formulas", false),
                    McpServer.GetInt(args, "max_rows", 200))));

            server.RegisterTool(
                "excel_compare_sheets",
                "Confronta due fogli Excel (stessa cartella o due cartelle aperte) su valori e formule in C#. Usa key_column (es. 'A' o il nome dell'intestazione codice/ID) per allineare le righe anche se ne sono state inserite o rimosse (confronto chiavi case-insensitive). Imposta has_header=false se i dati non hanno riga di intestazione.",
                McpServer.Schema((props, req) =>
                {
                    props["sheet1"] = McpServer.Prop("string", "Nome del primo foglio (originale).");
                    props["sheet2"] = McpServer.Prop("string", "Nome del secondo foglio (revisionato).");
                    props["workbook1"] = McpServer.Prop("string", "Opzionale: workbook_id o nome della prima cartella (default: cartella attiva).");
                    props["workbook2"] = McpServer.Prop("string", "Opzionale: workbook_id o nome della seconda cartella (default: cartella attiva).");
                    props["range"] = McpServer.Prop("string", "Opzionale: intervallo da confrontare (es. 'A1:H100'). Se omesso usa l'unione degli UsedRange (max 5000 righe x 50 colonne).");
                    props["key_column"] = McpServer.Prop("string", "Opzionale (consigliato per tabelle): lettera colonna (es. 'A') o nome intestazione nella prima riga (es. 'Codice' o 'ID') per allineare le righe per chiave ed evitare slittamenti. Il confronto delle chiavi ignora maiuscole/minuscole (es. 'ab1' == 'AB1') e spazi iniziali/finali.");
                    props["has_header"] = McpServer.Prop("boolean", "Usato con key_column: se true (default) la prima riga dell'intervallo è trattata come intestazione; imposta false se i dati partono subito dalla prima riga (in tal caso key_column va indicato come lettera colonna, es. 'A').");
                    props["compare_formulas"] = McpServer.Prop("boolean", "Se true (default) segnala anche differenze di formula oltre ai valori.");
                    props["max_diffs"] = McpServer.IntProp("Numero massimo di differenze da elencare (1-500, default 200).", 1, 500);
                    req.Add("sheet1");
                    req.Add("sheet2");
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => CompareSheets(
                    McpServer.GetString(args, "workbook1"),
                    McpServer.GetString(args, "sheet1"),
                    McpServer.GetString(args, "workbook2"),
                    McpServer.GetString(args, "sheet2"),
                    McpServer.GetString(args, "range"),
                    McpServer.GetString(args, "key_column"),
                    McpServer.GetBool(args, "has_header", true),
                    McpServer.GetBool(args, "compare_formulas", true),
                    McpServer.GetInt(args, "max_diffs", 200))));

            server.RegisterTool(
                "excel_write_cells",
                "Scrive una matrice rettangolare in Excel. Passare numeri come numeri JSON e date in formato ISO 'yyyy-MM-dd'. ATTENZIONE: la scrittura via COM azzera la cronologia Ctrl+Z di Excel. Se overwrite=true (o se expected_state_token è fornito), usa lo state_token del messaggio di blocco o di una lettura dello STESSO intervallo esatto.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_workbook"] = McpServer.Prop("string", "workbook_id (FullName) restituito da excel_status/excel_read_range.");
                    props["expected_sheet"] = McpServer.Prop("string", "Nome esatto del foglio di destinazione.");
                    props["start_cell"] = McpServer.Prop("string", "Singola cella iniziale (es. 'A1' o 'C5').");
                    props["rows"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Matrice rettangolare 2D di righe e valori (numeri JSON, stringhe, date ISO 'yyyy-MM-dd', booleani).",
                        ["items"] = new JsonObject { ["type"] = "array" }
                    };
                    props["allow_formulas"] = McpServer.Prop("boolean", "Se false (default) tutte le stringhe sono scritte come testo letterale sicuro. Se true, stringhe che iniziano con '=' vengono inviate come formule (e rifiutate se invalide).");
                    props["use_local_formulas"] = McpServer.Prop("boolean", "Solo se allow_formulas=true: se true usa FormulaLocal (es. '=SOMMA(...)'), default false (inglese '=SUM(...)').");
                    props["parse_iso_dates"] = McpServer.Prop("boolean", "Se true (default) converte stringhe ISO esatte 'yyyy-MM-dd' in vere date Excel ordinabili; se false le scrive come testo.");
                    props["overwrite"] = McpServer.Prop("boolean", "Se true consente sovrascrittura di celle occupate SOLO se expected_state_token coincide con lo stato attuale.");
                    props["expected_state_token"] = McpServer.Prop("string", "Obbligatorio quando overwrite=true su celle occupate; se fornito viene SEMPRE validato (anche su intervallo vuoto). Deve provenire dal messaggio di blocco o da una lettura dello STESSO intervallo.");
                    req.Add("expected_workbook");
                    req.Add("expected_sheet");
                    req.Add("start_cell");
                    req.Add("rows");
                }),
                McpServer.Mutating(destructive: true),
                args => WriteCells(args));

            server.RegisterTool(
                "excel_add_sheet",
                "Crea un nuovo foglio di lavoro in fondo alla cartella attiva di Excel (senza toccare i fogli esistenti). Richiede expected_workbook.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_workbook"] = McpServer.Prop("string", "workbook_id (FullName) da excel_status/excel_read_range.");
                    props["sheet_name"] = McpServer.Prop("string", "Nome del nuovo foglio (max 31 caratteri, senza \\ / ? * [ ] :).");
                    req.Add("expected_workbook");
                    req.Add("sheet_name");
                }),
                McpServer.Mutating(destructive: false),
                args => AddSheet(
                    McpServer.GetString(args, "expected_workbook"),
                    McpServer.GetString(args, "sheet_name")));

            server.RegisterTool(
                "excel_add_chart",
                "Crea un grafico nativo modificabile in Excel (AddChart2) collegato direttamente a un intervallo del foglio (si aggiorna automaticamente quando cambiano i numeri).",
                McpServer.Schema((props, req) =>
                {
                    props["expected_workbook"] = McpServer.Prop("string", "workbook_id (FullName) da excel_status/excel_read_range.");
                    props["source_sheet"] = McpServer.Prop("string", "Nome del foglio che contiene i dati sorgente.");
                    props["source_range"] = McpServer.Prop("string", "Intervallo dati (es. 'A1:C10', inclusa riga intestazioni e colonna categorie).");
                    props["chart_type"] = McpServer.EnumProp("Tipo di grafico: 'column', 'bar', 'line', 'pie', 'area'.", "column", "bar", "line", "pie", "area");
                    props["title"] = McpServer.Prop("string", "Titolo del grafico (opzionale).");
                    props["target_sheet"] = McpServer.Prop("string", "Foglio su cui posizionare il grafico (opzionale; default uguale a source_sheet).");
                    props["left"] = McpServer.IntProp("Posizione sinistra in punti (opzionale, default 320).", 0, 5000);
                    props["top"] = McpServer.IntProp("Posizione superiore in punti (opzionale, default 20).", 0, 5000);
                    props["width"] = McpServer.IntProp("Larghezza in punti (opzionale, default 480).", 100, 2000);
                    props["height"] = McpServer.IntProp("Altezza in punti (opzionale, default 300).", 80, 1500);
                    req.Add("expected_workbook");
                    req.Add("source_sheet");
                    req.Add("source_range");
                    req.Add("chart_type");
                }),
                McpServer.Mutating(destructive: false),
                args => AddChart(args));
        }

        /// <summary>
        /// Token canonico tipizzato: include dimensioni, tipo reale di ciascuna cella (numero vs stringa vs formula reale vs data vs errore) e lunghezze escapate.
        /// </summary>
        internal static string ComputeRangeToken(
            string wbFullName,
            string sheetName,
            string address,
            object[,] values,
            object[,] formulas,
            bool[,] hasFormulaMask)
        {
            int rows = values != null ? values.GetLength(0) : 0;
            int cols = values != null ? values.GetLength(1) : 0;
            var sb = new StringBuilder();
            sb.Append("wb:").Append(wbFullName?.Length ?? 0).Append(':').Append(wbFullName);
            sb.Append("|sh:").Append(sheetName?.Length ?? 0).Append(':').Append(sheetName);
            sb.Append("|ad:").Append(address?.Length ?? 0).Append(':').Append(address);
            sb.Append("|dim:").Append(rows).Append('x').Append(cols).Append('|');

            if (rows > 0 && cols > 0)
            {
                int vr0 = values.GetLowerBound(0), vc0 = values.GetLowerBound(1);
                int fr0 = formulas != null ? formulas.GetLowerBound(0) : 0;
                int fc0 = formulas != null ? formulas.GetLowerBound(1) : 0;

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        bool isFormula = hasFormulaMask != null && hasFormulaMask[r, c];
                        if (isFormula)
                        {
                            string f = formulas != null ? Convert.ToString(formulas[fr0 + r, fc0 + c], CultureInfo.InvariantCulture) ?? "" : "";
                            sb.Append("F:").Append(f.Length).Append(':').Append(f).Append(';');
                        }
                        else
                        {
                            object v = values[vr0 + r, vc0 + c];
                            sb.Append(EncodeTypedValue(v)).Append(';');
                        }
                    }
                    sb.Append('\n');
                }
            }
            return ComHelper.Sha256Short(sb.ToString());
        }

        private static string EncodeTypedValue(object val)
        {
            if (val == null || val == DBNull.Value) return "NULL";
            if (val is string s) return $"S:{s.Length}:{s}";
            if (val is bool b) return b ? "B:1" : "B:0";
            if (val is DateTime dt) return $"D:{dt.ToString("O", CultureInfo.InvariantCulture)}";
            if (val is int errCode && ExcelErrorCodes.TryGetValue(errCode, out var errName)) return $"E:{errName}";
            if (val is double d) return $"N:{d.ToString("R", CultureInfo.InvariantCulture)}";
            if (val is decimal m) return $"N:{m.ToString(CultureInfo.InvariantCulture)}";
            if (val is float fl) return $"N:{fl.ToString("R", CultureInfo.InvariantCulture)}";
            if (val is long l) return $"N:{l.ToString(CultureInfo.InvariantCulture)}";
            if (val is int iv) return $"N:{iv.ToString(CultureInfo.InvariantCulture)}";
            string fallback = Convert.ToString(val, CultureInfo.InvariantCulture) ?? string.Empty;
            return $"O:{val.GetType().Name}:{fallback.Length}:{fallback}";
        }

        private static bool[,] ReadHasFormulaMask(dynamic range, object[,] formulas, int rows, int cols)
        {
            var mask = new bool[rows, cols];
            object hfAll = null;
            try { hfAll = range.HasFormula; } catch { }

            if (hfAll is bool allBool)
            {
                if (allBool)
                {
                    for (int r = 0; r < rows; r++)
                        for (int c = 0; c < cols; c++)
                            mask[r, c] = true;
                }
                return mask;
            }

            // Miste (DBNull.Value): controlla HasFormula solo sulle celle la cui Formula inizia con '='.
            int fr0 = formulas.GetLowerBound(0), fc0 = formulas.GetLowerBound(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (formulas[fr0 + r, fc0 + c] is string fs && fs.StartsWith("="))
                    {
                        dynamic cell = range.Cells.Item(r + 1, c + 1);
                        try
                        {
                            object chf = cell.HasFormula;
                            mask[r, c] = chf is bool cb && cb;
                        }
                        finally { ComHelper.Release(cell); }
                    }
                }
            }
            return mask;
        }

        private static string GetStatus()
        {
            dynamic app = ComHelper.GetActiveApp("Excel.Application", "Microsoft Excel");
            dynamic wbs = null;
            dynamic wb = null;
            dynamic ws = null;
            dynamic sheets = null;
            dynamic sel = null;
            try
            {
                wbs = app.Workbooks;
                if ((int)wbs.Count == 0)
                    return "Microsoft Excel è aperto ma non ci sono cartelle di lavoro aperte.";

                wb = app.ActiveWorkbook;
                ws = app.ActiveSheet;
                var sb = new StringBuilder();
                sb.AppendLine($"workbook_id: {wb.FullName}");
                sb.AppendLine($"Cartella attiva: {wb.Name}");
                sb.AppendLine($"Foglio attivo: {ws.Name}");
                try
                {
                    sel = app.Selection;
                    sb.AppendLine($"Selezione corrente: {sel.Address(false, false)}");
                }
                catch { }

                sheets = wb.Worksheets;
                int sCount = (int)sheets.Count;
                var names = new List<string>(sCount);
                for (int i = 1; i <= sCount; i++)
                {
                    dynamic s = sheets.Item(i);
                    try { names.Add((string)s.Name); }
                    finally { ComHelper.Release(s); }
                }
                sb.AppendLine("Fogli: " + string.Join(", ", names));
                return sb.ToString().TrimEnd();
            }
            finally
            {
                ComHelper.Release(sel);
                ComHelper.Release(sheets);
                ComHelper.Release(ws);
                ComHelper.Release(wb);
                ComHelper.Release(wbs);
                ComHelper.Release(app);
            }
        }

        private static dynamic ResolveOpenWorkbook(dynamic wbs, string wbIdOrName)
        {
            int count = (int)wbs.Count;
            if (count == 0)
                throw new InvalidOperationException("Nessuna cartella di lavoro aperta in Excel.");

            string query = (wbIdOrName ?? string.Empty).Trim();
            for (int i = 1; i <= count; i++)
            {
                dynamic wb = wbs.Item(i);
                string fn = (string)wb.FullName;
                string nm = (string)wb.Name;
                if (string.Equals(fn, query, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(nm, query, StringComparison.OrdinalIgnoreCase))
                    return wb;
                ComHelper.Release(wb);
            }
            throw new InvalidOperationException($"Cartella '{wbIdOrName}' non trovata tra le cartelle aperte in Excel.");
        }

        private static string ReadRange(string workbookQuery, string rangeAddr, string sheetName, bool useUsedRange, bool includeFormulas, int maxRows)
        {
            dynamic app = ComHelper.GetActiveApp("Excel.Application", "Microsoft Excel");
            dynamic wbs = null;
            dynamic wb = null;
            dynamic sheets = null;
            dynamic ws = null;
            dynamic rng = null;
            dynamic areas = null;
            dynamic bounded = null;
            dynamic parentWs = null;
            try
            {
                wbs = app.Workbooks;
                if ((int)wbs.Count == 0)
                    throw new InvalidOperationException("Nessuna cartella di lavoro aperta in Excel.");

                bool wbSpecified = !string.IsNullOrWhiteSpace(workbookQuery);
                wb = wbSpecified ? ResolveOpenWorkbook(wbs, workbookQuery) : app.ActiveWorkbook;
                string wbFull = (string)wb.FullName;

                bool sheetSpecified = !string.IsNullOrWhiteSpace(sheetName);
                bool rangeSpecified = !string.IsNullOrWhiteSpace(rangeAddr);

                if ((sheetSpecified || wbSpecified) && !rangeSpecified && !useUsedRange)
                {
                    throw new ArgumentException(
                        "Quando specifichi 'sheet' o 'workbook', devi indicare anche 'range' (es. 'A1:D20') oppure impostare use_used_range=true.");
                }

                if (sheetSpecified)
                {
                    sheets = wb.Worksheets;
                    ws = sheets.Item(sheetName);
                }
                else
                {
                    ws = wb.ActiveSheet;
                }

                if (rangeSpecified)
                    rng = ws.Range(rangeAddr);
                else if (useUsedRange)
                    rng = ws.UsedRange;
                else
                    rng = app.Selection;

                // Rifiuta selezioni multiple non contigue (Ctrl+clic).
                areas = rng.Areas;
                if ((int)areas.Count > 1)
                    throw new InvalidOperationException("Selezione multipla non contigua rilevata: seleziona un singolo intervallo rettangolare o specifica 'range'.");

                parentWs = rng.Worksheet;
                string actualSheet = (string)parentWs.Name;

                int totalRows = (int)rng.Rows.Count;
                int totalCols = (int)rng.Columns.Count;
                int rowsToRead = Math.Min(totalRows, maxRows);
                int colsToRead = Math.Min(totalCols, 50);

                bounded = (rowsToRead < totalRows || colsToRead < totalCols)
                    ? rng.Resize(rowsToRead, colsToRead)
                    : rng;

                string addr = (string)bounded.Address(false, false);

                object[,] values = To2DArray(bounded.Value, rowsToRead, colsToRead);
                object[,] formulas = To2DArray(bounded.Formula, rowsToRead, colsToRead);
                bool[,] hasFormulaMask = ReadHasFormulaMask(bounded, formulas, rowsToRead, colsToRead);
                string token = ComputeRangeToken(wbFull, actualSheet, addr, values, formulas, hasFormulaMask);

                var sb = new StringBuilder();
                sb.AppendLine($"workbook_id: {wbFull}");
                sb.AppendLine($"sheet: {actualSheet}");
                sb.AppendLine($"range: {addr} (righe lette: {rowsToRead}/{totalRows}, colonne: {colsToRead}/{totalCols})");
                sb.AppendLine($"state_token: {token}");
                sb.AppendLine();

                int vr0 = values.GetLowerBound(0), vc0 = values.GetLowerBound(1);
                int fr0 = formulas.GetLowerBound(0), fc0 = formulas.GetLowerBound(1);
                for (int r = 0; r < rowsToRead; r++)
                {
                    var rowCells = new List<string>(colsToRead);
                    for (int c = 0; c < colsToRead; c++)
                    {
                        string text = FormatCellValue(values[vr0 + r, vc0 + c]);
                        if (includeFormulas && hasFormulaMask[r, c])
                        {
                            string f = formulas[fr0 + r, fc0 + c] as string;
                            if (!string.IsNullOrEmpty(f))
                                text = $"{text} [{f}]";
                        }
                        rowCells.Add(text);
                    }
                    sb.AppendLine(string.Join("\t", rowCells));
                }

                return sb.ToString().TrimEnd();
            }
            finally
            {
                if (!ReferenceEquals(bounded, rng)) ComHelper.Release(bounded);
                ComHelper.Release(parentWs);
                ComHelper.Release(areas);
                ComHelper.Release(rng);
                ComHelper.Release(ws);
                ComHelper.Release(sheets);
                ComHelper.Release(wb);
                ComHelper.Release(wbs);
                ComHelper.Release(app);
            }
        }

        private static string WriteCells(JsonObject args)
        {
            string expectedWb = McpServer.GetString(args, "expected_workbook").Trim();
            string expectedSheet = McpServer.GetString(args, "expected_sheet").Trim();
            string startCell = McpServer.GetString(args, "start_cell").Trim();
            bool allowFormulas = McpServer.GetBool(args, "allow_formulas", false);
            bool useLocalFormulas = McpServer.GetBool(args, "use_local_formulas", false);
            bool parseIsoDates = McpServer.GetBool(args, "parse_iso_dates", true);
            bool overwrite = McpServer.GetBool(args, "overwrite", false);
            string expectedToken = McpServer.GetString(args, "expected_state_token").Trim();

            if (string.IsNullOrEmpty(startCell))
                throw new ArgumentException("Parametro 'start_cell' obbligatorio.");

            var rowsArr = args["rows"] as JsonArray;
            if (rowsArr == null || rowsArr.Count == 0)
                throw new ArgumentException("Parametro 'rows' deve essere una matrice 2D non vuota.");

            var matrix = new List<List<object>>(rowsArr.Count);
            int expectedCols = -1;
            foreach (var rNode in rowsArr)
            {
                var cArr = rNode as JsonArray;
                if (cArr == null || cArr.Count == 0)
                    throw new ArgumentException("Ogni riga in 'rows' deve essere un array non vuoto.");
                if (expectedCols < 0) expectedCols = cArr.Count;
                else if (cArr.Count != expectedCols)
                    throw new ArgumentException($"Matrice non rettangolare: attesa larghezza {expectedCols} colonne, trovata riga da {cArr.Count}.");

                var row = new List<object>(cArr.Count);
                foreach (var cNode in cArr)
                    row.Add(JsonToExcelValue(cNode));
                matrix.Add(row);
            }

            int rowCount = matrix.Count;
            int colCount = expectedCols;
            if (rowCount * colCount > MaxWriteCells)
                throw new ArgumentException($"Troppe celle in una singola scrittura ({rowCount * colCount} > limite {MaxWriteCells}).");

            dynamic app = ComHelper.GetActiveApp("Excel.Application", "Microsoft Excel");
            dynamic wbs = null;
            dynamic wb = null;
            dynamic sheets = null;
            dynamic ws = null;
            dynamic topLeft = null;
            dynamic targetRng = null;
            try
            {
                wbs = app.Workbooks;
                if ((int)wbs.Count == 0)
                    throw new InvalidOperationException("Nessuna cartella di lavoro aperta in Excel.");

                wb = app.ActiveWorkbook;
                string actualWb = (string)wb.FullName;
                if (!string.Equals(actualWb, expectedWb, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Cartella cambiata: attesa '{expectedWb}', ma attiva '{actualWb}'.");

                sheets = wb.Worksheets;
                ws = sheets.Item(expectedSheet);
                string actualSheet = (string)ws.Name;

                topLeft = ws.Range(startCell);
                if ((int)topLeft.Cells.Count != 1)
                    throw new ArgumentException($"'start_cell' deve essere una singola cella ('{startCell}').");

                int startRow = (int)topLeft.Row;
                int startCol = (int)topLeft.Column;
                if (startRow + rowCount - 1 > MaxExcelRows || startCol + colCount - 1 > MaxExcelCols)
                    throw new ArgumentOutOfRangeException(nameof(startCell), "L'intervallo supera i limiti della griglia Excel.");

                targetRng = topLeft.Resize(rowCount, colCount);
                string targetAddr = (string)targetRng.Address(false, false);

                object mergeObj = targetRng.MergeCells;
                if (mergeObj == DBNull.Value || (mergeObj is bool mb && mb))
                    throw new InvalidOperationException($"L'intervallo '{targetAddr}' contiene celle unite: scrittura rifiutata.");

                bool sheetProtected = (bool)ws.ProtectContents;
                if (sheetProtected)
                {
                    object lockedObj = targetRng.Locked;
                    if (lockedObj == DBNull.Value || (lockedObj is bool lb && lb))
                        throw new InvalidOperationException($"Il foglio '{actualSheet}' è protetto e '{targetAddr}' contiene celle bloccate.");
                }

                // Snapshot tipizzato completo prima della scrittura
                object[,] prevValues = To2DArray(targetRng.Value, rowCount, colCount);
                object[,] prevFormulas = To2DArray(targetRng.Formula, rowCount, colCount);
                bool[,] prevHasFormula = ReadHasFormulaMask(targetRng, prevFormulas, rowCount, colCount);
                string currentToken = ComputeRangeToken(actualWb, actualSheet, targetAddr, prevValues, prevFormulas, prevHasFormula);

                // Se expected_state_token è fornito, validarlo SEMPRE (anche se il range è diventato vuoto!)
                if (!string.IsNullOrEmpty(expectedToken) && !string.Equals(expectedToken, currentToken, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Stato delle celle in '{actualSheet}!{targetAddr}' cambiato rispetto al token fornito (atteso '{expectedToken}', attuale '{currentToken}'). Rileggi lo stesso intervallo prima di scrivere.");
                }

                // Controllo celle occupate (valori o formule, inclusa ="")
                var occupied = new List<string>();
                int vr0 = prevValues.GetLowerBound(0), vc0 = prevValues.GetLowerBound(1);
                int fr0 = prevFormulas.GetLowerBound(0), fc0 = prevFormulas.GetLowerBound(1);
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        bool hasF = prevHasFormula[r, c];
                        object pv = prevValues[vr0 + r, vc0 + c];
                        bool hasV = pv != null && pv != DBNull.Value && !(pv is string s && s.Length == 0);
                        if (hasF || hasV)
                            occupied.Add($"{ColIndexToLetter(startCol + c)}{startRow + r}");
                    }
                }

                if (occupied.Count > 0)
                {
                    if (!overwrite)
                    {
                        string sample = string.Join(", ", occupied.GetRange(0, Math.Min(occupied.Count, 20)));
                        return $"SCRITTURA BLOCCATA per sicurezza: su '{actualSheet}' ({targetAddr}) risultano occupate {occupied.Count} celle ({sample}).\n" +
                               $"state_token: {currentToken}\n" +
                               "Chiedi conferma all'utente e, solo se autorizzato, richiama excel_write_cells con overwrite=true e questo expected_state_token.";
                    }
                    if (string.IsNullOrEmpty(expectedToken))
                        throw new ArgumentException("Quando overwrite=true su celle occupate, 'expected_state_token' è obbligatorio.");
                }

                // Preparazione payload
                object[,] payload = new object[rowCount, colCount];
                bool anyFormula = false;
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        object val = matrix[r][c];
                        if (val is string s)
                        {
                            if (allowFormulas && s.StartsWith("="))
                            {
                                if (!IsPlausibleFormulaStart(s))
                                    throw new ArgumentException($"Formula non valida alla riga {r + 1}, colonna {c + 1}: '{s}'. Se intendevi testo letterale, imposta allow_formulas=false.");
                                payload[r, c] = s;
                                anyFormula = true;
                            }
                            else if (parseIsoDates && TryParseIsoDate(s, out DateTime dt))
                            {
                                payload[r, c] = dt;
                            }
                            else
                            {
                                payload[r, c] = EscapeLiteralTextForExcel(s);
                            }
                        }
                        else
                        {
                            payload[r, c] = val;
                        }
                    }
                }

                // Rileva eventuali celle di testo occupate con formattazione mista intra-cella (es. parte in grassetto/colore diverso)
                // per aggiornarle chirurgicamente via cell.Characters(start, len) preservando gli stili misti!
                var mixedRichTextCells = new Dictionary<(int r, int c), (string oldText, string newText)>();
                if (occupied.Count > 0)
                {
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            if (prevHasFormula[r, c]) continue;
                            if (!(prevValues[vr0 + r, vc0 + c] is string oldStr) || oldStr.Length == 0) continue;
                            if (!(matrix[r][c] is string newStr) || (allowFormulas && newStr.StartsWith("="))) continue;
                            if (parseIsoDates && TryParseIsoDate(newStr, out _)) continue;
                            if (string.Equals(oldStr, newStr, StringComparison.Ordinal))
                            {
                                mixedRichTextCells[(r, c)] = (oldStr, newStr);
                                continue;
                            }

                            dynamic cell = targetRng.Cells.Item(r + 1, c + 1);
                            try
                            {
                                if (HasMixedCellFormatting(cell))
                                    mixedRichTextCells[(r, c)] = (oldStr, newStr);
                            }
                            finally { ComHelper.Release(cell); }
                        }
                    }
                }

                try
                {
                    if (mixedRichTextCells.Count == 0)
                    {
                        if (anyFormula && useLocalFormulas)
                            targetRng.FormulaLocal = payload;
                        else if (anyFormula)
                            targetRng.Formula = payload;
                        else
                            targetRng.Value = payload;
                    }
                    else
                    {
                        // Scrive le celle normali normalmente e applica edit chirurgici Characters() sulle celle rich-text miste
                        for (int r = 0; r < rowCount; r++)
                        {
                            for (int c = 0; c < colCount; c++)
                            {
                                dynamic cell = targetRng.Cells.Item(r + 1, c + 1);
                                try
                                {
                                    if (mixedRichTextCells.TryGetValue((r, c), out var rt))
                                    {
                                        if (!string.Equals(rt.oldText, rt.newText, StringComparison.Ordinal))
                                        {
                                            var edits = WordTools.ComputeSurgicalTextEdits(rt.oldText, rt.newText);
                                            foreach (var ed in edits)
                                            {
                                                // Excel Characters(Start, Length) è 1-based
                                                dynamic ch = cell.Characters(ed.start + 1, ed.length);
                                                try { ch.Text = ed.replacement; }
                                                finally { ComHelper.Release(ch); }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        object pv = payload[r, c];
                                        if (anyFormula && pv is string fs && fs.StartsWith("="))
                                        {
                                            if (useLocalFormulas) cell.FormulaLocal = fs;
                                            else cell.Formula = fs;
                                        }
                                        else
                                        {
                                            cell.Value = pv;
                                        }
                                    }
                                }
                                finally { ComHelper.Release(cell); }
                            }
                        }
                    }
                }
                catch (Exception writeEx)
                {
                    string rollbackStatus = AttemptVerifiedRollback(
                        targetRng, actualWb, actualSheet, targetAddr,
                        prevValues, prevFormulas, prevHasFormula, currentToken, rowCount, colCount);
                    throw new InvalidOperationException($"Scrittura su '{actualSheet}!{targetAddr}' fallita: {writeEx.Message}. {rollbackStatus}");
                }

                var resp = new StringBuilder();
                resp.AppendLine($"Scritte {rowCount} righe x {colCount} colonne su '{actualSheet}!{targetAddr}'.");
                resp.AppendLine("Nota: la scrittura via COM azzera la cronologia Ctrl+Z di Excel.");
                if (occupied.Count > 0)
                {
                    resp.AppendLine($"Valori/formule precedenti sovrascritti ({occupied.Count} celle):");
                    int listed = 0;
                    for (int r = 0; r < rowCount && listed < 25; r++)
                    {
                        for (int c = 0; c < colCount && listed < 25; c++)
                        {
                            if (prevHasFormula[r, c])
                            {
                                resp.AppendLine($"  {ColIndexToLetter(startCol + c)}{startRow + r}: [formula] {prevFormulas[fr0 + r, fc0 + c]}");
                                listed++;
                            }
                            else
                            {
                                string pvStr = FormatCellValue(prevValues[vr0 + r, vc0 + c]);
                                if (!string.IsNullOrEmpty(pvStr))
                                {
                                    resp.AppendLine($"  {ColIndexToLetter(startCol + c)}{startRow + r}: {pvStr}");
                                    listed++;
                                }
                            }
                        }
                    }
                }
                return resp.ToString().TrimEnd();
            }
            finally
            {
                ComHelper.Release(targetRng);
                ComHelper.Release(topLeft);
                ComHelper.Release(ws);
                ComHelper.Release(sheets);
                ComHelper.Release(wb);
                ComHelper.Release(wbs);
                ComHelper.Release(app);
            }
        }

        private static string AttemptVerifiedRollback(
            dynamic targetRng, string wbFull, string sheet, string addr,
            object[,] prevValues, object[,] prevFormulas, bool[,] prevHasFormula,
            string expectedPreToken, int rows, int cols)
        {
            try
            {
                int vr0 = prevValues.GetLowerBound(0), vc0 = prevValues.GetLowerBound(1);
                int fr0 = prevFormulas.GetLowerBound(0), fc0 = prevFormulas.GetLowerBound(1);
                var restore = new object[rows, cols];
                bool anyF = false;
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (prevHasFormula[r, c])
                        {
                            restore[r, c] = prevFormulas[fr0 + r, fc0 + c];
                            anyF = true;
                        }
                        else
                        {
                            object pv = prevValues[vr0 + r, vc0 + c];
                            restore[r, c] = pv is string ps ? EscapeLiteralTextForExcel(ps) : pv;
                        }
                    }
                }
                if (anyF) targetRng.Formula = restore;
                else targetRng.Value = restore;

                // Rilettura di verifica
                object[,] afterVal = To2DArray(targetRng.Value, rows, cols);
                object[,] afterFrm = To2DArray(targetRng.Formula, rows, cols);
                bool[,] afterMask = ReadHasFormulaMask(targetRng, afterFrm, rows, cols);
                string afterToken = ComputeRangeToken(wbFull, sheet, addr, afterVal, afterFrm, afterMask);

                return string.Equals(afterToken, expectedPreToken, StringComparison.Ordinal)
                    ? "Ripristino dello snapshot precedente verificato tramite rilettura."
                    : $"ATTENZIONE: tentativo di ripristino eseguito ma il token post-rilettura differisce ('{afterToken}' != '{expectedPreToken}'); verificare {addr} in Excel.";
            }
            catch (Exception rbEx)
            {
                return $"ATTENZIONE: ripristino fallito ({rbEx.Message}); possibili modifiche parziali in {addr}.";
            }
        }

        internal static bool IsPlausibleFormulaStart(string s)
        {
            if (string.IsNullOrEmpty(s) || s.Length < 2 || s[0] != '=') return false;
            char c1 = s[1];
            // Ammette riferimenti assoluti ($A$1), funzioni/nomi, parentesi, segni, stringhe, underscore
            return char.IsLetterOrDigit(c1) || c1 == '$' || c1 == '(' || c1 == '-' || c1 == '+' || c1 == '"' || c1 == '\'' || c1 == '_';
        }

        internal static bool TryParseIsoDate(string s, out DateTime dt)
        {
            dt = default;
            if (string.IsNullOrEmpty(s) || !IsoDateRegex.IsMatch(s)) return false;
            return DateTime.TryParseExact(
                s,
                new[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss" },
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out dt);
        }

        internal static string EscapeLiteralTextForExcel(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            char first = s[0];
            if (first == '=' || first == '+' || first == '-' || first == '@' || first == '\'')
                return "'" + s;
            if (first == '0' && s.Length > 1 && char.IsDigit(s[1]))
                return "'" + s;
            if (s.IndexOf('/') >= 0 || s.IndexOf('-') >= 0)
            {
                if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out _) ||
                    DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))
                    return "'" + s;
            }
            return s;
        }

        private static object[,] To2DArray(object raw, int rows, int cols)
        {
            if (raw is object[,] arr) return arr;
            var single = new object[1, 1];
            single[0, 0] = raw;
            return single;
        }

        private static object JsonToExcelValue(JsonNode node)
        {
            if (node == null) return string.Empty;
            if (node is JsonValue v)
            {
                if (v.TryGetValue<double>(out var d)) return d;
                if (v.TryGetValue<bool>(out var b)) return b;
                if (v.TryGetValue<string>(out var s)) return s;
            }
            throw new ArgumentException("Ogni cella in 'rows' deve essere stringa, numero, booleano o null.");
        }

        internal static string FormatCellValue(object val)
        {
            if (val == null || val == DBNull.Value) return string.Empty;
            if (val is DateTime dt)
            {
                return dt.TimeOfDay == TimeSpan.Zero
                    ? dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (val is int errCode && ExcelErrorCodes.TryGetValue(errCode, out var errName))
                return errName;
            if (val is double d) return d.ToString(CultureInfo.InvariantCulture);
            return Convert.ToString(val, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static string AddSheet(string expectedWb, string sheetName)
        {
            string name = (sheetName ?? string.Empty).Trim();
            if (name.Length == 0 || name.Length > 31 || name.IndexOfAny(new[] { '\\', '/', '?', '*', '[', ']', ':' }) >= 0)
                throw new ArgumentException($"Nome foglio non valido: '{name}' (1-31 caratteri, senza \\ / ? * [ ] :).");

            dynamic app = ComHelper.GetActiveApp("Excel.Application", "Microsoft Excel");
            dynamic wbs = null;
            dynamic wb = null;
            dynamic allSheets = null;
            dynamic prevActive = null;
            dynamic lastSheet = null;
            dynamic newSheet = null;
            bool renamed = false;
            try
            {
                wbs = app.Workbooks;
                if ((int)wbs.Count == 0)
                    throw new InvalidOperationException("Nessuna cartella di lavoro aperta in Excel.");

                wb = app.ActiveWorkbook;
                string actualWb = (string)wb.FullName;
                if (!string.Equals(actualWb, expectedWb, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Cartella cambiata: attesa '{expectedWb}', attiva '{actualWb}'.");

                prevActive = wb.ActiveSheet;
                // Controlla TUTTI gli Sheets (inclusi eventuali fogli grafico!)
                allSheets = wb.Sheets;
                int count = (int)allSheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic s = allSheets.Item(i);
                    try
                    {
                        if (string.Equals((string)s.Name, name, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException($"Esiste già un foglio chiamato '{name}'.");
                    }
                    finally { ComHelper.Release(s); }
                }

                lastSheet = allSheets.Item(count);
                newSheet = wb.Worksheets.Add(After: lastSheet);
                newSheet.Name = name;
                renamed = true;

                // Riattiva il foglio precedente così la selezione dell'utente non salta altrove
                try { prevActive.Activate(); } catch { }

                return $"Creato nuovo foglio '{name}' in fondo a '{wb.Name}' (foglio attivo preservato: '{prevActive.Name}').";
            }
            catch
            {
                if (!renamed && newSheet != null)
                {
                    bool prevAlerts = true;
                    try
                    {
                        prevAlerts = (bool)app.DisplayAlerts;
                        app.DisplayAlerts = false;
                        newSheet.Delete();
                    }
                    catch { }
                    finally { try { app.DisplayAlerts = prevAlerts; } catch { } }
                }
                throw;
            }
            finally
            {
                ComHelper.Release(newSheet);
                ComHelper.Release(lastSheet);
                ComHelper.Release(allSheets);
                ComHelper.Release(prevActive);
                ComHelper.Release(wb);
                ComHelper.Release(wbs);
                ComHelper.Release(app);
            }
        }

        private static string AddChart(JsonObject args)
        {
            string expectedWb = McpServer.GetString(args, "expected_workbook").Trim();
            string sourceSheetName = McpServer.GetString(args, "source_sheet").Trim();
            string sourceRangeAddr = McpServer.GetString(args, "source_range").Trim();
            string targetSheetName = McpServer.GetString(args, "target_sheet").Trim();
            if (string.IsNullOrEmpty(targetSheetName)) targetSheetName = sourceSheetName;
            string chartType = McpServer.GetString(args, "chart_type", "column");
            string title = McpServer.GetString(args, "title").Trim();
            bool hasLeft = args.ContainsKey("left");
            bool hasTop = args.ContainsKey("top");
            float width = McpServer.GetInt(args, "width", 480);
            float height = McpServer.GetInt(args, "height", 300);

            int xlChartType = chartType switch
            {
                "bar" => 57,
                "line" => 4,
                "pie" => 5,
                "area" => 1,
                _ => 51
            };

            dynamic app = ComHelper.GetActiveApp("Excel.Application", "Microsoft Excel");
            dynamic wbs = null;
            dynamic wb = null;
            dynamic sheets = null;
            dynamic srcWs = null;
            dynamic dstWs = null;
            dynamic srcRng = null;
            dynamic shapes = null;
            dynamic chartShape = null;
            dynamic chart = null;
            bool chartCompleted = false;
            try
            {
                wbs = app.Workbooks;
                if ((int)wbs.Count == 0)
                    throw new InvalidOperationException("Nessuna cartella di lavoro aperta in Excel.");

                wb = app.ActiveWorkbook;
                string actualWb = (string)wb.FullName;
                if (!string.Equals(actualWb, expectedWb, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Cartella cambiata: attesa '{expectedWb}', attiva '{actualWb}'.");

                sheets = wb.Worksheets;
                srcWs = sheets.Item(sourceSheetName);
                bool sameSheet = string.Equals(targetSheetName, sourceSheetName, StringComparison.OrdinalIgnoreCase);
                dstWs = sameSheet ? srcWs : sheets.Item(targetSheetName);

                srcRng = srcWs.Range(sourceRangeAddr);
                if ((int)srcRng.Rows.Count < 2 || (int)srcRng.Columns.Count < 2)
                    throw new ArgumentException("L'intervallo sorgente del grafico deve avere almeno 2 righe e 2 colonne.");

                // Posizionamento intelligente: se sullo stesso foglio dei dati, mette il grafico a DESTRA dell'intervallo per non coprirlo!
                float left = hasLeft
                    ? McpServer.GetInt(args, "left", 20)
                    : (sameSheet ? (float)srcRng.Left + (float)srcRng.Width + 20f : 20f);
                float top = hasTop
                    ? McpServer.GetInt(args, "top", 20)
                    : (sameSheet ? (float)srcRng.Top : 20f);

                shapes = dstWs.Shapes;
                chartShape = shapes.AddChart2(-1, xlChartType, left, top, width, height);
                chart = chartShape.Chart;
                chart.SetSourceData(srcRng);
                if (!string.IsNullOrEmpty(title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                }
                chartCompleted = true;

                return $"Creato grafico '{chartType}' su '{dstWs.Name}' (left={left:F0}, top={top:F0}) collegato a '{srcWs.Name}!{srcRng.Address(false, false)}'.";
            }
            catch
            {
                if (!chartCompleted && chartShape != null)
                {
                    try { chartShape.Delete(); } catch { }
                }
                throw;
            }
            finally
            {
                ComHelper.Release(chart);
                ComHelper.Release(chartShape);
                ComHelper.Release(shapes);
                ComHelper.Release(srcRng);
                if (!ReferenceEquals(dstWs, srcWs)) ComHelper.Release(dstWs);
                ComHelper.Release(srcWs);
                ComHelper.Release(sheets);
                ComHelper.Release(wb);
                ComHelper.Release(wbs);
                ComHelper.Release(app);
            }
        }

        private static string CompareSheets(
            string wb1Query, string sheet1Name,
            string wb2Query, string sheet2Name,
            string rangeAddr, string keyColumn, bool hasHeader, bool compareFormulas, int maxDiffs)
        {
            dynamic app = ComHelper.GetActiveApp("Excel.Application", "Microsoft Excel");
            dynamic wbs = null;
            dynamic wb1 = null, wb2 = null;
            dynamic ws1 = null, ws2 = null;
            dynamic r1 = null, r2 = null;
            try
            {
                wbs = app.Workbooks;
                if ((int)wbs.Count == 0)
                    throw new InvalidOperationException("Nessuna cartella di lavoro aperta in Excel.");

                wb1 = string.IsNullOrWhiteSpace(wb1Query) ? app.ActiveWorkbook : ResolveOpenWorkbook(wbs, wb1Query);
                wb2 = string.IsNullOrWhiteSpace(wb2Query) ? app.ActiveWorkbook : ResolveOpenWorkbook(wbs, wb2Query);

                ws1 = wb1.Worksheets.Item(sheet1Name);
                ws2 = wb2.Worksheets.Item(sheet2Name);

                int startRow = 1, startCol = 1, rows, cols;
                int usedRows1 = 0, usedRows2 = 0;
                bool usedRangeCapped = false;
                if (!string.IsNullOrWhiteSpace(rangeAddr))
                {
                    r1 = ws1.Range(rangeAddr);
                    r2 = ws2.Range(rangeAddr);
                    startRow = (int)r1.Row;
                    startCol = (int)r1.Column;
                    int rawR = (int)r1.Rows.Count, rawC = (int)r1.Columns.Count;
                    rows = Math.Min(rawR, 5000);
                    cols = Math.Min(rawC, 50);
                    usedRangeCapped = rawR > 5000 || rawC > 50;
                    usedRows1 = rows; usedRows2 = rows;
                }
                else
                {
                    dynamic u1 = ws1.UsedRange;
                    dynamic u2 = ws2.UsedRange;
                    try
                    {
                        usedRows1 = (int)u1.Row + (int)u1.Rows.Count - 1;
                        usedRows2 = (int)u2.Row + (int)u2.Rows.Count - 1;
                        int endR = Math.Max(usedRows1, usedRows2);
                        int endC = Math.Max((int)u1.Column + (int)u1.Columns.Count - 1, (int)u2.Column + (int)u2.Columns.Count - 1);
                        usedRangeCapped = endR > 5000 || endC > 50;
                        rows = Math.Min(Math.Max(endR, 1), 5000);
                        cols = Math.Min(Math.Max(endC, 1), 50);
                    }
                    finally
                    {
                        ComHelper.Release(u2);
                        ComHelper.Release(u1);
                    }
                    string unionAddr = $"A1:{ColIndexToLetter(cols)}{rows}";
                    r1 = ws1.Range(unionAddr);
                    r2 = ws2.Range(unionAddr);
                }

                object[,] v1 = To2DArray(r1.Value, rows, cols);
                object[,] f1 = To2DArray(r1.Formula, rows, cols);
                bool[,] m1 = ReadHasFormulaMask(r1, f1, rows, cols);

                object[,] v2 = To2DArray(r2.Value, rows, cols);
                object[,] f2 = To2DArray(r2.Formula, rows, cols);
                bool[,] m2 = ReadHasFormulaMask(r2, f2, rows, cols);

                string label1 = $"{wb1.Name}!{ws1.Name}";
                string label2 = $"{wb2.Name}!{ws2.Name}";
                string report = !string.IsNullOrWhiteSpace(keyColumn)
                    ? ComputeKeyedSheetDiff(label1, label2, startRow, startCol, rows, cols, v1, f1, m1, v2, f2, m2, keyColumn.Trim(), hasHeader, compareFormulas, maxDiffs)
                    : ComputeSheetDiff(label1, label2, startRow, startCol, rows, cols, v1, f1, m1, v2, f2, m2, compareFormulas, maxDiffs, usedRows1 != usedRows2);

                if (usedRangeCapped)
                    report = $"[Avviso: UsedRange tagliato al limite massimo di 5000 righe x 50 colonne; usa 'range' per leggere oltre.]\n\n" + report;
                return report;
            }
            finally
            {
                ComHelper.Release(r2);
                ComHelper.Release(r1);
                ComHelper.Release(ws2);
                ComHelper.Release(ws1);
                if (!ReferenceEquals(wb2, wb1)) ComHelper.Release(wb2);
                ComHelper.Release(wb1);
                ComHelper.Release(wbs);
                ComHelper.Release(app);
            }
        }

        internal static string ComputeKeyedSheetDiff(
            string label1, string label2, int startRow, int startCol, int rows, int cols,
            object[,] v1, object[,] f1, bool[,] m1,
            object[,] v2, object[,] f2, bool[,] m2,
            string keyColumn, bool hasHeader, bool compareFormulas, int maxDiffs)
        {
            int r10 = v1.GetLowerBound(0), c10 = v1.GetLowerBound(1);
            int r20 = v2.GetLowerBound(0), c20 = v2.GetLowerBound(1);
            int fr10 = f1.GetLowerBound(0), fc10 = f1.GetLowerBound(1);
            int fr20 = f2.GetLowerBound(0), fc20 = f2.GetLowerBound(1);

            // Risolvi indice colonna chiave:
            // - se hasHeader=true: prima per nome intestazione nella riga 0, poi per lettera colonna (es. 'A')
            // - se hasHeader=false: solo per lettera colonna (la riga 0 contiene dati)
            int keyColIdx = -1;
            string[] headers = new string[cols];
            var headerDiffs = new List<string>();
            for (int c = 0; c < cols; c++)
            {
                string colLetter = ColIndexToLetter(startCol + c);
                if (hasHeader)
                {
                    string h1 = FormatCellValue(v1[r10, c10 + c]).Trim();
                    string h2 = FormatCellValue(v2[r20, c20 + c]).Trim();
                    string h = !string.IsNullOrEmpty(h1) ? h1 : h2;
                    headers[c] = string.IsNullOrEmpty(h) ? colLetter : h;
                    if (!string.Equals(h1, h2, StringComparison.OrdinalIgnoreCase))
                        headerDiffs.Add($"col {colLetter}: '{(string.IsNullOrEmpty(h1) ? "[vuota]" : h1)}' -> '{(string.IsNullOrEmpty(h2) ? "[vuota]" : h2)}'");
                    if (keyColIdx < 0 && (!string.IsNullOrEmpty(h1) && string.Equals(h1, keyColumn, StringComparison.OrdinalIgnoreCase) ||
                                          !string.IsNullOrEmpty(h2) && string.Equals(h2, keyColumn, StringComparison.OrdinalIgnoreCase)))
                        keyColIdx = c;
                }
                else
                {
                    headers[c] = colLetter;
                }
            }
            if (keyColIdx < 0)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (string.Equals(ColIndexToLetter(startCol + c), keyColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        keyColIdx = c;
                        break;
                    }
                }
            }
            if (keyColIdx < 0)
            {
                string hint = hasHeader
                    ? $"tra le intestazioni ({string.Join(", ", headers)}) né come lettera di colonna"
                    : "come lettera di colonna (con has_header=false indica es. 'A')";
                throw new ArgumentException($"Colonna chiave '{keyColumn}' non trovata {hint}.");
            }

            var map1 = new Dictionary<string, (int row, string displayKey)>(StringComparer.OrdinalIgnoreCase);
            var map2 = new Dictionary<string, (int row, string displayKey)>(StringComparer.OrdinalIgnoreCase);
            var occ1 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var occ2 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var emptyKeyRows1 = new List<int>();
            var emptyKeyRows2 = new List<int>();
            int dup1 = 0, dup2 = 0;
            int firstDataRow = hasHeader ? 1 : 0;

            bool RowHasData(object[,] vals, object[,] frms, bool[,] mask, int vr0, int vc0, int fr0, int fc0, int rIdx)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (mask[rIdx, c] && !string.IsNullOrWhiteSpace(frms[fr0 + rIdx, fc0 + c] as string))
                        return true;
                    if (!string.IsNullOrWhiteSpace(FormatCellValue(vals[vr0 + rIdx, vc0 + c])))
                        return true;
                }
                return false;
            }

            for (int r = firstDataRow; r < rows; r++)
            {
                string k1 = FormatCellValue(v1[r10 + r, c10 + keyColIdx]).Trim();
                if (!string.IsNullOrEmpty(k1))
                {
                    occ1.TryGetValue(k1, out int c1);
                    c1++;
                    occ1[k1] = c1;
                    if (c1 > 1) dup1++;
                    string composite1 = c1 == 1 ? k1 : $"{k1}\0#{c1}";
                    string disp1 = c1 == 1 ? k1 : $"{k1} [occ. #{c1}]";
                    map1[composite1] = (r, disp1);
                }
                else if (RowHasData(v1, f1, m1, r10, c10, fr10, fc10, r))
                {
                    emptyKeyRows1.Add(startRow + r);
                }

                string k2 = FormatCellValue(v2[r20 + r, c20 + keyColIdx]).Trim();
                if (!string.IsNullOrEmpty(k2))
                {
                    occ2.TryGetValue(k2, out int c2);
                    c2++;
                    occ2[k2] = c2;
                    if (c2 > 1) dup2++;
                    string composite2 = c2 == 1 ? k2 : $"{k2}\0#{c2}";
                    string disp2 = c2 == 1 ? k2 : $"{k2} [occ. #{c2}]";
                    map2[composite2] = (r, disp2);
                }
                else if (RowHasData(v2, f2, m2, r20, c20, fr20, fc20, r))
                {
                    emptyKeyRows2.Add(startRow + r);
                }
            }

            var lines = new List<string>();
            int added = 0, removed = 0, modified = 0;

            // Righe rimosse o modificate (confronta TUTTE le occorrenze, incluse le chiavi duplicate)
            foreach (var kv in map1)
            {
                string compositeKey = kv.Key;
                int row1 = kv.Value.row;
                string dispKey = kv.Value.displayKey;
                if (!map2.TryGetValue(compositeKey, out var entry2))
                {
                    removed++;
                    if (lines.Count < maxDiffs)
                        lines.Add($"- RIMOSSA chiave '{dispKey}' (era riga {startRow + row1} in {label1})");
                }
                else
                {
                    int row2 = entry2.row;
                    var cellDiffs = new List<string>();
                    for (int c = 0; c < cols; c++)
                    {
                        if (c == keyColIdx) continue;
                        string s1 = FormatCellValue(v1[r10 + row1, c10 + c]);
                        string s2 = FormatCellValue(v2[r20 + row2, c20 + c]);
                        string sf1 = m1[row1, c] ? (f1[fr10 + row1, fc10 + c] as string ?? "") : "";
                        string sf2 = m2[row2, c] ? (f2[fr20 + row2, fc20 + c] as string ?? "") : "";
                        if (!string.Equals(s1, s2, StringComparison.Ordinal) ||
                            (compareFormulas && !string.Equals(sf1, sf2, StringComparison.Ordinal)))
                        {
                            string d1 = string.IsNullOrEmpty(s1) ? "[vuota]" : s1;
                            string d2 = string.IsNullOrEmpty(s2) ? "[vuota]" : s2;
                            if (m1[row1, c]) d1 += $" [{sf1}]";
                            if (m2[row2, c]) d2 += $" [{sf2}]";
                            cellDiffs.Add($"{headers[c]}: {d1} -> {d2}");
                        }
                    }
                    if (cellDiffs.Count > 0)
                    {
                        modified++;
                        if (lines.Count < maxDiffs)
                            lines.Add($"~ MODIFICATA chiave '{dispKey}' (r.{startRow + row1} -> r.{startRow + row2}): " + string.Join("; ", cellDiffs));
                    }
                }
            }

            // Righe aggiunte
            foreach (var kv in map2)
            {
                if (!map1.ContainsKey(kv.Key))
                {
                    added++;
                    if (lines.Count < maxDiffs)
                        lines.Add($"+ AGGIUNTA chiave '{kv.Value.displayKey}' (riga {startRow + kv.Value.row} in {label2})");
                }
            }

            string headerMode = hasHeader
                ? "has_header=true (prima riga = intestazione; imposta has_header=false se i dati partono dalla riga 1)"
                : "has_header=false (prima riga inclusa nei dati)";

            bool hasEmptyKeyData = emptyKeyRows1.Count > 0 || emptyKeyRows2.Count > 0;
            if (added == 0 && removed == 0 && modified == 0 && headerDiffs.Count == 0 && dup1 == 0 && dup2 == 0 && !hasEmptyKeyData)
                return $"Confronto per chiave '{headers[keyColIdx]}' ('{label1}' vs '{label2}', {headerMode}): NESSUNA DIFFERENZA ({map1.Count} righe allineate identiche).";

            var sb = new StringBuilder();
            sb.AppendLine($"Confronto per chiave '{headers[keyColIdx]}' ('{label1}' vs '{label2}', {headerMode}):");
            sb.AppendLine($"Righe: +{added} aggiunte, -{removed} rimosse, ~{modified} modificate.");
            if (headerDiffs.Count > 0)
                sb.AppendLine($"Avviso intestazioni diverse (colonne probabilmente aggiunte, rimosse o spostate; confronto eseguito per posizione): {string.Join("; ", headerDiffs)}.");
            if (dup1 > 0 || dup2 > 0)
                sb.AppendLine($"Avviso chiavi duplicate (confrontate per ordine di occorrenza #1, #2, ...): {dup1} extra in {label1}, {dup2} extra in {label2}.");
            if (hasEmptyKeyData)
            {
                string FormatRows(List<int> list)
                {
                    if (list.Count <= 15) return string.Join(", ", list);
                    return string.Join(", ", list.GetRange(0, 15)) + $" ... (+{list.Count - 15} altre)";
                }
                var parts = new List<string>();
                if (emptyKeyRows1.Count > 0) parts.Add($"{label1}: righe {FormatRows(emptyKeyRows1)}");
                if (emptyKeyRows2.Count > 0) parts.Add($"{label2}: righe {FormatRows(emptyKeyRows2)}");
                sb.AppendLine($"Avviso confronto incompleto: trovate righe con dati/formule ma chiave vuota (non confrontabili per chiave): {string.Join("; ", parts)}.");
            }
            if (added + removed + modified > lines.Count)
                sb.AppendLine($"(Mostrate le prime {lines.Count} differenze su {added + removed + modified})");
            if (lines.Count > 0)
            {
                sb.AppendLine();
                foreach (var l in lines) sb.AppendLine(l);
            }
            return ComHelper.Truncate(sb.ToString().TrimEnd(), 80000, out _);
        }

        internal static string ComputeSheetDiff(
            string label1, string label2, int startRow, int startCol, int rows, int cols,
            object[,] v1, object[,] f1, bool[,] m1,
            object[,] v2, object[,] f2, bool[,] m2,
            bool compareFormulas, int maxDiffs, bool rowCountsDiffer = false)
        {
            int r10 = v1.GetLowerBound(0), c10 = v1.GetLowerBound(1);
            int r20 = v2.GetLowerBound(0), c20 = v2.GetLowerBound(1);
            int fr10 = f1.GetLowerBound(0), fc10 = f1.GetLowerBound(1);
            int fr20 = f2.GetLowerBound(0), fc20 = f2.GetLowerBound(1);

            int diffCount = 0;
            var lines = new List<string>();

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    string s1 = FormatCellValue(v1[r10 + r, c10 + c]);
                    string s2 = FormatCellValue(v2[r20 + r, c20 + c]);
                    string sf1 = m1[r, c] ? (f1[fr10 + r, fc10 + c] as string ?? "") : "";
                    string sf2 = m2[r, c] ? (f2[fr20 + r, fc20 + c] as string ?? "") : "";

                    bool valDiff = !string.Equals(s1, s2, StringComparison.Ordinal);
                    bool frmDiff = compareFormulas && !string.Equals(sf1, sf2, StringComparison.Ordinal);

                    if (valDiff || frmDiff)
                    {
                        diffCount++;
                        if (lines.Count < maxDiffs)
                        {
                            string addr = $"{ColIndexToLetter(startCol + c)}{startRow + r}";
                            string d1 = string.IsNullOrEmpty(s1) ? "[vuota]" : s1;
                            string d2 = string.IsNullOrEmpty(s2) ? "[vuota]" : s2;
                            if (m1[r, c]) d1 += $" [{sf1}]";
                            if (m2[r, c]) d2 += $" [{sf2}]";
                            lines.Add($"  {addr}: {d1}  ->  {d2}");
                        }
                    }
                }
            }

            if (diffCount == 0)
                return $"Confronto '{label1}' vs '{label2}' ({rows}x{cols}): NESSUNA DIFFERENZA trovata.";

            var sb = new StringBuilder();
            sb.AppendLine($"Confronto '{label1}' vs '{label2}' ({rows} righe x {cols} colonne):");
            sb.AppendLine($"Celle diverse trovate: {diffCount}");
            if (rowCountsDiffer && diffCount >= 10)
                sb.AppendLine("Suggerimento: i due fogli hanno un numero di righe diverso; se è stata inserita o eliminata una riga, richiama excel_compare_sheets passando key_column (es. 'A' o il nome della colonna codice/ID).");
            if (diffCount > lines.Count)
                sb.AppendLine($"(Mostrate le prime {lines.Count} differenze su {diffCount})");
            sb.AppendLine();
            foreach (var l in lines) sb.AppendLine(l);
            return ComHelper.Truncate(sb.ToString().TrimEnd(), 80000, out _);
        }

        private static bool HasMixedCellFormatting(dynamic cell)
        {
            dynamic font = null;
            try
            {
                font = cell.Font;
                // In Excel COM, se una proprietà del font varia all'interno della stessa cella, restituisce DBNull.Value (null)
                object bold = font.Bold;
                object italic = font.Italic;
                object color = font.Color;
                object size = font.Size;
                object name = font.Name;
                return bold == null || bold == DBNull.Value ||
                       italic == null || italic == DBNull.Value ||
                       color == null || color == DBNull.Value ||
                       size == null || size == DBNull.Value ||
                       name == null || name == DBNull.Value;
            }
            catch { return false; }
            finally { ComHelper.Release(font); }
        }

        private static string ColIndexToLetter(int col)
        {
            var sb = new StringBuilder();
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                col = (col - 1) / 26;
            }
            return sb.ToString();
        }
    }
}
