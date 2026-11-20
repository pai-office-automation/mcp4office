using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json.Nodes;

namespace Mcp4Office.Tools
{
    internal static class PowerPointTools
    {
        public static void Register(McpServer server)
        {
            server.RegisterTool(
                "ppt_status",
                "Mostra presentation_id (percorso), numero di slide e slide selezionata in PowerPoint.",
                McpServer.EmptySchema(),
                McpServer.ReadOnly(),
                _ => ComHelper.RetryRead(GetStatus));

            server.RegisterTool(
                "ppt_read_slides",
                "Legge slide_id stabile (SlideID), titoli, testi (incluse tabelle e forme raggruppate) e note relatore delle slide in PowerPoint, con notes_state_token.",
                McpServer.Schema((props, _) =>
                {
                    props["from_slide"] = McpServer.IntProp("Numero della prima slide (1-based, default 1).", 1, 2000);
                    props["to_slide"] = McpServer.IntProp("Numero dell'ultima slide (opzionale; default tutte).", 1, 2000);
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ReadSlides(
                    McpServer.GetInt(args, "from_slide", 1),
                    McpServer.GetInt(args, "to_slide", 0))));

            server.RegisterTool(
                "ppt_add_slide",
                "Aggiunge una nuova slide IN FONDO alla presentazione attiva (testo semplice). Usa layout='title_only' prima di disegnare grafici/tabelle/blocchi, oppure 'title_and_content' (default) per testo.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_presentation"] = McpServer.Prop("string", "presentation_id (FullName) restituito da ppt_status/ppt_read_slides.");
                    props["title"] = McpServer.Prop("string", "Titolo della nuova slide.");
                    props["layout"] = McpServer.EnumProp("Layout: 'title_and_content' (default) o 'title_only' (ideale per slide grafiche).", "title_and_content", "title_only");
                    props["bullets"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Elenco di punti testuali per il corpo della slide (solo con layout='title_and_content').",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    };
                    props["notes"] = McpServer.Prop("string", "Note relatore opzionali.");
                    req.Add("expected_presentation");
                    req.Add("title");
                }),
                McpServer.Mutating(destructive: false),
                args => AddSlide(args));

            server.RegisterTool(
                "ppt_set_slide_notes",
                "Imposta le note relatore della slide identificata dallo stabile slide_id (da ppt_read_slides). Se esistono già note e overwrite!=true blocca restituendo notes_state_token; se expected_state_token è fornito viene SEMPRE validato anche se le note sono vuote.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_presentation"] = McpServer.Prop("string", "presentation_id (FullName) restituito da ppt_status/ppt_read_slides.");
                    props["slide_id"] = McpServer.IntProp("ID stabile della slide (slide_id restituito da ppt_read_slides; non cambia se si riordinano le slide).", 1, int.MaxValue);
                    props["notes"] = McpServer.Prop("string", "Testo semplice delle note relatore.");
                    props["overwrite"] = McpServer.Prop("boolean", "Se true sovrascrive note esistenti SOLO se expected_state_token coincide (default false).");
                    props["expected_state_token"] = McpServer.Prop("string", "Obbligatorio quando overwrite=true su note presenti; se fornito viene sempre validato.");
                    req.Add("expected_presentation");
                    req.Add("slide_id");
                    req.Add("notes");
                }),
                McpServer.Mutating(destructive: true),
                args => SetSlideNotes(
                    McpServer.GetString(args, "expected_presentation"),
                    McpServer.GetInt(args, "slide_id", 0),
                    McpServer.GetString(args, "notes"),
                    McpServer.GetBool(args, "overwrite", false),
                    McpServer.GetString(args, "expected_state_token")));

            server.RegisterTool(
                "ppt_add_table",
                "Aggiunge una tabella nativa modificabile (AddTable) su una slide esistente identificata da slide_id.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_presentation"] = McpServer.Prop("string", "presentation_id (FullName) da ppt_status/ppt_read_slides.");
                    props["slide_id"] = McpServer.IntProp("slide_id stabile della slide su cui aggiungere la tabella.", 1, int.MaxValue);
                    props["rows"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Matrice rettangolare 2D di stringhe (max 20 righe x 12 colonne; prima riga = intestazione).",
                        ["items"] = new JsonObject { ["type"] = "array" }
                    };
                    props["left"] = McpServer.IntProp("Posizione sinistra in punti (opzionale, default 50).", 0, 2000);
                    props["top"] = McpServer.IntProp("Posizione superiore in punti (opzionale, default 130).", 0, 2000);
                    props["width"] = McpServer.IntProp("Larghezza in punti (opzionale, default 860).", 50, 2000);
                    props["height"] = McpServer.IntProp("Altezza in punti (opzionale, default 320).", 30, 2000);
                    req.Add("expected_presentation");
                    req.Add("slide_id");
                    req.Add("rows");
                }),
                McpServer.Mutating(destructive: false),
                args => AddTable(args));

            server.RegisterTool(
                "ppt_add_visual_blocks",
                "Disegna uno schema grafico nativo con forme modificabili (AddShape) su una slide: card KPI ('kpi_cards'), flusso a passi con frecce ('process_steps') o griglia di box ('grid_boxes').",
                McpServer.Schema((props, req) =>
                {
                    props["expected_presentation"] = McpServer.Prop("string", "presentation_id (FullName) da ppt_status/ppt_read_slides.");
                    props["slide_id"] = McpServer.IntProp("slide_id stabile della slide.", 1, int.MaxValue);
                    props["layout"] = McpServer.EnumProp("Stile grafico: 'kpi_cards', 'process_steps' o 'grid_boxes'.", "kpi_cards", "process_steps", "grid_boxes");
                    props["items"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Elenco di 1-6 blocchi: ciascuno è un oggetto {\"title\": \"...\", \"body\": \"...\"}.",
                        ["items"] = new JsonObject { ["type"] = "object" }
                    };
                    req.Add("expected_presentation");
                    req.Add("slide_id");
                    req.Add("layout");
                    req.Add("items");
                }),
                McpServer.Mutating(destructive: false),
                args => AddVisualBlocks(args));

            server.RegisterTool(
                "ppt_add_chart",
                "Aggiunge un grafico nativo modificabile (AddChart2) su una slide esistente con categorie e serie numeriche.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_presentation"] = McpServer.Prop("string", "presentation_id (FullName) da ppt_status/ppt_read_slides.");
                    props["slide_id"] = McpServer.IntProp("slide_id stabile della slide.", 1, int.MaxValue);
                    props["chart_type"] = McpServer.EnumProp("Tipo di grafico: 'column', 'bar', 'line', 'pie'.", "column", "bar", "line", "pie");
                    props["title"] = McpServer.Prop("string", "Titolo del grafico (opzionale).");
                    props["categories"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Etichette asse categorie (es. [\"Q1\",\"Q2\",\"Q3\",\"Q4\"]).",
                        ["items"] = new JsonObject { ["type"] = "string" }
                    };
                    props["series"] = new JsonObject
                    {
                        ["type"] = "array",
                        ["description"] = "Serie numeriche: [{\"name\": \"Vendite\", \"values\": [10, 20, 15, 30]}].",
                        ["items"] = new JsonObject { ["type"] = "object" }
                    };
                    req.Add("expected_presentation");
                    req.Add("slide_id");
                    req.Add("chart_type");
                    req.Add("categories");
                    req.Add("series");
                }),
                McpServer.Mutating(destructive: false),
                args => AddChart(args));

            server.RegisterTool(
                "ppt_add_picture",
                "Inserisce un'immagine locale (.png, .jpg, .jpeg, .bmp, .gif; max 25 MB, nelle cartelle consentite da --pdf-roots/Documenti/Desktop/Download) su una slide esistente preservando le proporzioni.",
                McpServer.Schema((props, req) =>
                {
                    props["expected_presentation"] = McpServer.Prop("string", "presentation_id (FullName) da ppt_status/ppt_read_slides.");
                    props["slide_id"] = McpServer.IntProp("slide_id stabile della slide.", 1, int.MaxValue);
                    props["path"] = McpServer.Prop("string", "Percorso locale del file immagine (no UNC/symlink).");
                    props["left"] = McpServer.IntProp("Posizione sinistra in punti (opzionale, default 80).", 0, 2000);
                    props["top"] = McpServer.IntProp("Posizione superiore in punti (opzionale, default 120).", 0, 2000);
                    props["width"] = McpServer.IntProp("Larghezza desiderata in punti (opzionale, default 480; altezza proporzionale).", 30, 2000);
                    req.Add("expected_presentation");
                    req.Add("slide_id");
                    req.Add("path");
                }),
                McpServer.Mutating(destructive: false),
                args => AddPicture(args));
        }

        internal static string ComputeNotesToken(string presFullName, int slideId, string notes)
        {
            string n = notes ?? string.Empty;
            return ComHelper.Sha256Short($"{presFullName?.Length ?? 0}:{presFullName}|sid:{slideId}|n:{n.Length}:{n}");
        }

        private static string GetStatus()
        {
            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    return "Microsoft PowerPoint è aperto ma non ci sono presentazioni aperte.";

                pres = app.ActivePresentation;
                slides = pres.Slides;
                var sb = new StringBuilder();
                sb.AppendLine($"presentation_id: {pres.FullName}");
                sb.AppendLine($"Presentazione attiva: {pres.Name}");
                sb.AppendLine($"Numero slide: {slides.Count}");
                try
                {
                    dynamic curSlide = app.ActiveWindow.View.Slide;
                    try
                    {
                        sb.AppendLine($"Slide corrente: {curSlide.SlideIndex} (slide_id: {curSlide.SlideID})");
                    }
                    finally { ComHelper.Release(curSlide); }
                }
                catch { }
                return sb.ToString().TrimEnd();
            }
            finally
            {
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static string ReadSlides(int fromSlide, int toSlide)
        {
            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string fullName = (string)pres.FullName;
                string name = (string)pres.Name;
                slides = pres.Slides;
                int total = (int)slides.Count;
                if (total == 0)
                    return $"presentation_id: {fullName}\n(La presentazione '{name}' non contiene slide.)";

                int start = ComHelper.Clamp(fromSlide, 1, total);
                int end = toSlide > 0 ? Math.Min(total, toSlide) : total;

                var sb = new StringBuilder();
                sb.AppendLine($"presentation_id: {fullName}");
                sb.AppendLine($"Presentazione: {name} | Slide {start}-{end} di {total}");

                for (int i = start; i <= end; i++)
                {
                    dynamic slide = slides.Item(i);
                    dynamic shapes = null;
                    try
                    {
                        int slideId = (int)slide.SlideID;
                        sb.AppendLine($"\n--- Slide {i} (slide_id: {slideId}) ---");
                        shapes = slide.Shapes;
                        int shapeCount = (int)shapes.Count;
                        for (int s = 1; s <= shapeCount; s++)
                        {
                            dynamic shape = shapes.Item(s);
                            try { ExtractShapeText(shape, sb); }
                            finally { ComHelper.Release(shape); }
                        }

                        string notes = ExtractNotes(slide);
                        string notesToken = ComputeNotesToken(fullName, slideId, notes);
                        if (!string.IsNullOrWhiteSpace(notes))
                            sb.AppendLine($"[Note (notes_state_token={notesToken})]: {notes}");
                        else
                            sb.AppendLine($"[Note vuote (notes_state_token={notesToken})]");
                    }
                    finally
                    {
                        ComHelper.Release(shapes);
                        ComHelper.Release(slide);
                    }
                }

                return ComHelper.Truncate(sb.ToString().TrimEnd(), 80000, out _);
            }
            finally
            {
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static void ExtractShapeText(dynamic shape, StringBuilder sb)
        {
            try
            {
                if ((int)shape.Type == 6) // msoGroup
                {
                    dynamic groupItems = shape.GroupItems;
                    try
                    {
                        int gCount = (int)groupItems.Count;
                        for (int g = 1; g <= gCount; g++)
                        {
                            dynamic child = groupItems.Item(g);
                            try { ExtractShapeText(child, sb); }
                            finally { ComHelper.Release(child); }
                        }
                    }
                    finally { ComHelper.Release(groupItems); }
                    return;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic tbl = shape.Table;
                    try
                    {
                        int rows = (int)tbl.Rows.Count;
                        int cols = (int)tbl.Columns.Count;
                        sb.AppendLine("[Tabella]:");
                        for (int r = 1; r <= rows; r++)
                        {
                            var rowVals = new List<string>(cols);
                            for (int c = 1; c <= cols; c++)
                            {
                                dynamic cell = tbl.Cell(r, c);
                                try
                                {
                                    string t = (string)(cell.Shape.TextFrame.TextRange.Text ?? string.Empty);
                                    rowVals.Add(t.Replace("\r", " ").Trim());
                                }
                                catch { rowVals.Add(string.Empty); }
                                finally { ComHelper.Release(cell); }
                            }
                            sb.AppendLine("  " + string.Join(" | ", rowVals));
                        }
                    }
                    finally { ComHelper.Release(tbl); }
                    return;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    string txt = (string)(shape.TextFrame.TextRange.Text ?? string.Empty);
                    txt = txt.Replace("\r", "\n").Trim();
                    if (!string.IsNullOrEmpty(txt))
                        sb.AppendLine(txt);
                }
            }
            catch { }
        }

        private static string AddSlide(JsonObject args)
        {
            string expectedPres = McpServer.GetString(args, "expected_presentation").Trim();
            string title = McpServer.GetString(args, "title").Trim();
            string layout = McpServer.GetString(args, "layout", "title_and_content");
            string notes = McpServer.GetString(args, "notes");
            int ppLayout = layout == "title_only" ? 11 : 2; // 11 = ppLayoutTitleOnly, 2 = ppLayoutText
            var bullets = new List<string>();
            if (args.TryGetPropertyValue("bullets", out var bNode) && bNode is JsonArray arr)
            {
                foreach (var item in arr)
                {
                    if (item != null)
                    {
                        var s = item.ToString().Trim();
                        if (s.Length > 0) bullets.Add(s);
                    }
                }
            }

            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            dynamic slide = null;
            dynamic shapes = null;
            dynamic placeholders = null;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string actualPres = (string)pres.FullName;
                if (!string.Equals(actualPres, expectedPres, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Presentazione cambiata: attesa '{expectedPres}', ma attiva '{actualPres}'.");

                slides = pres.Slides;
                int newIndex = (int)slides.Count + 1;
                slide = slides.Add(newIndex, ppLayout);
                int newSlideId = (int)slide.SlideID;
                shapes = slide.Shapes;
                placeholders = shapes.Placeholders;

                var warnings = new List<string>();
                bool titleWritten = false;
                bool bodyWritten = bullets.Count == 0;
                bool notesWritten = string.IsNullOrWhiteSpace(notes);

                int phCount = (int)placeholders.Count;
                for (int i = 1; i <= phCount; i++)
                {
                    dynamic ph = placeholders.Item(i);
                    try
                    {
                        int phType = (int)ph.PlaceholderFormat.Type;
                        if (!titleWritten && (phType == 1 || phType == 3))
                        {
                            ph.TextFrame.TextRange.Text = title;
                            titleWritten = true;
                        }
                        else if (!bodyWritten && (phType == 2 || phType == 7))
                        {
                            ph.TextFrame.TextRange.Text = string.Join("\r", bullets);
                            bodyWritten = true;
                        }
                    }
                    finally { ComHelper.Release(ph); }
                }

                if (!titleWritten) warnings.Add("segnaposto titolo non trovato nel layout corrente");
                if (!bodyWritten) warnings.Add("segnaposto corpo/contenuto non trovato nel layout corrente");

                if (!notesWritten)
                {
                    notesWritten = TryWriteNotes(slide, notes);
                    if (!notesWritten) warnings.Add("segnaposto note relatore non trovato");
                }

                if (!titleWritten && !bodyWritten)
                {
                    try { slide.Delete(); } catch { }
                    throw new InvalidOperationException($"Slide vuota rimossa: {string.Join("; ", warnings)}.");
                }

                string msg = $"Aggiunta slide {newIndex} (slide_id: {newSlideId}, '{title}') in fondo a '{pres.Name}'.";
                if (warnings.Count > 0)
                    msg += " Avvisi parziali: " + string.Join("; ", warnings) + ".";
                return msg;
            }
            finally
            {
                ComHelper.Release(placeholders);
                ComHelper.Release(shapes);
                ComHelper.Release(slide);
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static string SetSlideNotes(string expectedPres, int slideId, string notes, bool overwrite, string expectedToken)
        {
            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            dynamic slide = null;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string actualPres = (string)pres.FullName;
                if (!string.Equals(actualPres, expectedPres, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Presentazione cambiata: attesa '{expectedPres}', ma attiva '{actualPres}'.");

                slides = pres.Slides;
                try
                {
                    slide = slides.FindBySlideID(slideId);
                }
                catch
                {
                    throw new InvalidOperationException($"Slide con slide_id={slideId} non trovata in '{pres.Name}'. Rileggi con ppt_read_slides.");
                }

                string currentNotes = ExtractNotes(slide);
                string currentToken = ComputeNotesToken(actualPres, slideId, currentNotes);

                // Valida SEMPRE expected_state_token se fornito (anche se le note sono diventate vuote!).
                if (!string.IsNullOrWhiteSpace(expectedToken) && !string.Equals(expectedToken.Trim(), currentToken, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Note della slide (slide_id={slideId}) cambiate (atteso token '{expectedToken}', attuale '{currentToken}'). Rileggi prima di scrivere.");
                }

                if (!string.IsNullOrWhiteSpace(currentNotes))
                {
                    if (!overwrite)
                    {
                        return $"SCRITTURA NOTE BLOCCATA: la slide (slide_id={slideId}) contiene già note relatore.\n" +
                               $"notes_state_token: {currentToken}\n" +
                               $"Note attuali: {ComHelper.Truncate(currentNotes, 500, out _)}\n" +
                               "Chiedi conferma all'utente e richiama ppt_set_slide_notes con overwrite=true e questo expected_state_token.";
                    }
                    if (string.IsNullOrWhiteSpace(expectedToken))
                        throw new ArgumentException("Quando overwrite=true su note già presenti, 'expected_state_token' è obbligatorio.");
                }

                if (!TryWriteNotes(slide, notes ?? string.Empty))
                    throw new InvalidOperationException($"Impossibile scrivere le note su slide_id={slideId}: segnaposto note relatore non trovato.");

                return $"Note relatore aggiornate per la slide {slide.SlideIndex} (slide_id={slideId}) di '{pres.Name}'.";
            }
            finally
            {
                ComHelper.Release(slide);
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        /// <summary>
        /// Determina l'area contenuto utile sulla slide: se esiste un segnaposto corpo/oggetto VUOTO,
        /// ne riusa esattamente Left/Top/Width/Height e lo rimuove per evitare sovrapposizioni;
        /// altrimenti calcola margini proporzionali su PageSetup.SlideWidth/SlideHeight (sicuro anche su 4:3).
        /// </summary>
        private static (float left, float top, float width, float height, float slideW, float slideH) ResolveContentArea(dynamic pres, dynamic slide)
        {
            float slideW = 960f, slideH = 540f;
            try
            {
                slideW = (float)pres.PageSetup.SlideWidth;
                slideH = (float)pres.PageSetup.SlideHeight;
            }
            catch { }

            dynamic shapes = slide.Shapes;
            dynamic placeholders = null;
            try
            {
                placeholders = shapes.Placeholders;
                int phCount = (int)placeholders.Count;
                for (int i = 1; i <= phCount; i++)
                {
                    dynamic ph = placeholders.Item(i);
                    try
                    {
                        int phType = (int)ph.PlaceholderFormat.Type;
                        if (phType == 2 || phType == 7) // Body / Object
                        {
                            bool isEmpty = true;
                            try
                            {
                                if ((int)ph.HasTextFrame == -1 && (int)ph.TextFrame.HasText == -1)
                                {
                                    string t = ((string)(ph.TextFrame.TextRange.Text ?? string.Empty)).Trim();
                                    if (t.Length > 0) isEmpty = false;
                                }
                            }
                            catch { }

                            if (isEmpty)
                            {
                                float l = (float)ph.Left;
                                float tp = (float)ph.Top;
                                float w = (float)ph.Width;
                                float h = (float)ph.Height;
                                try { ph.Delete(); } catch { }
                                if (w > 80f && h > 60f)
                                    return (l, tp, w, h, slideW, slideH);
                            }
                        }
                    }
                    finally { ComHelper.Release(ph); }
                }
            }
            catch { }
            finally
            {
                ComHelper.Release(placeholders);
                ComHelper.Release(shapes);
            }

            float marginX = Math.Max(36f, slideW * 0.05f);
            float topY = Math.Min(135f, slideH * 0.24f);
            float areaW = Math.Max(240f, slideW - marginX * 2f);
            float areaH = Math.Max(160f, slideH - topY - 48f);
            return (marginX, topY, areaW, areaH, slideW, slideH);
        }

        private static string AddTable(JsonObject args)
        {
            string expectedPres = McpServer.GetString(args, "expected_presentation").Trim();
            int slideId = McpServer.GetInt(args, "slide_id", 0);
            bool hasLeft = args.ContainsKey("left");
            bool hasTop = args.ContainsKey("top");
            bool hasWidth = args.ContainsKey("width");
            bool hasHeight = args.ContainsKey("height");

            var rowsArr = args["rows"] as JsonArray;
            if (rowsArr == null || rowsArr.Count == 0 || rowsArr.Count > 20)
                throw new ArgumentException("Parametro 'rows' deve avere tra 1 e 20 righe.");

            var matrix = new List<List<string>>(rowsArr.Count);
            int cols = -1;
            foreach (var rNode in rowsArr)
            {
                var cArr = rNode as JsonArray;
                if (cArr == null || cArr.Count == 0 || cArr.Count > 12)
                    throw new ArgumentException("Ogni riga della tabella deve avere tra 1 e 12 colonne.");
                if (cols < 0) cols = cArr.Count;
                else if (cArr.Count != cols)
                    throw new ArgumentException($"Tabella non rettangolare: attese {cols} colonne, trovate {cArr.Count}.");

                var row = new List<string>(cols);
                foreach (var cNode in cArr)
                    row.Add(cNode == null ? string.Empty : cNode.ToString());
                matrix.Add(row);
            }

            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            dynamic slide = null;
            dynamic shapes = null;
            dynamic tblShape = null;
            dynamic tbl = null;
            bool completed = false;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string actualPres = (string)pres.FullName;
                if (!string.Equals(actualPres, expectedPres, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Presentazione cambiata: attesa '{expectedPres}', attiva '{actualPres}'.");

                slides = pres.Slides;
                try { slide = slides.FindBySlideID(slideId); }
                catch { throw new InvalidOperationException($"Slide con slide_id={slideId} non trovata."); }

                var area = ResolveContentArea(pres, slide);
                float left = hasLeft ? Math.Min(McpServer.GetInt(args, "left", (int)area.left), area.slideW - 80f) : area.left;
                float top = hasTop ? Math.Min(McpServer.GetInt(args, "top", (int)area.top), area.slideH - 60f) : area.top;
                float width = hasWidth ? Math.Min(McpServer.GetInt(args, "width", (int)area.width), area.slideW - left - 16f) : area.width;
                float height = hasHeight ? Math.Min(McpServer.GetInt(args, "height", (int)area.height), area.slideH - top - 16f) : area.height;

                shapes = slide.Shapes;
                tblShape = shapes.AddTable(matrix.Count, cols, left, top, width, height);
                tbl = tblShape.Table;
                float fontSize = matrix.Count > 8 ? 12f : (matrix.Count > 5 ? 14f : 16f);

                for (int r = 0; r < matrix.Count; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        dynamic cell = tbl.Cell(r + 1, c + 1);
                        try
                        {
                            dynamic tr = cell.Shape.TextFrame.TextRange;
                            try
                            {
                                tr.Text = matrix[r][c];
                                tr.Font.Size = fontSize;
                            }
                            finally { ComHelper.Release(tr); }
                        }
                        finally { ComHelper.Release(cell); }
                    }
                }
                completed = true;

                string warn = string.Empty;
                try
                {
                    float finalBottom = (float)tblShape.Top + (float)tblShape.Height;
                    if (finalBottom > area.slideH)
                        warn = $" Avviso: l'altezza finale della tabella ({finalBottom:F0} pt) supera la slide ({area.slideH:F0} pt).";
                }
                catch { }

                return $"Aggiunta tabella {matrix.Count}x{cols} sulla slide {slide.SlideIndex} (slide_id={slideId}) di '{pres.Name}'.{warn}";
            }
            catch
            {
                if (!completed && tblShape != null) { try { tblShape.Delete(); } catch { } }
                throw;
            }
            finally
            {
                ComHelper.Release(tbl);
                ComHelper.Release(tblShape);
                ComHelper.Release(shapes);
                ComHelper.Release(slide);
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static string AddVisualBlocks(JsonObject args)
        {
            string expectedPres = McpServer.GetString(args, "expected_presentation").Trim();
            int slideId = McpServer.GetInt(args, "slide_id", 0);
            string layout = McpServer.GetString(args, "layout", "kpi_cards");
            var itemsArr = args["items"] as JsonArray;
            if (itemsArr == null || itemsArr.Count < 1 || itemsArr.Count > 6)
                throw new ArgumentException("Parametro 'items' deve contenere tra 1 e 6 blocchi.");

            var items = new List<(string title, string body)>(itemsArr.Count);
            foreach (var n in itemsArr)
            {
                var obj = n as JsonObject;
                if (obj == null) throw new ArgumentException("Ogni blocco in 'items' deve essere un oggetto {title, body}.");
                string t = McpServer.GetString(obj, "title").Trim();
                string b = McpServer.GetString(obj, "body").Trim();
                if (string.IsNullOrEmpty(t) && string.IsNullOrEmpty(b))
                    throw new ArgumentException("Ogni blocco in 'items' deve avere almeno 'title' o 'body'.");
                items.Add((t, b));
            }

            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            dynamic slide = null;
            dynamic shapes = null;
            var createdNames = new List<string>();
            bool completed = false;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string actualPres = (string)pres.FullName;
                if (!string.Equals(actualPres, expectedPres, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Presentazione cambiata: attesa '{expectedPres}', attiva '{actualPres}'.");

                slides = pres.Slides;
                try { slide = slides.FindBySlideID(slideId); }
                catch { throw new InvalidOperationException($"Slide con slide_id={slideId} non trovata."); }

                var area = ResolveContentArea(pres, slide);
                shapes = slide.Shapes;

                int nItems = items.Count;
                if (layout == "grid_boxes" && nItems >= 4)
                {
                    int cols = (nItems <= 4) ? 2 : 3;
                    int rows = (int)Math.Ceiling(nItems / (double)cols);
                    float gap = 16f;
                    float boxW = (area.width - gap * (cols - 1)) / cols;
                    float boxH = (area.height - gap * (rows - 1)) / rows;
                    for (int i = 0; i < nItems; i++)
                    {
                        int r = i / cols, c = i % cols;
                        float x = area.left + c * (boxW + gap);
                        float y = area.top + r * (boxH + gap);
                        createdNames.Add(DrawThemeCardShape(shapes, 5, x, y, boxW, boxH, false, items[i].title, items[i].body));
                    }
                }
                else
                {
                    bool arrows = layout == "process_steps";
                    float gap = arrows ? 28f : 16f;
                    float boxW = (area.width - gap * (nItems - 1)) / nItems;
                    float boxH = Math.Min(250f, area.height);
                    float y = area.top + (area.height - boxH) / 2f;

                    for (int i = 0; i < nItems; i++)
                    {
                        float x = area.left + i * (boxW + gap);
                        bool useAccent = layout == "kpi_cards";
                        string titleTxt = arrows ? $"{i + 1}. {items[i].title}" : items[i].title;
                        createdNames.Add(DrawThemeCardShape(shapes, 5, x, y, boxW, boxH, useAccent, titleTxt, items[i].body));

                        if (arrows && i < nItems - 1)
                        {
                            dynamic arrow = shapes.AddShape(33, x + boxW + 4f, y + boxH / 2f - 12f, gap - 8f, 24f);
                            try
                            {
                                arrow.Fill.Solid();
                                try { arrow.Fill.ForeColor.ObjectThemeColor = 5; } // msoThemeColorAccent1
                                catch { arrow.Fill.ForeColor.RGB = 0xA16903; }
                                arrow.Line.Visible = 0;
                                createdNames.Add((string)arrow.Name);
                            }
                            finally { ComHelper.Release(arrow); }
                        }
                    }
                }

                // Raggruppa tutte le forme create così l'utente può selezionarle/spostarle/cancellarle in un solo clic
                string groupName = string.Empty;
                if (createdNames.Count > 1)
                {
                    try
                    {
                        dynamic sr = shapes.Range(createdNames.ToArray());
                        dynamic grp = sr.Group();
                        try
                        {
                            groupName = $"mcp4office {layout}";
                            grp.Name = groupName;
                        }
                        finally
                        {
                            ComHelper.Release(grp);
                            ComHelper.Release(sr);
                        }
                    }
                    catch { }
                }
                completed = true;

                return $"Disegnati {nItems} blocchi grafici ({layout}) sulla slide {slide.SlideIndex} (slide_id={slideId}) di '{pres.Name}'" +
                       (string.IsNullOrEmpty(groupName) ? "." : $" (raggruppati come '{groupName}').");
            }
            catch
            {
                if (!completed && shapes != null && createdNames.Count > 0)
                {
                    foreach (var n in createdNames)
                    {
                        try { dynamic s = shapes.Item(n); s.Delete(); ComHelper.Release(s); } catch { }
                    }
                }
                throw;
            }
            finally
            {
                ComHelper.Release(shapes);
                ComHelper.Release(slide);
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static string DrawThemeCardShape(dynamic shapes, int shapeType, float x, float y, float w, float h, bool accentFill, string title, string body)
        {
            dynamic shp = shapes.AddShape(shapeType, x, y, w, h);
            try
            {
                shp.Fill.Solid();
                // Usa i colori del tema della presentazione (Accent1=5, Light2=4, Light1=2, Dark1=1) e eredita il font del tema.
                try
                {
                    shp.Fill.ForeColor.ObjectThemeColor = accentFill ? 5 : 4; // Accent1 / Light2
                    shp.Line.ForeColor.ObjectThemeColor = 5; // Accent1
                }
                catch
                {
                    shp.Fill.ForeColor.RGB = accentFill ? 0xA16903 : 0xF9F5F1;
                    shp.Line.ForeColor.RGB = 0xE2D8CE;
                }
                shp.Line.Weight = 1.25f;
                shp.TextFrame.WordWrap = -1;
                shp.TextFrame.MarginLeft = 12f;
                shp.TextFrame.MarginRight = 12f;
                shp.TextFrame.MarginTop = 10f;
                shp.TextFrame.MarginBottom = 10f;

                string combined = string.IsNullOrEmpty(body) ? title : $"{title}\r{body}";
                dynamic tr = shp.TextFrame.TextRange;
                try
                {
                    tr.Text = combined;
                    tr.Font.Size = 14f;
                    try { tr.Font.Color.ObjectThemeColor = accentFill ? 2 : 1; } // Light1 / Dark1
                    catch { tr.Font.Color.RGB = accentFill ? 0xFFFFFF : 0x2A170F; }
                    if (!string.IsNullOrEmpty(title))
                    {
                        dynamic p1 = tr.Paragraphs(1, 1);
                        try
                        {
                            p1.Font.Bold = -1;
                            p1.Font.Size = 18f;
                        }
                        finally { ComHelper.Release(p1); }
                    }
                }
                finally { ComHelper.Release(tr); }
                return (string)shp.Name;
            }
            finally { ComHelper.Release(shp); }
        }

        private static string AddChart(JsonObject args)
        {
            string expectedPres = McpServer.GetString(args, "expected_presentation").Trim();
            int slideId = McpServer.GetInt(args, "slide_id", 0);
            string chartType = McpServer.GetString(args, "chart_type", "column");
            string title = McpServer.GetString(args, "title").Trim();

            var catsArr = args["categories"] as JsonArray;
            var seriesArr = args["series"] as JsonArray;
            if (catsArr == null || catsArr.Count == 0 || catsArr.Count > 50)
                throw new ArgumentException("Parametro 'categories' deve avere tra 1 e 50 voci.");
            if (seriesArr == null || seriesArr.Count == 0 || seriesArr.Count > 10)
                throw new ArgumentException("Parametro 'series' deve avere tra 1 e 10 serie.");
            if (chartType == "pie" && seriesArr.Count > 1)
                throw new ArgumentException("Il grafico a torta ('pie') supporta una sola serie di dati.");

            var categories = new List<string>(catsArr.Count);
            foreach (var c in catsArr) categories.Add(c?.ToString() ?? string.Empty);

            var series = new List<(string name, List<double> values)>(seriesArr.Count);
            foreach (var sNode in seriesArr)
            {
                var obj = sNode as JsonObject;
                if (obj == null) throw new ArgumentException("Ogni serie deve essere un oggetto {name, values}.");
                string sName = McpServer.GetString(obj, "name", "Serie");
                var valsArr = obj.TryGetPropertyValue("values", out var vn) ? vn as JsonArray : null;
                if (valsArr == null || valsArr.Count != categories.Count)
                    throw new ArgumentException($"La serie '{sName}' deve avere esattamente {categories.Count} valori numerici.");
                var vals = new List<double>(valsArr.Count);
                foreach (var v in valsArr)
                {
                    if (!(v is JsonValue jv) || !jv.TryGetValue<double>(out var d) || double.IsNaN(d) || double.IsInfinity(d))
                        throw new ArgumentException($"Valore non numerico nella serie '{sName}'.");
                    vals.Add(d);
                }
                series.Add((sName, vals));
            }

            int xlChartType = chartType switch
            {
                "bar" => 57,
                "line" => 4,
                "pie" => 5,
                _ => 51
            };

            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            dynamic slide = null;
            dynamic shapes = null;
            dynamic chartShape = null;
            dynamic chart = null;
            dynamic chartData = null;
            dynamic wb = null;
            dynamic ws = null;
            bool completed = false;
            string closeWarning = string.Empty;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string actualPres = (string)pres.FullName;
                if (!string.Equals(actualPres, expectedPres, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Presentazione cambiata: attesa '{expectedPres}', attiva '{actualPres}'.");

                slides = pres.Slides;
                try { slide = slides.FindBySlideID(slideId); }
                catch { throw new InvalidOperationException($"Slide con slide_id={slideId} non trovata."); }

                var area = ResolveContentArea(pres, slide);
                shapes = slide.Shapes;
                chartShape = shapes.AddChart2(-1, xlChartType, area.left, area.top, area.width, area.height);
                chart = chartShape.Chart;
                chartData = chart.ChartData;
                try { chartData.ActivateChartDataWindow(); } catch { chartData.Activate(); }
                wb = chartData.Workbook;
                ws = wb.Worksheets.Item(1);

                ws.Cells.Clear();
                ws.Cells.Item(1, 1).Value2 = "Categoria";
                for (int s = 0; s < series.Count; s++)
                    ws.Cells.Item(1, s + 2).Value2 = ExcelTools.EscapeLiteralTextForExcel(series[s].name);

                for (int r = 0; r < categories.Count; r++)
                {
                    ws.Cells.Item(r + 2, 1).Value2 = ExcelTools.EscapeLiteralTextForExcel(categories[r]);
                    for (int s = 0; s < series.Count; s++)
                        ws.Cells.Item(r + 2, s + 2).Value2 = series[s].values[r];
                }

                string lastCol = ((char)('A' + series.Count)).ToString();
                string dataRange = $"='{ws.Name}'!$A$1:${lastCol}${categories.Count + 1}";
                chart.SetSourceData(dataRange);

                if (!string.IsNullOrEmpty(title))
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = title;
                }
                completed = true;

                return $"Aggiunto grafico '{chartType}' ({series.Count} serie, {categories.Count} categorie) sulla slide {slide.SlideIndex} (slide_id={slideId}) di '{pres.Name}'.{closeWarning}";
            }
            catch
            {
                if (!completed && chartShape != null)
                {
                    try { chartShape.Delete(); } catch { }
                }
                throw;
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(true); }
                    catch (Exception cex) { closeWarning = $" (Avviso chiusura workbook grafico: {cex.Message})"; }
                }
                ComHelper.Release(ws);
                ComHelper.Release(wb);
                ComHelper.Release(chartData);
                ComHelper.Release(chart);
                ComHelper.Release(chartShape);
                ComHelper.Release(shapes);
                ComHelper.Release(slide);
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static string AddPicture(JsonObject args)
        {
            string expectedPres = McpServer.GetString(args, "expected_presentation").Trim();
            int slideId = McpServer.GetInt(args, "slide_id", 0);
            string rawPath = McpServer.GetString(args, "path");
            float left = McpServer.GetInt(args, "left", 80);
            float top = McpServer.GetInt(args, "top", 120);
            float width = McpServer.GetInt(args, "width", 480);

            string fullPath = PdfTools.ValidateLocalFilePath(
                rawPath,
                new[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif" },
                25L * 1024 * 1024,
                "immagine");

            dynamic app = ComHelper.GetActiveApp("PowerPoint.Application", "Microsoft PowerPoint");
            dynamic presAll = null;
            dynamic pres = null;
            dynamic slides = null;
            dynamic slide = null;
            dynamic shapes = null;
            dynamic pic = null;
            try
            {
                presAll = app.Presentations;
                if ((int)presAll.Count == 0)
                    throw new InvalidOperationException("Nessuna presentazione aperta in PowerPoint.");

                pres = app.ActivePresentation;
                string actualPres = (string)pres.FullName;
                if (!string.Equals(actualPres, expectedPres, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException($"Presentazione cambiata: attesa '{expectedPres}', attiva '{actualPres}'.");

                slides = pres.Slides;
                try { slide = slides.FindBySlideID(slideId); }
                catch { throw new InvalidOperationException($"Slide con slide_id={slideId} non trovata."); }

                shapes = slide.Shapes;
                // LinkToFile=0 (msoFalse), SaveWithDocument=-1 (msoTrue), Width/Height=-1 (dimensione nativa poi ridimensionata proporzionalmente)
                pic = shapes.AddPicture(fullPath, 0, -1, left, top, -1, -1);
                pic.LockAspectRatio = -1; // msoTrue
                pic.Width = width;

                return $"Immagine '{System.IO.Path.GetFileName(fullPath)}' inserita sulla slide {slide.SlideIndex} (slide_id={slideId}) di '{pres.Name}'.";
            }
            finally
            {
                ComHelper.Release(pic);
                ComHelper.Release(shapes);
                ComHelper.Release(slide);
                ComHelper.Release(slides);
                ComHelper.Release(pres);
                ComHelper.Release(presAll);
                ComHelper.Release(app);
            }
        }

        private static string ExtractNotes(dynamic slide)
        {
            dynamic notesPage = slide.NotesPage;
            dynamic shapes = null;
            dynamic placeholders = null;
            try
            {
                shapes = notesPage.Shapes;
                placeholders = shapes.Placeholders;
                int count = (int)placeholders.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic ph = placeholders.Item(i);
                    try
                    {
                        if ((int)ph.PlaceholderFormat.Type == 2 && (int)ph.HasTextFrame == -1)
                            return ((string)(ph.TextFrame.TextRange.Text ?? string.Empty)).Replace("\r", "\n").Trim();
                    }
                    finally { ComHelper.Release(ph); }
                }
                return string.Empty;
            }
            finally
            {
                ComHelper.Release(placeholders);
                ComHelper.Release(shapes);
                ComHelper.Release(notesPage);
            }
        }

        private static bool TryWriteNotes(dynamic slide, string text)
        {
            dynamic notesPage = slide.NotesPage;
            dynamic shapes = null;
            dynamic placeholders = null;
            try
            {
                shapes = notesPage.Shapes;
                placeholders = shapes.Placeholders;
                int count = (int)placeholders.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic ph = placeholders.Item(i);
                    try
                    {
                        if ((int)ph.PlaceholderFormat.Type == 2 && (int)ph.HasTextFrame == -1)
                        {
                            ph.TextFrame.TextRange.Text = (text ?? string.Empty).Replace("\r\n", "\r").Replace("\n", "\r");
                            return true;
                        }
                    }
                    finally { ComHelper.Release(ph); }
                }
                return false;
            }
            finally
            {
                ComHelper.Release(placeholders);
                ComHelper.Release(shapes);
                ComHelper.Release(notesPage);
            }
        }
    }
}
