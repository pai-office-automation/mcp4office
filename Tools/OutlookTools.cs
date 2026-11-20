using System;
using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace Mcp4Office.Tools
{
    internal static class OutlookTools
    {
        private static readonly Regex BodyOpenTagRegex =
            new Regex(@"<body\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static void Register(McpServer server)
        {
            server.RegisterTool(
                "outlook_list_selected",
                "Legge le email attualmente selezionate in Outlook classico (entry_id, store_id, mittente, oggetto, data, corpo testuale; mai allegati). Usare entry_id/store_id per outlook_create_reply_draft.",
                McpServer.Schema((props, _) =>
                {
                    props["max_chars_per_mail"] = McpServer.IntProp("Caratteri massimi per email (1-80000, default 15000).", 1, 80000);
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ListSelected(McpServer.GetInt(args, "max_chars_per_mail", 15000))));

            server.RegisterTool(
                "outlook_list_inbox",
                "Elenca le ultime email della Posta in arrivo di Outlook classico (entry_id, store_id, oggetto, mittente, data e breve anteprima).",
                McpServer.Schema((props, _) =>
                {
                    props["count"] = McpServer.IntProp("Numero di email da elencare (1-30, default 10).", 1, 30);
                }),
                McpServer.ReadOnly(),
                args => ComHelper.RetryRead(() => ListInbox(McpServer.GetInt(args, "count", 10))));

            server.RegisterTool(
                "outlook_create_draft",
                "Crea una NUOVA BOZZA email in Outlook (testo semplice, niente Markdown) e la apre a schermo. Non invia mai automaticamente; l'utente deve sempre verificare destinatari To/CC prima di inviare.",
                McpServer.Schema((props, req) =>
                {
                    props["to"] = McpServer.Prop("string", "Destinatari 'A' separati da punto e virgola (opzionale).");
                    props["cc"] = McpServer.Prop("string", "Destinatari 'CC' (opzionale).");
                    props["subject"] = McpServer.Prop("string", "Oggetto dell'email.");
                    props["body"] = McpServer.Prop("string", "Corpo del messaggio in testo semplice.");
                    req.Add("subject");
                    req.Add("body");
                }),
                McpServer.Mutating(destructive: false),
                args => CreateDraft(
                    McpServer.GetString(args, "to"),
                    McpServer.GetString(args, "cc"),
                    McpServer.GetString(args, "subject"),
                    McpServer.GetString(args, "body")));

            server.RegisterTool(
                "outlook_create_reply_draft",
                "Crea una BOZZA DI RISPOSTA all'email esatta identificata da entry_id (e opzionalmente store_id) ottenuti da outlook_list_selected/outlook_list_inbox. Non dipende più dalla selezione corrente e non invia mai automaticamente.",
                McpServer.Schema((props, req) =>
                {
                    props["entry_id"] = McpServer.Prop("string", "entry_id dell'email a cui rispondere (da outlook_list_selected o outlook_list_inbox).");
                    props["store_id"] = McpServer.Prop("string", "store_id opzionale dell'email (consigliato se hai più caselle in Outlook).");
                    props["body"] = McpServer.Prop("string", "Testo semplice della risposta da inserire sopra il messaggio originale.");
                    props["reply_all"] = McpServer.Prop("boolean", "Se true crea 'Rispondi a tutti', altrimenti solo al mittente (default false).");
                    req.Add("entry_id");
                    req.Add("body");
                }),
                McpServer.Mutating(destructive: false),
                args => CreateReplyDraft(
                    McpServer.GetString(args, "entry_id"),
                    McpServer.GetString(args, "store_id"),
                    McpServer.GetString(args, "body"),
                    McpServer.GetBool(args, "reply_all", false)));
        }

        /// <summary>
        /// Inserisce un blocco HTML subito dopo il tag di apertura &lt;body ...&gt;, preservando &lt;html&gt;, &lt;head&gt;, firma e citazione.
        /// </summary>
        internal static string InsertAfterBody(string existingHtml, string snippetHtml)
        {
            if (string.IsNullOrEmpty(existingHtml))
                return "<html><body>" + (snippetHtml ?? string.Empty) + "</body></html>";

            var m = BodyOpenTagRegex.Match(existingHtml);
            if (m.Success)
            {
                int pos = m.Index + m.Length;
                return existingHtml.Substring(0, pos) + (snippetHtml ?? string.Empty) + existingHtml.Substring(pos);
            }
            return (snippetHtml ?? string.Empty) + existingHtml;
        }

        private static string TryGetStoreId(dynamic item)
        {
            dynamic parent = null;
            try
            {
                parent = item.Parent;
                return (string)(parent.StoreID ?? string.Empty);
            }
            catch { return string.Empty; }
            finally { ComHelper.Release(parent); }
        }

        private static string ListSelected(int maxCharsPerMail)
        {
            dynamic app = ComHelper.GetActiveApp("Outlook.Application", "Microsoft Outlook classico");
            dynamic explorer = null;
            dynamic selection = null;
            try
            {
                explorer = app.ActiveExplorer();
                if (explorer == null)
                    throw new InvalidOperationException("Nessuna finestra di Outlook attiva.");

                selection = explorer.Selection;
                int count = (int)selection.Count;
                if (count == 0)
                    return "(Nessuna email selezionata in Outlook.)";

                var sb = new StringBuilder();
                int totalCap = 80000;
                int readCount = 0;

                for (int i = 1; i <= count; i++)
                {
                    dynamic item = selection.Item(i);
                    try
                    {
                        if ((int)item.Class != 43) continue; // 43 = olMail
                        readCount++;
                        string entryId = (string)(item.EntryID ?? string.Empty);
                        string storeId = TryGetStoreId(item);
                        string sender = (string)(item.SenderName ?? "(sconosciuto)");
                        string subject = (string)(item.Subject ?? "(senza oggetto)");
                        string received = Convert.ToString(item.ReceivedTime);
                        string body = (string)(item.Body ?? string.Empty);
                        body = ComHelper.Truncate(body.Trim(), maxCharsPerMail, out _);

                        sb.AppendLine($"=== Email {readCount} ===");
                        sb.AppendLine($"entry_id: {entryId}");
                        if (!string.IsNullOrEmpty(storeId)) sb.AppendLine($"store_id: {storeId}");
                        sb.AppendLine($"Da: {sender}");
                        sb.AppendLine($"Data: {received}");
                        sb.AppendLine($"Oggetto: {subject}");
                        sb.AppendLine();
                        sb.AppendLine(body);
                        sb.AppendLine();
                    }
                    finally { ComHelper.Release(item); }

                    if (sb.Length >= totalCap) break;
                }

                if (readCount == 0)
                    return "(Gli elementi selezionati in Outlook non sono messaggi email.)";

                return ComHelper.Truncate(sb.ToString().TrimEnd(), totalCap, out _);
            }
            finally
            {
                ComHelper.Release(selection);
                ComHelper.Release(explorer);
                ComHelper.Release(app);
            }
        }

        private static string ListInbox(int count)
        {
            dynamic app = ComHelper.GetActiveApp("Outlook.Application", "Microsoft Outlook classico");
            dynamic ns = null;
            dynamic inbox = null;
            dynamic items = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                inbox = ns.GetDefaultFolder(6); // 6 = olFolderInbox
                string storeId = string.Empty;
                try { storeId = (string)(inbox.StoreID ?? string.Empty); } catch { }

                items = inbox.Items;
                items.Sort("[ReceivedTime]", true);

                int total = (int)items.Count;
                int limit = Math.Min(count, total);
                var sb = new StringBuilder();
                sb.AppendLine($"Posta in arrivo ({total} elementi totali, mostrate le ultime {limit}):");
                if (!string.IsNullOrEmpty(storeId)) sb.AppendLine($"store_id: {storeId}");

                int shown = 0;
                for (int i = 1; i <= total && shown < limit; i++)
                {
                    dynamic item = items.Item(i);
                    try
                    {
                        if ((int)item.Class != 43) continue;
                        shown++;
                        string entryId = (string)(item.EntryID ?? string.Empty);
                        string sender = (string)(item.SenderName ?? "(sconosciuto)");
                        string subject = (string)(item.Subject ?? "(senza oggetto)");
                        string received = Convert.ToString(item.ReceivedTime);
                        string body = ((string)(item.Body ?? string.Empty)).Replace("\r", " ").Replace("\n", " ").Trim();
                        if (body.Length > 140) body = body.Substring(0, 140) + "…";
                        sb.AppendLine($"{shown}. [{received}] {sender} — {subject}");
                        sb.AppendLine($"   entry_id: {entryId}");
                        if (body.Length > 0) sb.AppendLine($"   {body}");
                    }
                    finally { ComHelper.Release(item); }
                }

                return sb.ToString().TrimEnd();
            }
            finally
            {
                ComHelper.Release(items);
                ComHelper.Release(inbox);
                ComHelper.Release(ns);
                ComHelper.Release(app);
            }
        }

        private static string CreateDraft(string to, string cc, string subject, string body)
        {
            dynamic app = ComHelper.GetActiveApp("Outlook.Application", "Microsoft Outlook classico");
            dynamic mail = null;
            try
            {
                mail = app.CreateItem(0); // 0 = olMailItem
                if (!string.IsNullOrWhiteSpace(to)) mail.To = to;
                if (!string.IsNullOrWhiteSpace(cc)) mail.CC = cc;
                mail.Subject = subject ?? string.Empty;
                // Display() PRIMA di scrivere il corpo così Outlook carica la firma predefinita.
                mail.Display(false);
                string htmlIntro = "<div style=\"font-family:Segoe UI,Calibri,sans-serif;font-size:11pt;\">" +
                                   WebUtility.HtmlEncode(body ?? string.Empty).Replace("\r\n", "<br>").Replace("\n", "<br>") +
                                   "</div><br>";
                string existingHtml = (string)(mail.HTMLBody ?? string.Empty);
                mail.HTMLBody = InsertAfterBody(existingHtml, htmlIntro);
                return $"Bozza creata e aperta in Outlook: '{subject}'. Verifica sempre destinatari (A/CC) e testo prima di inviare.";
            }
            finally
            {
                ComHelper.Release(mail);
                ComHelper.Release(app);
            }
        }

        private static string CreateReplyDraft(string entryId, string storeId, string body, bool replyAll)
        {
            if (string.IsNullOrWhiteSpace(entryId))
                throw new ArgumentException("Parametro 'entry_id' obbligatorio.");

            dynamic app = ComHelper.GetActiveApp("Outlook.Application", "Microsoft Outlook classico");
            dynamic ns = null;
            dynamic item = null;
            dynamic reply = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                item = string.IsNullOrWhiteSpace(storeId)
                    ? ns.GetItemFromID(entryId)
                    : ns.GetItemFromID(entryId, storeId);

                if (item == null || (int)item.Class != 43)
                    throw new InvalidOperationException("L'elemento indicato da entry_id non esiste o non è un'email.");

                reply = replyAll ? item.ReplyAll() : item.Reply();
                string htmlIntro = "<div style=\"font-family:Segoe UI,Calibri,sans-serif;font-size:11pt;\">" +
                                   WebUtility.HtmlEncode(body ?? string.Empty).Replace("\r\n", "<br>").Replace("\n", "<br>") +
                                   "</div><br>";
                string existingHtml = (string)(reply.HTMLBody ?? string.Empty);
                reply.HTMLBody = InsertAfterBody(existingHtml, htmlIntro);
                reply.Display(false);

                string subj = (string)(reply.Subject ?? "");
                string origSender = (string)(item.SenderName ?? "");
                return $"Bozza di risposta ({(replyAll ? "Rispondi a tutti" : "Rispondi")}) aperta in Outlook per '{origSender}' — '{subj}'. (Non inviata.)";
            }
            finally
            {
                ComHelper.Release(reply);
                ComHelper.Release(item);
                ComHelper.Release(ns);
                ComHelper.Release(app);
            }
        }
    }
}
