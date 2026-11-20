using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Mcp4Office.Tools
{
    internal static class PdfTools
    {
        private const long MaxPdfBytes = 50L * 1024 * 1024; // 50 MB
        private const int MaxPagesPerCall = 100;

        private const uint IO_REPARSE_TAG_MOUNT_POINT = 0xA0000003;
        private const uint IO_REPARSE_TAG_SYMLINK = 0xA000000C;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WIN32_FIND_DATA
        {
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            public uint dwReserved0; // Reparse tag quando FILE_ATTRIBUTE_REPARSE_POINT è impostato
            public uint dwReserved1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
            public string cAlternateFileName;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr FindFirstFile(string lpFileName, out WIN32_FIND_DATA lpFindFileData);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FindClose(IntPtr hFindFile);

        private static readonly IntPtr InvalidHandleValue = new IntPtr(-1);

        private static List<string> _allowedRoots = BuildDefaultRoots();

        public static void ConfigureAllowedRoots(IEnumerable<string> customRoots, bool isExplicit = false)
        {
            var list = new List<string>();
            if (customRoots != null)
            {
                foreach (var r in customRoots)
                {
                    if (string.IsNullOrWhiteSpace(r)) continue;
                    if (r.StartsWith(@"\\") || r.StartsWith("//"))
                    {
                        if (isExplicit)
                            throw new ArgumentException($"Radice PDF di rete/UNC non consentita: '{r}'.");
                        continue;
                    }
                    string full;
                    try
                    {
                        full = Path.GetFullPath(r).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                    }
                    catch (Exception ex)
                    {
                        if (isExplicit) throw new ArgumentException($"Radice PDF non valida '{r}': {ex.Message}");
                        continue;
                    }
                    if (full.StartsWith(@"\\"))
                    {
                        if (isExplicit) throw new ArgumentException($"Radice PDF UNC non consentita: '{r}'.");
                        continue;
                    }
                    list.Add(full);
                }
            }

            if (isExplicit && list.Count == 0)
                throw new ArgumentException("Nessuna radice PDF valida specificata in --pdf-roots / MCP4OFFICE_PDF_ROOTS.");

            _allowedRoots = list.Count > 0 ? list : BuildDefaultRoots();
        }

        private static List<string> BuildDefaultRoots()
        {
            var roots = new List<string>();
            void Add(string dir)
            {
                if (string.IsNullOrWhiteSpace(dir)) return;
                try
                {
                    string full = Path.GetFullPath(dir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                    if (!roots.Any(x => string.Equals(x, full, StringComparison.OrdinalIgnoreCase)))
                        roots.Add(full);
                }
                catch { }
            }

            Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrWhiteSpace(userProfile))
                Add(Path.Combine(userProfile, "Downloads"));
            return roots;
        }

        public static void Register(McpServer server)
        {
            server.RegisterTool(
                "pdf_info",
                "Restituisce metadati e numero di pagine di un file PDF nelle cartelle consentite (--pdf-roots / Documenti / Desktop / Download). Funziona anche se aperto in Acrobat Reader.",
                McpServer.Schema((props, req) =>
                {
                    props["path"] = McpServer.Prop("string", "Percorso locale completo del file .pdf (no percorsi di rete UNC).");
                    req.Add("path");
                }),
                McpServer.ReadOnly(),
                args => GetPdfInfo(McpServer.GetString(args, "path")));

            server.RegisterTool(
                "pdf_read_file",
                "Estrae il testo con spaziatura e ordine di lettura (ContentOrderTextExtractor) di un file PDF nelle cartelle consentite (max 100 pagine per chiamata).",
                McpServer.Schema((props, req) =>
                {
                    props["path"] = McpServer.Prop("string", "Percorso locale completo del file .pdf (no percorsi di rete UNC).");
                    props["from_page"] = McpServer.IntProp("Prima pagina da leggere (1-based, default 1).", 1, 10000);
                    props["to_page"] = McpServer.IntProp("Ultima pagina da leggere (opzionale; max 100 pagine per chiamata).", 1, 10000);
                    props["max_chars"] = McpServer.IntProp("Lunghezza massima in caratteri (1-200000, default 80000).", 1, ComHelper.MaxAllowedChars);
                    req.Add("path");
                }),
                McpServer.ReadOnly(),
                args => ReadPdfFile(
                    McpServer.GetString(args, "path"),
                    McpServer.GetInt(args, "from_page", 1),
                    McpServer.GetInt(args, "to_page", 0),
                    McpServer.GetInt(args, "max_chars", 80000)));

            server.RegisterTool(
                "acrobat_read_active",
                "Se Adobe Acrobat Pro è aperto, rileva il PDF attivo via COM e ne estrae il testo (soggetto alle stesse restrizioni di cartella/dimensione di pdf_read_file). Non supportato su Acrobat Reader gratuito.",
                McpServer.Schema((props, _) =>
                {
                    props["from_page"] = McpServer.IntProp("Prima pagina (1-based, default 1).", 1, 10000);
                    props["to_page"] = McpServer.IntProp("Ultima pagina (opzionale; max 100 pagine per chiamata).", 1, 10000);
                    props["max_chars"] = McpServer.IntProp("Lunghezza massima in caratteri (1-200000, default 80000).", 1, ComHelper.MaxAllowedChars);
                }),
                McpServer.ReadOnly(),
                args => ReadActiveAcrobat(
                    McpServer.GetInt(args, "from_page", 1),
                    McpServer.GetInt(args, "to_page", 0),
                    McpServer.GetInt(args, "max_chars", 80000)));
        }

        internal static string ValidatePdfPath(string path)
        {
            return ValidateLocalFilePath(path, new[] { ".pdf" }, MaxPdfBytes, "PDF");
        }

        internal static string ValidateLocalFilePath(string path, string[] allowedExtensions, long maxBytes, string kindLabel)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("Parametro 'path' obbligatorio.");

            string trimmed = path.Trim();
            if (trimmed.StartsWith(@"\\") || trimmed.StartsWith("//"))
                throw new UnauthorizedAccessException("Percorsi di rete UNC o device path non sono ammessi per sicurezza.");

            string fullPath;
            try
            {
                fullPath = Path.GetFullPath(trimmed);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Percorso non valido: {ex.Message}");
            }

            if (fullPath.StartsWith(@"\\"))
                throw new UnauthorizedAccessException("Percorsi UNC non ammessi.");

            string ext = Path.GetExtension(fullPath);
            if (!allowedExtensions.Any(a => string.Equals(ext, a, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException($"Estensione non ammessa per {kindLabel} ('{ext}'). Ammesse: {string.Join(", ", allowedExtensions)}.");

            bool inAllowedRoot = _allowedRoots.Any(root =>
                fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase));
            if (!inAllowedRoot)
            {
                throw new UnauthorizedAccessException(
                    $"Percorso fuori dalle cartelle locali consentite ({string.Join("; ", _allowedRoots)}). " +
                    "Avvia mcp4office con --pdf-roots \"C:\\tua_cartella\" o imposta MCP4OFFICE_PDF_ROOTS per autorizzarla.");
            }

            // Controlla symlink/junction PRIMA di File.Exists lungo tutti i segmenti del percorso.
            CheckNoSymlinkOrJunction(fullPath);

            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"File {kindLabel} non trovato: {fullPath}", fullPath);

            var fi = new FileInfo(fullPath);
            if (fi.Length > maxBytes)
                throw new InvalidOperationException($"File {kindLabel} troppo grande ({fi.Length / (1024 * 1024)} MB > limite {maxBytes / (1024 * 1024)} MB).");

            return fullPath;
        }

        private static void CheckNoSymlinkOrJunction(string fullPath)
        {
            // Risale dalla radice fino al file: se un segmento esiste ed è symlink o junction, rifiuta.
            // Consente i reparse tag cloud di OneDrive ("File su richiesta").
            string root = Path.GetPathRoot(fullPath);
            string rel = fullPath.Substring(root.Length);
            string[] parts = rel.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

            string current = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            foreach (var part in parts)
            {
                current = current + Path.DirectorySeparatorChar + part;
                IntPtr h = FindFirstFile(current, out var data);
                if (h == InvalidHandleValue)
                    break; // segmento inesistente: verrà gestito da File.Exists
                FindClose(h);

                bool isReparse = (data.dwFileAttributes & (uint)FileAttributes.ReparsePoint) != 0;
                if (isReparse && IsDangerousReparseTag(data.dwReserved0))
                {
                    throw new UnauthorizedAccessException(
                        $"Rilevato collegamento simbolico o junction in '{current}': lettura rifiutata per sicurezza.");
                }
            }
        }

        internal static bool IsDangerousReparseTag(uint tag)
        {
            // Rifiuta symlink e junction; consente tag cloud OneDrive (0x9000xxxx / 0x8000xxxx).
            return tag == IO_REPARSE_TAG_SYMLINK || tag == IO_REPARSE_TAG_MOUNT_POINT;
        }

        private static string GetPdfInfo(string path)
        {
            string fullPath = ValidatePdfPath(path);
            var fi = new FileInfo(fullPath);
            using (var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var doc = PdfDocument.Open(fs))
            {
                var sb = new StringBuilder();
                sb.AppendLine($"File: {fi.Name}");
                sb.AppendLine($"Percorso: {fi.FullName}");
                sb.AppendLine($"Dimensione: {fi.Length / 1024.0:F1} KB");
                sb.AppendLine($"Pagine: {doc.NumberOfPages}");
                sb.AppendLine($"Versione PDF: {doc.Version}");
                if (doc.Information != null)
                {
                    if (!string.IsNullOrWhiteSpace(doc.Information.Title)) sb.AppendLine($"Titolo: {doc.Information.Title}");
                    if (!string.IsNullOrWhiteSpace(doc.Information.Author)) sb.AppendLine($"Autore: {doc.Information.Author}");
                    if (!string.IsNullOrWhiteSpace(doc.Information.Subject)) sb.AppendLine($"Oggetto: {doc.Information.Subject}");
                }
                return sb.ToString().TrimEnd();
            }
        }

        private static string ReadPdfFile(string path, int fromPage, int toPage, int maxChars)
        {
            string fullPath = ValidatePdfPath(path);
            using (var fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var doc = PdfDocument.Open(fs))
            {
                int total = doc.NumberOfPages;
                if (total == 0)
                    return "(Il PDF non contiene pagine.)";

                int start = Math.Max(1, fromPage);
                if (start > total)
                    throw new ArgumentOutOfRangeException(nameof(fromPage), $"Il PDF ha solo {total} pagine.");

                int requestedEnd = toPage > 0 ? Math.Min(total, toPage) : total;
                if (requestedEnd < start)
                    throw new ArgumentException($"Intervallo pagine non valido: from_page ({start}) > to_page ({requestedEnd}).");

                int end = Math.Min(requestedEnd, start + MaxPagesPerCall - 1);
                bool pagesCapped = end < requestedEnd;

                var sb = new StringBuilder();
                sb.AppendLine($"[PDF: {Path.GetFileName(fullPath)} | Pagine {start}-{end} di {total}]");
                if (pagesCapped)
                    sb.AppendLine($"(Limite {MaxPagesPerCall} pagine per chiamata: continua da from_page={end + 1})");

                for (int p = start; p <= end; p++)
                {
                    var page = doc.GetPage(p);
                    string extracted = ContentOrderTextExtractor.GetText(page) ?? string.Empty;
                    sb.AppendLine($"\n--- Pagina {p} ---");
                    sb.AppendLine(extracted);
                    if (sb.Length >= maxChars + 500) break;
                }

                string result = sb.ToString().TrimEnd();
                if (result.Length <= 80 && total > 0)
                    result += "\n\n(Nota: testo estratto quasi vuoto; potrebbe essere un PDF scansionato come immagine senza OCR.)";

                return ComHelper.Truncate(result, maxChars, out _);
            }
        }

        private static string ReadActiveAcrobat(int fromPage, int toPage, int maxChars)
        {
            dynamic acroApp = null;
            dynamic avDoc = null;
            dynamic pdDoc = null;
            dynamic js = null;
            try
            {
                try
                {
                    acroApp = Marshal.GetActiveObject("AcroExch.App");
                }
                catch (COMException)
                {
                    throw new InvalidOperationException(
                        "Adobe Acrobat Pro non risulta aperto (o stai usando Acrobat Reader gratuito, che non espone COM). " +
                        "Usa 'pdf_read_file' passando il percorso locale del file PDF.");
                }

                avDoc = acroApp.GetActiveDoc();
                if (avDoc == null)
                    throw new InvalidOperationException("Adobe Acrobat Pro è aperto ma non c'è alcun documento attivo.");

                pdDoc = avDoc.GetPDDoc();
                string fileName = (string)pdDoc.GetFileName();

                string fullPath = null;
                try
                {
                    js = pdDoc.GetJSObject();
                    string jsPath = (string)js.path;
                    if (!string.IsNullOrEmpty(jsPath) && jsPath.StartsWith("/") && jsPath.Length > 3 && jsPath[2] == '/')
                        fullPath = jsPath[1] + ":" + jsPath.Substring(2).Replace('/', '\\');
                }
                catch { }

                if (!string.IsNullOrEmpty(fullPath))
                    return ReadPdfFile(fullPath, fromPage, toPage, maxChars);

                throw new InvalidOperationException(
                    $"Documento attivo in Acrobat rilevato ('{fileName}'), ma percorso locale non determinabile. Usa 'pdf_read_file'.");
            }
            finally
            {
                ComHelper.Release(js);
                ComHelper.Release(pdDoc);
                ComHelper.Release(avDoc);
                ComHelper.Release(acroApp);
            }
        }
    }
}
