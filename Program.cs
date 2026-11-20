using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Mcp4Office
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            Console.InputEncoding = utf8NoBom;
            Console.OutputEncoding = utf8NoBom;

            bool readOnly = false;
            bool listTools = false;
            bool selfTest = false;
            bool configureClients = false;
            bool unconfigureClients = false;
            string reportPath = null;
            string testRoot = null;
            var pdfRoots = new List<string>();

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (string.Equals(a, "--read-only", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(a, "--sola-lettura", StringComparison.OrdinalIgnoreCase))
                    readOnly = true;
                else if (string.Equals(a, "--list-tools", StringComparison.OrdinalIgnoreCase))
                    listTools = true;
                else if (string.Equals(a, "--self-test", StringComparison.OrdinalIgnoreCase))
                    selfTest = true;
                else if (string.Equals(a, "--configure-clients", StringComparison.OrdinalIgnoreCase))
                    configureClients = true;
                else if (string.Equals(a, "--unconfigure-clients", StringComparison.OrdinalIgnoreCase))
                    unconfigureClients = true;
                else if (string.Equals(a, "--report", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                    reportPath = args[++i];
                else if (string.Equals(a, "--test-root", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                    testRoot = args[++i];
                else if (string.Equals(a, "--pdf-roots", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    foreach (var p in args[++i].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                        pdfRoots.Add(p.Trim());
                }
            }

            if (configureClients || unconfigureClients)
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                return ClientConfigurator.Run(configureClients, exePath, reportPath, testRoot, Console.Out);
            }

            bool explicitPdfRoots = pdfRoots.Count > 0;
            if (!explicitPdfRoots)
            {
                string envRoots = Environment.GetEnvironmentVariable("MCP4OFFICE_PDF_ROOTS");
                if (!string.IsNullOrWhiteSpace(envRoots))
                {
                    explicitPdfRoots = true;
                    foreach (var p in envRoots.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                        pdfRoots.Add(p.Trim());
                }
            }

            try
            {
                Tools.PdfTools.ConfigureAllowedRoots(pdfRoots, explicitPdfRoots);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[mcp4office] Configurazione --pdf-roots non valida: " + ex.Message);
                return 2;
            }

            if (selfTest)
                return SelfTests.RunAll(Console.Out);

            var server = new McpServer(readOnly);
            RegisterAllTools(server);

            if (listTools)
            {
                Console.WriteLine(server.DumpToolsJson());
                return 0;
            }

            try
            {
                server.RunStdio(Console.In, Console.Out, Console.Error);
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[mcp4office] Errore fatale: " + ex);
                return 1;
            }
        }

        internal static void RegisterAllTools(McpServer server)
        {
            Tools.WordTools.Register(server);
            Tools.ExcelTools.Register(server);
            Tools.PowerPointTools.Register(server);
            Tools.OutlookTools.Register(server);
            Tools.PdfTools.Register(server);
        }
    }
}
