using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace Mcp4Office
{
    internal static class ComHelper
    {
        private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
        private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);
        public const int MaxAllowedChars = 200000;

        public static dynamic GetActiveApp(string progId, string friendlyName)
        {
            try
            {
                return Marshal.GetActiveObject(progId);
            }
            catch (COMException)
            {
                throw new InvalidOperationException(
                    $"{friendlyName} non risulta aperto (o è in esecuzione con privilegi diversi, es. Amministratore). Aprine una finestra come utente normale prima di usare questo comando.");
            }
        }

        /// <summary>
        /// Ritenta solo operazioni di LETTURA su errori COM transitori ("applicazione occupata").
        /// Non usare mai per mutazioni.
        /// </summary>
        public static T RetryRead<T>(Func<T> action, int maxAttempts = 3)
        {
            int delayMs = 250;
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    return action();
                }
                catch (COMException ex) when (IsTransientComError(ex))
                {
                    if (attempt >= maxAttempts)
                    {
                        throw new InvalidOperationException(
                            "L'applicazione Office è occupata (es. cella in modifica con F2 o finestra di dialogo aperta): premi Invio o Esc in Office e riprova.",
                            ex);
                    }
                    Thread.Sleep(delayMs);
                    delayMs *= 2;
                }
            }
        }

        public static bool IsTransientComError(COMException ex)
        {
            return ex != null && (ex.HResult == RPC_E_CALL_REJECTED || ex.HResult == RPC_E_SERVERCALL_RETRYLATER);
        }

        public static void Release(object comObj)
        {
            if (comObj == null) return;
            try
            {
                if (Marshal.IsComObject(comObj))
                    Marshal.ReleaseComObject(comObj);
            }
            catch
            {
                // Ignora errori di rilascio durante la chiusura.
            }
        }

        public static void CollectComGarbage()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public static string Truncate(string text, int maxChars, out bool truncated)
        {
            int effectiveMax = Clamp(maxChars, 1, MaxAllowedChars);
            if (string.IsNullOrEmpty(text))
            {
                truncated = false;
                return string.Empty;
            }
            if (text.Length > effectiveMax)
            {
                truncated = true;
                return text.Substring(0, effectiveMax) +
                       $"\n\n[... troncato a {effectiveMax} caratteri su {text.Length} totali ...]";
            }
            truncated = false;
            return text;
        }

        public static int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        public static string Sha256Short(string input)
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(input ?? string.Empty));
                var sb = new StringBuilder(16);
                for (int i = 0; i < 8; i++)
                    sb.Append(bytes[i].ToString("x2"));
                return sb.ToString();
            }
        }
    }
}
