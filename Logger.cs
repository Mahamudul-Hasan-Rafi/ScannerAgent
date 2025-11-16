using System;
using System.IO;
using System.Text;
using System.Threading;

static class Logger
{
    static readonly object Sync = new object();
    static StreamWriter writer;
    static string logFilePath;

    public static void Init(string directory = null, string baseFileName = "ScannerAgent", bool append = true)
    {
        try
        {
            if (directory == null)
            {
                directory = AppDomain.CurrentDomain.BaseDirectory;
            }

            Directory.CreateDirectory(directory);

            // Daily file name with timestamp
            string fileName = $"{baseFileName}-{DateTime.Now:yyyyMMdd}.log";
            logFilePath = Path.Combine(directory, fileName);

            var fs = new FileStream(logFilePath, append ? FileMode.Append : FileMode.Create, FileAccess.Write, FileShare.Read);
            writer = new StreamWriter(fs, Encoding.UTF8) { AutoFlush = true };
            WriteInternal("Logger initialized");
        }
        catch
        {
            // Swallow initialization errors to avoid crashing the app on startup
        }
    }

    static void WriteInternal(string level, string message = null)
    {
        if (writer == null) return;
        try
        {
            lock (Sync)
            {
                string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                int threadId = Thread.CurrentThread.ManagedThreadId;
                writer.WriteLine($"{time} [{level}] (T{threadId}) {message}");
            }
        }
        catch { }
    }

    public static void Info(string message) => WriteInternal("INFO", message);
    public static void Warn(string message) => WriteInternal("WARN", message);
    public static void Error(string message) => WriteInternal("ERROR", message);
    public static void Exception(Exception ex) => WriteInternal("ERROR", ex == null ? "null" : $"{ex.Message}\n{ex.StackTrace}");
    public static void Close()
    {
        try
        {
            lock (Sync)
            {
                writer?.Flush();
                writer?.Close();
                writer = null;
            }
        }
        catch { }
    }
}

