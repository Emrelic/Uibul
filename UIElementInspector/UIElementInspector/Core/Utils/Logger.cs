using System;
using System.IO;
using System.Text;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Simple file-based logger for application operations
    /// </summary>
    public class Logger : IDisposable
    {
        private readonly string _logFilePath;
        private readonly object _lockObject = new object();
        private StreamWriter _writer;

        public Logger()
        {
            // Create Logs directory in application folder
            var appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var logsDirectory = Path.Combine(appDirectory, "Logs");

            if (!Directory.Exists(logsDirectory))
            {
                Directory.CreateDirectory(logsDirectory);
            }

            // Create log file with timestamp
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            var fileName = $"UIInspector_{timestamp}.log";
            _logFilePath = Path.Combine(logsDirectory, fileName);

            // Initialize writer
            _writer = new StreamWriter(_logFilePath, append: true, Encoding.UTF8)
            {
                AutoFlush = true
            };

            // Write header
            WriteHeader();
        }

        private void WriteHeader()
        {
            _writer.WriteLine("================================================================================");
            _writer.WriteLine($"Universal UI Element Inspector - Log Started");
            _writer.WriteLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _writer.WriteLine($"Version: 1.0.0");
            _writer.WriteLine($"OS: {Environment.OSVersion}");
            _writer.WriteLine($".NET: {Environment.Version}");
            _writer.WriteLine("================================================================================");
            _writer.WriteLine();
        }

        public void Log(string message, LogLevel level = LogLevel.Info)
        {
            lock (_lockObject)
            {
                try
                {
                    var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
                    var levelStr = level.ToString().ToUpper().PadRight(7);
                    var logEntry = $"[{timestamp}] [{levelStr}] {message}";

                    _writer?.WriteLine(logEntry);
                }
                catch (Exception ex)
                {
                    // Silently fail - don't let logging crash the app
                    System.Diagnostics.Debug.WriteLine($"Logging error: {ex.Message}");
                }
            }
        }

        public void LogInfo(string message) => Log(message, LogLevel.Info);
        public void LogWarning(string message) => Log(message, LogLevel.Warning);
        public void LogError(string message) => Log(message, LogLevel.Error);
        public void LogDebug(string message) => Log(message, LogLevel.Debug);

        public void LogException(Exception ex, string context = "")
        {
            lock (_lockObject)
            {
                try
                {
                    _writer?.WriteLine();
                    _writer?.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    _writer?.WriteLine($"EXCEPTION: {context}");
                    _writer?.WriteLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                    _writer?.WriteLine($"Type: {ex.GetType().Name}");
                    _writer?.WriteLine($"Message: {ex.Message}");
                    _writer?.WriteLine($"Stack Trace:");
                    _writer?.WriteLine(ex.StackTrace);

                    if (ex.InnerException != null)
                    {
                        _writer?.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                        _writer?.WriteLine($"Inner Stack Trace:");
                        _writer?.WriteLine(ex.InnerException.StackTrace);
                    }

                    _writer?.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
                    _writer?.WriteLine();
                }
                catch
                {
                    // Silently fail
                }
            }
        }

        public void LogSection(string sectionName)
        {
            lock (_lockObject)
            {
                try
                {
                    _writer?.WriteLine();
                    _writer?.WriteLine($"═══════════════════════════════════════════════════════════════════════════════");
                    _writer?.WriteLine($"  {sectionName}");
                    _writer?.WriteLine($"═══════════════════════════════════════════════════════════════════════════════");
                    _writer?.WriteLine();
                }
                catch
                {
                    // Silently fail
                }
            }
        }

        public string GetLogFilePath() => _logFilePath;

        public void Dispose()
        {
            lock (_lockObject)
            {
                try
                {
                    _writer?.WriteLine();
                    _writer?.WriteLine("================================================================================");
                    _writer?.WriteLine($"Log Ended: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    _writer?.WriteLine("================================================================================");
                    _writer?.Flush();
                    _writer?.Dispose();
                    _writer = null;
                }
                catch
                {
                    // Silently fail
                }
            }
        }
    }

    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }
}
