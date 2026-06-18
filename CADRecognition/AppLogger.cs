using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace CADRecognition
{
    /// <summary>
    /// 应用程序日志服务，支持按天分文件的日志记录。
    /// </summary>
    internal sealed class AppLogger : IDisposable
    {
        private static readonly Lazy<AppLogger> _instance = new(() => new AppLogger());
        public static AppLogger Instance => _instance.Value;

        private readonly string _logDirectory;
        private readonly object _lock = new();
        private StreamWriter? _writer;
        private string _currentDate = string.Empty;
        private bool _disposed;
        private bool _enableDebug = false; // 默认关闭 DEBUG 日志，减少噪音

        private AppLogger()
        {
            _logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CADRecognition",
                "Logs");

            Directory.CreateDirectory(_logDirectory);
            InitializeWriter();
        }

        private void InitializeWriter()
        {
            var today = DateTime.Now.ToString("yyyy-MM-dd");
            var logFile = Path.Combine(_logDirectory, $"CADRecognition_{today}.log");

            lock (_lock)
            {
                if (_writer != null)
                {
                    if (_currentDate == today && _writer.BaseStream.CanWrite)
                    {
                        return;
                    }
                    _writer.Flush();
                    _writer.Dispose();
                    _writer = null;
                }

                _currentDate = today;
                _writer = new StreamWriter(logFile, append: true, encoding: Encoding.UTF8)
                {
                    AutoFlush = true
                };
            }
        }

        /// <summary>
        /// 启动时调用，清除非今日的所有日志文件。
        /// </summary>
        public void CleanOldLogs()
        {
            var today = DateTime.Now.ToString("yyyy-MM-dd");
            try
            {
                var files = Directory.GetFiles(_logDirectory, "CADRecognition_*.log");
                foreach (var file in files)
                {
                    var fileName = Path.GetFileNameWithoutExtension(file);
                    var datePart = fileName.Replace("CADRecognition_", "");
                    if (!string.Equals(datePart, today, StringComparison.Ordinal))
                    {
                        File.Delete(file);
                    }
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// 记录信息日志。
        /// </summary>
        public void Info(string message) => WriteLog("INFO", message);

        /// <summary>
        /// 记录警告日志。
        /// </summary>
        public void Warn(string message) => WriteLog("WARN", message);

        /// <summary>
        /// 记录错误日志。
        /// </summary>
        public void Error(string message, Exception? ex = null)
        {
            var sb = new StringBuilder(message);
            if (ex != null)
            {
                sb.AppendLine();
                sb.Append("  异常: ").Append(ex.GetType().Name).Append(": ").AppendLine(ex.Message);
                if (ex.StackTrace != null)
                {
                    sb.AppendLine("  堆栈:");
                    foreach (var line in ex.StackTrace.Split('\n'))
                    {
                        var trimmed = line.Trim();
                        if (!string.IsNullOrEmpty(trimmed))
                        {
                            sb.AppendLine("    ").Append(trimmed);
                        }
                    }
                }
                if (ex.InnerException != null)
                {
                    sb.Append("  内部异常: ").Append(ex.InnerException.GetType().Name)
                      .Append(": ").AppendLine(ex.InnerException.Message);
                }
            }
            WriteLog("ERROR", sb.ToString());
        }

        /// <summary>
        /// 记录调试日志（默认关闭，可通过 EnableDebug 开启）。
        /// </summary>
        public void Debug(string message)
        {
            if (_enableDebug)
            {
                WriteLog("DEBUG", message);
            }
        }

        /// <summary>
        /// 开启或关闭调试日志。
        /// </summary>
        public void EnableDebug(bool enable) => _enableDebug = enable;

        /// <summary>
        /// 记录操作节点日志，带详细描述。
        /// </summary>
        public void LogOperation(string operation, string description)
        {
            var sb = new StringBuilder();
            sb.AppendLine("========================================");
            sb.Append("【操作】").AppendLine(operation);
            sb.Append("【描述】").AppendLine(description);
            sb.Append("【时间】").AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            sb.AppendLine("========================================");
            WriteLogRaw(sb.ToString());
        }

        /// <summary>
        /// 记录状态变化。
        /// </summary>
        public void LogStatus(string status) => WriteLog("STATUS", status);

        /// <summary>
        /// 记录 Modbus 操作。
        /// </summary>
        public void LogModbus(string action, string register, string? result = null, string? error = null)
        {
            var sb = new StringBuilder();
            sb.Append($"Modbus[{action}] {register}");
            if (!string.IsNullOrEmpty(result))
            {
                sb.Append($" => {result}");
            }
            if (!string.IsNullOrEmpty(error))
            {
                sb.Append($" [ERROR: {error}]");
            }
            WriteLog("MODBUS", sb.ToString());
        }

        private void WriteLog(string level, string message)
        {
            EnsureWriter();
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            var line = $"[{timestamp}] [{level,-5}] {message}";
            lock (_lock)
            {
                _writer?.WriteLine(line);
            }
        }

        private void WriteLogRaw(string message)
        {
            EnsureWriter();
            lock (_lock)
            {
                _writer?.WriteLine(message);
            }
        }

        private void EnsureWriter()
        {
            var today = DateTime.Now.ToString("yyyy-MM-dd");
            if (_currentDate != today || _writer == null)
            {
                InitializeWriter();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            lock (_lock)
            {
                _writer?.Flush();
                _writer?.Dispose();
                _writer = null;
            }
        }
    }
}
