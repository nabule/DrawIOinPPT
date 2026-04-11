using System;
using System.IO;
using System.Text;
using DrawioPpt.Core;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class PluginTraceLog
    {
        private const long MaxLogFileBytes = 1024 * 1024;
        private static readonly object SyncRoot = new object();
        private readonly string _logPath;

        public PluginTraceLog()
            : this(GetDefaultLogPath())
        {
        }

        public PluginTraceLog(string logPath)
        {
            _logPath = logPath ?? string.Empty;
        }

        public string LogPath
        {
            get { return _logPath; }
        }

        public void Info(string scope, string message)
        {
            Write("INFO", scope, message, null);
        }

        public void Warn(string scope, string message)
        {
            Write("WARN", scope, message, null);
        }

        public void Error(string scope, string message, Exception exception)
        {
            Write("ERROR", scope, message, exception);
        }

        private void Write(string level, string scope, string message, Exception exception)
        {
            if (string.IsNullOrWhiteSpace(_logPath))
            {
                return;
            }

            try
            {
                lock (SyncRoot)
                {
                    string directory = Path.GetDirectoryName(_logPath);
                    if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    RotateIfNeeded();

                    StringBuilder builder = new StringBuilder();
                    builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    builder.Append(" [");
                    builder.Append(level ?? "INFO");
                    builder.Append("] ");
                    builder.Append(string.IsNullOrWhiteSpace(scope) ? "General" : scope);
                    builder.Append(" - ");
                    builder.AppendLine(message ?? string.Empty);

                    if (exception != null)
                    {
                        builder.AppendLine(exception.ToString());
                    }

                    File.AppendAllText(_logPath, builder.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
            }
        }

        private void RotateIfNeeded()
        {
            try
            {
                if (!File.Exists(_logPath))
                {
                    return;
                }

                FileInfo info = new FileInfo(_logPath);
                if (info.Length < MaxLogFileBytes)
                {
                    return;
                }

                string backupPath = _logPath + ".1";
                if (File.Exists(backupPath))
                {
                    File.Delete(backupPath);
                }

                File.Move(_logPath, backupPath);
            }
            catch
            {
            }
        }

        private static string GetDefaultLogPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, KnownMetadata.SettingsFolderName, "Logs", "drawioppt.log");
        }
    }
}
