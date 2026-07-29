using System;
using System.Diagnostics;
using System.IO;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class DesktopSvgExporter
    {
        private const int ProcessTimeoutMilliseconds = 60000;

        public bool TryExport(string editorExecutablePath, string drawioFilePath, string outputSvgPath)
        {
            if (string.IsNullOrWhiteSpace(editorExecutablePath) ||
                string.IsNullOrWhiteSpace(drawioFilePath) ||
                string.IsNullOrWhiteSpace(outputSvgPath) ||
                !File.Exists(editorExecutablePath) ||
                !File.Exists(drawioFilePath))
            {
                return false;
            }

            string directory = Path.GetDirectoryName(outputSvgPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string[] argumentsToTry = new string[]
            {
                string.Format("-x -f svg -e --embed-svg-images --svg-theme light -o {0} {1}", Quote(outputSvgPath), Quote(drawioFilePath)),
                string.Format("--export --format svg --embed-diagram --embed-svg-images --svg-theme light --output {0} {1}", Quote(outputSvgPath), Quote(drawioFilePath))
            };

            int index;
            for (index = 0; index < argumentsToTry.Length; index++)
            {
                if (RunProcess(editorExecutablePath, argumentsToTry[index]) && File.Exists(outputSvgPath))
                {
                    return true;
                }
            }

            return false;
        }

        public bool TryExportPng(
            string editorExecutablePath,
            string drawioFilePath,
            string outputPngPath,
            int width)
        {
            if (string.IsNullOrWhiteSpace(editorExecutablePath) ||
                string.IsNullOrWhiteSpace(drawioFilePath) ||
                string.IsNullOrWhiteSpace(outputPngPath) ||
                width <= 0 ||
                !File.Exists(editorExecutablePath) ||
                !File.Exists(drawioFilePath))
            {
                return false;
            }

            try
            {
                string directory = Path.GetDirectoryName(outputPngPath);
                if (!string.IsNullOrEmpty(directory) &&
                    !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (File.Exists(outputPngPath))
                {
                    File.Delete(outputPngPath);
                }

                string arguments = string.Format(
                    "--export --format png --page-index 1 --border 0 --width {0} --output {1} {2}",
                    width,
                    Quote(outputPngPath),
                    Quote(drawioFilePath));
                return RunProcess(editorExecutablePath, arguments) &&
                    File.Exists(outputPngPath) &&
                    new FileInfo(outputPngPath).Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool RunProcess(string fileName, string arguments)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = fileName;
                startInfo.Arguments = arguments;
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;

                using (Process process = Process.Start(startInfo))
                {
                    if (process == null)
                    {
                        return false;
                    }

                    if (!process.WaitForExit(
                            ProcessTimeoutMilliseconds))
                    {
                        try
                        {
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                        catch
                        {
                        }

                        return false;
                    }

                    return process.ExitCode == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }
    }
}
