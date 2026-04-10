using System;
using System.Diagnostics;
using System.IO;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class DesktopSvgExporter
    {
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
                string.Format("--export --format svg --output {0} {1}", Quote(outputSvgPath), Quote(drawioFilePath)),
                string.Format("-x -f svg -o {0} {1}", Quote(outputSvgPath), Quote(drawioFilePath))
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

        private static bool RunProcess(string fileName, string arguments)
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

                process.WaitForExit(60000);
                return process.ExitCode == 0;
            }
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }
    }
}

