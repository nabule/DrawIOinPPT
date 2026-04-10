using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using DrawioPpt.Core.Models;
using DrawioPpt.Core.Services;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class DesktopEditorLauncher
    {
        private readonly PresentationSidecarPathBuilder _pathBuilder;

        public DesktopEditorLauncher()
            : this(new PresentationSidecarPathBuilder())
        {
        }

        public DesktopEditorLauncher(PresentationSidecarPathBuilder pathBuilder)
        {
            _pathBuilder = pathBuilder;
        }

        public string PrepareWorkingFile(DiagramEnvelope envelope, PluginSettings settings, string presentationPath)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }

            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            string targetPath = envelope.SidecarPath;
            if (string.IsNullOrWhiteSpace(targetPath))
            {
                if (settings.KeepSidecarFile && !string.IsNullOrWhiteSpace(presentationPath))
                {
                    targetPath = _pathBuilder.BuildPath(presentationPath, envelope.DiagramId, envelope.DiagramName, settings.SidecarFolderName);
                }
                else
                {
                    targetPath = BuildTemporaryPath(envelope.DiagramId);
                }
            }

            string directory = Path.GetDirectoryName(targetPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(targetPath, envelope.DrawioXml ?? string.Empty, Encoding.UTF8);
            return targetPath;
        }

        public void Launch(string editorExecutablePath, string drawioFilePath)
        {
            if (string.IsNullOrWhiteSpace(editorExecutablePath))
            {
                throw new ArgumentException("Editor executable path is required.", "editorExecutablePath");
            }

            if (string.IsNullOrWhiteSpace(drawioFilePath))
            {
                throw new ArgumentException("Draw.io file path is required.", "drawioFilePath");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = editorExecutablePath;
            startInfo.Arguments = Quote(drawioFilePath);
            startInfo.UseShellExecute = true;
            Process.Start(startInfo);
        }

        private static string BuildTemporaryPath(string diagramId)
        {
            string baseDirectory = Path.Combine(Path.GetTempPath(), "DrawioPpt");
            return Path.Combine(baseDirectory, diagramId + ".drawio");
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }
    }
}

