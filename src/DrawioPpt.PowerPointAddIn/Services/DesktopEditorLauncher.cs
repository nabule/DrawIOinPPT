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
        private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);
        private readonly PresentationSidecarPathBuilder _pathBuilder;

        public DesktopEditorLauncher()
            : this(new PresentationSidecarPathBuilder())
        {
        }

        public DesktopEditorLauncher(PresentationSidecarPathBuilder pathBuilder)
        {
            _pathBuilder = pathBuilder;
        }

        public string ResolvePreferredSidecarPath(DiagramEnvelope envelope, PluginSettings settings, string presentationPath)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException("envelope");
            }

            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            if (settings.KeepSidecarFile && _pathBuilder.CanBuildPath(presentationPath))
            {
                return _pathBuilder.BuildPath(presentationPath, envelope.DiagramId, envelope.DiagramName, settings.SidecarFolderName);
            }

            string currentPath = envelope.SidecarPath;
            if (!string.IsNullOrWhiteSpace(currentPath) && Path.IsPathRooted(currentPath))
            {
                return currentPath;
            }

            return string.Empty;
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

            string targetPath = ResolvePreferredSidecarPath(envelope, settings, presentationPath);
            if (string.IsNullOrWhiteSpace(targetPath))
            {
                targetPath = BuildTemporaryPath(envelope.DiagramId);
            }

            string directory = Path.GetDirectoryName(targetPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(targetPath, envelope.DrawioXml ?? string.Empty, Utf8WithoutBom);
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

            string fullDrawioFilePath = Path.GetFullPath(drawioFilePath);
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = editorExecutablePath;
            startInfo.Arguments = Quote(fullDrawioFilePath);
            startInfo.WorkingDirectory = Path.GetDirectoryName(fullDrawioFilePath);
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
