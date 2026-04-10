using System;
using System.IO;

namespace DrawioPpt.Core.Services
{
    public class PresentationSidecarPathBuilder
    {
        public string BuildPath(string presentationPath, string diagramId, string diagramName, string sidecarFolderName)
        {
            if (string.IsNullOrWhiteSpace(presentationPath))
            {
                throw new ArgumentException("Presentation path is required.", "presentationPath");
            }

            if (string.IsNullOrWhiteSpace(diagramId))
            {
                throw new ArgumentException("Diagram id is required.", "diagramId");
            }

            string presentationDirectory = Path.GetDirectoryName(presentationPath);
            string presentationBaseName = Path.GetFileNameWithoutExtension(presentationPath);
            string folderName = string.IsNullOrWhiteSpace(sidecarFolderName) ? "drawio" : sidecarFolderName;
            string targetDirectory = Path.Combine(presentationDirectory, folderName);

            string safeName = MakeSafeFileName(diagramName);
            if (string.IsNullOrWhiteSpace(safeName))
            {
                safeName = diagramId;
            }

            return Path.Combine(targetDirectory, string.Format("{0}.{1}.drawio", presentationBaseName, safeName));
        }

        private static string MakeSafeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            char[] invalidChars = Path.GetInvalidFileNameChars();
            char[] buffer = value.ToCharArray();
            int i;
            for (i = 0; i < buffer.Length; i++)
            {
                if (Array.IndexOf(invalidChars, buffer[i]) >= 0)
                {
                    buffer[i] = '_';
                }
            }

            return new string(buffer).Trim();
        }
    }
}

