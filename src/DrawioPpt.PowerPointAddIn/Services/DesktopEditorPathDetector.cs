using System;
using System.IO;
using Microsoft.Win32;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class DesktopEditorPathDetector
    {
        public string Detect()
        {
            string path = DetectFromRegistry();
            if (!string.IsNullOrWhiteSpace(path))
            {
                return path;
            }

            string[] candidates = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "draw.io", "draw.io.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "draw.io", "drawio.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "diagrams.net", "diagrams.net.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "draw.io", "draw.io.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "draw.io", "drawio.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "diagrams.net", "diagrams.net.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "draw.io", "draw.io.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "draw.io", "drawio.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "diagrams.net", "diagrams.net.exe")
            };

            int index;
            for (index = 0; index < candidates.Length; index++)
            {
                if (File.Exists(candidates[index]))
                {
                    return candidates[index];
                }
            }

            return string.Empty;
        }

        private static string DetectFromRegistry()
        {
            string[] registryPaths = new string[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\draw.io.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\drawio.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\diagrams.net.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\draw.io.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\drawio.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\diagrams.net.exe"
            };

            int index;
            for (index = 0; index < registryPaths.Length; index++)
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryPaths[index]))
                {
                    if (key == null)
                    {
                        continue;
                    }

                    object value = key.GetValue(string.Empty);
                    string candidate = value as string;
                    if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return string.Empty;
        }
    }
}

