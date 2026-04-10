using System;
using System.IO;
using System.Xml.Linq;
using DrawioPpt.Core.Contracts;
using DrawioPpt.Core.Models;

namespace DrawioPpt.Core.Services
{
    public class FilePluginSettingsStore : IPluginSettingsStore
    {
        private readonly string _settingsPath;

        public FilePluginSettingsStore()
            : this(GetDefaultSettingsPath())
        {
        }

        public FilePluginSettingsStore(string settingsPath)
        {
            _settingsPath = settingsPath;
        }

        public PluginSettings Load()
        {
            if (!File.Exists(_settingsPath))
            {
                return new PluginSettings();
            }

            XDocument document = XDocument.Load(_settingsPath);
            XElement root = document.Root;
            if (root == null)
            {
                return new PluginSettings();
            }

            PluginSettings settings = new PluginSettings();
            settings.DesktopEditorPath = Read(root, "desktopEditorPath", settings.DesktopEditorPath);
            settings.EditorUrl = Read(root, "editorUrl", settings.EditorUrl);
            settings.AutoOpenOnSelection = Read(root, "autoOpenOnSelection", settings.AutoOpenOnSelection);
            settings.AutoUpdateOnSave = Read(root, "autoUpdateOnSave", settings.AutoUpdateOnSave);
            settings.KeepSidecarFile = Read(root, "keepSidecarFile", settings.KeepSidecarFile);
            settings.SidecarFolderName = Read(root, "sidecarFolderName", settings.SidecarFolderName);

            string editorModeText = Read(root, "editorMode", settings.EditorMode.ToString());
            EditorMode editorMode;
            if (Enum.TryParse(editorModeText, true, out editorMode))
            {
                settings.EditorMode = editorMode;
            }

            return settings;
        }

        public void Save(PluginSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            string directory = Path.GetDirectoryName(_settingsPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            XDocument document = new XDocument(
                new XElement(
                    "pluginSettings",
                    new XElement("editorMode", settings.EditorMode.ToString()),
                    new XElement("desktopEditorPath", settings.DesktopEditorPath ?? string.Empty),
                    new XElement("editorUrl", settings.EditorUrl ?? string.Empty),
                    new XElement("autoOpenOnSelection", settings.AutoOpenOnSelection),
                    new XElement("autoUpdateOnSave", settings.AutoUpdateOnSave),
                    new XElement("keepSidecarFile", settings.KeepSidecarFile),
                    new XElement("sidecarFolderName", settings.SidecarFolderName ?? string.Empty)
                )
            );

            document.Save(_settingsPath);
        }

        public static string GetDefaultSettingsPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, KnownMetadata.SettingsFolderName, KnownMetadata.SettingsFileName);
        }

        private static string Read(XElement root, string name, string defaultValue)
        {
            XElement element = root.Element(name);
            if (element == null)
            {
                return defaultValue;
            }

            return element.Value;
        }

        private static bool Read(XElement root, string name, bool defaultValue)
        {
            XElement element = root.Element(name);
            if (element == null)
            {
                return defaultValue;
            }

            bool value;
            if (!bool.TryParse(element.Value, out value))
            {
                return defaultValue;
            }

            return value;
        }
    }
}

