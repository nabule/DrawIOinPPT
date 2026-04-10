using DrawioPpt.Core.Models;

namespace DrawioPpt.Core.Contracts
{
    public interface IPluginSettingsStore
    {
        PluginSettings Load();
        void Save(PluginSettings settings);
    }
}

