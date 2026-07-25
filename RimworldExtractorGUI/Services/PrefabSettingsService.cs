using RimworldExtractorInternal.Core;

namespace RimworldExtractorGUI.Services;

public interface IPrefabSettingsService
{
    void Load();
    void Save();
    void Reset();
    string AutoDetectVersion();
}

public class PrefabSettingsService : IPrefabSettingsService
{
    public void Load()
    {
        if (File.Exists("Prefabs.dat"))
        {
            Prefabs.Load();
        }
        else
        {
            Prefabs.Init();
            Prefabs.Save();
        }
    }

    public void Save()
    {
        Prefabs.Save();
    }

    public void Reset()
    {
        Prefabs.Init();
        Prefabs.Save();
    }

    public string AutoDetectVersion()
    {
        return Prefabs.AutoDetectRimworldVersion();
    }
}