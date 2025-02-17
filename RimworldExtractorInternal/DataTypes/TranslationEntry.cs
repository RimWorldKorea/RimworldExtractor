﻿using System.Diagnostics;

namespace RimworldExtractorInternal.DataTypes
{
    /// <summary>
    /// 번역 데이터
    /// </summary>
    /// <param name="ClassName">번역 데이터의 종류. ○○Def, Keyed, Strings, Patches.○○Def</param>
    /// <param name="Node">위치</param>
    /// <param name="Original">원문</param>
    /// <param name="Translated">번역문</param>
    /// <param name="RequiredMods">요구 모드</param>
    public record TranslationEntry(string ClassName, string Node, string Original, string? Translated,
        RequiredMods? RequiredMods, string? SourceFile)
    {
        public TranslationEntry(TranslationEntry other)
        {
            ClassName = other.ClassName;
            Node = other.Node;
            Original = other.Original;
            Translated = other.Translated;
            if (other.RequiredMods != null)
            {
                this.RequiredMods = new RequiredMods(other.RequiredMods);
            }
            SourceFile = other.SourceFile;

            _extensions = new Dictionary<string, object>();
            foreach (var otherExtension in other._extensions)
            {
                _extensions.Add(otherExtension.Key, otherExtension.Value);
            }
        }

        private readonly Dictionary<string, object> _extensions = new();
        public bool TryGetExtension(string key, out object? extension)
        {
            extension = null;
            if (_extensions.TryGetValue(key, out extension) == true)
            {
                return true;
            }

            return false;
        }

        public bool HasRequiredMods()
        {
            return RequiredMods == null || RequiredMods.CountAllowed > 0 || RequiredMods.CountDisallowed > 0;
        }

        public TranslationEntry AddExtension(string key, object extension)
        {
            _extensions.Add(key, extension);
            return this;
        }

        public string ClassNode => $"{ClassName}+{Node}";

    }
}
