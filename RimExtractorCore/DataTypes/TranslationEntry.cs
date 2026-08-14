namespace RimExtractorCore.DataTypes;

/// <summary>
/// 번역 데이터를 다루는 매개 타입입니다. 기본적으로 XML LanguageData 구문에 일대일 대응합니다.
/// </summary>
/// <param name="ClassName">번역 데이터의 종류 (*Def, Keyed, Strings)</param>
/// <param name="Node">바인딩 위치</param>
/// <param name="Original">원문</param>
/// <param name="Translated">번역문</param>
/// <param name="RequiredMods">요구 모드</param>
public record TranslationEntry(
    string ClassName,
    string Node,
    string Original,
    string? Translated,
    RequiredMods? RequiredMods,
    string? SourceFile)
{
    public bool MayNotTranslate { get; init; } = false;
    public bool FullListTranslate { get; init; } = false;

    public TranslationEntry(TranslationEntry other)
    {
        ClassName = other.ClassName;
        Node = other.Node;
        Original = other.Original;
        Translated = other.Translated;
        MayNotTranslate = other.MayNotTranslate;
        FullListTranslate = other.FullListTranslate;

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
    public string DefName => Node.Contains('.') ? Node[..Node.IndexOf('.')] : Node;
    public string RealNode => Node.Contains('.') ? Node[(Node.IndexOf('.') + 1)..] : Node;

    /// <summary>
    /// 현재 Entry를 Grid 타입으로 변환합니다.
    /// </summary>
    public List<string> ToGridRow()
    {
        var noticeTags = new List<string>();
        if (MayNotTranslate) noticeTags.Add(Constants.AttrMayNotTranslate);
        if (FullListTranslate) noticeTags.Add(Constants.AttrTranslationCanChangeCount);
        
        return new List<string>
        {
            $"{ClassName}+{Node}",
            ClassName,
            Node,
            RequiredMods?.ToString() ?? string.Empty,
            string.Join(", ", noticeTags),
            Original,
            Translated ?? string.Empty
        };
    }
}
