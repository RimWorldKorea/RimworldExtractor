using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore.Extractor;

/// <summary>
/// 여기 있는 것들 대부분 찢어서 정리할 예정
/// </summary>
public static partial class ExtractorEngine
{
    //TODO 이거 프레임워크 모드 지원하려고 있는건가?
    private static IEnumerable<TranslationEntry> FindExtractableNodesXmlExtensionSettings(string defName,
        string className,
        XElement rootNode, string? curNormalizedPath = null)
    {
        var extractableTagsXmlExtensionSettings = new[] { "label", "text", "tooltip" };
        // (CurrentNode, CurrentPath)
        var q = new Queue<(XElement, string)>();
        if (curNormalizedPath == null)
        {
            foreach (var node in rootNode.Elements())
            {
                q.Enqueue((node, node.IsListNode() ? GetIdxOfListNode(node).ToString() : node.Name.LocalName));
            }
        }
        else
        {
            q.Enqueue((rootNode, curNormalizedPath + "." + (rootNode.IsListNode()
                ? GetIdxOfListNode(rootNode).ToString()
                : rootNode.Name.LocalName)));
        }

        while (q.Count > 0)
        {
            var (curNode, curPath) = q.Dequeue();
            var token = curPath.Split('.');
            var lastTag = token[^1];

            if (curNode.IsTextNode())
            {
                var isListNode = token.Length > 1 && int.TryParse(lastTag, out _) &&
                                 extractableTagsXmlExtensionSettings.Contains(token[^2]);

                if (extractableTagsXmlExtensionSettings.Contains(lastTag) || isListNode)
                {
                    var tKey = curNode.Parent?.Element("tKey")?.Value;
                    var tKeyTip = curNode.Parent?.Element("tKeyTip")?.Value;
                    if (lastTag is "label" or "text" && tKey != null)
                    {
                        yield return new TranslationEntry("Keyed", tKey, curNode.Value, null, null, null);
                    }
                    else if (lastTag == "tooltip" && tKeyTip != null)
                    {
                        yield return new TranslationEntry("Keyed", tKeyTip, curNode.Value, null, null, null);
                    }
                    else
                    {
                        yield return new TranslationEntry(className, $"{defName}.{curPath}", curNode.Value, null, null,
                            null);
                    }
                }

                continue;
            }

            foreach (var childNode in curNode.Elements())
            {
                string path;
                if (childNode.IsListNode())
                {
                    path = $"{curPath}.{GetIdxOfListNode(childNode)}";
                }
                else
                    path = $"{curPath}.{childNode.Name.LocalName}";

                q.Enqueue((childNode, path));
            }
        }
    }
    
    //TODO 이거 림월드 기본 동작이랑 다시 비교해보기 (그냥 패키지에서 바로 가져올까?)
    /// <summary>
    /// 번역 핸들을 생성합니다.
    /// </summary>
    /// <param name="handle"></param>
    /// <returns></returns>
    private static string NormalizedHandle(string handle)
    {
        if (string.IsNullOrEmpty(handle))
        {
            return handle;
        }

        handle = handle.Trim();
        handle = handle.Replace(' ', '_');
        handle = handle.Replace('\n', '_');
        handle = handle.Replace("\r", "");
        handle = handle.Replace('\t', '_');
        handle = handle.Replace(".", "");

        if (handle.IndexOf('-') >= 0)
        {
            handle = handle.Replace('-'.ToString(), "");
        }

        if (handle.IndexOf('{') >= 0)
        {
            handle = new Regex("{.*?}").Replace(handle, "");
        }

        var sb = new StringBuilder();
        for (int i = 0; i < handle.Length; i++)
        {
            if ("qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890-_".IndexOf(handle[i]) >= 0)
            {
                sb.Append(handle[i]);
            }
        }

        handle = sb.ToString();
        sb.Length = 0;
        for (int j = 0; j < handle.Length; j++)
        {
            if (j == 0 || handle[j] != '_' || handle[j - 1] != '_')
            {
                sb.Append(handle[j]);
            }
        }

        handle = sb.ToString();
        handle = handle.Trim(new char[]
        {
            '_'
        });
        if (!string.IsNullOrEmpty(handle) && handle.All(char.IsDigit))
        {
            handle = "_" + handle;
        }

        return handle;
    }

    public static XElement? GetRootDefNode(XElement node, out string? nodeName)
    {
        if (node.Element("defName") != null)
        {
            nodeName = node.Element("defName")!.Value;
            return node;
        }
        else if (node.Name.LocalName == "Defs")
        {
            nodeName = null;
            return null;
        }

        var parentNode = node;
        nodeName = node.IsListNode() ? GetIdxOfListNode(node).ToString() : node.Name.LocalName;

        do
        {
            parentNode = parentNode.Parent;
            if (parentNode == null)
                throw new InvalidOperationException("Couldn't find root Def node");

            if (parentNode.Element("defName") != null)
            {
                nodeName = $"{parentNode.Element("defName")!.Value}.{nodeName}";
                break;
            }

            nodeName =
                $"{(parentNode.IsListNode() ? GetIdxOfListNode(parentNode).ToString() : parentNode.Name.LocalName)}.{nodeName}";
        } while (true);

        return parentNode;
    }

    private static int GetIdxOfListNode(XElement curNode)
    {
        var nodes = curNode.Parent?.Elements().ToList();
        if (nodes == null)
            throw new InvalidOperationException("ParentNode was null.");

        int i;
        for (i = 0; i < nodes.Count; i++)
        {
            if (nodes[i] == curNode)
                return i;
        }

        throw new InvalidOperationException("Couldn't find idx of list node.");
    }
}