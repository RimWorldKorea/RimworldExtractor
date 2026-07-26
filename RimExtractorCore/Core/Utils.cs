using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;
using ClosedXML.Excel;
using RimExtractorCore.DataTypes;

namespace RimExtractorCore
{
    public static partial class Utils
    {
        /// <param name="ModName">"함수 이름 - 창작마당 일련번호" 형식으로 된 문자열</param>
        /// <param name="TypeName">DefInjected 폴더 아래의 ThingDef 등 하위 클래스 폴더 이름</param>
        public static string GenerateFileName(string ModName, string TypeName)
        {
#if DEBUG
            Log.Msg("[입력 번수] ModName: \"" + ModName + "\" TypeName: \"" + TypeName + "\"");
#endif
            return ToBase36(GetDeterministicHash(ModName, TypeName));
        }
        
        /// <param name="ModName">"함수 이름 - 창작마당 일련번호" 형식으로 된 문자열</param>
        /// <param name="TypeName">DefInjected 폴더 아래의 ThingDef 등 하위 클래스 폴더 이름</param>
        /// <param name="DefName">ModName과 TypeName만으로 파일 이름을 결정할 수 없을 때 사용합니다. 해당 파일의 첫번째 노드 이름을 추천합니다.</param>
        public static string GenerateFileName(string ModName, string TypeName, string DefName)
        {
#if DEBUG
            Log.Msg("[입력 번수] ModName: \"" + ModName + "\" TypeName: \"" + TypeName + "\" DefName: \"" + DefName +"\"");
#endif
            return ToBase36(GetDeterministicHash(ModName, TypeName, DefName));
        }

        public static uint GetDeterministicHash(params string[] inputs)
        {
            string inputCombined = string.Join("::", inputs);
            byte[] inputBytes = Encoding.UTF8.GetBytes(inputCombined);
            
            // Deterministic Hashing
            byte[] hashBytes = SHA256.HashData(inputBytes);
            
            return BitConverter.ToUInt32(hashBytes, 0);
        }
        
        public static string ToBase36(uint value)
        {
            const string chars = "0123456789abcdefghijklmnopqrstuvwxyz";
            const int resultLength = 6;
            uint charsLength = (uint)chars.Length;
            
            value = value % 2176782336u;
            char[] result = new char[resultLength];
            
            for (int i = resultLength - 1; i >= 0; i--)
            {
                result[i] = chars[(int)(value % charsLength)];
                value /= charsLength;
            }
            
            return new string(result);
        }

        public static XElement Append(this XElement parent, Action<XElement> work)
        {
            work(parent);
            return parent;
        }

        public static XElement AppendElement(this XContainer parent, string name, string? innerText = null)
        {
            var child = new XElement(name);
            if (innerText != null)
            {
                child.Value = innerText;
            }
            parent.Add(child);
            return child;
        }

        public static XElement AppendElement(this XContainer parent, string name, Action<XElement> work)
        {
            var child = parent.AppendElement(name);
            work(child);
            return child;
        }

        public static XAttribute? AppendAttribute(this XElement parent, string name, string? value)
        {
            parent.SetAttributeValue(name, value);
            return parent.Attribute(name);
        }

        public static void AppendComment(this XContainer parent, string comment)
        {
            parent.Add(new XComment(comment));
        }

        public static List<T> Combine<T>(this IEnumerable<T>? first, IEnumerable<T>? second)
        {
            var newList = new List<T>();
            if (first != null)
            {
                newList.AddRange(first);
            }

            if (second != null)
            {
                newList.AddRange(second);
            }
            return newList;
        }

        public static bool HasSameElements<T>(this IEnumerable<T> node1, IEnumerable<T>? node2)
        {
            if (node2 == null)
                return false;
            var node1Array = node1 as T[] ?? node1.ToArray();
            var node2Array = node2 as T[] ?? node2.ToArray();
            return !node1Array.Except(node2Array).Any() && !node2Array.Except(node1Array).Any();
        }

        public static bool HasAttribute(this XElement node, string attributeName)
        {
            return node.Attribute(attributeName) != null;
        }

        public static bool HasAttribute(this XElement node, string attributeName, string value)
        {
            return node.Attribute(attributeName)?.Value == value;
        }

        public static bool TryGetAttritube(this XElement node, string attritubeName, out string? value)
        {
            value = node.Attribute(attritubeName)?.Value;
            return value != null;
        }

        public static string GetXpath(string className, string nodeName)
        {
            var defName = nodeName.Split('.')[0];
            var tokens = nodeName[(defName.Length + 1)..].Split('.');
            for (int i = 0; i < tokens.Length; i++)
            {
                // 리스트 노드일 경우
                if (int.TryParse(tokens[i], out var k))
                {
                    tokens[i] = $"li[{k + 1}]";
                }
                // TranslationHandle을 사용한 경우
                else if (!char.IsLower(tokens[i][0]))
                {
                    tokens[i] = $"*[.//*[contains(text(), '{tokens[i]}')]]";
                }
            }

            nodeName = $"/Defs/{className}[defName=\"{defName}\"]/";
            nodeName += string.Join('/', tokens);
            return nodeName;
        }

        public static string StrVal(this IXLCell cell)
        {
            try
            {
                var value = cell.Value;
                if (value.TryGetText(out string str))
                    return str;
            }
            catch (Exception e)
            {
                Log.Msg($"엑셀 파일 속 텍스트를 읽는 중 에러 발생: {cell.Address}-{e.Message}");
            }
            return string.Empty;
        }

        public static bool IsListNode(this XElement? curNode) => curNode?.Name.LocalName == "li";

        public static bool IsTextNode(this XElement? curNode) =>
            curNode != null && !curNode.HasElements;

        public static IEnumerable<XElement>? SelectNodesSafe(this XContainer? doc, string? xpath)
        {
            if (doc == null || xpath == null) return null;
            try
            {
                return doc.XPathSelectElements(xpath);
            }
            catch (Exception e)
            {
                Log.Err(e.Message);
            }
            return null;
        }

        public static string StripInvaildChars(this string str)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                str = str.Replace(c.ToString(), "");
            }
            return StripSpace().Replace(str.Trim(), " ");
        }

        public static (int cntDefs, int cntKeyed, int cntStrings, int cntPatches) Count(
            this IEnumerable<TranslationEntry> entries)
        {
            int cntDefs = 0, cntKeyed = 0, cntStrings = 0, cntPatches = 0;
            foreach (var entry in entries)
            {
                if (entry.ClassName.StartsWith("Keyed"))
                    ++cntKeyed;
                else if (entry.ClassName.StartsWith("Strings"))
                    ++cntStrings;
                else if (entry.ClassName.StartsWith("Patches"))
                    ++cntPatches;
                else
                    ++cntDefs;
            }
            return (cntDefs, cntKeyed, cntStrings, cntPatches);
        }

        [GeneratedRegex("\\s+")]
        private static partial Regex StripSpace();

        /// <summary>
        /// RimWorld Mod Korean용 빌드 파일을 생성합니다.
        /// </summary>
        public static string WriteBuildYamlText(ModMetadata ModInfo)
        {
            return
                $"BuildRule:\n  Binding:\n    PackageID: [\"{ModInfo.PackageId}\"]\n    Mode: \"None\"\n    Dependency: \"Independent\"\n  Order:\n    After: \n    Before: \n  Version:\n    Default: \"{ConfigManager.Current.CurrentVersion}\"\n    LeftBoundary: \n    RightBoundary: \n    Designate: \n    Ban: \nMetadata:\n  WorkshopID: \"{ModInfo.Id}\"\n  ModName: \"{ModInfo.ModName}\"\n";
        }
    }
}