using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using DocumentFormat.OpenXml.Office.Word;
using RimworldExtractorInternal.DataTypes;
using Crc32 = System.IO.Hashing.Crc32;

namespace RimworldExtractorInternal
{
    public static partial class Utils
    {
        public static XmlElement Append(this XmlElement parent, Action<XmlElement> work)
        {
            work(parent);
            return parent;
        }

        public static XmlElement AppendElement(this XmlNode parent, string name, string? innerText = null)
        {
            var child = (XmlElement?)parent.AppendChild(
                (parent.NodeType == XmlNodeType.Document ? (XmlDocument)parent : parent.OwnerDocument!)
                .CreateElement(name)) ?? throw new NullReferenceException();
            if (innerText != null)
            {
                child.InnerText = innerText;
            }

            return child;
        }
        public static XmlElement AppendElement(this XmlElement parent, string name, string? innerText = null)
        {
            var child = (XmlElement?)parent.AppendChild(parent.OwnerDocument.CreateElement(name)) ??
                        throw new NullReferenceException();

            if (innerText != null)
            {
                child.InnerText = innerText;
            }
            return child;
        }

        public static XmlElement AppendElement(this XmlNode parent, string name, Action<XmlElement> work)
        {
            var child = parent.AppendElement(name);
            work(child);
            return child;
        }

        public static XmlAttribute? AppendAttribute(this XmlNode parent, string name, string? value)
        {
            if (parent is XmlElement e)
                return e.AppendAttribute(name, value);
            else
                return null;
        }
        public static XmlAttribute AppendAttribute(this XmlElement parent, string name, string? value)
        {
            var attr = parent.Attributes.Append(parent.OwnerDocument.CreateAttribute(name));
            if (value != null)
            {
                attr.Value = value;
            }
            return attr;
        }

        public static XmlComment AppendComment(this XmlElement parent, string comment)
        {
            var child = (XmlComment)parent.AppendChild(parent.OwnerDocument.CreateComment(comment))!;
            return child;
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

        public static IEnumerable<XmlNode> Where(this XmlNodeList nodes, Predicate<XmlNode> predicate)
        {
            return nodes.OfType<XmlNode>().Where(x => predicate(x));
        }

        public static IEnumerable<T> Select<T>(this XmlNodeList nodes, Func<XmlNode, T> selector)
        {
            return nodes.OfType<XmlNode>().Select(selector);
        }

        public static XmlNode? FirstOrDefault(this XmlNodeList nodes, Predicate<XmlNode> predicate)
        {
            return nodes.OfType<XmlNode>().FirstOrDefault(x => predicate(x));
        }

        public static bool HasSameElements<T>(this IEnumerable<T> node1, IEnumerable<T>? node2)
        {
            if (node2 == null)
                return false;
            var node1Array = node1 as T[] ?? node1.ToArray();
            var node2Array = node2 as T[] ?? node2.ToArray();
            return !node1Array.Except(node2Array).Any() && !node2Array.Except(node1Array).Any();
        }

        public static bool HasAttribute(this XmlNode node, string attributeName)
        {
            return node.Attributes?[attributeName] != null;
        }

        public static bool HasAttribute(this XmlNode node, string attributeName, string value)
        {
            return node.Attributes?[attributeName]?.Value == value;
        }

        public static bool TryGetAttritube(this XmlNode node, string attritubeName, out string? value)
        {
            value = node.Attributes?[attritubeName]?.Value;
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

        internal static bool IsListNode(this XmlNode? curNode) => curNode?.Name == "li";

        internal static bool IsTextNode(this XmlNode? curNode) =>
            curNode?.ChildNodes.Count == 1 && (curNode.FirstChild!.NodeType == XmlNodeType.Text ||
                                               curNode.FirstChild!.NodeType == XmlNodeType.CDATA);

        public static XmlNodeList? SelectNodesSafe(this XmlDocument? doc, string? xpath)
        {
            if (doc == null || xpath == null) return null;
            try
            {
                return doc.SelectNodes(xpath);
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

        /** 긴 파일명을 줄이고 중복을 피하기 위해 사용
         *  
         */
        public class StringShortener
        {
            private const string Base36Chars = "0123456789abcdefghijklmnopqrstuvwxyz";

            public static UInt32 GetCRC32Hash(string input)
            {
                Span<byte> inputBytes = Encoding.UTF8.GetBytes(input);
                
                Span<byte> hashBytes = new byte[4];
                Crc32.TryHash(inputBytes, hashBytes, out _); // TryHash는 리틀 엔디언을 반환
                
                return BinaryPrimitives.ReadUInt32LittleEndian(hashBytes);
            }

            public static string ToBase36(UInt32 input, int length)
            {
                Span<char> result = stackalloc char[length];
                for (int i = length - 1; i >= 0; i--)
                {
                    result[i] = Base36Chars[(int)(input % 36)];
                    input /= 36;
                }
                
                return new string(result);
            }
            
            public static string Run(string input, int length = 7)
            {
                return ToBase36(GetCRC32Hash(input), length);
            }
        }
    }
}
