using System.Diagnostics;
using System.Collections.Concurrent;
using System.Runtime.CompilerServices;

namespace RimExtractorCore
{
    public static class Log
    {
        public static TextWriter Out { private get; set; } = Console.Out;

        public static IEnumerable<string> Messages => _logQueue;
        
        private static readonly ConcurrentQueue<string> _logQueue = new();
        private static readonly HashSet<int> _hashes = new();
        private const int MAX_COUNT = 10000;
        
        private static readonly ConcurrentDictionary<string, string> _classNameCache = new();

        /// <summary>기본 공통 로그 메서드입니다.</summary>
        /// <param name="filePath">호출자의 파일 경로입니다. 컴파일 시점에 자동으로 결정됩니다.</param>
        /// <param name="memberName">호출자의 이름입니다. 컴파일 시점에 자동으로 결정됩니다.</param>
        public static void Default(string message, string colorHexOrName = "White", bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
        {
            string className = _classNameCache.GetOrAdd(filePath, path => Path.GetFileNameWithoutExtension(path));
            
            // [COLOR:색상값] 접두를 붙이면 LogPrinter가 색깔로 넘겨줍니다.

            string str = $"[COLOR:{colorHexOrName}]";
            if (!hideTrace) str += $"{className}\\{memberName}::";
            str += $"{message}";

            Out.WriteLine(str);
            StoreMessage(str);
        }
        
        public static void Err(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "") 
            => Default(message, "Tomato", hideTrace, filePath, memberName);
        
        public static void Wrn(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "") 
            => Default(message, "Gold", hideTrace, filePath, memberName);
        
        public static void Msg(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
            => Default(message, "White", hideTrace, filePath, memberName );
        
        public static void Good(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
            => Default(message, "LimeGreen", hideTrace, filePath, memberName );
        
        public static void ErrOnce(string message, int hash,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
        {
            if (_hashes.Contains(hash)) return;
            _hashes.Add(hash);
            Err(message, false, filePath, memberName);
        }
        
        public static void WrnOnce(string message, int hash,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
        {
            if (_hashes.Contains(hash)) return;
            _hashes.Add(hash);
            Wrn(message, false, filePath, memberName);
        }
        
        private static void StoreMessage(string message)
        {
            if (_logQueue.Count > MAX_COUNT)
            {
                _logQueue.TryDequeue(out _);
            }
            _logQueue.Enqueue(message);
        }
        
        /// <summary>마지막 로그의 내용을 message로 대체합니다.</summary>
        public static void ReplaceLast(string message)
        {
            var str = $"[REPLACE]{message}"; // [REPLACE] 접두를 달아서 보내면 Face 모듈이 알아서 합니다.
            Out.WriteLine(str);
        }
        
        /// <summary>
        /// 꼬리에 . .. ... .... 효과를 비동기적으로 나타내는 로그를 특정 작업과 함께 띄웁니다.
        /// </summary>
        /// <param name="baseMessage"></param>
        /// <param name="work"></param>
        /// <param name="colorHexOrName"></param>
        /// <param name="filePath"></param>
        /// <param name="memberName"></param>
        public static void TaskWithCyclingDot(string baseMessage, string colorHexOrName, bool hideTrace, Action work,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
        {
            string className = _classNameCache.GetOrAdd(filePath, path => Path.GetFileNameWithoutExtension(path));
            string prefix = $"{className}\\{memberName}::";
            
            // 🟢 처음 한 번만 색상 태그와 호출자 정보를 포함해 [REPLACE] 텍스트를 구성합니다.
            Out.WriteLine($"[REPLACE][COLOR:{colorHexOrName}]{prefix}{baseMessage}");

            using var cts = new CancellationTokenSource();
            
            var progressTask = Task.Run(async () =>
            {
                string[] dots = { ".", "..", "..." };
                int i = 0;
                while (!cts.Token.IsCancellationRequested)
                {
                    // 이후부터는 색상 태그 없이 텍스트만 업데이트
                    ReplaceLast($"{prefix}{baseMessage}{dots[i % 3]}");
                    i++;
                    try { await Task.Delay(500, cts.Token); } catch { break; }
                }
            });

            try
            {
                work(); 
            }
            finally
            {
                cts.Cancel();
                progressTask.Wait();
            }
        }
        
        /*
        private static string GetCallStack()
        {
            var curMethod = new StackTrace().GetFrame(2)?.GetMethod();
            if (curMethod?.DeclaringType == typeof(Log))
            {
                curMethod = new StackTrace().GetFrame(3)?.GetMethod();
            }
            return $"{curMethod?.DeclaringType?.Name ?? "UNKNOWN Type"}.{curMethod?.Name ?? "UNKNOWN Method"}()";
        }
        */
    }
}
