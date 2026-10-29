using System.Collections.Concurrent;
using System.Runtime.CompilerServices;

namespace RimExtractorCore
{
    public static class Log
    {
        public static TextWriter Out { private get; set; } = Console.Out;

        public static IEnumerable<string> Messages => _logQueue;
        
        private static readonly ConcurrentQueue<string> _logQueue = new();

        private const int MAX_COUNT = 10000;
        
        private static readonly ConcurrentDictionary<string, string> _classNameCache = new();

        /// <summary>기본 공통 로그 메서드입니다.</summary>
        /// <param name="message">출력할 메시지입니다.</param>
        /// <param name="colorHexOrName">출력 메시지 색상입니다. 자체적으로 색을 넣도록 되어있진 않고, 실제 출력 모듈이 참고할 태그로 가공됩니다.</param>
        /// <param name="hideTrace">로그에 자동으로 주입되는 호출자 정보를 숨깁니다.</param>
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
        
        /// <summary>빨간색 로그입니다.</summary>
        public static void Err(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "") 
            => Default(message, "Tomato", hideTrace, filePath, memberName);
        
        /// <summary>노란색 로그입니다.</summary>
        public static void Wrn(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "") 
            => Default(message, "Gold", hideTrace, filePath, memberName);
        
        /// <summary>하얀색 로그입니다.</summary>
        public static void Msg(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
            => Default(message, "White", hideTrace, filePath, memberName );
        
        /// <summary>초록색 로그입니다.</summary>
        public static void Good(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
            => Default(message, "LimeGreen", hideTrace, filePath, memberName );
        
        /// <summary>연속된 중복 로그 요청을 무시하는 Err입니다.</summary>
        public static void ErrOnce(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
        {
            if (message.Equals(_logQueue.LastOrDefault())) return;
            Err(message, hideTrace, filePath, memberName);
        }
        
        /// <summary>연속된 중복 로그 요청을 무시하는 Wrn입니다.</summary>
        public static void WrnOnce(string message, bool hideTrace = false,
            [CallerFilePath] string filePath = "", [CallerMemberName] string memberName = "")
        {
            if (message.Equals(_logQueue.LastOrDefault())) return;
            Wrn(message, hideTrace, filePath, memberName);
        }
        
        private static void StoreMessage(string message)
        {
            if (_logQueue.Count >= MAX_COUNT)
            {
                if(!_logQueue.TryDequeue(out _))
                    throw new InvalidOperationException();
            }
            _logQueue.Enqueue(message);
        }
        
        /// <summary>마지막 로그의 내용을 message로 대체합니다.</summary>
        /// <remarks>대체 메시지는 로그 큐에는 저장되지 않습니다.</remarks>
        public static void ReplaceLast(string message)
        {
            var str = $"[REPLACE]{message}"; // [REPLACE] 접두를 달아서 보내면 Face 모듈이 알아서 합니다.
            Out.WriteLine(str);
        }
        
        /// <summary>작업을 실행하고 마지막 로그의 꼬리에 진행 중 효과를 비동기적으로 실행합니다.</summary>
        /// <param name="work">효과를 나타낼 작업입니다.</param>
        /// <param name="dotLength">따라붙는 점의 최대 개수입니다.</param>
        /// <param name="millisecondsInterval">점이 늘어나는 시간 간격입니다.</param>
        /// <remarks>진행 중간에 다른 로그가 끼어드는 경우에는 대비되지 않았습니다.</remarks>
        public static void RunWithCyclingDot(Action work, int dotLength = 19, int millisecondsInterval = 30)
        {
            // _logQueue에 방금 막 추가된 마지막 문자열(색상 태그 포함)을 가져옵니다.
            string baseMessage = _logQueue.LastOrDefault() ?? "";
            
            using var cts = new CancellationTokenSource();
            
            var progressTask = Task.Run(async () =>
            {
                string[] dotPack = new string[dotLength + 1];
                
                dotPack[0] = string.Empty;
                for (int i = 1; i < dotPack.Length; i++)
                    dotPack[i] = dotPack[i - 1] + ".";
                
                int j = 0;
                while (!cts.Token.IsCancellationRequested)
                {
                    ReplaceLast($"{baseMessage}{dotPack[j % dotPack.Length]}");
                    j++;
                    try { await Task.Delay(millisecondsInterval, cts.Token); } catch { break; }
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
                
                // 작업이 끝나면 점이 없는 원래 텍스트 상태로 되돌립니다.
                ReplaceLast(baseMessage);
            }
        }
    }
}
