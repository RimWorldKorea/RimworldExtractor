using System.Text;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.Threading;

namespace RimExtractorFace.Utils;

/// <summary>로그창 출력기 클래스입니다. Core의 Log 클래스에 등록해서 사용합니다.</summary>
public class LogPrinter : TextWriter
{
    private readonly SelectableTextBlock _textBlock;
    private readonly ScrollViewer _scrollViewer;
    private readonly StreamWriter _logFileWriter;
    private readonly object _lock = new();
    
    /// <summary>마지막 로그의 캐시입니다.</summary>
    private Run? _lastReplaceableRun;
    
    public override Encoding Encoding => Encoding.UTF8;

    public LogPrinter(SelectableTextBlock textBlock, ScrollViewer scrollViewer)
    {
        _textBlock = textBlock;
        _scrollViewer = scrollViewer;
        _logFileWriter = File.CreateText("log.txt");
    }

    public override void WriteLine(string? value)
    {
        if (value == null) return;

        lock (_lock)
        {
            _logFileWriter.WriteLine(value);
            _logFileWriter.Flush();

            var line = value + Environment.NewLine;
            
            // UI 스레드 안전 호출
            Dispatcher.UIThread.Post(() => AppendToTextBlock(line, value));
        }
    }

    private void AppendToTextBlock(string line, string originalValue)
    {
        // 텍스트 누적으로 인한 렌더링 부하 방지 (초기화)
        if (_textBlock.Inlines!.Count > 10000)
        {
            _textBlock.Inlines.Clear();
            _textBlock.Inlines.Add(new Run { Text = "Log clean-up done!" + Environment.NewLine, Foreground = Brushes.Gray });
            _lastReplaceableRun = null;
        }

        // [REPLACE] 태그 확인 및 제거
        bool isReplace = originalValue.StartsWith("[REPLACE]");
        if (isReplace)
        {
            originalValue = originalValue.Substring(9);
            line = line.Substring(9);
        }

        var run = new Run { Foreground = Brushes.White };
        string displayText = line;
        bool hasColorTag = false;

        // [COLOR: ] 색상 태그 파싱 및 제거
        if (originalValue.StartsWith("[COLOR:"))
        {
            int endIdx = originalValue.IndexOf(']');
            if (endIdx > 7)
            {
                string colorStr = originalValue.Substring(7, endIdx - 7);
                try { run.Foreground = Brush.Parse(colorStr); } catch { }
                
                displayText = line.Substring(endIdx + 1);
                hasColorTag = true;
            }
        }

        // REPLACE 처리
        if (isReplace)
        {
            Run? lastInline = _textBlock.Inlines.Count > 0 ? _textBlock.Inlines[_textBlock.Inlines.Count - 1] as Run : null;
            
            // 중간에 다른 로그가 끼어들지 않았다면 기존 줄 텍스트만 교체
            if (_lastReplaceableRun != null && lastInline == _lastReplaceableRun)
            {
                _lastReplaceableRun.Text = displayText;
                // 새 색상 태그가 명시되었을 때만 덮어쓰고, 없으면 이전 색상 유지
                if (hasColorTag) 
                {
                    _lastReplaceableRun.Foreground = run.Foreground;
                }
                return; // UI에 새로 Add하지 않고 여기서 종료
            }
            else
            {
                // 중간에 에러 로그 등이 끼어들었다면 덮어쓰기를 포기하고 새 줄로 추가
                run.Text = displayText;
                _textBlock.Inlines.Add(run);
                _lastReplaceableRun = run;
                _scrollViewer.ScrollToEnd();
                return;
            }
        }

        // 5. 일반 로그 처리
        run.Text = displayText;
        _textBlock.Inlines.Add(run);
        
        // 일반 로그가 들어오더라도, 다음 틱의 ReplaceLast가 이 줄에 점(...)을 찍을 수 있도록 기억해 둡니다.
        _lastReplaceableRun = run; 
        
        _scrollViewer.ScrollToEnd();
    }

    protected override void Dispose(bool disposing)
    {
        lock (_lock)
        {
            _logFileWriter.Flush();
            _logFileWriter.Close();
            _logFileWriter.Dispose();
            base.Dispose(disposing);
        }
    }
}