using System.Text;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.Threading;
using RimExtractorCore;

namespace RimExtractorFace.Utils;

/// <summary>로그창 출력기 클래스입니다. Core의 Log 클래스에 등록해서 사용합니다.</summary>
public class LogPrinter : TextWriter
{
    private readonly SelectableTextBlock _textBlock;
    private readonly ScrollViewer _scrollViewer;
    private readonly StreamWriter _logFileWriter;
    private readonly object _lock = new();
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
        if (_textBlock.Inlines!.Count > 5000)
        {
            _textBlock.Inlines.Clear();
            _textBlock.Inlines.Add(new Run { Text = "Log clean-up done!" + Environment.NewLine, Foreground = Brushes.Gray });
        }

        var run = new Run { Foreground = Brushes.White };
        string displayText = line;

        // 색상 태그 파싱 및 화면 출력 텍스트에서 태그 제거
        if (originalValue.StartsWith("[COLOR:"))
        {
            int endIdx = originalValue.IndexOf(']');
            if (endIdx > 7)
            {
                string colorStr = originalValue.Substring(7, endIdx - 7);
                try
                {
                    run.Foreground = Brush.Parse(colorStr);
                }
                catch { /* 파싱 실패 시 기본값(White) 유지 */ }

                // [COLOR:xxx] 부분(endIdx + 1)을 잘라내고 순수 메시지만 출력
                displayText = line.Substring(endIdx + 1);
            }
        }

        run.Text = displayText;
        _textBlock.Inlines.Add(run);
        
        // 스크롤을 맨 아래로 자동 이동
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