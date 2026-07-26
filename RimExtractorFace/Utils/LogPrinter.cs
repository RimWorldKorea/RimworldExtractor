using System.Text;
using Avalonia.Controls;
using Avalonia.Threading;

namespace RimExtractorFace.Utils;

public class LogPrinter : TextWriter
{
    private readonly TextBox _textBox;
    private readonly StreamWriter _logFileWriter;
    private readonly object _lock = new();
    public override Encoding Encoding => Encoding.UTF8;

    public LogPrinter(TextBox textBox)
    {
        _textBox = textBox;
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
            
            // UI 스레드 안전 호출 (WinForms의 InvokeRequired/Invoke 대체)
            Dispatcher.UIThread.Post(() => AppendToTextBox(line));
        }
    }

    private void AppendToTextBox(string line)
    {
        if ((_textBox.Text?.Length ?? 0) + line.Length > 327670)
        {
            _textBox.Text = "Log clean-up done!" + Environment.NewLine; // 텍스트 초기화
        }

        _textBox.Text += line;
        
        // 커서를 맨 아래로 이동
        _textBox.CaretIndex = _textBox.Text?.Length ?? 0;
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