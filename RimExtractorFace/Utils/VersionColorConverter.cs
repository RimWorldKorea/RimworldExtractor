using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace RimExtractorFace.Utils;

/// <summary>MainWindow에서 사용할 용도로 만들었습니다.</summary>
public class VersionColorConverter : IValueConverter
{
    /// <summary>
    /// bool인 value의 값에 따라서 다른 색상의 Brush를 반환합니다.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="targetType"></param>
    /// <param name="parameter"></param>
    /// <param name="culture"></param>
    /// <returns></returns>
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var isLatest = value as bool?;

        return isLatest switch
        {
            true => Brush.Parse("#388E3C"), // Green
            false => Brush.Parse("#FFD54F"), // Yellow
            _ => Brush.Parse("#E57373") // Red
        };
    }

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
